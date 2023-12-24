#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
import json
import re
import time
import urllib.parse
from collections import defaultdict
from typing import NamedTuple, Optional

import openpyxl
import requests
from Levenshtein import distance


def now_ms():
    return round(time.time() * 1000)


HEADER = ["Team ID", "Название", "Город"]


class TeamTuple(NamedTuple):
    id: int
    town: str
    name: Optional[str]


def find_team_by_name(name):
    print(f"trying to find team {name}")
    req = requests.get(f"https://api.rating.chgk.net/teams.json?name={name}")
    time.sleep(0.5)
    for team in req.json():
        if team["name"] == name:
            return TeamTuple(team["id"], team["town"]["name"], None)
    return TeamTuple(None, None, None)


def rating_get_results(id_):
    return requests.get(
        f"https://api.rating.chgk.net/tournaments/{id_}/results.json"
    ).json()


class WbMaker:
    def __init__(self, results, args):
        self.results = results
        self.args = args
        self.team_dict = {}
        if self.args.tournament_id:
            print(f"getting results from {self.args.tournament_id}")
            rating_results = rating_get_results(self.args.tournament_id)
            for res in rating_results:
                sr = res.get("synchRequest")
                if not sr or sr["venue"]["name"] != self.args.venue_name:
                    continue
                name = res["current"]["name"]
                self.team_dict[name.lower()] = TeamTuple(
                    id=res["team"]["id"],
                    name=name,
                    town=res["current"]["town"]["name"]
                )

    def try_to_find_team(self, team_name):
        if self.team_dict:
            lower = team_name.lower()
            if lower in self.team_dict:
                return self.team_dict[lower]
            else:
                srt = sorted([
                    (distance(lower, k), k)
                    for k in self.team_dict
                ])
                if srt[0][0] >= 5:
                    print(f"team {team_name} not found in results, searching in rating...")
                    return find_team_by_name(team_name)
                else:
                    return self.team_dict[srt[0][1]]
        else:
            return find_team_by_name(team_name)

    def make(self):
        results = self.results
        key_to_question = {
            json.dumps(q["Key"]).replace(" ", ""): int(q["Name"])
            for q in results["Questions"]
        }
        team_to_res = defaultdict(dict)
        for team in results["Teams"]:
            name = team["TeamName"]
            details = team["Details"]
            for key, value in details.items():
                question = key_to_question[key]
                point = value["Result"]
                team_to_res[name][question] = point
        question_header = sorted(key_to_question.values())
        header = HEADER + question_header
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(header)
        for team in team_to_res:
            res = team_to_res[team]
            id_, town, new_name = self.try_to_find_team(team)
            row = [id_, new_name or team, town] + [res.get(q) for q in question_header]
            ws.append(row)
        return wb


def sanitize_title(title):
    return re.sub("[^a-zA-Zа-яА-ЯЁё0-9\\-\\.]", "_", title)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("url")
    parser.add_argument("--tournament_id", "-t", type=int)
    parser.add_argument("--venue_name", "-v", default="ХВИП")
    args = parser.parse_args()

    parsed = urllib.parse.urlparse(args.url)
    qs = dict(urllib.parse.parse_qsl(parsed.query))
    quiz_id = qs["quiz"]
    token = qs["token"]
    title = qs["quizName"]
    static_url = f"https://www.open-quiz.com/static/{quiz_id}-{token}/results.json?nocache={now_ms()}"

    results = requests.get(static_url).json()

    wb = WbMaker(results, args).make()
    filename = sanitize_title(title) + ".xlsx"
    print(filename)
    wb.save(filename)


if __name__ == "__main__":
    main()
