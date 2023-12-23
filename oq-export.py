#!/usr/bin/env python
# -*- coding: utf-8 -*-
import argparse
import json
import re
import time
from collections import defaultdict

import requests
import openpyxl
from playwright.sync_api import sync_playwright


class TournData:
    def __init__(self):
        self.results = None
        self.title = None

    def handle_route(self, route, request):
        response = route.fetch()
        json_ = response.json()
        self.results = json_
        print("route handled")
        route.fulfill(response=response, json=json_)


HEADER = ["Team ID", "Название", "Город"]


def try_to_find_team(name):
    print(f"trying to find team {name}")
    req = requests.get(f"https://api.rating.chgk.net/teams.json?name={name}")
    time.sleep(0.5)
    for team in req.json():
        if team["name"] == name:
            return (team["id"], team["town"]["name"])
    return (None, None)


def make_workbook(results):
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
        id_, town = try_to_find_team(team)
        row = [id_, team, town] + [res.get(q) for q in question_header]
        ws.append(row)
    return wb


def sanitize_title(title):
    return re.sub("[^a-zA-Zа-яА-ЯЁё0-9\\-\\.]", "_", title)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("url")
    args = parser.parse_args()

    td = TournData()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page()
        page.route(
            "**/static/**", lambda route, request: td.handle_route(route, request)
        )
        # page.on("route", lambda route: handle_request(route))
        page.goto(args.url)
        title = page.query_selector("#quizName").inner_text()
        td.title = title
        print(f"title: {title}")
        time.sleep(2)
        browser.close()

    wb = make_workbook(td.results)
    filename = sanitize_title(td.title) + ".xlsx"
    print(filename)
    wb.save(filename)


if __name__ == "__main__":
    main()
