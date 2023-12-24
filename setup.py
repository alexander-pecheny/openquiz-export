from setuptools import setup


def get_version():
    version = {}
    with open("openquiz_export/version.py") as f:
        exec(f.read(), version)
    return version["__version__"]


long_description = "Openquiz-export exports quiz results from open-quiz.com to Excel"


setup(
    name="openquiz_export",
    version=get_version(),
    author="Alexander Pecheny",
    author_email="ap@pecheny.me",
    description=long_description,
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://gitlab.com/peczony/openquiz-export",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    packages=["openquiz_export"],
    entry_points={
        "console_scripts": [
            "oq-export = openquiz_export.__main__:main",
        ]
    },
    install_requires=["requests", "openpyxl", "levenshtein"],
)
