#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pathlib
import pytest
import shutil
import subprocess
import typing

ADAM_SCRIPT_PATH = pathlib.Path(__file__).parents[1] / "adam-script.py"
SAMPLE_SHEET_FILE = pathlib.Path("Sample Sheet.zip")
SAMPLE_SHEET_DIR = pathlib.Path("Sample Sheet")
POINT_FILE = SAMPLE_SHEET_DIR / "points.json"
CONFIG_INDIVIDUAL = pathlib.Path("config-individual.json")
CONFIG_SHARED_PREFIX = "config-shared-"

UNZIPPED_FILES = {
    "exercise": [
        ".sheet_info",
        "11910_Schmid_Stucki_Studer",
        "11910_feedback_tatiana_ex1_3.pdf.todo",
        "11910_feedback_tatiana_ex1_3_code-submission.cc",
        "11911_Sommer",
        "11911_feedback_tatiana_ex1_3.pdf.todo",
        "11911_feedback_tatiana_ex1_3_sabines-code.cc",
        "Sample Sheet",
        "code-submission.cc",
        "config-individual.json",
        "config-shared-exercise.json",
        "feedback",
        "feedback",
        "points.json",
        "sabines-code.cc",
        "sabines-submission.pdf",
        "samiras-submission.pdf",
    ],
    "random": [
        ".sheet_info",
        "11910_Schmid_Stucki_Studer",
        "DO_NOT_MARK_11911_Sommer",
        "Sample Sheet",
        "code-submission.cc",
        "config-individual.json",
        "config-shared-random.json",
        "feedback",
        "feedback_tatiana.pdf.todo",
        "feedback_tatiana_code-submission.cc",
        "points.json",
        "sabines-code.cc",
        "sabines-submission.pdf",
        "samiras-submission.pdf",
    ],
    "static": [
        ".sheet_info",
        "11910_Schmid_Stucki_Studer",
        "DO_NOT_MARK_11911_Sommer",
        "Sample Sheet",
        "code-submission.cc",
        "config-individual.json",
        "config-shared-static.json",
        "feedback",
        "feedback.pdf.todo",
        "feedback_code-submission.cc",
        "points.json",
        "sabines-code.cc",
        "sabines-submission.pdf",
        "samiras-submission.pdf",
    ],
}


def cd_to_test_dir() -> None:
    os.chdir(pathlib.Path(__file__).parent)


def test_help() -> None:
    subprocess.check_call(["python3", ADAM_SCRIPT_PATH, "-h"])


def subtest_init(
    path: pathlib.Path, config_shared: pathlib.Path, mode: str, args: list[str]
) -> None:
    # Call 'init'.
    subprocess.check_call(
        [
            "python3",
            str(ADAM_SCRIPT_PATH),
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(config_shared),
            "init",
            str(SAMPLE_SHEET_FILE),
        ]
        + args
    )
    # Remove zip.
    (path / SAMPLE_SHEET_FILE).unlink()
    # Get all unzipped files.
    files = sorted([f.name for f in path.glob("**/*")])
    print(files)
    assert files == UNZIPPED_FILES[mode]


def subtest_collect(path: pathlib.Path, config_shared: pathlib.Path) -> None:
    # Give feedback.
    for todo_file in path.glob("**/*.todo"):
        shutil.move(todo_file, todo_file.with_suffix(""))
    for todo_file in path.glob("**/*.todo"):
        print(todo_file.name)
    # Enter points.
    with open(POINT_FILE, "r") as file:
        data = file.read()
    data = data.replace(': ""', ': "1.5"')
    with open(POINT_FILE, "w") as file:
        file.write(data)
    # Call 'collect'.
    subprocess.check_call(
        [
            "python3",
            str(ADAM_SCRIPT_PATH),
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(config_shared),
            "collect",
            str(SAMPLE_SHEET_DIR),
        ]
    )


@pytest.mark.parametrize(
    "mode, args",
    [
        ("random", ["-n", "5"]),
        ("static", []),
        ("exercise", ["-e", "1", "3"]),
    ],
)
def test_modes(tmp_path: pathlib.Path, mode: str, args: list[str]) -> None:
    cd_to_test_dir()
    # Setup temporary directory.
    shutil.copy(SAMPLE_SHEET_FILE, tmp_path)
    shutil.copy(CONFIG_INDIVIDUAL, tmp_path)
    config_shared = pathlib.Path(CONFIG_SHARED_PREFIX + mode + ".json")
    shutil.copy(config_shared, tmp_path)
    os.chdir(tmp_path)
    # Test subcommands.
    subtest_init(tmp_path, config_shared, mode, args)
    # TODO: Test 'collect' for mode 'exercise' once supported.
    if (mode in ["random", "static"]):
        subtest_collect(tmp_path, config_shared)
    # TODO: Test 'send' command?
