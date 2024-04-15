#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pathlib
import pytest
import re
import shutil
import subprocess
import typing

ADAM_SCRIPT = pathlib.Path("adam-script.py")
CONFIG_INDIVIDUAL = pathlib.Path("config-individual.json")
CONFIG_RANDOM = pathlib.Path("config-shared-random.json")
CONFIG_STATIC = pathlib.Path("config-shared-static.json")
CONFIG_EXERCISE = pathlib.Path("config-shared-exercise.json")
SAMPLE_SHEET = pathlib.Path("Sample Sheet.zip")
SAMPLE_SHEET_DIR = pathlib.Path("Sample Sheet")
SAMPLE_SHEET_SUB_DIR = SAMPLE_SHEET_DIR / "SUB_DIR"


@pytest.fixture(autouse=True)
def change_test_dir(request, monkeypatch):
    """
    Make the directory where the test module lies the current working directory.
    """
    monkeypatch.chdir(request.fspath.dirname)


@pytest.fixture(params=["Abgaben", "Submissions"])
def setup_test_directory(request, tmp_path: pathlib.Path):
    # Copy sample sheet directory, create zip file in ADAM's format, and remove
    # sheet directory.
    shutil.copytree(SAMPLE_SHEET_DIR, tmp_path / SAMPLE_SHEET_DIR.name)
    shutil.move(
        tmp_path / SAMPLE_SHEET_SUB_DIR,
        tmp_path / SAMPLE_SHEET_DIR / request.param,
    )
    os.makedirs(tmp_path / "temp")
    shutil.move(tmp_path / SAMPLE_SHEET_DIR, tmp_path / "temp")
    shutil.move(tmp_path / "temp", tmp_path / SAMPLE_SHEET_DIR)
    shutil.make_archive(
        str(tmp_path / SAMPLE_SHEET_DIR), "zip", tmp_path / SAMPLE_SHEET_DIR
    )
    shutil.rmtree(tmp_path / SAMPLE_SHEET_DIR)
    # Copy individual config file.
    shutil.copy(CONFIG_INDIVIDUAL, tmp_path)
    # Copy python files.
    for file in pathlib.Path.cwd().glob(r"../*.py"):
        shutil.copy(file, tmp_path)
    return tmp_path


@pytest.fixture
def config_shared(request, monkeypatch, setup_test_directory: pathlib.Path):
    """
    This fixture copies all relevant files into a testing directory.
    """
    shutil.copy(request.param, setup_test_directory)
    monkeypatch.chdir(setup_test_directory)
    return request.param


@pytest.fixture(params=["tamara", "tatiana", "terence"])
def tutor_name(request):
    with open("config-individual.json", "r") as f:
        filled_in = f.read().replace("PLACEHOLDER_NAME", request.param)
    with open("config-individual.json", "w") as f:
        f.write(filled_in)
    return request.param


def get_point_file_path(tutor_name: str) -> pathlib.Path:
    return SAMPLE_SHEET_DIR / f"points_{tutor_name}_sample_sheet.json"


def give_feedback(tutor_name: str):
    # Remove '.todo' suffixes.
    for todo_file in pathlib.Path.cwd().glob("**/*.todo"):
        shutil.move(todo_file, todo_file.with_suffix(""))
    point_file = get_point_file_path(tutor_name)
    # Enter points.
    with open(point_file, "r") as file:
        data = file.read()
    data = data.replace(': ""', ': "1.5"')
    with open(point_file, "w") as file:
        file.write(data)


@pytest.mark.parametrize(
    "config_shared, args",
    [
        (CONFIG_RANDOM, ["-n", "5"]),
        (CONFIG_STATIC, []),
        (CONFIG_EXERCISE, ["-e", "1", "3"]),
    ],
    indirect=["config_shared"],
)
def test(capfd, config_shared: pathlib.Path, tutor_name: str, args: list[str]):
    # Call 'init'.
    subprocess.check_call(
        [
            "python3",
            str(ADAM_SCRIPT),
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(config_shared),
            "init",
            str(SAMPLE_SHEET),
        ]
        + args
    )
    # Remove zip.
    SAMPLE_SHEET.unlink()

    # Verify 'init' ran successfully.
    out, err = capfd.readouterr()
    assert "Command 'init' terminated successfully." in out

    # Prepare for 'collect'.
    give_feedback(tutor_name)

    # Call 'collect'.
    subprocess.check_call(
        [
            "python3",
            str(ADAM_SCRIPT),
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(config_shared),
            "collect",
            str(SAMPLE_SHEET_DIR),
        ]
    )

    # Verify 'collect' ran successfully.
    out, err = capfd.readouterr()
    assert "Command 'collect' terminated successfully." in out

    # Call 'combine' for the configurations using the mode 'exercise'.
    if config_shared == CONFIG_EXERCISE:
        subprocess.check_call(
            [
                "python3",
                str(ADAM_SCRIPT),
                "-i",
                str(CONFIG_INDIVIDUAL),
                "-s",
                str(config_shared),
                "combine",
                str(SAMPLE_SHEET_DIR),
            ]
        )

        # Verify 'combine' ran successfully.
        out, err = capfd.readouterr()
        assert "Command 'combine' terminated successfully." in out

    # Call 'send' for all configurations except 'exercise'.
    if config_shared != CONFIG_EXERCISE:
        subprocess.check_call(
            [
                "python3",
                str(ADAM_SCRIPT),
                "-i",
                str(CONFIG_INDIVIDUAL),
                "-s",
                str(config_shared),
                "send",
                str(SAMPLE_SHEET_DIR),
                "-d",
            ]
        )
        out, err = capfd.readouterr()
        assert "Command 'send' terminated successfully." in out

    if config_shared == CONFIG_STATIC:
        if tutor_name == 'tamara':
            assert "Ready to send 2 email(s)" in out
            assert "To: sabine.sommer@stud.unibas.ch" in out
        elif tutor_name == 'tatiana':
            assert "Ready to send 2 email(s)" in out
            assert "To: salome.schmid@stud.unibas.ch, samira.studer@stud.unibas.ch, silvan.stucki@stud.unibas.ch" in out
        elif tutor_name == 'terence':
            assert "Ready to send 1 email(s)" in out
    if config_shared == CONFIG_RANDOM:
        assert re.search(r"Ready to send [12] email\(s\)", out)
