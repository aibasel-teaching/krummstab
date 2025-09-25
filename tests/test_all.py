#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import gzip
import os
import pathlib
import pytest
import shutil
import subprocess
import typing

CONFIG_INDIVIDUAL = pathlib.Path("config-individual.json")
CONFIG_STATIC = pathlib.Path("config-shared-static.json")
CONFIG_EXERCISE = pathlib.Path("config-shared-exercise.json")
SAMPLE_SHEET = pathlib.Path("Sample Sheet.zip")
SAMPLE_SHEET_DIR = pathlib.Path("Sample Sheet")
SAMPLE_SHEET_SUB_DIR = SAMPLE_SHEET_DIR / "SUB_DIR"
SAMPLE_INDIVIDUAL_POINTS_FILES_DIR_STATIC = pathlib.Path("points-files-static")
SAMPLE_INDIVIDUAL_POINTS_FILES_DIR_EXERCISE = pathlib.Path(
    "points-files-exercise"
)


@pytest.fixture(autouse=True)
def change_test_dir(request, monkeypatch):
    """
    Make the directory where the test module lies the current working directory.
    """
    monkeypatch.chdir(request.fspath.dirname)


@pytest.fixture(params=["Abgaben", "Submissions"])
def setup_test_directory(request, tmp_path: pathlib.Path):
    """
    Copy sample sheet directory, create zip file in ADAM's format, and remove
    sheet directory.
    """
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
    return tmp_path


@pytest.fixture
def mode_dict(request, monkeypatch, setup_test_directory: pathlib.Path):
    """
    This fixture copies all relevant files into a testing directory.
    """
    shutil.copy(request.param["config_shared"], setup_test_directory)
    shutil.copytree(
        request.param["individual_point_file_dir"],
        setup_test_directory / request.param["individual_point_file_dir"].name,
    )
    monkeypatch.chdir(setup_test_directory)
    return request.param


@pytest.fixture(params=["tamara", "tatiana", "terence"])
def insert_tutor_name(request):
    with open("config-individual.json", "r") as f:
        filled_in = f.read().replace("PLACEHOLDER_NAME", request.param)
    with open("config-individual.json", "w") as f:
        f.write(filled_in)


@pytest.fixture(
    params=[
        ("true", '["xournalpp", "{xopp_file}"]'),
        ("false", '["ls", "{all_pdf_files}"]'),
    ]
)
def insert_xopp_setting(request):
    with open("config-individual.json", "r") as f:
        filled_in = (
            f.read()
            .replace("PLACEHOLDER_XOPP_SETTING", request.param[0])
            .replace("PLACEHOLDER_MARKING_COMMAND", request.param[1])
        )
    with open("config-individual.json", "w") as f:
        f.write(filled_in)


def give_feedback():
    # Enter points.
    for point_file in pathlib.Path.cwd().glob("**/points*.json"):
        with open(point_file, "r") as file:
            data = file.read()
        data = data.replace(': ""', ': "1.5"')
        with open(point_file, "w") as file:
            file.write(data)
    # gzip .xopp files to pretend that we opened and saved the xopp files with
    # Xournal++. If we don't do this 'collect' will complain that we did not
    # actually give any feedback.
    for xopp_file in pathlib.Path.cwd().glob("**/*.xopp"):
        with open(xopp_file, "rb") as file_in:
            content = file_in.read()
        with gzip.open(xopp_file, "wb") as file_out:
            file_out.write(content)


@pytest.mark.parametrize(
    "mode_dict, args",
    [
        (
            {
                "config_shared": CONFIG_STATIC,
                "individual_point_file_dir": SAMPLE_INDIVIDUAL_POINTS_FILES_DIR_STATIC,
            },
            [],
        ),
        (
            {
                "config_shared": CONFIG_EXERCISE,
                "individual_point_file_dir": SAMPLE_INDIVIDUAL_POINTS_FILES_DIR_EXERCISE,
            },
            ["-e", "1", "3"],
        ),
    ],
    indirect=["mode_dict"],
)
def test(
    capfd,
    mode_dict: dict,
    insert_tutor_name,
    insert_xopp_setting,
    skip_mark_test,
    args: list[str],
):
    # Call 'init'.
    subprocess.check_call(
        [
            "krummstab",
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(mode_dict["config_shared"]),
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

    if not skip_mark_test:
        # Call 'mark'.
        subprocess.check_call(
            [
                "krummstab",
                "-i",
                str(CONFIG_INDIVIDUAL),
                "-s",
                str(mode_dict["config_shared"]),
                "mark",
                "--dry-run",
                str(SAMPLE_SHEET_DIR),
            ]
        )

    # Prepare for 'collect'.
    give_feedback()

    # Call 'collect'.
    subprocess.check_call(
        [
            "krummstab",
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(mode_dict["config_shared"]),
            "collect",
            str(SAMPLE_SHEET_DIR),
        ]
    )
    out, err = capfd.readouterr()
    assert "Command 'collect' terminated successfully." in out

    # Call 'combine' for the configurations using the mode 'exercise'.
    if mode_dict["config_shared"] == CONFIG_EXERCISE:
        subprocess.check_call(
            [
                "krummstab",
                "-i",
                str(CONFIG_INDIVIDUAL),
                "-s",
                str(mode_dict["config_shared"]),
                "combine",
                str(SAMPLE_SHEET_DIR),
            ]
        )
        out, err = capfd.readouterr()
        assert "Command 'combine' terminated successfully." in out

    # Call 'send'.
    subprocess.check_call(
        [
            "krummstab",
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(mode_dict["config_shared"]),
            "send",
            "--dry-run",
            str(SAMPLE_SHEET_DIR),
        ]
    )
    out, err = capfd.readouterr()
    assert "Command 'send' terminated successfully." in out

    # Call 'summarize'.
    subprocess.check_call(
        [
            "krummstab",
            "-i",
            str(CONFIG_INDIVIDUAL),
            "-s",
            str(mode_dict["config_shared"]),
            "summarize",
            str(mode_dict["individual_point_file_dir"]),
        ]
    )
    out, err = capfd.readouterr()
    assert (
        "Command 'summarize' terminated successfully." in out
        and len(list(pathlib.Path.cwd().glob("*.xlsx"))) == 1
    )
