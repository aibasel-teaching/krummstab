import json
import logging
import openpyxl
import pathlib
import shutil
import sys
import tempfile
import jsonschema

from collections import defaultdict
from typing import Any
from zipfile import ZipFile

from .students import Student
from .teams import Team


# Assignment -------------------------------------------------------------------


def create_submission_team_to_tutors_dict(
    submission_teams: list[Team],
    student_email_to_tutor_dict: dict[str, set[str]],
    tutor_list: list[str],
) -> dict[str, set[str]]:
    """
    Create a dictionary that maps submission team IDs to a set of tutors, used
    to determine which tutor has to mark which submissions.
    """
    team_to_tutors = defaultdict(set)
    for team in submission_teams:
        candidate_tutors = {
            tutor
            for member in team
            for tutor in student_email_to_tutor_dict[member.email]
        }
        # In case of a team where none of its member appear in the config,
        # candidate_tutors will be empty here. We add all tutors as candidates.
        if not candidate_tutors:
            candidate_tutors = set(tutor_list)
        team_to_tutors[team.adam_id] = candidate_tutors
    return team_to_tutors


# ADAM Input -------------------------------------------------------------------


def read_teams_from_adam_spreadsheet(file: pathlib.Path) -> dict[str, Team]:
    """
    Reads the teams from the ADAM Excel spreadsheet and returns a dictionary
    with the team IDs as keys and the teams as values.
    """
    assert file.is_file()
    wb = openpyxl.load_workbook(file)
    sheet = wb.active
    col_last_name = 0
    col_first_name = 1
    col_email = 2
    col_team_id = 4
    teams_data = defaultdict(list)
    for row in sheet.iter_rows(min_row=2, values_only=True):
        team_id = str(row[col_team_id])
        first_name = row[col_first_name]
        last_name = row[col_last_name]
        email = row[col_email]
        teams_data[team_id].append((first_name, last_name, email))
    for team in teams_data.values():
        team.sort()
    teams = {}
    for team_id, team in teams_data.items():
        teams[team_id] = Team([Student(*student) for student in team], team_id)
    return teams


# Logging ----------------------------------------------------------------------


def configure_logging(level=logging.INFO):
    class ColoredFormatter(logging.Formatter):
        FORMATS = {
            logging.DEBUG: "\033[0;37m[{levelname}]\033[0m {message}",
            logging.INFO: "\033[0;34m[{levelname}]\033[0m {message}",
            logging.WARNING: "\033[0;33m[{levelname}]\033[0m {message}",
            logging.ERROR: "\033[0;31m[{levelname}]\033[0m {message}",
            logging.CRITICAL: "\033[0;31m[{levelname}]\033[0m {message}",
        }

        def format(self, record):
            formatter = logging.Formatter(
                ColoredFormatter.FORMATS[record.levelno], style="{"
            )
            return formatter.format(record)

    class LevelFilter:
        def __init__(self, min_level, max_level):
            self.min_level = min_level
            self.max_level = max_level

        def filter(self, record):
            return self.min_level <= record.levelno <= self.max_level

    class CustomHandler(logging.StreamHandler):
        def __init__(
            self, stream, min_level=logging.DEBUG, max_level=logging.CRITICAL
        ):
            logging.StreamHandler.__init__(self, stream)
            self.setFormatter(ColoredFormatter())
            self.addFilter(LevelFilter(min_level, max_level))

        def emit(self, record):
            logging.StreamHandler.emit(self, record)
            if record.levelno >= logging.CRITICAL:
                sys.exit("aborting")

    root_logger = logging.getLogger("")
    # Remove old handlers
    for handler in root_logger.handlers:
        root_logger.removeHandler(handler)
    root_logger.addHandler(CustomHandler(sys.stdout, max_level=logging.WARNING))
    root_logger.addHandler(CustomHandler(sys.stderr, min_level=logging.ERROR))
    root_logger.setLevel(level)


# Printing ---------------------------------------------------------------------


def query_yes_no(text: str, default: bool = True) -> bool:
    """
    Ask the user a yes/no question and return answer.
    """
    valid = {"yes": True, "y": True, "ye": True, "no": False, "n": False}
    options = "[Y/n]" if default else "[y/N]"
    print("\033[0;35m[Query]\033[0m " + text + f" {options}")
    choice = input().lower()
    if choice == "":
        return default
    elif choice in valid:
        return valid[choice]
    else:
        logging.warning(
            f"Invalid choice '{choice}'. Please respond with 'yes' or 'no'."
        )
        return query_yes_no(text, default)


# JSON parsing -------------------------------------------------------


def validate_json(
    data: dict,
    schema: dict,
    source: str = "file",
    schema_version=jsonschema.Draft7Validator,
) -> None:
    """
    Validates a JSON object against a given schema.
    """
    try:
        jsonschema.validate(data, schema, schema_version)
    except jsonschema.exceptions.ValidationError as error:
        logging.critical(
            f"Validation error: {source} does not have the right format: "
            f"{error.message}"
        )


def read_json(source: str | pathlib.Path, source_name: str = "file") -> dict:
    """
    Reads a JSON file and returns its contents.
    """
    data = {}
    try:
        if isinstance(source, pathlib.Path):
            source_name = source
            json_str = source.read_text(encoding="utf-8")
        else:
            json_str = source
        data = json.loads(json_str)
    except json.decoder.JSONDecodeError as error:
        logging.critical(f"Wrong JSON format in {source_name}: {error}")
    return data


# File handling ----------------------------------------------------------------


def is_hidden_path(path: pathlib.Path) -> bool:
    """
    Check if a given file is a hidden file.
    """
    # Using path.resolve().parts would give us all parts of the absolute path
    # instead of only the parts of the relative path, but that would make the
    # check too strict. We ideally don't want to care whether some parent
    # directory outside the scope of Krummstab is hidden or not.
    return is_superfluous_macos_path(path) or any(
        part.startswith(".") for part in path.parts
    )


def is_superfluous_macos_path(path: pathlib.Path) -> bool:
    """
    Check if the given path is a non-essential file created by MacOS, such as
    .DS_Store created by Finder or __MACOSX folders created when creating zip
    archives.
    """
    return any(
        part == magic_string
        for part in path.parts
        for magic_string in ["__MACOSX", ".DS_Store"]
    )


def filtered_extract(zip_file: ZipFile, dest: pathlib.Path) -> None:
    """
    Extract all files except for MACOS helper files.
    """
    zip_content = zip_file.namelist()
    for file_str in zip_content:
        if is_superfluous_macos_path(pathlib.Path(file_str)):
            continue
        zip_file.extract(file_str, dest)


def move_content_and_delete(src: pathlib.Path, dst: pathlib.Path) -> None:
    """
    Move all content of source directory to destination directory.
    This does not complain if the dst directory already exists.
    """
    assert src.is_dir() and dst.is_dir()
    with tempfile.TemporaryDirectory() as temp_dir:
        shutil.copytree(src, temp_dir, dirs_exist_ok=True)
        shutil.rmtree(src)
        shutil.copytree(temp_dir, dst, dirs_exist_ok=True)


def unzip_or_move_adam_zip(
    adam_zip_path: pathlib.Path, destination: pathlib.Path
) -> None:
    """
    Unzip ADAM zip contents to the destination. On Apple systems the zip is
    extracted automatically, so we instead just move the content.
    """
    if adam_zip_path.is_file():
        # Unzip to the directory within the zip file.
        # Should be the name of the exercise sheet,
        # for example "Exercise Sheet 2".
        with ZipFile(adam_zip_path, mode="r") as zip_file:
            filtered_extract(zip_file, destination)
    else:
        # Assume the directory is an extracted ADAM zip.
        unzipped_path = pathlib.Path(adam_zip_path)
        unzipped_destination_path = (
            pathlib.Path(destination) / unzipped_path.name
        )
        shutil.copytree(unzipped_path, unzipped_destination_path)


# Type juggling ----------------------------------------------------------------


def represents_float(value: Any) -> bool:
    try:
        float(value)
    except ValueError:
        return False
    return True


def convert_to_float_if_possible(value: Any) -> Any:
    if represents_float(value):
        return float(value)
    return value


def make_lower_case_if_possible(value: Any) -> Any:
    try:
        return value.lower()
    except AttributeError:
        return value
