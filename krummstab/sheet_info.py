from collections.abc import Iterator
from pathlib import Path
import logging
import json

from . import config, submission_info

SHEET_INFO_FILE_NAME = "sheet.json"
FEEDBACK_FILE_PREFIX = "feedback_"
COMBINED_DIR_NAME = "feedback_combined"


class SheetInfo:
    def __init__(self, sheet_root_dir: Path):
        self._sheet_info_path = sheet_root_dir / SHEET_INFO_FILE_NAME
        if not sheet_root_dir.is_dir() or not self._sheet_info_path.exists():
            logging.critical("Could not find a directory with marking "
                             f"information at '{sheet_root_dir}'. Use the command "
                             "'init' to set it up.")
        self.root_dir = sheet_root_dir
        self._load()

    def _load(self):
        with open(self._sheet_info_path, "r", encoding="utf-8") as sheet_info_file:
            sheet_info = json.load(sheet_info_file)
            self.adam_sheet_name = sheet_info.get('adam_sheet_name')
            self.exercises = sheet_info.get('exercises')

    def get_adam_sheet_name_string(self) -> str:
        """
        Turn the sheet name given by ADAM into a string usable for file names.
        """
        return self.adam_sheet_name.replace(" ", "_").lower()

    def get_feedback_file_name(self, _the_config: config.Config) -> str:
        file_name = FEEDBACK_FILE_PREFIX + self.get_adam_sheet_name_string() + "_"
        if _the_config.marking_mode == "exercise":
            file_name += _the_config.tutor_name + "_"
            file_name += "_".join([f"ex{exercise}" for exercise in self.exercises])
        elif _the_config.marking_mode == "static":
            # Remove trailing underscore.
            file_name = file_name[:-1]
        else:
            _the_config.unsupported_marking_mode_error()
        return file_name

    def get_combined_feedback_file_name(self) -> str:
        return FEEDBACK_FILE_PREFIX + self.get_adam_sheet_name_string()

    def get_combined_feedback_dir(self):
        return self.root_dir / COMBINED_DIR_NAME

    def get_marks_file_path(self, _the_config: config.Config):
        return (
                self.root_dir
                / f"points_{_the_config.tutor_name.lower()}_{self.get_adam_sheet_name_string()}.json"
        )

    def get_individual_marks_file_path(self, _the_config: config.Config):
        marks_file_path = self.get_marks_file_path(_the_config)
        individual_marks_file_path = marks_file_path.with_name(
            marks_file_path.stem + "_individual" + marks_file_path.suffix
        )
        return individual_marks_file_path

    def get_all_team_submission_info(self) -> Iterator[submission_info.SubmissionInfo]:
        """
        Return all team submission info. Exclude other directories that may be created
        in the sheet root directory, such as one containing combined feedback.
        """
        for team_dir in self.root_dir.iterdir():
            if team_dir.is_dir() and team_dir != self.get_combined_feedback_dir():
                yield submission_info.SubmissionInfo(team_dir)

    def get_relevant_team_submission_info(self) -> Iterator[submission_info.SubmissionInfo]:
        """
        Return the submission info of the teams whose submission has to be
        corrected by the tutor running the script.
        """
        for submission in self.get_all_team_submission_info():
            if submission.relevant:
                yield submission


def create_sheet_info_file(sheet_root_dir: Path, adam_sheet_name: str,
                           _the_config: config.Config, exercises=None) -> SheetInfo:
    """
    Write information generated during the execution of the `init` command in a
    sheet info file. In particular the name of the exercise sheet as given by
    ADAM and which exercises should be marked if the marking mode is `exercise`.
    Later commands (e.g. `collect`, or `send`) are meant to load the information
    stored in this file when initializing the class sheet_info.SheetInfo
    and access it that way.
    """
    info_dict = {}
    info_dict.update({"adam_sheet_name": adam_sheet_name})
    if _the_config.marking_mode == "exercise":
        info_dict.update({"exercises": exercises})
    with open(
            sheet_root_dir / SHEET_INFO_FILE_NAME, "w", encoding="utf-8"
    ) as sheet_info_file:
        json.dump(
            info_dict,
            sheet_info_file,
            indent=4,
            ensure_ascii=False,
            sort_keys=True,
        )
    return SheetInfo(sheet_root_dir=sheet_root_dir)
