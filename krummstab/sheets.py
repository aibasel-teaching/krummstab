from collections.abc import Iterator
from pathlib import Path
import logging
import json

from . import config, submissions, errors

SHEET_INFO_FILE_NAME = "sheet.json"
FEEDBACK_FILE_PREFIX = "feedback_"
COMBINED_DIR_NAME = "feedback_combined"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"
SHARE_ARCHIVE_PREFIX = "share_archive"


class Sheet:
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
            self.name = sheet_info.get('adam_sheet_name')
            self.exercises = sheet_info.get('exercises')

    def get_adam_sheet_name_string(self) -> str:
        """
        Turn the sheet name given by ADAM into a string usable for file names.
        """
        return self.name.replace(" ", "_").lower()

    def get_feedback_file_name(self, _the_config: config.Config) -> str:
        file_name = FEEDBACK_FILE_PREFIX + self.get_adam_sheet_name_string() + "_"
        if _the_config.marking_mode == "exercise":
            file_name += _the_config.tutor_name + "_"
            file_name += "_".join([f"ex{exercise}" for exercise in self.exercises])
        elif _the_config.marking_mode == "static":
            # Remove trailing underscore.
            file_name = file_name[:-1]
        else:
            errors.unsupported_marking_mode_error(_the_config.marking_mode)
        return file_name

    def get_combined_feedback_file_name(self) -> str:
        return FEEDBACK_FILE_PREFIX + self.get_adam_sheet_name_string()

    def get_combined_feedback_path(self) -> Path:
        return self.root_dir / COMBINED_DIR_NAME

    def get_marks_file_path(self, _the_config: config.Config) -> Path:
        return (
                self.root_dir
                / f"points_{_the_config.tutor_name.lower()}_{self.get_adam_sheet_name_string()}.json"
        )

    def get_individual_marks_file_path(self, _the_config: config.Config) -> Path:
        marks_file_path = self.get_marks_file_path(_the_config)
        individual_marks_file_path = marks_file_path.with_name(
            marks_file_path.stem + "_individual" + marks_file_path.suffix
        )
        return individual_marks_file_path

    def get_all_team_submission_info(self) -> Iterator[submissions.Submission]:
        """
        Return all team submission info. Exclude other directories that may be created
        in the sheet root directory, such as one containing combined feedback.
        """
        for sub_dir in self.root_dir.iterdir():
            if sub_dir.is_dir() and sub_dir != self.get_combined_feedback_path():
                yield submissions.Submission(sub_dir)

    def get_relevant_submissions(self) -> Iterator[submissions.Submission]:
        """
        Return the submission info of the teams whose submission has to be
        corrected by the tutor running the script.
        """
        for submission in self.get_all_team_submission_info():
            if submission.relevant:
                yield submission

    def get_share_archive_file_path(self) -> Path:
        return self.root_dir / (
                SHARE_ARCHIVE_PREFIX
                + f"_{self.get_adam_sheet_name_string()}_"
                + "_".join([f"ex{num}" for num in self.exercises])
                + ".zip"
        )

    def get_share_archive_files(self) -> Iterator[Path]:
        """
        Return all share archive files under the current sheet root dir.
        """
        for share_archive_file in self.root_dir.glob(
                SHARE_ARCHIVE_PREFIX + "*.zip"
        ):
            yield share_archive_file


def create_sheet_info_file(sheet_root_dir: Path, adam_sheet_name: str,
                           _the_config: config.Config, exercises=None) -> Sheet:
    """
    Write information generated during the execution of the `init` command in a
    sheet info file. In particular the name of the exercise sheet as given by
    ADAM and which exercises should be marked if the marking mode is `exercise`.
    Later commands (e.g. `collect`, or `send`) are meant to load the information
    stored in this file when initializing the class sheet_info.Sheet
    and access it that way.
    """
    info_dict = {}
    info_dict["adam_sheet_name"] = adam_sheet_name
    if _the_config.marking_mode == "exercise":
        info_dict["exercises"] = exercises
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
    return Sheet(sheet_root_dir=sheet_root_dir)
