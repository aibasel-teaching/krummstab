import config
import logging
from pathlib import Path
import re
import shutil
from typing import Iterator

from utils import Team, write_json, read_json

FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"
FEEDBACK_FILE_PREFIX = "feedback_"
SHEET_INFO_FILE_NAME = ".sheet_info.json"
SHARE_ARCHIVE_PREFIX = "share_archive"
COMBINED_DIR_NAME = "feedback_combined"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"

# temporary patch. In the long term, we want to avoid this global variable and
# pass the SheetDirectory to the commands as a normal argument. Having it as a
# global variable just makes the transition easier and the diff more readable.
the_sheet_dir = None
def load(root):
    global the_sheet_dir
    the_sheet_dir = SheetDirectory(root)
def get():
    assert the_sheet_dir
    return the_sheet_dir


class TeamDirectory:
    def __init__(self, root: Path, sheet_dir: "SheetDirectory") -> None:
        self.root = root
        self.sheet_dir = sheet_dir

    def get_feedback_dir(self) -> Path:
        return self.root / FEEDBACK_DIR_NAME

    def get_collected_feedback_dir(self) -> Path:
        return self.root / FEEDBACK_COLLECTED_DIR_NAME

    def get_collected_feedback_file(self) -> Path:
        """
        Given a team directory, return the collected feedback file. This can be
        either a single pdf file, or a single zip archive. Throw an error if neither
        exists.
        """
        collected_feedback_dir = self.get_collected_feedback_dir()
        assert collected_feedback_dir.is_dir()
        collected_feedback_files = list(collected_feedback_dir.iterdir())
        assert (
            len(collected_feedback_files) == 1
            and collected_feedback_files[0].is_file()
            and collected_feedback_files[0].suffix in [".pdf", ".zip"]
        )
        return collected_feedback_files[0]
    
    def get_team(self):
        # TODO: copy team into TeamDirectory rather than looking it up.
        return self.sheet_dir.team_dir_to_team[self.root.name]
    
    def __lt__(self, other):
        return self.root < other.root


class SheetDirectory:
    def __init__(self, root: Path) -> None:
        self._sheet_info_path = root / SHEET_INFO_FILE_NAME
        if not root.is_dir() or not self._sheet_info_path.exists():
            logging.critical("Could not find a directory with marking "
                             f"information at '{root}'. Use the command "
                             "'init' to set it up.")
        self.root = root
        self._load()
    
    def _load(self):
        sheet_info = read_json(self.root / SHEET_INFO_FILE_NAME)
        for key, value in sheet_info.items():
            setattr(self, key, value)

    def get_adam_sheet_name(self):
        return self.adam_sheet_name

    def _get_sanitized_adam_sheet_name(self):
        """
        Turn the sheet name given by ADAM into a string usable for file names.
        """
        return self.adam_sheet_name.replace(" ", "_").lower()


    def get_feedback_file_name(self, args):
        file_name = FEEDBACK_FILE_PREFIX + self._get_sanitized_adam_sheet_name() + "_"
        if config.get().marking_mode == "exercise":
            file_name += config.get().tutor_name + "_"
            file_name += "_".join([f"ex{exercise}" for exercise in args.exercises])
        elif config.get().marking_mode == "random":
            file_name += config.get().tutor_name
        elif config.get().marking_mode == "static":
            # Remove trailing underscore.
            file_name = file_name[:-1]
        else:
            logging.critical(f"Unsupported marking mode {config.get().marking_mode}!")
        return file_name

    def get_combined_feedback_file_name(self):
        return FEEDBACK_FILE_PREFIX + self._get_sanitized_adam_sheet_name()

    def get_combined_feedback_dir(self):
        return self.root / COMBINED_DIR_NAME

    def get_marks_file_path(self):
        return (
            self.root
            / f"points_{config.get().tutor_name.lower()}_{self._get_sanitized_adam_sheet_name()}.json"
        )

    def get_share_archive_path(self, exercises):
        share_archive_file_name = (
            SHARE_ARCHIVE_PREFIX
            + f"_{self._get_sanitized_adam_sheet_name()}_"
            + "_".join([f"ex{num}" for num in exercises])
            + ".zip"
        )
        return self.root / share_archive_file_name

    def get_share_archive_files(self) -> Iterator[Path]:
        """
        Return all share archive files under the current sheet root dir.
        """
        for share_archive_file in self.root.glob(
            SHARE_ARCHIVE_PREFIX + "*.zip"
        ):
            yield share_archive_file

    def get_all_team_dirs(self):
        """
        Return all team directories within the sheet root directory. It is assumed
        that all team directory names start with some digits, followed by an
        underscore, followed by more characters. In particular this excludes
        other directories that may be created in the sheet root directory, such as
        one containing combined feedback.
        """
        return _get_all_team_dirs(self.root)

    def get_relevant_team_dirs(self) -> Iterator[TeamDirectory]:
        """
        Return the team directories of the teams whose submission has to be
        corrected by the tutor running the script.
        """
        for team_dir in self.get_all_team_dirs():
            if not DO_NOT_MARK_PREFIX in team_dir.name:
                yield TeamDirectory(team_dir, self)

def do_not_mark(team_dir: Path):
    shutil.move(
        team_dir, team_dir.with_name(DO_NOT_MARK_PREFIX + team_dir.name)
    )


def _get_all_team_dirs(root: Path):
    for team_dir in root.iterdir():
        if team_dir.is_dir() and re.match(r"[0-9]+_.+", team_dir.name):
            yield team_dir


def create_sheet_info_file(root: Path, adam_sheet_name, team_dir_to_team: dict[Path, Team], exercises=None) -> SheetDirectory:
    """
    Write information generated during the execution of the 'init' command in a
    sheet info file. In particular a mapping from team directory names to teams
    and the name of the exercise sheet as given by ADAM. The latter is used as
    the seed to make random assignment of submissions to tutors consistent
    between tutors, but still vary from sheet to sheet. Later commands (e.g.
    'collect', or 'send') are meant to load the information stored in this file
    into the 'args' object and access it that way.
    """
    info_dict = {
        "team_dir_to_team": team_dir_to_team,
        "adam_sheet_name": adam_sheet_name,
    }
    if config.get().marking_mode == "exercise":
        info_dict["exercises"] = exercises
    write_json(root / SHEET_INFO_FILE_NAME, info_dict)

    return SheetDirectory(root)