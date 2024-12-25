import json
import logging
from importlib import resources
from pathlib import Path
import jsonschema

from . import config, schemas, sheets
from .students import Student
from .teams import Team

SUBMISSION_INFO_FILE_NAME = "submission.json"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"


class Submission:
    def __init__(self, team_dir: Path):
        self.root_dir = team_dir
        self.sheet = sheets.Sheet(self.root_dir.parent)
        self._load()

    def get_feedback_dir(self) -> Path:
        return self.root_dir / FEEDBACK_DIR_NAME

    def get_collected_feedback_dir(self) -> Path:
        return self.root_dir / FEEDBACK_COLLECTED_DIR_NAME

    def get_collected_feedback_path(self) -> Path:
        """
        Return the collected feedback file. This can be
        either a single pdf file, or a single zip archive. Throw an error if neither
        exists.
        """
        collected_feedback_dir = self.get_collected_feedback_dir()
        if not collected_feedback_dir.is_dir():
            logging.critical(
                "The directory for collected feedback at"
                f" '{str(collected_feedback_dir)}' does not exist. You probably"
                " have to run the 'collect' command first."
            )
        collected_feedback_files = list(collected_feedback_dir.iterdir())
        assert (
                len(collected_feedback_files) == 1
                and collected_feedback_files[0].is_file()
                and collected_feedback_files[0].suffix in [".pdf", ".zip"]
        )
        return collected_feedback_files[0]

    def get_combined_feedback_file(self) -> Path:
        """
        Return the combined feedback file. This is always a
        zip archive because in the usual case it will contain feedback from multiple
        tutors.
        """
        combined_feedback_dir = self.sheet.get_combined_feedback_path()
        if not combined_feedback_dir.is_dir():
            logging.critical(
                "The directory for combined feedback at"
                f" '{str(combined_feedback_dir)}' does not exist. You probably"
                " have to run the 'combine' command first."
            )
        combined_feedback_files = list(
            (combined_feedback_dir / self.root_dir.name).iterdir()
        )
        assert (
                len(combined_feedback_files) == 1
                and combined_feedback_files[0].is_file()
                and combined_feedback_files[0].suffix == ".zip"
        )
        return combined_feedback_files[0]

    def _load(self):
        """
        Load the submission info of the team directory. In particular the
        team, the ADAM ID and if it is a relevant team.
        """
        try:
            with open(
                    self.root_dir / SUBMISSION_INFO_FILE_NAME, "r", encoding="utf-8"
            ) as submission_info_file:
                submission_info = json.load(submission_info_file)
                submission_info_schema = json.loads(
                    resources.files(schemas).joinpath("submission-info-schema.json").read_text(encoding="utf-8"))
                jsonschema.validate(submission_info, submission_info_schema, jsonschema.Draft7Validator)
                self.team = Team(
                    [Student(*student) for student
                     in submission_info.get("team")],
                    submission_info.get("adam_id")
                )
                self.relevant = submission_info.get("relevant")
        except FileNotFoundError:
            logging.critical("The submission.json file does not exist.")
        except NotADirectoryError:
            logging.critical(f"The path '{self.root_dir}' is not a team directory.")
        except jsonschema.exceptions.ValidationError as error:
            logging.critical(f"The submission.json file does not have the right format: {error.message}")

    def __lt__(self, other):
        return self.root_dir < other.root_dir


def create_submission_info_file(_the_config: config.Config, team: Team,
                                team_id: str, team_dir: Path) -> None:
    """
    Write in each team directory a JSON file which contains the team,
    the ADAM ID of the team which ADAM sets anew with each exercise sheet,
    and if the tutor specified in the config has to mark this team.
    """
    team = team.to_tuples()
    relevant_teams = [team.to_tuples() for team in
                      _the_config.get_relevant_teams()]
    is_relevant = False
    submission_info = {}
    if team in relevant_teams:
        is_relevant = True
    submission_info.update({"team": team})
    submission_info.update({"adam_id": team_id})
    submission_info.update({"relevant": is_relevant})
    with open(
            team_dir / SUBMISSION_INFO_FILE_NAME, "w", encoding="utf-8"
    ) as submission_info_file:
        json.dump(
            submission_info,
            submission_info_file,
            indent=4,
            ensure_ascii=False,
        )
