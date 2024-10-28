import json
import logging
import shutil
from importlib import resources
from pathlib import Path
import jsonschema

from . import config, schemas

SUBMISSION_INFO_FILE_NAME = "submission.json"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"


class SubmissionInfo:
    def __init__(self, team_dir: Path):
        self.team_dir = team_dir
        self._load()

    def get_feedback_dir(self) -> Path:
        return self.team_dir / FEEDBACK_DIR_NAME

    def get_collected_feedback_dir(self) -> Path:
        return self.team_dir / FEEDBACK_COLLECTED_DIR_NAME

    def get_collected_feedback_file(self) -> Path:
        """
        Given a team directory, return the collected feedback file. This can be
        either a single pdf file, or a single zip archive. Throw an error if neither
        exists.
        """
        collected_feedback_dir = self.team_dir / FEEDBACK_COLLECTED_DIR_NAME
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

    def get_team_key(self) -> str:
        """
        Create a string representation of the team with the ADAM ID: team_id_LastName1-LastName2
        """
        return self.adam_id + "_" + config.team_to_string(self.team)

    def _load(self):
        """
        Load the submission info of the team directory. In particular the
        team, the ADAM ID and if it is a relevant team.
        """
        try:
            with open(
                    self.team_dir / SUBMISSION_INFO_FILE_NAME, "r", encoding="utf-8"
            ) as submission_info_file:
                submission_info = json.load(submission_info_file)
                submission_info_schema = json.loads(
                    resources.files(schemas).joinpath("submission-info-schema.json").read_text(encoding="utf-8"))
                jsonschema.validate(submission_info, submission_info_schema, jsonschema.Draft7Validator)
                self.team = submission_info.get("team")
                self.adam_id = submission_info.get("adam_id")
                self.relevant = submission_info.get("relevant")
        except FileNotFoundError:
            logging.critical("The submission.json file does not exist.")
        except NotADirectoryError:
            logging.critical(f"The path '{self.team_dir}' is not a team directory.")
        except jsonschema.exceptions.ValidationError:
            logging.critical("The submission.json file has not the right format.")

    def __lt__(self, other):
        return self.team_dir < other.team_dir


def create_submission_info_files(_the_config: config.Config, sheet_root_dir: Path) -> None:
    """
    Write in each team directory a JSON file which contains the team,
    the ADAM ID of the team which ADAM sets anew with each exercise sheet,
    and if the tutor specified in the config has to mark this team. At
    the same time, the "Team " prefix is removed from directory names.
    """
    relevant_teams = _the_config.get_relevant_teams()
    adam_id_to_team = {}
    for team_dir in sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        team_id = team_dir.name.split(" ")[1]
        submission_dir = list(team_dir.iterdir())[0]
        submission_email = submission_dir.name.split("_")[-2]
        teams = [
            team
            for team in _the_config.teams
            if any(submission_email in student for student in team)
        ]
        if len(teams) == 0:
            logging.critical(
                f"The student with the email '{submission_email}' is not "
                "assigned to a team. Your config file is likely out of date."
                "\n"
                "Please update the config file such that it reflects the team "
                "assignments of this week correctly and share the updated "
                "config file with your fellow tutors and the teaching "
                "assistant."
            )
        # The case that a student is assigned to multiple teams would already be
        # caught when reading in the config file, so we just assert that this is
        # not the case here.
        assert len(teams) == 1
        # TODO: if team[0] in adam_id_to_team.values(): -> multiple separate
        # submissions
        # Catch the case where multiple members of a team independently submit
        # solutions without forming a team on ADAM and print a warning.
        for existing_id, existing_team in adam_id_to_team.items():
            if existing_team == teams[0]:
                logging.warning(
                    f"There are multiple submissions for team '{teams[0]}'"
                    f" under separate ADAM IDs ({existing_id} and {team_id})!"
                    " This probably means that multiple members of a team"
                    " submitted solutions without forming a team on ADAM. You"
                    " will have to combine the submissions manually."
                )
        adam_id_to_team.update({team_id: teams[0]})
        team_dir = Path(
            shutil.move(team_dir, team_dir.with_name(team_id))
        )
        is_relevant = False
        submission_info = {}
        if teams[0] in relevant_teams:
            is_relevant = True
        submission_info.update({"team": teams[0]})
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
