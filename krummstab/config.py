import json
import logging
from importlib import resources
from pathlib import Path
import jsonschema

from . import errors, schemas
from .teams import *


# Within this class, Team objects are created with their adam_id set to None
# because the adam_id is not available at the time of the Config class
# instantiation.
class Config:
    def __init__(self, config_paths: list[Path]) -> None:
        data = {}
        for path in config_paths:
            logging.info(f"Reading config file '{path}'")
            data.update(json.loads(path.read_text(encoding="utf-8")))
        config_schema = json.loads(resources.files(schemas).joinpath("config-schema.json").read_text(encoding="utf-8"))
        jsonschema.validate(data, config_schema, jsonschema.Draft7Validator)

        for key, value in data.items():
            setattr(self, key, value)

        # We currently plan to support the following marking modes.
        # static:   Every tutor corrects the submissions of the teams assigned to
        #           that tutor. These will usually be the teams in that tutors
        #           exercise class.
        # exercise: Every tutor corrects some exercise(s) on all sheets.
        if self.marking_mode == "static":
            if self.tutor_name not in self.classes:
                logging.critical(f"Did not find a class for '{self.tutor_name}' in the config.")
            self.teams = [team for classs in self.classes.values() for team in classs]
        else:
            if self.tutor_name not in self.tutor_list:
                logging.critical(f"Did not find '{self.tutor_name}' in tutor_list in the config.")

        # Sort teams and their students to make iterating over them
        # predictable, independent of their order in config.json.
        for team in self.teams:
            team.sort()
        self.teams.sort()

        # Create Team objects with their adam_id set to None because
        # the adam_id is not available here
        if self.marking_mode == "static":
            self.classes = {
                tutor: [Team([Student(*student) for student in team], None)
                        for team in teams]
                for tutor, teams in self.classes.items()
            }
        self.teams = [Team([Student(*student) for student in team], None)
                      for team in self.teams]
        _validate_teams(self.teams, self.max_team_size)
        logging.info("Processed config successfully.")

    def get_relevant_teams(self) -> list[Team]:
        """
        Get a list of teams that the tutor specified in the config has to mark.
        We rename the directories using the `DO_NOT_MARK_PREFIX` and thereafter only
        access relevant teams via `get_relevant_submissions()`.
        """
        if self.marking_mode == "static":
            return self.classes[self.tutor_name]
        elif self.marking_mode == "exercise":
            return self.teams
        else:
            errors.unsupported_marking_mode_error(self.marking_mode)
            return []


def _validate_teams(teams: list[Team], max_team_size) -> None:
    """
    Check for duplicate entries and maximal team size.
    """
    all_students: list[tuple[str, str]] = []
    all_emails: list[str] = []
    for team in teams:
        if len(team.members) > max_team_size:
            logging.critical(f"Team with size {len(team.members)} violates maximal "
                             f"team size.")
        for member in team.members:
            all_students.append((member.first_name, member.last_name))
            all_emails.append(member.email)
    if len(all_students) != len(set(all_students)):
        logging.critical("There are duplicate students in the config file!")
    if len(all_emails) != len(set(all_emails)):
        logging.critical("There are duplicate student emails in the config file!")
