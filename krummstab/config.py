import json
import logging
from importlib import resources
from pathlib import Path
import jsonschema
from . import schemas

Student = tuple[str, str, str]
Team = list[Student]


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

        _validate_teams(self.teams, self.max_team_size)
        # Sort teams and their students to make iterating over them
        # predictable, independent of their order in config.json.
        for team in self.teams:
            team.sort()
        self.teams.sort()
        logging.info("Processed config successfully.")


def _validate_teams(teams: list[Team], max_team_size) -> None:
    """
    Check for duplicate entries and maximal team size.
    """
    all_students: list[tuple[str, str]] = []
    all_emails: list[str] = []
    for team in teams:
        if len(team) > max_team_size:
            logging.critical(f"Team with size {len(team)} violates maximal team size.")
        for first, last, email in team:
            all_students.append((first, last))
            all_emails.append(email)
    if len(all_students) != len(set(all_students)):
        logging.critical("There are duplicate students in the config file!")
    if len(all_emails) != len(set(all_emails)):
        logging.critical("There are duplicate student emails in the config file!")
