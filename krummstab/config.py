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

    def unsupported_marking_mode_error(self) -> None:
        """
        Throw an error if a marking mode is encountered that is not handled
        correctly. This is primarily used in the else case of if-elif-else branches
        on the marking_mode, which shouldn't be reached in regular operation anyway
        because the config settings should be validated first. But this function
        offers an easy way to find sections to consider if we ever want to add a new
        marking_mode.
        """
        logging.critical(f"Unsupported marking mode {self.marking_mode}!")

    def get_relevant_teams(self) -> list[Team]:
        """
        Get a list of teams that the tutor specified in the config has to mark.
        We rename the directories using the `DO_NOT_MARK_PREFIX` and thereafter only
        access relevant teams via `get_relevant_team_dirs()`.
        """
        if self.marking_mode == "static":
            return self.classes[self.tutor_name]
        elif self.marking_mode == "exercise":
            return self.teams
        else:
            self.unsupported_marking_mode_error()
            return []


def team_to_string(team: Team) -> str:
    """
    Concatenate the last names of students to get a pretty-ish string
    representation of teams.
    """
    return "_".join(sorted([student[1].replace(" ", "-") for student in team]))


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
