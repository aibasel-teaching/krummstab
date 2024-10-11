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
        schema_config_shared = json.loads(
            resources.read_text(schemas, "config-shared-schema.json", encoding="utf-8"))
        schema_config_individual = json.loads(
            resources.read_text(schemas, "config-individual-schema.json", encoding="utf-8"))
        jsonschema.validate(json.loads(config_paths[0].read_text(encoding="utf-8")),
                            schema_config_shared, jsonschema.Draft7Validator)
        jsonschema.validate(json.loads(config_paths[1].read_text(encoding="utf-8")),
                            schema_config_individual, jsonschema.Draft7Validator)
        data = {}
        for path in config_paths:
            logging.info(f"Reading config file '{path}'")
            data.update(json.loads(path.read_text(encoding="utf-8")))

        self.tutor_name = data.get("tutor_name")
        self.ignore_feedback_suffix = data.get("ignore_feedback_suffix", [])

        self.tutor_email = data.get("tutor_email")
        self.assistant_email = data.get("assistant_email")
        self.email_signature = data.get("email_signature")
        self.feedback_email_cc = data.get("feedback_email_cc", [])
        self.smtp_url = data.get("smtp_url")
        self.smtp_port = data.get("smtp_port")
        self.smtp_user = data.get("smtp_user")

        self.lecture_title = data.get("lecture_title")
        self.marking_mode = data.get("marking_mode")

        self.max_team_size = data.get("max_team_size")
        self.use_marks_file = data.get("use_marks_file")
        self.points_per = data.get("points_per")
        self.min_point_unit = data.get("min_point_unit")

        # We currently plan to support the following marking modes.
        # static:   Every tutor corrects the submissions of the teams assigned to
        #           that tutor. These will usually be the teams in that tutors
        #           exercise class.
        # exercise: Every tutor corrects some exercise(s) on all sheets.
        if self.marking_mode == "static":
            self.classes = data.get("classes")
            if self.tutor_name not in self.classes:
                logging.critical(f"Did not find a class for '{self.tutor_name}' in the config.")
            self.teams = [team for classs in self.classes.values() for team in classs]
        else:
            self.tutor_list = data.get("tutor_list")
            if self.tutor_name not in self.tutor_list:
                logging.critical(f"Did not find '{self.tutor_name}' in tutor_list in the config.")
            self.teams = data.get("teams")

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
