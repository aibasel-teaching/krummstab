import json
import logging
from pathlib import Path
import re
import types
from typing import Union, Callable, get_origin, get_args

Student = tuple[str, str, str]
Team = list[Student]

def ensure_list(value):
    if not value:
        return []
    elif isinstance(value, list):
        return value
    else:
        return [value]

_the_config = None

def load(*config_paths: list[Path]) -> None:
    global _the_config
    _the_config = Config(config_paths)

def get():
    global _the_config
    if _the_config is None:
        logging.critical("Accessed config before loading it.")
    return _the_config

class Config:
    def __init__(self, config_paths: list[Path]) -> None:
        data = {}
        for path in config_paths:
            logging.info(f"Reading config file '{path}'")
            data.update(json.loads(path.read_text(encoding="utf-8")))

        self._extract_value(data, "tutor_name", str),
        self._extract_value(data, "ignore_feedback_suffix", list[str], 
            _validate_elements(_validate_suffix), optional=True, default=[]),

            # Email settings, currently all optional because not fully functional.
        self._extract_value(data, "tutor_email", str,
            _validate_email, optional=True),
        self._extract_value(data, "assistant_email", str,
            _validate_email, optional=True),
        self._extract_value(data, "feedback_email_cc", list[str], 
            _validate_elements(_validate_email), optional=True, default=[]),
        self._extract_value(data, "smtp_url", str, optional=True),
        self._extract_value(data, "smtp_port", int, optional=True),
        self._extract_value(data, "smtp_user", str, optional=True),

        self._extract_value(data, "lecture_title", str, _validate_non_empty),
        self._extract_value(data, "marking_mode", str,
            _validate_choices("static", "random", "exercise")),

        self._extract_value(data, "max_team_size", int, _validate_positive),
        self._extract_value(data, "use_marks_file", bool),
        self._extract_value(data, "points_per", str,
            _validate_choices("sheet", "exercise")),
        self._extract_value(data, "min_point_unit", Union[float, int],
            _validate_positive),


        # We currently plan to support the following marking modes.
        # static:   Every tutor corrects the submissions of the teams assigned to
        #           that tutor. These will usually be the teams in that tutors
        #           exercise class.
        # random:   Every tutor corrects some submissions which are assigned
        #           randomly with every sheet.
        # exercise: Every tutor corrects some exercise(s) on all sheets.
        if self.marking_mode == "static":
            self._extract_value(data, "classes", dict[str, list[Team]])
            if self.tutor_name not in self.classes:
                logging.critical(f"Did not find a class for '{self.tutor_name}' in the config.")
            self.teams = [team for classs in self.classes.values() for team in classs]
        else:
            self._extract_value(data, "tutor_list", list[str])
            if self.tutor_name not in self.tutor_list:
                logging.critical(f"Did not find '{self.tutor_name}' in tutor_list in the config.")
            self._extract_value(data, "teams", list[Team])

        _validate_teams(self.teams, self.max_team_size)
        # Sort teams and their students to make iterating over them
        # predictable, independent of their order in config.json.
        for team in self.teams:
            team.sort()
        self.teams.sort()
        logging.info("Processed config successfully.")

    def _extract_value(self, data, key, expected_type, validators=None, optional=False, default=None):
        if not optional and key not in data:
            logging.critical(f"Expected option '{key}' to be configured but "
                             "did not find it in any config file.")
        value = data.get(key, default)
        _validate_expected_type(key, value, expected_type)
        for validator in ensure_list(validators):
            validator(key, value)
        setattr(self, key, value)

def _validate_expected_type(key: str, value: str, expected_type: type) -> None:
    if isinstance(expected_type, types.GenericAlias):
        if get_origin(expected_type) == list:
            if not isinstance(value, list):
                logging.critical(f"Expected option '{key}' to be a list but "
                                f"got '{type(value).__name__}'.")
            element_type = get_args(expected_type)[0]
            for i, element in enumerate(value):
                _validate_expected_type(f"{key}[{i}]", element, element_type)
        elif get_origin(expected_type) == dict:
            if not isinstance(value, dict):
                logging.critical(f"Expected option '{key}' to be a dict but "
                                f"got '{type(value).__name__}'.")
            key_type = get_args(expected_type)[0]
            value_type = get_args(expected_type)[1]
            for i, (d_key, d_value) in enumerate(value.items()):
                _validate_expected_type(f"Key {i} of {key}", d_key, key_type)
                _validate_expected_type(f"{key}[{d_key}]", d_value, value_type)
        elif get_origin(expected_type) == tuple:
            if not (isinstance(value, list) or isinstance(value, tuple)):
                logging.critical(f"Expected option '{key}' to be a tuple but "
                                f"got '{type(value).__name__}'.")
            elmenent_types = get_args(expected_type)
            if len(value) != len(elmenent_types):
                logging.critical(f"Expected option '{key}' to be a tuple of "
                                    f"length {len(elmenent_types)} but got a tuple of "
                                    f"length {len(value)}.")
            for i, (element, elmenent_type) in enumerate(zip(value, elmenent_types)):
                _validate_expected_type(f"Element {i} of {key}", element, elmenent_type)
        else:
            logging.critical("Unsupported Generic type")
    elif not isinstance(value, expected_type):
        logging.critical(f"Expected option '{key}' to be "
                        f"'{expected_type.__name__}' but "
                        f"got '{type(value).__name__}'.")


def _validate_email(key: str, value: str) -> None:
    if not re.match(r"[^@]+@[^@]+\.[^@]+", value):
        logging.critical(f"Value '{value}' for option '{key}' does not match "
                         "the format of an email address.")


def _validate_suffix(key: str, value: str) -> None:
    if not (value.startswith(".") and len(value) > 1):
        logging.critical(f"Value '{value}' for option '{key}' does not match "
                         "the format of a suffix (e.g. '.xopp').")


def _validate_elements(element_validator) -> Callable[[list], None]:
    def validator(key: str, value: list) -> bool:
        for i, element in enumerate(value):
            return element_validator(f"{key}[{i}]", element)
    return validator


def _validate_non_empty(key: str, value: str) -> bool:
    if not value:
        logging.critical(f"Expected nonempty value for option '{key}'.")


def _validate_choices(*choices: list[str]) -> Callable[[str], bool]:
    def validator(key: str, value: str) -> bool:
        if value.lower() not in [c.lower() for c in choices]:
            logging.critical(f"Value '{value}' for option '{key}' must be one "
                             f"of {{{', '.join(choices)}}}.")
    return validator


def _validate_positive(key: str, value: Union[int, float]) -> bool:
    if value <= 0:
        logging.critical(f"Expected positive value for option '{key}' but got '{value}'.")

def _validate_teams(teams: list[Team], max_team_size) -> None:
    """
    Check for duplicate entries and maximal team size.
    """
    all_students: list[tuple[str, str]] = []
    all_emails: list[str] = []
    for team in teams:
        if  len(team) > max_team_size:
            logging.critical(f"Team with size {len(team)} violates maximal team size.")
        for first, last, email in team:
            _validate_email("Email of {first} {last}", email)
            all_students.append((first, last))
            all_emails.append(email)
    if len(all_students) != len(set(all_students)):
        logging.critical("There are duplicate students in the config file!")
    if len(all_emails) != len(set(all_emails)):
        logging.critical("There are duplicate student emails in the config file!")
