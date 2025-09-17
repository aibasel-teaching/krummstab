import logging
from typing import Optional

from .students import *


class Team:
    def __init__(self, members: list[Student], adam_id: Optional[str] = None):
        self.members = sorted(members)
        self.adam_id = adam_id

    def __eq__(self, other) -> bool:
        return sorted(self.members) == sorted(other.members)

    def __format__(self, spec) -> str:
        """
        Get a pretty printed string representation of a team.
        """
        return ", ".join(f"{member}" for member in self.members)

    def get_first_names(self) -> list[str]:
        """
        Get a list of the first names of all team members.
        """
        return [member.first_name for member in self.members]

    def get_emails(self) -> list[str]:
        """
        Get a list of the emails of all team members.
        """
        return [member.email for member in self.members]

    def get_team_key(self) -> str:
        """
        Create a string representation of the team with the ADAM ID:
        team_id_LastName1-LastName2
        """
        if not self.adam_id:
            logging.critical(
                "Internal error for developers: get_team_key() cannot be used "
                "when the team's adam_id is None."
            )
        return self.adam_id + "_" + self.last_names_to_string()

    def last_names_to_string(self) -> str:
        """
        Concatenate the last names of students to get a pretty-ish string
        representation of a team.
        """
        return "_".join(
            sorted(
                [member.last_name.replace(" ", "-") for member in self.members]
            )
        )

    def to_tuples(self) -> list[tuple[str, str, str]]:
        """
        Get a tuples of strings representation of a team.
        """
        return [member.to_tuple() for member in self.members]


def create_email_to_name_dict(teams: list[Team]) -> dict[str, tuple[str, str]]:
    """
    Maps the students' email addresses to their first and last name.
    """
    email_to_name = {}
    for team in teams:
        for member in team.members:
            email_to_name[member.email] = (member.first_name, member.last_name)
    return email_to_name
