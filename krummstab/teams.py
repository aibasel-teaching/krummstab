class Team:
    def __init__(self, members, adam_id):
        self.members = members
        self.adam_id = adam_id

    def __eq__(self, other) -> bool:
        return self.members == other.members

    def get_first_names(self) -> list[str]:
        """
        Get a list of the first names of all team members.
        """
        return [member[0] for member in self.members]

    def get_emails(self) -> list[str]:
        """
        Get a list of the emails of all team members.
        """
        return [member[2] for member in self.members]

    def get_team_key(self) -> str:
        """
        Create a string representation of the team with the ADAM ID:
        team_id_LastName1-LastName2
        """
        return self.adam_id + "_" + self.last_names_to_string()

    def last_names_to_string(self) -> str:
        """
        Concatenate the last names of students to get a pretty-ish string
        representation of a team.
        """
        return "_".join(
            sorted([member[1].replace(" ", "-") for member in self.members])
        )

    def to_tuples(self) -> list[tuple[str, str, str]]:
        """
        Get a tuples of strings representation of a team.
        """
        return [(member[0], member[1], member[2]) for member in self.members]


def create_email_to_name_dict(teams: list[Team]) -> dict[str, tuple[str, str]]:
    """
    Maps the students' email addresses to their first and last name.
    """
    email_to_name = {}
    for team in teams:
        for first_name, last_name, email in team.members:
            email_to_name[email] = (first_name, last_name)
    return email_to_name
