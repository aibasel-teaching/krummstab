class Team:
    def __init__(self, members, adam_id):
        self.members = members
        self.adam_id = adam_id

    def get_team_key(self) -> str:
        """
        Create a string representation of the team with the ADAM ID:
        team_id_LastName1-LastName2
        """
        return self.adam_id + "_" + self.team_to_string()

    def team_to_string(self) -> str:
        """
        Concatenate the last names of students to get a pretty-ish string
        representation of teams.
        """
        return "_".join(
            sorted([member[1].replace(" ", "-") for member in self.members])
        )


def create_email_to_name_dict(teams) -> dict[str, tuple[str, str]]:
    """
    Maps the students' email addresses to their first and last name.
    """
    email_to_name = {}
    for team in teams:
        for first_name, last_name, email in team:
            email_to_name[email] = (first_name, last_name)
    return email_to_name
