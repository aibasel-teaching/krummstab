import logging
import pathlib
import tempfile
from collections import defaultdict

from .. import config, errors, strings, utils
from ..teams import Team


def validate_team_size(
    max_team_size: int, submission_teams: list[Team]
) -> None:
    """
    Checks if the team size of the submission teams does not exceed the
    maximum allowed team size.
    """
    teams = [
        team for team in submission_teams if len(team.members) > max_team_size
    ]
    if not teams:
        return
    logging.warning(
        "The following teams that have submitted have more members than "
        "allowed."
    )
    print(strings.SEPARATOR_LINE)
    for team in teams:
        print(f"* {team}")
        print(strings.SEPARATOR_LINE)


def warn_about_restructured_teams(
    config_teams: list[Team], restructured_teams: list[Team]
) -> None:
    logging.warning(
        "The following team(s) have submitted but are structured differently "
        "in the config."
    )
    print(strings.SEPARATOR_LINE)
    for restructured_team in restructured_teams:
        print(f"{restructured_team}\n")
        # Get config teams that share a member with the submission team.
        matching_config_teams = [
            config_team
            for config_team in config_teams
            if any(member in config_team for member in restructured_team)
        ]
        if matching_config_teams:
            print("Matching config team(s):")
            for matching_team in matching_config_teams:
                print(f"* {matching_team}")
        new_students = [
            member
            for member in restructured_team
            if not any(member in config_team for config_team in config_teams)
        ]
        if new_students:
            print(
                "Member(s) of the restructured team that do not appear in the "
                "config:"
            )
            for student in new_students:
                print(f"* {student}")
        print(strings.SEPARATOR_LINE)


def check_team_consistency(
    config_teams: list[Team],
    submission_teams: list[Team],
) -> None:
    """
    Checks if the teams defined in the config `config_teams` are consistent with
    the teams that submitted `submission_teams` and prints warnings in case of
    inconsistencies.
    """
    # Get teams that submitted but are not in the config and contain at least
    # one member that is mentioned in the config.
    restructured_teams = [
        submission_team
        for submission_team in submission_teams
        if submission_team not in config_teams
        and any(
            member in config_team
            for member in submission_team
            for config_team in config_teams
        )
    ]
    if restructured_teams:
        warn_about_restructured_teams(config_teams, restructured_teams)
    # Get teams that submitted but none of its members are mentioned in the
    # config.
    new_teams = [
        submission_team
        for submission_team in submission_teams
        if all(
            member not in config_team
            for member in submission_team
            for config_team in config_teams
        )
    ]
    if new_teams:
        logging.warning(
            "The following team(s) have submitted and their members do not "
            "appear in the config."
        )
        print(strings.SEPARATOR_LINE)
        for new_team in new_teams:
            print(f"{new_team}")
            print(strings.SEPARATOR_LINE)


def print_count_line(
    entry: str,
    tutor_assignment_count,
    tutor_field_width: int,
    count_field_width: int,
) -> None:
    print(
        f"* {entry + ':':{tutor_field_width}} {tutor_assignment_count[entry]:{count_field_width}d}"
    )


def check(_the_config: config.Config, args) -> None:
    """
    When submissions for a new sheet are out but before tutors start giving
    feedback, the assistant may want to verify whether the marking load is
    evenly distributed among the tutors.
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        utils.unzip_or_move_adam_zip(args.adam_zip_path, temp_dir)
        # Check if the zip has the expected structure.
        children = list(pathlib.Path(temp_dir).iterdir())
        # We expect the zip to contain a single subdirectory with the
        # exercise sheet name given on ADAM.
        if len(children) != 1 or not children[0].is_dir():
            errors.unexpected_zip_structure(args.adam_zip_path)
        temp_sheet_root_dir = children[0]
        grand_children = list(pathlib.Path(temp_sheet_root_dir).iterdir())
        # Within its single subdirectory, we expect the zip to contain a
        # single subsubdirectory named either "Abgaben" or "Submissions" and
        # a single spreadsheet with information about the submissions.
        if (
            len(grand_children) != 2
            or not any(
                grand_child.is_file() and grand_child.suffix == ".xlsx"
                for grand_child in grand_children
            )
            or not any(grand_child.is_dir() for grand_child in grand_children)
        ):
            errors.unexpected_zip_structure(args.adam_zip_path)
        # Read teams from spreadsheet.
        spreadsheet_file = list(temp_sheet_root_dir.glob("*.xlsx"))[0]
        submission_teams = utils.read_teams_from_adam_spreadsheet(
            spreadsheet_file
        ).values()

    # Check consistency.
    validate_team_size(_the_config.max_team_size, submission_teams)
    check_team_consistency(_the_config.teams, submission_teams)

    if _the_config.marking_mode != "static":
        return

    # Calculate balance.
    student_email_to_tutor_dict = (
        _the_config.create_student_email_to_tutor_dict()
    )
    tutor_list = _the_config.classes.keys()
    team_to_tutors = utils.create_submission_team_to_tutors_dict(
        submission_teams, student_email_to_tutor_dict, tutor_list
    )
    tutor_assignment_count = defaultdict(int)
    tutor_assignment_count["total"] = len(submission_teams)
    for _, tutors in team_to_tutors.items():
        if len(tutors) == 1:
            tutor_assignment_count[list(tutors)[0]] += 1
        elif len(tutors) == len(tutor_list):
            tutor_assignment_count["unassigned"] += 1
        else:
            tutor_assignment_count["unclear"] += 1
    # Print balance overview table.
    logging.info("Marking load balance:")
    tutor_field_width = (
        max([len(tutor) for tutor in tutor_assignment_count.keys()]) + 1
    )
    count_field_width = max(
        [len(str(count)) for count in tutor_assignment_count.values()]
    )
    total_width = tutor_field_width + count_field_width + 3
    for tutor in tutor_list:
        print_count_line(
            tutor, tutor_assignment_count, tutor_field_width, count_field_width
        )
    print("-" * total_width)
    print_count_line(
        "unclear", tutor_assignment_count, tutor_field_width, count_field_width
    )
    print_count_line(
        "unassigned",
        tutor_assignment_count,
        tutor_field_width,
        count_field_width,
    )
    print("=" * total_width)
    print_count_line(
        "total", tutor_assignment_count, tutor_field_width, count_field_width
    )
