import json
import logging
import os
import pathlib
import shutil
import tempfile
import textwrap
from typing import Union
from zipfile import ZipFile

from .. import config, sheets, submissions, utils
from ..teams import *


def extract_adam_zip(args) -> tuple[pathlib.Path, str]:
    """
    Unzips the given ADAM zip file and renames the directory to *target* if one
    is given. This is done stupidly right now, it would be better to extract to
    a temporary folder and then move once to the right location.
    """
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract ADAM zip contents to a temporary directory (or move in case it
        # was extracted automatically, this happened on an Apple system).
        if args.adam_zip_path.is_file():
            # Unzip to the directory within the zip file.
            # Should be the name of the exercise sheet,
            # for example "Exercise Sheet 2".
            with ZipFile(args.adam_zip_path, mode="r") as zip_file:
                zip_content = zip_file.namelist()
                sheet_root_dir = pathlib.Path(temp_dir) / zip_content[0]
                utils.filtered_extract(zip_file, pathlib.Path(temp_dir))
        else:
            # Assume the directory is an extracted ADAM zip.
            unzipped_path = pathlib.Path(args.adam_zip_path)
            unzipped_destination_path = (
                pathlib.Path(temp_dir) / unzipped_path.name
            )
            sheet_root_dir = pathlib.Path(
                shutil.copytree(unzipped_path, unzipped_destination_path)
            )
        # Store ADAM exercise sheet name.
        adam_sheet_name = sheet_root_dir.name
        destination = pathlib.Path(
            args.target if args.target else adam_sheet_name
        )
        if destination.exists():
            logging.critical(
                f"Extraction failed because the path '{destination}' exists"
                " already!"
            )
        sheet_root_dir = shutil.move(sheet_root_dir, destination)
    # Flatten intermediate directory.
    sub_dirs = list(sheet_root_dir.glob("*"))
    sub_dirs = [item for item in sub_dirs if item.suffix != '.xlsx']
    if len(sub_dirs) != 1:
        logging.critical(
            "The ADAM zip file contains an unexpected number of"
            f" directories/files ({len(sub_dirs)}), expected exactly 1"
            " subdirectory named either 'Abgaben' or 'Submissions'."
        )
    if sub_dirs[0].name not in ["Abgaben", "Submissions"]:
        logging.warning(
            "It looks like the format of the zip file created by ADAM has"
            " changed. This may require adaptions to this script, but I will"
            " try anyway."
        )
    utils.move_content_and_delete(sub_dirs[0], sheet_root_dir)
    return sheet_root_dir, adam_sheet_name


def mark_irrelevant_team_dirs(_the_config: config.Config, sheet: sheets.Sheet) -> None:
    """
    Indicate which team directories do not have to be marked by adding the
    `DO_NOT_MARK_PREFIX` to their directory name.
    """
    for submission in sheet.get_all_team_submission_info():
        if not submission.relevant:
            shutil.move(
                submission.root_dir,
                submission.root_dir.with_name(sheets.DO_NOT_MARK_PREFIX +
                                              submission.root_dir.name)
            )


def rename_team_dirs(sheet: sheets.Sheet) -> None:
    """
    The team directories are renamed to: team_id_LastName1-LastName2
    The team ID can be helpful to identify a team on the ADAM web interface.
    """
    for submission in sheet.get_all_team_submission_info():
        team_key = submission.team.get_team_key()
        team_dir = pathlib.Path(
            shutil.move(submission.root_dir, submission.root_dir.with_name(team_key))
        )


def flatten_team_dirs(sheet: sheets.Sheet) -> None:
    """
    There can be multiple directories within a "Team 00000" directory. This
    happens when multiple members of the team upload solutions. Sometimes, only
    one directory contains submitted files, in this case we remove the empty
    ones silently. In case multiple submissions exist, we put the files within
    them next to each other and print a warning.
    """
    for submission in sheet.get_all_team_submission_info():
        # Remove empty subdirectories.
        for team_submission_dir in submission.root_dir.iterdir():
            if team_submission_dir.is_dir() and len(list(team_submission_dir.iterdir())) == 0:
                team_submission_dir.rmdir()
        # Store the list of team submission directories in variable, because the
        # generator may include subdirectories of team submission directories
        # that have already been flattened.
        team_submission_dirs = [path for path in submission.root_dir.iterdir()
                                if not path.name == submissions.SUBMISSION_INFO_FILE_NAME]
        if len(team_submission_dirs) > 1:
            logging.warning(
                f"There are multiple submissions for group '{submission.root_dir.name}'!"
            )
        if len(team_submission_dirs) < 1:
            logging.warning(
                f"The submission of group '{submission.root_dir.name}' is empty!"
            )
        for team_submission_dir in team_submission_dirs:
            if team_submission_dir.is_dir():
                utils.move_content_and_delete(team_submission_dir,
                                              submission.root_dir)


def unzip_internal_zips(sheet: sheets.Sheet) -> None:
    """
    If multiple files are uploaded to ADAM, the submission becomes a single zip
    file. Here we extract this zip. I'm not sure if nested zip files are also
    extracted. Additionally, we flatten the directory as long as a level only
    consists of a single directory.
    """
    for submission in sheet.get_all_team_submission_info():
        for zip_file in submission.root_dir.glob("**/*.zip"):
            with ZipFile(zip_file, mode="r") as zf:
                utils.filtered_extract(zf, zip_file.parent)
            os.remove(zip_file)
        sub_dirs = [path for path in submission.root_dir.iterdir()
                    if not path.name == submissions.SUBMISSION_INFO_FILE_NAME]
        while len(sub_dirs) == 1 and sub_dirs[0].is_dir():
            utils.move_content_and_delete(sub_dirs[0], submission.root_dir)
            sub_dirs = [path for path in submission.root_dir.iterdir()
                        if not path.name == submissions.SUBMISSION_INFO_FILE_NAME]


def create_marks_file(_the_config: config.Config, sheet: sheets.Sheet,
                      args) -> None:
    """
    Write a json file to add the marks for all relevant teams and exercises.
    """
    exercise_dict: Union[str, dict[str, str]] = ""
    if _the_config.points_per == "exercise":
        if _the_config.marking_mode == "static":
            exercise_dict = {
                f"exercise_{i}": "" for i in range(1, args.num_exercises + 1)
            }
        elif _the_config.marking_mode == "exercise":
            exercise_dict = {f"exercise_{i}": "" for i in sheet.exercises}
    else:
        exercise_dict = ""

    marks_dict = {}
    for submission in sorted(sheet.get_relevant_submissions()):
        team_key = submission.team.get_team_key()
        marks_dict.update({team_key: exercise_dict})

    with open(sheet.get_marks_file_path(_the_config), "w", encoding="utf-8") as marks_json:
        json.dump(marks_dict, marks_json, indent=4, ensure_ascii=False)


def create_feedback_directories(_the_config: config.Config, sheet: sheets.Sheet) -> None:
    """
    Create a directory for every team that should be corrected by the tutor
    specified in the config. A copy of every non-PDF file is prefixed and placed
    in the feedback folder. The idea is that feedback on code files or similar
    can be added to these copies directly and files without feedback can simply
    be deleted. PDFs are not copied, instead a placeholder file with the correct
    name is created so that it can be overwritten by a real PDF file containing
    the feedback.
    """
    for submission in sheet.get_relevant_submissions():
        feedback_dir = submission.get_feedback_dir()
        feedback_dir.mkdir()

        feedback_file_name = sheet.get_feedback_file_name(_the_config)
        dummy_pdf_name = feedback_file_name + ".pdf.todo"
        pathlib.Path(feedback_dir / dummy_pdf_name).touch(exist_ok=True)

        # Copy non-pdf submission files into feedback directory with added
        # prefix.
        for submission_file in submission.root_dir.glob("*"):
            if submission_file.is_dir() or submission_file.suffix == ".pdf" \
                    or submission_file.name == submissions.SUBMISSION_INFO_FILE_NAME:
                continue
            this_feedback_file_name = (
                feedback_file_name + "_" + submission_file.name
            )
            shutil.copy(submission_file, feedback_dir / this_feedback_file_name)


def generate_xopp_files(sheet: sheets.Sheet) -> None:
    """
    Generate xopp files in the feedback directories that point to the single pdf
    in the submission directory and skip if multiple PDF files exist.
    """
    from pypdf import PdfReader

    def write_to_file(f, string):
        f.write(textwrap.dedent(string))

    logging.info("Generating .xopp files...")
    for submission in sheet.get_relevant_submissions():
        pdf_paths = list(submission.root_dir.glob("*.pdf"))
        if len(pdf_paths) != 1:
            logging.warning(
                f"Skipping .xopp file generation for {submission.root_dir.name}: No or"
                " multiple PDF files."
            )
            continue
        pdf_path = pdf_paths[0]
        feedback_dir = submission.get_feedback_dir()
        todo_paths = list(feedback_dir.glob("*.pdf.todo"))
        assert len(todo_paths) == 1
        todo_path = todo_paths[0]
        pages = PdfReader(pdf_path).pages
        # Strips the ".todo" and replaces ".pdf" by ".xopp".
        xopp_path = todo_path.with_suffix("").with_suffix(".xopp")
        if xopp_path.is_file():
            logging.warning(
                f"Skipping .xopp file generation for {submission.root_dir.name}: xopp file"
                " exists."
            )
            continue
        xopp_file = open(xopp_path, "w", encoding="utf-8")
        for i, page in enumerate(pages, start=1):
            width = page.mediabox.width
            height = page.mediabox.height
            if i == 1:  # Special entry for first page
                write_to_file(
                    xopp_file,
                    f"""\
                    <?xml version="1.0" standalone="no"?>
                    <xournal creator="Xournal++ 1.1.1" fileversion="4">
                    <title>Xournal++ document - see https://github.com/xournalpp/xournalpp</title>
                    <page width="{width}" height="{height}">
                    <background type="pdf" domain="absolute" filename="{pdf_path.resolve()}" pageno="{i}"/>
                    <layer/>
                    </page>""",  # noqa
                )
            else:
                write_to_file(
                    xopp_file,
                    f"""\
                    <page width="{width}" height="{height}">
                    <background type="pdf" pageno="{i}"/>
                    <layer/>
                    </page>""",
                )
        write_to_file(xopp_file, "</xournal>")
        xopp_file.close()
        todo_path.unlink()  # Delete placeholder todo file.
    logging.info("Done generating .xopp files.")


def print_missing_submissions(_the_config: config.Config, sheet: sheets.Sheet) -> None:
    """
    Print all teams that are listed in the config file, but whose submission is
    not present in the zip downloaded from ADAM.
    """
    teams_who_submitted = []
    for submission in sheet.get_all_team_submission_info():
        teams_who_submitted.append(submission.team)
    missing_teams = [
        team for team in _the_config.teams if team not in
        teams_who_submitted
    ]
    if missing_teams:
        logging.warning("There are no submissions for the following team(s):")
        for missing_team in missing_teams:
            print(f"* {missing_team.last_names_to_string()}")


def lookup_teams(_the_config: config.Config,
                 team_dir: pathlib.Path) -> tuple[str, list[Team]]:
    """
    Extracts the team ID from the directory name and searches for teams
    based on the extracted email address from the subdirectory name.
    """
    team_id = team_dir.name.split(" ")[1]
    submission_dir = list(team_dir.iterdir())[0]
    submission_email = submission_dir.name.split("_")[-2]
    teams = [
        team
        for team in _the_config.teams
        if any(student.email == submission_email for student in team.members)
    ]
    return team_id, teams


def validate_team_dirs(_the_config: config.Config,
                       sheet_root_dir: pathlib.Path) -> None:
    """
    Checks whether all students are assigned to a team and checks for
    multiple submissions from teams under different IDs.
    """
    adam_id_to_team = {}
    for team_dir in sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        team_id, teams = lookup_teams(_the_config, team_dir)
        if len(teams) == 0:
            submission_email = list(team_dir.iterdir())[0].name.split("_")[-2]
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
                    f"There are multiple submissions for team '{teams[0].members}'"
                    f" under separate ADAM IDs ({existing_id} and {team_id})!"
                    " This probably means that multiple members of a team"
                    " submitted solutions without forming a team on ADAM. You"
                    " will have to combine the submissions manually."
                )
        adam_id_to_team.update({team_id: teams[0]})


def create_all_submission_info_files(_the_config: config.Config,
                                     sheet_root_dir: pathlib.Path) -> None:
    """
    Creates the submission info JSON files in all team directories.
    """
    for team_dir in sheet_root_dir.iterdir():
        if team_dir.is_dir():
            team_id, teams = lookup_teams(_the_config, team_dir)
            submissions.create_submission_info_file(
                _the_config, teams[0], team_id, team_dir
            )


def init(_the_config: config.Config, args) -> None:
    """
    Prepares the directory structure holding the submissions.
    """
    # Catch wrong combinations of marking_mode/points_per/-n/-e.
    # Not possible eariler because marking_mode and points_per are given by the
    # config file.
    if _the_config.points_per == "exercise":
        if _the_config.marking_mode == "exercise" and not args.exercises:
            logging.critical(
                "You must provide a list of exercise numbers to be marked with "
                "the '-e' flag, for example '-e 1 3 4'. In case the '-e' flag "
                "is the last option before the ADAM zip path, make sure to "
                "separate exercise numbers from the path using '--'."
            )
        if _the_config.marking_mode == "static" and not args.num_exercises:
            logging.critical(
                "You must provide the number of exercises in the sheet with "
                "the '-n' flag, for example '-n 5'."
            )
    else:  # points per sheet
        if _the_config.marking_mode == "exercise":
            logging.critical(
                "Points must be given per exercise if marking is done per "
                "exercise. Set the value of 'poins_per' to 'exercise' or "
                "change the marking mode."
            )
        if args.num_exercises or args.exercises:
            logging.warning(
                "'points_per' is 'sheet', so the flags '-n' and '-e' are "
                "ignored."
            )

    # Sort the list exercises of flag "-e" to make printing nicer later.
    if args.exercises:
        args.exercises.sort()

    # The adam zip file is expected to have the following structure:
    # ADAM Exercise Sheet Name.zip
    # └── ADAM Exercise Sheet Name
    #     └── Abgaben
    #         ├── Team 12345
    #         .   └── Muster_Hans_hans.muster@unibas.ch_000000
    #         .       └── submission.pdf or submission.zip
    sheet_root_dir, adam_sheet_name = extract_adam_zip(args)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── Team 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    validate_team_dirs(_the_config, sheet_root_dir)
    create_all_submission_info_files(_the_config, sheet_root_dir)
    sheet = sheets.create_sheet_info_file(
        sheet_root_dir, adam_sheet_name, _the_config, args.exercises
    )
    print_missing_submissions(_the_config, sheet)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── Team 12345
    # .   ├── Muster_Hans_hans.muster@unibas.ch_000000
    # .   │   └── submission.pdf or submission.zip
    # .   └── submission.json
    # └── sheet.json
    rename_team_dirs(sheet)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── Muster_Hans_hans.muster@unibas.ch_000000
    # .   │   └── submission.pdf or submission.zip
    # .   └── submission.json
    # └── sheet.json
    flatten_team_dirs(sheet)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── submission.pdf or submission.zip
    # .   └── submission.json
    # └── sheet.json
    unzip_internal_zips(sheet)

    # From here on, we need information about relevant teams.
    mark_irrelevant_team_dirs(_the_config, sheet)

    if _the_config.use_marks_file:
        create_marks_file(_the_config, sheet, args)

    create_feedback_directories(_the_config, sheet)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── feedback
    # .   │   ├── feedback.pdf.todo
    # .   │   └── feedback_copy-of-submitted-file.cc
    # .   ├── submission.pdf or submission files
    # .   └── submission.json
    # ├── sheet.json
    # └── points.json
    if _the_config.xopp:
        generate_xopp_files(sheet)
