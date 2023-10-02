#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
This script uses the following, potentially ambiguous, terminology:
team     - set of students that submit together, as defined on ADAM
class    - set of teams that are in the same exercise class
relevant - a team whose submission has to be marked by the tutor running the
           script is considered 'relevant'
to mark  - grading/correcting a sheet; giving feedback
marks    - points awarded for a sheet or exercise
"""
import argparse
import hashlib
import json
import mimetypes
import os
import pathlib
import random
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap

from zipfile import ZipFile

# For typing annotations.
from typing import Any, Optional, Union
from collections.abc import Iterator

# Needed for email stuff.
import smtplib
import ssl
from email.message import EmailMessage
from getpass import getpass

Student = tuple[str, str, str]
Team = list[Student]

DEFAULT_SHARED_CONFIG_FILE = "config-shared.json"
DEFAULT_INDIVIDUAL_CONFIG_FILE = "config-individual.json"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"
FEEDBACK_FILE_PREFIX = "feedback_"
SHEET_INFO_FILE_NAME = ".sheet_info"
MARKS_FILE_NAME = "points.json"
PRINT_INDENT_WIDTH = 4
SHARE_ARCHIVE_PREFIX = "share_archive"
COMBINED_DIR_NAME = "feedback_combined"

# Might be necessary to make colored output work on Windows.
os.system("")


# ============================= Utility Functions ==============================


# Printing ---------------------------------------------------------------------
def print_indented(text: str) -> None:
    """
    Print a message that belongs to a previously printed info or warning
    message.
    """
    print(" " * PRINT_INDENT_WIDTH + text)


def print_info(text: str, bare: bool = False) -> None:
    """
    Print an info message with or without leading '[Info]'.
    """
    if args.quiet:
        return
    prefix: str = "" if bare else "\033[0;34m[Info]\033[0m "
    print(prefix + text)


def print_warning(text: str) -> None:
    """
    Print a warning message.
    """
    if args.quiet:
        return
    print("\033[0;33m[Warning]\033[0m " + text)


def query_yes_no(text: str, default: bool = True) -> bool:
    """
    Ask the user a yes/no question and return answer.
    """
    valid = {"yes": True, "y": True, "ye": True, "no": False, "n": False}
    options = "[Y/n]" if default else "[y/N]"
    print("\033[0;35m[Query]\033[0m " + text + f" {options}")
    choice = input().lower()
    if choice == "":
        return default
    elif choice in valid:
        return valid[choice]
    else:
        print_warning(
            f"Invalid choice '{choice}'. Please respond with 'yes' or 'no'."
        )
        return query_yes_no(text, default)


# Errors -----------------------------------------------------------------------
def throw_error(text: str) -> None:
    """
    Print an error message and exit.
    """
    print("\033[0;31m[Error]\033[0m " + text)
    sys.exit(1)


def abort(text: str) -> None:
    """
    Gracefully terminate the script without success because something went
    wrong in an expected way.
    """
    if not args.quiet:
        print_warning(text)
    print_warning(f"Aborting the subcommand '{args.sub_command}'.")
    sys.exit(1)


# String things ----------------------------------------------------------------
def team_to_string(team: Team) -> str:
    """
    Concatenate the last names of students to get a pretty-ish string
    representation of teams.
    """
    return "_".join(sorted([student[1].replace(" ", "-") for student in team]))


def get_adam_sheet_name_string() -> str:
    """
    Turn the sheet name given by ADAM into a string usable for file names.
    """
    return args.adam_sheet_name.replace(" ", "_").lower()


def get_feedback_file_name() -> str:
    file_name = FEEDBACK_FILE_PREFIX + get_adam_sheet_name_string() + "_"
    if args.marking_mode == "exercise":
        # TODO: I'm not sure why I added the team_id here. Add it back in if
        # it's necessary, remove these lines otherwise.
        # team_id = team_dir.name.split("_")[0]
        # prefix = team_id + "_" + prefix
        file_name += args.tutor_name + "_"
        file_name += "_".join([f"ex{exercise}" for exercise in args.exercises])
    elif args.marking_mode == "random":
        file_name += args.tutor_name
    elif args.marking_mode == "static":
        # Remove trailing underscore.
        file_name = file_name[:-1]
    else:
        throw_error(f"Unsupported marking mode {args.marking_mode}!")
    return file_name


# Miscellaneous ----------------------------------------------------------------
def is_email(email: str) -> bool:
    """
    Check if a string more or less matches the format of an email address.
    """
    return type(email) is str and bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))


def is_macos_path(path: str) -> bool:
    """
    Check if the given path is non-essential file created by MacOS.
    """
    return "__MACOSX" in path or ".DS_Store" in path


def filtered_extract(zip_file: ZipFile, dest: pathlib.Path) -> None:
    """
    Extract all files except for MACOS helper files.
    """
    zip_content = zip_file.namelist()
    for file_path in zip_content:
        if is_macos_path(file_path):
            continue
        zip_file.extract(file_path, dest)


def move_content_and_delete(src: pathlib.Path, dst: pathlib.Path) -> None:
    """
    Move all content of src directory to dest directory.
    This does not complain if the dst directory already exists.
    """
    assert src.is_dir() and dst.is_dir()
    shutil.copytree(src, dst, dirs_exist_ok=True)
    shutil.rmtree(src)


def verify_sheet_root_dir() -> None:
    """
    Ensure that the given sheet root directory is valid. Needed for multiple
    sub-commands such as 'collect', or 'send'.
    """
    if not args.sheet_root_dir.is_dir():
        throw_error("The given sheet directory is not valid!")


def get_all_team_dirs() -> Iterator[pathlib.Path]:
    """
    Return all team directories within the sheet root directory. It is assumed
    that all team directory names start with some digits, followed by an
    underscore, followed by more characters. In particular this excludes
    other directories that may be created in the sheet root directory, such as
    one containing combined feedback.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if team_dir.is_dir() and re.match(r"[0-9]+_.+", team_dir.name):
            yield team_dir


def get_relevant_team_dirs() -> Iterator[pathlib.Path]:
    """
    Return the team directories of the teams whose submission has to be
    corrected by the tutor running the script.
    """
    for team_dir in get_all_team_dirs():
        if not DO_NOT_MARK_PREFIX in team_dir.name:
            yield team_dir


def get_share_archive_files() -> Iterator[pathlib.Path]:
    """
    Return all share archive files under the current sheet root dir.
    """
    for share_archive_file in args.sheet_root_dir.glob(
        SHARE_ARCHIVE_PREFIX + "*.zip"
    ):
        yield share_archive_file


def get_collected_feedback_file(team_dir: pathlib.Path) -> pathlib.Path:
    """
    Given a team directory, return the collected feedback file. This can be
    either a single pdf file, or a single zip archive. Throw an error if neither
    exists.
    """
    collected_feedback_dir = team_dir / FEEDBACK_COLLECTED_DIR_NAME
    assert collected_feedback_dir.is_dir()
    collected_feedback_files = list(collected_feedback_dir.iterdir())
    assert (
        len(collected_feedback_files) == 1
        and collected_feedback_files[0].is_file()
        and collected_feedback_files[0].suffix in [".pdf", ".zip"]
    )
    return collected_feedback_files[0]


def load_sheet_info() -> None:
    """
    Load the information stored in the sheet info file into the args object.
    """
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "r", encoding="utf-8"
    ) as sheet_info_file:
        sheet_info = json.load(sheet_info_file)
    for key, value in sheet_info.items():
        add_to_args(key, value)


# ============================== Send Sub-Command ==============================


def add_attachment(mail: EmailMessage, path: pathlib.Path) -> None:
    """
    Add a file as attachment to an email.
    This is copied from Patrick's/Silvan's script, not entirely sure what's
    going on here.
    """
    assert path.exists()
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(path, "rb") as fp:
        mail.add_attachment(
            fp.read(),
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(path),
        )


def construct_email(
    receivers: list[str],
    cc: list[str],
    subject: str,
    content: str,
    sender: str,
    attachment: pathlib.Path,
) -> EmailMessage:
    """
    Construct an email object.
    """
    assert isinstance(receivers, list)
    assert subject and content and sender and attachment
    mail = EmailMessage()
    mail.set_content(content)
    mail["Subject"] = subject
    mail["From"] = sender
    mail["To"] = ", ".join(receivers)
    mail["Cc"] = ", ".join(cc)
    add_attachment(mail, attachment)
    return mail


def send_messages(emails: list[EmailMessage]) -> None:
    with smtplib.SMTP(args.smtp_url, args.smtp_port) as smtp:
        smtp.starttls(context=ssl.create_default_context())
        password = getpass("Email password: ")
        print_info(
            f"Authentication: {smtp.login(args.smtp_user, password)}", True
        )
        print_info("Sending emails...")
        for email in emails:
            print_info(f"... to {email['To']}", True)
            smtp.send_message(email)
    print_info("Done sending emails.")


def get_email_subject() -> str:
    """
    Builds the email subject.
    """
    return f"Feedback {args.adam_sheet_name} | {args.lecture_title}"


def get_email_greeting(name_list: list[str]) -> str:
    """
    Builds the first line of the email.
    """
    # Only keep one name per entry, "Hans Jakob" becomes "Hans".
    name_list = [name.split(" ")[0] for name in name_list]
    name_list.sort()
    assert len(name_list) > 0
    if len(name_list) == 1:
        names = name_list[0]
    elif len(name_list) == 2:
        names = name_list[0] + " and " + name_list[1]
    else:
        names = ", ".join(name_list[:-1]) + ", and " + name_list[-1]
    return "Dear " + names + ","


def get_email_content(name_list: list[str]) -> str:
    """
    Builds the body of the email.
    """
    return textwrap.dedent(
        f"""
    {get_email_greeting(name_list)}

    Please find feedback on your submission for {args.adam_sheet_name} in the attachment.
    If you have any questions, you can contact us in the exercise session or by replying to this email (reply to all).

    Best,
    Your Tutors
    """
    )[
        1:
    ]  # Removes the leading newline.


def send() -> None:
    """
    After the collection step finished successfully, send the feedback to the
    students via email. This currently only works if the tutor's email account
    is whitelisted for the smpt-ext.unibas.ch server.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()
    if args.marking_mode == "exercise":
        throw_error(
            "Sending for marking mode 'exercise' is not implemented because "
            "collection is not yet figured out."
        )
    # Send emails.
    emails: list[EmailMessage] = []
    for team_dir in get_relevant_team_dirs():
        team = args.team_dir_to_team[team_dir.name]
        team_first_names, _, team_emails = zip(*team)
        email = construct_email(
            list(team_emails),
            args.feedback_email_cc,
            get_email_subject(),
            get_email_content(team_first_names),
            args.tutor_email,
            get_collected_feedback_file(team_dir),
        )
        emails.append(email)
    print_info(f"Ready to send {len(emails)} email(s).")
    send_messages(emails)


# ============================ Collect Sub-Command =============================


def validate_marks_json() -> None:
    """
    Verify that all necessary marks are present in the MARK_FILE_NAME file and
    adhere to the granularity defined in the config file.
    """
    marks_json_file = args.sheet_root_dir / MARKS_FILE_NAME
    if not marks_json_file.is_file():
        throw_error(
            f"Missing points file in directory '{args.sheet_root_dir}'!"
        )
    with open(marks_json_file, "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)
    relevant_teams = [
        relevant_team.name for relevant_team in get_relevant_team_dirs()
    ]
    marked_teams = list(marks.keys())
    if sorted(relevant_teams) != sorted(marked_teams):
        throw_error(
            "There is no 1-to-1 mapping between team directories "
            "that need to be marked and entries in the "
            f"'{MARKS_FILE_NAME}' "
            "file! Make sure that it contains exactly one entry for every team "
            "directory that needs to be marked, and that directory name and "
            "key are the same."
        )
    if args.points_per == "exercise":
        marks_list = [
            mark
            for team_marks in marks.values()
            for mark in team_marks.values()
        ]
    else:
        marks_list = marks.values()
    if not all(marks_list):
        throw_error(
            f"There are missing points in the '{MARKS_FILE_NAME}' file!"
        )
    if not all(
        (float(mark) / args.min_point_unit).is_integer() for mark in marks_list
    ):
        throw_error(
            f"'{MARKS_FILE_NAME}' contains marks that are more fine-grained "
            "than allowed! You may only award points in "
            f"'{args.min_point_unit}' increments."
        )


def collect_feedback_files(team_dir: pathlib.Path) -> None:
    """
    Take the contents of a {team_dir}/feedback directory and collect the files
    that actually contain feedback (e.g., no .xopp files). If there are
    multiple, add them to a zip archive and save it to
    {team_dir}/feedback_collected. If there is only a single pdf, copy it to
    {team_dir}/feedback_collected.
    """
    feedback_dir = team_dir / FEEDBACK_DIR_NAME
    collected_feedback_dir = team_dir / FEEDBACK_COLLECTED_DIR_NAME
    collected_feedback_zip_name = get_feedback_file_name() + ".zip"
    # Error handling.
    if not feedback_dir.exists():
        throw_error(f"Missing feedback directory for team {team_dir.name}!")
    content = list(feedback_dir.iterdir())
    if any(".todo" in file_or_dir.name for file_or_dir in content):
        throw_error(
            f"Feedback for {team_dir.name} contains placeholder TODO file!"
        )
    # The directory for collected feedback should exist and be empty. Either it
    # was created new, or the user chose to overwrite and previously existing
    # directories have been removed and replaced by empty ones.
    assert collected_feedback_dir.is_dir() and not any(
        collected_feedback_dir.iterdir()
    )
    # Create list of feedback files. Those are all files in the feedback
    # directory which do not have an ignored suffix.
    feedback_files = [
        file
        for file in feedback_dir.rglob("*")
        if file.is_file() and not file.suffix in args.ignore_feedback_suffix
    ]
    if len(feedback_files) <= 0:
        throw_error(
            f"Feedback archive for team {team_dir.name} is empty! "
            "Did you forget the '-x' flag to export .xopp files?"
        )
    # If there is exactly one pdf in the feedback directory, we do not need to
    # create a zip archive.
    if len(feedback_files) == 1 and feedback_files[0].suffix == ".pdf":
        shutil.copy(feedback_files[0], collected_feedback_dir)
        return
    # Otherwise, zip up feedback files.
    feedback_contains_pdf = False
    with ZipFile(
        collected_feedback_dir / collected_feedback_zip_name, "w"
    ) as zip_file:
        for file_to_zip in feedback_files:
            if file_to_zip.suffix == ".pdf":
                feedback_contains_pdf = True
            zip_file.write(
                file_to_zip, arcname=file_to_zip.relative_to(feedback_dir)
            )
    if not feedback_contains_pdf:
        print_warning(f"The feedback for {team_dir.name} contains no PDF file!")


def delete_collected_feedback_directories() -> None:
    """
    Removes existing collected feedback directories. Does not care about
    non-existing ones.
    """
    for team_dir in get_relevant_team_dirs():
        collected_feedback_dir = team_dir / FEEDBACK_COLLECTED_DIR_NAME
        shutil.rmtree(collected_feedback_dir, ignore_errors=True)


def create_collected_feedback_directories() -> None:
    """
    Create an empty directory in each relevant team directory. The collected
    feedback will be saved to these directories.
    """
    for team_dir in get_relevant_team_dirs():
        collected_feedback_dir = team_dir / FEEDBACK_COLLECTED_DIR_NAME
        assert not collected_feedback_dir.is_dir() or not any(
            collected_feedback_dir.iterdir()
        )
        collected_feedback_dir.mkdir(exist_ok=True)


def export_xopp_files() -> None:
    """
    Exports all xopp feedback files.
    """
    print_info("Exporting .xopp files...")
    for team_dir in get_relevant_team_dirs():
        feedback_dir = team_dir / FEEDBACK_DIR_NAME
        xopp_files = [
            file for file in feedback_dir.rglob("*") if file.suffix == ".xopp"
        ]
        for xopp_file in xopp_files:
            dest = xopp_file.with_suffix(".pdf")
            subprocess.run(["xournalpp", "-p", dest, xopp_file])
    print_info("Done exporting .xopp files.")


def print_marks() -> None:
    """
    Prints the marks so that they can be easily copy-pasted to the file where
    marks are collected.
    """
    # Read marks file.
    marks_json_file = args.sheet_root_dir / MARKS_FILE_NAME
    # Don't check whether the marks file exists because `validate_marks_json()`
    # would have already complained.
    with open(marks_json_file, "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)

    # Print marks.
    print_info("Start of copy-paste marks...")
    # We want all teams printed, not just the marked ones.
    for team_to_print in args.teams:
        for team_dir, team in args.team_dir_to_team.items():
            # Every team should only be the value of at most one entry in
            # `team_dir_to_team`.
            if team == team_to_print:
                for student in team:
                    full_name = f"{student[0]} {student[1]}"
                    output_str = f"{full_name:>35};"
                    if args.points_per == "exercise":
                        # The value `marks` assigns to the team_dir key is a
                        # dict with (exercise name, mark) pairs.
                        team_marks = marks.get(team_dir, {"null": ""})
                        _, exercise_marks = zip(*team_marks.items())
                        for mark in exercise_marks:
                            output_str += f"{mark:>3};"
                    else:
                        sheet_mark = marks.get(team_dir, "")
                        output_str += f"{sheet_mark:>3}"
                    print_info(output_str, True)
    print_info("End of copy-paste marks.")


def create_share_archive(overwrite: Optional[bool]) -> None:
    """
    In case the marking mode is exercise, the final feedback the teams get is
    made up of multiple sets of PDFs (and potentially other files) made by
    multiple tutors. This function stores all feedback made by a single tutor in
    a ZIP file which she/he can then share with the tutor that will send the
    combined feedback. Refer to the `combine` subcommand to see how to process
    the resulting "share archives".
    """
    # This function in only used when the correction mode is 'exercise'.
    # Consequently, exercises must be provided when running 'init', which should
    # be written to .sheet_info and then read in by a load_sheet_info().
    assert args.exercises
    # Build share archive file name.
    share_archive_file_name = (
        SHARE_ARCHIVE_PREFIX
        + f"_{get_adam_sheet_name_string()}_"
        + "_".join([f"ex{num}" for num in args.exercises])
        + ".zip"
    )
    share_archive_file = args.sheet_root_dir / share_archive_file_name
    if share_archive_file.is_file():
        # If the user has already chosen to overwrite when considering feedback
        # zips, then overwrite here too. Otherwise ask here.
        # We should not be here if the user chose 'No' (making overwrite False)
        # before, so we catch this case.
        assert overwrite != False
        if overwrite == None:
            overwrite = query_yes_no(
                (
                    "There already exists a share archive. Do you want to"
                    " overwrite it?"
                ),
                default=False,
            )
        if overwrite:
            share_archive_file.unlink(missing_ok=True)
        else:
            abort(f"Could not write share archive.")
    # Take all feedback.zip files and add them to the share archive. The file
    # structure should be similar to the following. In particular, collected
    # feedback that consists of only a single pdf should be zipped to achieve
    # the structure below, whereas collected feedback that is already an archive
    # simply needs to be written under a new name.
    # share_archive_sample_sheet_ex1_ex2.zip
    # ├── 12345_Muster-Meier-Mueller.zip
    # │   └── feedback_tutor1_ex1.pdf
    # ├── 12345_Muster-Meier-Mueller.zip
    # │   ├── feedback_tutor1_ex1.pdf
    # │   └── feedback_tutor1_ex1_code_submission.cc
    # └── ...
    with ZipFile(share_archive_file, "w") as zip_file:
        # The relevant team directories should always be *all* team directories
        # here, because we only need share archives for the 'exercise' marking
        # mode.
        for team_dir in get_relevant_team_dirs():
            collected_feedback_file = get_collected_feedback_file(team_dir)
            sub_zip_name = f"{team_dir.name}.zip"
            if collected_feedback_file.suffix == ".pdf":
                # Create a temporary zip file in the collected feedback
                # directory and add the single pdf.
                temp_zip_file = (
                    team_dir / FEEDBACK_COLLECTED_DIR_NAME / "temp_zip.zip"
                )
                with ZipFile(temp_zip_file, "w") as temp_zip:
                    temp_zip.write(
                        collected_feedback_file,
                        arcname=collected_feedback_file.name,
                    )
                # Add the temporary zip file to the share archive.
                zip_file.write(temp_zip_file, arcname=sub_zip_name)
                # Remove the temporary zip file.
                temp_zip_file.unlink()
            elif collected_feedback_file.suffix == ".zip":
                zip_file.write(collected_feedback_file, arcname=sub_zip_name)
            else:
                throw_error(
                    "Collected feedback must be either a single pdf file or a"
                    " single zip archive."
                )


def collect() -> None:
    """
    After marking is done, add feedback files to archives and print marks to be
    copy-pasted to shared point spreadsheet.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()
    # Collect feedback.

    # Check if there is a collected feedback directory with files inside
    # already.
    collected_feedback_exists = any(
        (team_dir / FEEDBACK_COLLECTED_DIR_NAME).is_dir()
        and any((team_dir / FEEDBACK_COLLECTED_DIR_NAME).iterdir())
        for team_dir in get_relevant_team_dirs()
    )
    # Ask the user whether collected feedback should be overwritten in case it
    # exists already.
    overwrite = None
    if collected_feedback_exists:
        overwrite = query_yes_no(
            (
                "There already exists collected feedback. Do you want to"
                " overwrite it?"
            ),
            default=False,
        )
        if overwrite:
            delete_collected_feedback_directories()
        else:
            abort(f"Could not write collected feedback archives.")
    if args.xopp:
        export_xopp_files()
    create_collected_feedback_directories()
    for team_dir in get_relevant_team_dirs():
        collect_feedback_files(team_dir)
    if args.marking_mode == "exercise":
        create_share_archive(overwrite)
    if args.use_marks_file:
        validate_marks_json()
        print_marks()


# ============================ Combine Sub-Command =============================


def combine() -> None:
    """
    Combine multiple share archives so that in the end we have one zip archive
    per team containing all feedback for that team.
    TODO: At the moment we end up with one directory per team containing all
    feedback, but we need to zip this up again.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()

    share_archive_files = get_share_archive_files()
    instructions = (
        "Run `collect` to generate the share archive for your own feedback and"
        " save the share archives you received from the other tutors under"
        f" {args.sheet_root_dir}."
    )
    if len(list(share_archive_files)) == 0:
        abort(
            f"No share archives exist in {args.sheet_root_dir}. " + instructions
        )
    if len(list(share_archive_files)) == 1:
        print_warning(
            "Only a single share archive is being combined. " + instructions
        )

    # Create directory to store combined feedback in.
    combined_dir = args.sheet_root_dir / COMBINED_DIR_NAME
    if combined_dir.exists() and combined_dir.is_dir():
        overwrite = query_yes_no(
            (
                f"The directory {combined_dir} exists already. Do you want to"
                " overwrite it?"
            ),
            default=False,
        )
        if overwrite:
            shutil.rmtree(combined_dir)
        else:
            abort(f"Could not write to '{combined_dir}'.")
    combined_dir.mkdir()

    # Create subdirectories for teams.
    for team_dir in get_relevant_team_dirs():
        combined_team_dir = combined_dir / team_dir.name
        combined_team_dir.mkdir()

    teams_all = [team_dir.name for team_dir in get_relevant_team_dirs()]
    # Extract feedback from share archives into the combined directory.
    for share_archive_file in args.sheet_root_dir.glob(
        SHARE_ARCHIVE_PREFIX + "*.zip"
    ):
        with ZipFile(share_archive_file, mode="r") as share_archive:
            # Check if this share archive is missing feedback archives for any
            # teams.
            teams_present = [
                pathlib.Path(team).stem for team in share_archive.namelist()
            ]
            teams_not_present = list(set(teams_all) - set(teams_present))
            for team_not_present in teams_not_present:
                print_warning(
                    f"The shared archive {share_archive_file} contains no"
                    f" feedback for team {team_not_present}."
                )
            for team in teams_present:
                # Extract feedback file from share_archive.
                share_archive.extract(team + ".zip", combined_dir / team)
                feedback_file_names = list((combined_dir / team).glob("*"))
                assert (
                    len(feedback_file_names) == 1
                    and feedback_file_names[0].is_file()
                )
                feedback_file = feedback_file_names[0]
                # If the feedback is not an archive but a single pdf, move on to
                # the next team.
                if feedback_file.suffix == ".pdf":
                    continue
                assert feedback_file.suffix == ".zip"
                # Otherwise, extract feedback from feedback archive.
                print(f"{feedback_file=}")
                with ZipFile(feedback_file, mode="r") as feedback_archive:
                    feedback_archive.extractall(path=combined_dir / team)
                # Remove feedback archive.
                feedback_file.unlink()
                # TODO: Test if this works when there are actually multiple different share archives.


# ============================== Init Sub-Command ==============================


def extract_adam_zip() -> tuple[pathlib.Path, str]:
    """
    Unzips the given ADAM zip file and renames the directory to *target* if one
    is given. This is done stupidly right now, it would be better to extract to
    a temporary folder and then move once to the right location.
    """
    if args.adam_zip_path.is_file():
        # Unzip to the directory within the zip file.
        # Should be the name of the exercise sheet, for example "Exercise Sheet 2".
        with ZipFile(args.adam_zip_path, mode="r") as zip_file:
            zip_content = zip_file.namelist()
            sheet_root_dir = pathlib.Path(zip_content[0])
            # TODO: Do this with tempfile.
            if sheet_root_dir.exists():
                throw_error(
                    "Extraction failed because the extraction path "
                    f"'{sheet_root_dir}' exists already!"
                )
            filtered_extract(zip_file, pathlib.Path("."))
    else:
        # Assume the directory is an extracted ADAM zip.
        sheet_root_dir = args.adam_zip_path
    # Flatten intermediate directory.
    sub_dirs = list(sheet_root_dir.glob("*"))
    if len(sub_dirs) != 1:
        throw_error(
            "The ADAM zip file contains an unexpected number of"
            f" directories/files ({len(sub_dirs)}), expected exactly 1"
            " subdirectory named either 'Abgaben' or 'Submissions'."
        )
    if sub_dirs[0].name not in ["Abgaben", "Submissions"]:
        print_warning(
            "It looks like the format of the zip file created by ADAM has"
            " changed. This may require adaptions to this script, but I will"
            " try anyway."
        )
    move_content_and_delete(sub_dirs[0], sheet_root_dir)
    # Store ADAM exercise sheet name to use as random seed.
    adam_sheet_name = sheet_root_dir.name
    if args.target:
        target_path = pathlib.Path(args.target)
        if target_path.exists():
            throw_error(
                f"Extraction failed because the path '{target_path}' "
                "exists already!"
            )
        sheet_root_dir = shutil.move(sheet_root_dir, target_path)
    return sheet_root_dir, adam_sheet_name


def get_adam_id_to_team_dict() -> dict[str, Team]:
    """
    ADAM assigns every team a new ID with every exercise sheet. This dict maps
    from that ID to the team represented by a list of [name, email] pairs. At
    the same time, the "Team " prefix is removed from directory names.
    """
    adam_id_to_team = {}
    for team_dir in args.sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        team_id = team_dir.name.split(" ")[1]
        submission_dir = list(team_dir.iterdir())[0]
        submission_email = submission_dir.name.split("_")[-2]
        teams = [
            team
            for team in args.teams
            if any(submission_email in student for student in team)
        ]
        if len(teams) == 0:
            throw_error(
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
                print_warning(
                    f"There are multiple submissions for team '{teams[0]}'"
                    f" under separate ADAM IDs ({existing_id} and {team_id})!"
                    " This probably means that multiple members of a team"
                    " submitted solutions without forming a team on ADAM. You"
                    " will have to combine the submissions manually."
                )
        adam_id_to_team.update({team_id: teams[0]})
        team_dir = pathlib.Path(
            shutil.move(team_dir, team_dir.with_name(team_id))
        )
    return adam_id_to_team


def mark_irrelevant_team_dirs() -> None:
    """
    Indicate which team directories do not have to be marked by adding the
    `DO_NOT_MARK_PREFIX` to their directory name.
    """
    relevant_teams = get_relevant_teams()
    for team_dir_name, team in args.team_dir_to_team.items():
        if not team in relevant_teams:
            team_dir = args.sheet_root_dir / team_dir_name
            shutil.move(
                team_dir, team_dir.with_name(DO_NOT_MARK_PREFIX + team_dir_name)
            )


def get_relevant_teams() -> list[Team]:
    """
    Get a list of teams that the tutor specified in the config has to mark.
                                     !DANGER!
    Calling this function multiple times in marking_mode == 'random' will return
    different results. We only run it once, rename the directories using the
    `DO_NOT_MARK_PREFIX`, and thereafter only access relevant teams via
    `get_relevant_team_dirs()`.
    """
    if args.marking_mode == "static":
        return args.classes[args.tutor_name]
    elif args.marking_mode == "random":
        # Here not all teams are assigned to a tutor, but only those that
        # submitted something. This is to ensure that submissions can be
        # distributed fairly among tutors.
        num_tutors = len(args.tutor_list)
        seed = int(
            hashlib.sha256(args.adam_sheet_name.encode("utf-8")).hexdigest(), 16
        )
        shuffled_teams = [team for _, team in args.team_dir_to_team.items()]
        random.Random(seed).shuffle(shuffled_teams)
        chunks = [shuffled_teams[i::num_tutors] for i in range(num_tutors)]
        assert len(chunks) == num_tutors
        assert all(
            abs(len(this) - len(that)) <= 1
            for this in chunks
            for that in chunks
        )
        shuffled_tutors = args.tutor_list.copy()
        random.Random(seed).shuffle(shuffled_tutors)
        return chunks[shuffled_tutors.index(args.tutor_name)]
    elif args.marking_mode == "exercise":
        return args.teams
    else:
        throw_error(f"Unsupported marking mode {args.marking_mode}!")
        return []


def rename_team_dirs(adam_id_to_team: dict[str, Team]) -> None:
    """
    The team directories are renamed to: team_id_LastName1-LastName2
    The team ID can be helpful to identify a team on the ADAM web interface.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        team_id = team_dir.name
        team = adam_id_to_team[team_id]
        dir_name = team_id + "_" + team_to_string(team)
        team_dir = pathlib.Path(
            shutil.move(team_dir, team_dir.with_name(dir_name))
        )


def flatten_team_dirs() -> None:
    """
    There can be multiple directories within a "Team 00000" directory. This
    happens when multiple members of the team upload solutions. Sometimes, only
    one directory contains submitted files, in this case we remove the empty
    ones silently. In case multiple submissions exist, we put the files within
    them next to each other and print a warning.
    """
    for team_dir in get_all_team_dirs():
        # Remove empty subdirectories.
        for team_submission_dir in team_dir.iterdir():
            if len(list(team_submission_dir.iterdir())) == 0:
                team_submission_dir.rmdir()
        # Store the list of team submission directories in variable, because the
        # generator may include subdirectories of team submission directories
        # that have already been flattened.
        team_submission_dirs = list(team_dir.iterdir())
        if len(team_submission_dirs) > 1:
            print_warning(
                f"There are multiple submissions for group '{team_dir.name}'!"
            )
        if len(team_submission_dirs) < 1:
            print_warning(
                f"The submission of group '{team_dir.name}' is empty!"
            )
        for team_submission_dir in team_submission_dirs:
            move_content_and_delete(team_submission_dir, team_dir)


def unzip_internal_zips() -> None:
    """
    If multiple files are uploaded to ADAM, the submission becomes a single zip
    file. Here we extract this zip. I'm not sure if nested zip files are also
    extracted. Additionally we flatten the directory by one level if the zip
    contains only a single directory. Doing so recursively would be nicer.
    """
    for team_dir in get_all_team_dirs():
        if not team_dir.is_dir():
            continue
        for zip_file in team_dir.glob("**/*.zip"):
            with ZipFile(zip_file, mode="r") as zf:
                filtered_extract(zf, zip_file.parent)
            os.remove(zip_file)
        sub_dirs = list(team_dir.iterdir())
        if len(sub_dirs) == 1 and sub_dirs[0].is_dir():
            move_content_and_delete(sub_dirs[0], team_dir)


def create_marks_file() -> None:
    """
    Write a json file to add the marks for all relevant teams and exercises.
    """
    exercise_dict: Union[str, dict[str, str]] = ""
    if args.points_per == "exercise":
        if args.marking_mode == "static" or args.marking_mode == "random":
            exercise_dict = {
                f"exercise_{i}": "" for i in range(1, args.num_exercises + 1)
            }
        elif args.marking_mode == "exercise":
            exercise_dict = {f"exercise_{i}": "" for i in args.exercises}
    else:
        exercise_dict = ""

    marks_dict = {}
    for team_dir in sorted(list(get_relevant_team_dirs())):
        marks_dict.update({team_dir.name: exercise_dict})

    with open(
        args.sheet_root_dir / MARKS_FILE_NAME, "w", encoding="utf-8"
    ) as marks_json:
        json.dump(marks_dict, marks_json, indent=4, ensure_ascii=False)


def create_feedback_directories() -> None:
    """
    Create a directory for every team that should be corrected by the tutor
    specified in the config. A copy of every non-PDF file is prefixed and placed
    in the feedback folder. The idea is that feedback on code files or similar
    can be added to these copies directly and files without feedback can simply
    be deleted. PDFs are not copied, instead a placeholder file with the correct
    name is created so that it can be overwritten by a real PDF file containing
    the feedback.
    """
    for team_dir in get_relevant_team_dirs():
        feedback_dir = team_dir / FEEDBACK_DIR_NAME
        feedback_dir.mkdir()

        feedback_file_name = get_feedback_file_name()
        dummy_pdf_name = feedback_file_name + ".pdf.todo"
        pathlib.Path(feedback_dir / dummy_pdf_name).touch(exist_ok=True)

        # Copy non-pdf submission files into feedback directory with added
        # prefix.
        for submission_file in team_dir.glob("*"):
            if submission_file.is_dir() or submission_file.suffix == ".pdf":
                continue
            this_feedback_file_name = (
                feedback_file_name + "_" + submission_file.name
            )
            shutil.copy(submission_file, feedback_dir / this_feedback_file_name)


def generate_xopp_files() -> None:
    """
    Generate xopp files in the feedback directories that point to the single pdf
    in the submission directory and skip if multiple PDF files exist.
    """
    from PyPDF2 import PdfReader

    def write_to_file(f, string):
        f.write(textwrap.dedent(string))

    print_info("Generating .xopp files...")
    for team_dir in get_relevant_team_dirs():
        pdf_paths = list(team_dir.glob("*.pdf"))
        if len(pdf_paths) != 1:
            print_warning(
                f"Skipping .xopp file generation for {team_dir.name}: No or"
                " multiple PDF files."
            )
            continue
        pdf_path = pdf_paths[0]
        feedback_dir = team_dir / FEEDBACK_DIR_NAME
        todo_paths = list(feedback_dir.glob("*.pdf.todo"))
        assert len(todo_paths) == 1
        todo_path = todo_paths[0]
        pages = PdfReader(pdf_path).pages
        # Strips the ".todo" and replaces ".pdf" by ".xopp".
        xopp_path = todo_path.with_suffix("").with_suffix(".xopp")
        if xopp_path.is_file():
            print_warning(
                f"Skipping .xopp file generation for {team_dir.name}: xopp file"
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
                    </page>""",
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
    print_info("Done generating .xopp files.")


def create_sheet_info_file(adam_id_to_team: dict[str, Team]) -> None:
    """
    Write information generated during the execution of the 'init' command in a
    sheet info file. In particular a mapping from team directory names to teams
    and the name of the exercise sheet as given by ADAM. The latter is used as
    the seed to make random assignment of submissions to tutors consistent
    between tutors, but still vary from sheet to sheet. Later commands (e.g.
    'collect', or 'send') are meant to load the information stored in this file
    into the 'args' object and access it that way.
    """
    info_dict: dict[str, Union[str, dict[str, Team]]] = {}
    # Build the dict from team directory names to teams.
    team_dir_to_team = {}
    for team_dir in get_all_team_dirs():
        if team_dir.is_file():
            continue
        # Get ADAM ID from directory name.
        adam_id_match = re.search(r"\d+", team_dir.name)
        assert adam_id_match
        adam_id = adam_id_match.group()
        team = adam_id_to_team[adam_id]
        team_dir_to_team.update({team_dir.name: team})
    info_dict.update({"team_dir_to_team": team_dir_to_team})
    info_dict.update({"adam_sheet_name": args.adam_sheet_name})
    if args.marking_mode == "exercise":
        info_dict.update({"exercises": args.exercises})
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "w", encoding="utf-8"
    ) as sheet_info_file:
        # Sorting the keys here is essential because the order of teams here
        # will influence the assignment returned by `get_relevant_teams()` in
        # case args.marking_mode == "random".
        json.dump(
            info_dict,
            sheet_info_file,
            indent=4,
            ensure_ascii=False,
            sort_keys=True,
        )
    # Immediately load the info back into args.
    load_sheet_info()


def print_missing_submissions(adam_id_to_team: dict[str, Team]) -> None:
    """
    Print all teams that are listed in the config file, but whose submission is
    not present in the zip downloaded from ADAM.
    """
    missing_teams = [
        team for team in args.teams if not team in adam_id_to_team.values()
    ]
    if missing_teams:
        print_info("There are no submissions for the following team(s):")
        for missing_team in missing_teams:
            print_indented(f"{team_to_string(missing_team)}")


def init() -> None:
    """
    Prepares the directory structure holding the submissions.
    """
    # Catch wrong combinations of marking_mode/points_per/-n/-e.
    # Not possible eariler because marking_mode and points_per are given by the
    # config file.
    if args.points_per == "exercise":
        if args.marking_mode == "exercise" and not args.exercises:
            throw_error(
                "You must provide a list of exercise numbers to be marked with "
                "the '-e' flag, for example '-e 1 3 4'."
            )
        if (
            args.marking_mode == "random" or args.marking_mode == "static"
        ) and not args.num_exercises:
            throw_error(
                "You must provide the number of exercises in the sheet with "
                "the '-n' flag, for example '-n 5'."
            )
    else:  # points per sheet
        if args.marking_mode == "exercise":
            throw_error(
                "Points must be given per exercise if marking is done per "
                "exercise. Set the value of 'poins_per' to 'exercise' or "
                "change the marking mode."
            )
        if args.num_exercises or args.exercises:
            print_warning(
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
    sheet_root_dir, adam_sheet_name = extract_adam_zip()
    add_to_args("sheet_root_dir", sheet_root_dir)
    add_to_args("adam_sheet_name", adam_sheet_name)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── Team 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    adam_id_to_team = get_adam_id_to_team_dict()
    print_missing_submissions(adam_id_to_team)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    rename_team_dirs(adam_id_to_team)

    # From here on, get_all_team_dirs() should work.

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    flatten_team_dirs()

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   └── submission.pdf or submission.zip
    unzip_internal_zips()

    # From here on, we need information about relevant teams.
    # The function `get_relevant_teams()` depends on the sheet info file
    # (because `adam_sheet_name` from .sheet_info seeds the random assignment of
    # submissions to tutors).
    # That's why we create the sheet info file first...
    create_sheet_info_file(adam_id_to_team)
    # then rename the irrelevant team directories...
    mark_irrelevant_team_dirs()
    # and finally recreate the sheet info file to reflect the final team
    # directory names.
    create_sheet_info_file(adam_id_to_team)

    if args.use_marks_file:
        create_marks_file()

    create_feedback_directories()

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── feedback
    # .   │   ├── feedback.pdf.todo
    # .   │   └── feedback_copy-of-submitted-file.cc
    # .   └── submission.pdf or submission files
    # ├── .sheet_info
    # └── points.json
    if args.xopp:
        generate_xopp_files()


# ============================= Config Processing ==============================


def validate_teams(teams: list[Team]) -> None:
    """
    Verify that teams and its (first_name, last_name, email) triples are well
    formed. Also sort teams and their students to make iterating over them
    predictable, independent of their order in config.json.
    """
    assert type(teams) is list
    all_students: list[tuple[str, str]] = []
    all_emails: list[str] = []
    for team in teams:
        team.sort()
        assert len(team) <= args.max_team_size
        first_names, last_names, emails = list(zip(*team))
        assert all(type(first_name) is str for first_name in first_names)
        assert all(type(last_name) is str for last_name in last_names)
        assert all(is_email(email) for email in emails)
        all_students += list(zip(first_names, last_names))
        all_emails += emails
    if len(all_students) != len(set(all_students)):
        throw_error("There are duplicate students in the config file!")
    if len(all_emails) != len(set(all_emails)):
        throw_error("There are duplicate student emails in the config file!")
    teams.sort()


def process_static_config(data: dict[str, Any]) -> None:
    """
    Extracts and checks the config values necessary for the static correction
    marking mode.
    """
    classes = data["teams"]
    assert type(classes) is dict
    assert args.tutor_name in classes.keys()
    add_to_args("classes", classes)

    teams = [team for classs in classes.values() for team in classs]
    validate_teams(teams)
    add_to_args("teams", teams)


def process_dynamic_config(data: dict[str, Any]) -> None:
    """
    Extract and check the config values necessary for the dynamic correction
    marking modes, i.e., 'random' and 'exercise'.
    """
    tutor_list = data["tutor_list"]
    assert type(tutor_list) is list
    assert all(type(tutor) is str for tutor in tutor_list)
    assert args.tutor_name in tutor_list
    add_to_args("tutor_list", sorted(tutor_list))

    teams = data["teams"]
    validate_teams(teams)
    add_to_args("teams", teams)


def process_general_config(
    data_individual: dict[str, Any], data_shared: dict[str, Any]
) -> None:
    """
    Extract and check config values that are necessary in all marking modes.
    This includes both individual and shared settings.
    """
    # Individual settings
    tutor_name = data_individual["your_name"]
    assert type(tutor_name) is str
    add_to_args("tutor_name", tutor_name)

    # Use `get` because this config setting is optional.
    ignore_feedback_suffix = data_individual.get("ignore_feedback_suffix", [])
    assert type(ignore_feedback_suffix) is list
    assert all(
        type(suffix) is str and suffix[0] == "."
        for suffix in ignore_feedback_suffix
    )
    add_to_args("ignore_feedback_suffix", ignore_feedback_suffix + [".xopp"])

    # Email settings, currently all optional because not fully functional.
    tutor_email = data_individual.get("your_email", "")
    assert (tutor_email == "") or (
        type(tutor_email) is str and is_email(tutor_email)
    )
    add_to_args("tutor_email", tutor_email)

    feedback_email_cc = data_individual.get("feedback_email_cc", [])
    assert type(feedback_email_cc) is list
    assert (feedback_email_cc == []) or all(
        type(email) is str and is_email(email) for email in feedback_email_cc
    )
    add_to_args("feedback_email_cc", feedback_email_cc)

    smtp_url = data_individual.get("smtp_url", "smtp-ext.unibas.ch")
    assert type(smtp_url) is str
    add_to_args("smtp_url", smtp_url)

    smtp_port = data_individual.get("smtp_port", 587)
    assert type(smtp_port) is int
    add_to_args("smtp_port", smtp_port)

    smtp_user = data_individual.get("smtp_user", "")
    assert type(smtp_user) is str
    add_to_args("smtp_user", smtp_user)

    # Shared settings
    lecture_title = data_shared["lecture_title"]
    assert lecture_title and type(lecture_title) is str
    add_to_args("lecture_title", lecture_title)

    marking_mode = data_shared["marking_mode"]
    assert marking_mode in ["static", "random", "exercise"]
    add_to_args("marking_mode", marking_mode)

    max_team_size = data_shared["max_team_size"]
    assert type(max_team_size) is int and max_team_size > 0
    add_to_args("max_team_size", max_team_size)

    use_marks_file = data_shared["use_marks_file"]
    assert type(use_marks_file) is str and use_marks_file.lower() in [
        "true",
        "false",
    ]
    add_to_args("use_marks_file", use_marks_file.lower() == "true")

    points_per = data_shared["points_per"]
    assert type(points_per) is str
    assert points_per in ["sheet", "exercise"]
    add_to_args("points_per", points_per)

    min_point_unit = data_shared["min_point_unit"]
    assert type(min_point_unit) is float or type(min_point_unit) is int
    assert min_point_unit > 0
    add_to_args("min_point_unit", min_point_unit)


def add_to_args(key: str, value: Any) -> None:
    """
    Settings defined in the config file are parsed, checked, and then added to
    the args object. Similar with information that is calculated using things
    defined in the config file. After parsing, all necessary information should
    be contained in the args object.
    """
    vars(args).update({key: value})


# =============================== Main Function ================================


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="")
    # Main command arguments ---------------------------------------------------
    parser.add_argument(
        "-s",
        "--config-shared",
        default=DEFAULT_SHARED_CONFIG_FILE,
        type=pathlib.Path,
        help=(
            "path to the json config file containing shared settings such as "
            "the student/email list"
        ),
    )
    parser.add_argument(
        "-i",
        "--config-individual",
        default=DEFAULT_INDIVIDUAL_CONFIG_FILE,
        type=pathlib.Path,
        help=(
            "path to the json config file containing individual settings such "
            "as tutor name and email configuration"
        ),
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="only print errors, no infos or warnings",
    )
    # Subcommands ==============================================================
    subparsers = parser.add_subparsers(
        required=True,
        help="available sub-commands",
        dest="sub_command",
    )
    # Init command and arguments -----------------------------------------------
    parser_init = subparsers.add_parser(
        "init",
        help="unpack zip file from ADAM and prepare directory structure",
    )
    parser_init.add_argument(
        "adam_zip_path",
        type=pathlib.Path,
        help="path to the zip file downloaded from ADAM",
    )
    parser_init.add_argument(
        "-t",
        "--target",
        type=pathlib.Path,
        required=False,
        help="path to the directory that will contain the submissions",
    )
    group = parser_init.add_mutually_exclusive_group(required=False)
    group.add_argument(
        "-n",
        "--num-exercises",
        dest="num_exercises",
        type=int,
        help="the number of exercises in the sheet",
    )
    group.add_argument(
        "-e",
        "--exercises",
        dest="exercises",
        nargs="+",
        type=int,
        help="the exercises you have to mark",
    )
    parser_init.add_argument(
        "-x",
        "--xopp",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="generate .xopp files",
    )
    parser_init.set_defaults(func=init)
    # Collect command and arguments --------------------------------------------
    parser_collect = subparsers.add_parser(
        "collect",
        help="collect feedback files after marking is done",
    )
    parser_collect.add_argument(
        "-x",
        "--xopp",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="export .xopp files",
    )
    parser_collect.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_collect.set_defaults(func=collect)
    # Combine command and arguments  -------------------------------------------
    parser_combine = subparsers.add_parser(
        "combine",
        help=(
            "combine multiple share archives, only necessary if tutors mark per"
            " exercise and have to integrate their individual feedback into a"
            " single ZIP file to send to the students"
        ),
    )
    parser_combine.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_combine.set_defaults(func=combine)
    # Send command and arguments -----------------------------------------------
    parser_send = subparsers.add_parser(
        "send",
        help="send feedback via email",
    )
    parser_send.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_send.set_defaults(func=send)

    args = parser.parse_args()

    # Process config files =====================================================
    print_info("Processing config:")
    print_indented(f"Reading shared config file '{args.config_shared}'...")
    with open(args.config_shared, "r", encoding="utf-8") as config_file:
        data_shared = json.load(config_file)
    print_indented(
        f"Reading individual config file '{args.config_individual}'..."
    )
    with open(args.config_individual, "r", encoding="utf-8") as config_file:
        data_individual = json.load(config_file)
    assert data_shared.keys().isdisjoint(data_individual)

    # We currently plan to support the following marking modes.
    # static:   Every tutor corrects the submissions of the teams assigned to
    #           that tutor. These will usually be the teams in that tutors
    #           exercise class.
    # random:   Every tutor corrects some submissions which are assigned
    #           randomly with every sheet.
    # exercise: Every tutor corrects some exercise(s) on all sheets.
    process_general_config(data_individual, data_shared)

    if args.marking_mode == "static":
        process_static_config(data_shared)
    else:
        process_dynamic_config(data_shared)
    print_info("Processed config successfully.")

    # Execute subcommand =======================================================
    print_info(f"Running command '{args.sub_command}'...")
    # This calls the function set as default in the parser.
    # For example, `func` is set to `init` if the subcommand is "init".
    args.func()
    print_info(f"Command '{args.sub_command}' terminated successfully. 🎉")

# Only in Python 3.7+ are dicts order preserving, using older Pythons may cause
# the random assignments to not match up.
