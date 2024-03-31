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
import logging
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
from email.message import EmailMessage
from getpass import getpass

import config

Student = tuple[str, str, str]
Team = list[Student]

DEFAULT_SHARED_CONFIG_FILE = "config-shared.json"
DEFAULT_INDIVIDUAL_CONFIG_FILE = "config-individual.json"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"
FEEDBACK_FILE_PREFIX = "feedback_"
SHEET_INFO_FILE_NAME = ".sheet_info"
SHARE_ARCHIVE_PREFIX = "share_archive"
COMBINED_DIR_NAME = "feedback_combined"

# Might be necessary to make colored output work on Windows.
os.system("")


# ============================= Utility Functions ==============================

# Logging ----------------------------------------------------------------------

def configure_logging(level=logging.INFO):
    class ColoredFormatter(logging.Formatter):
        FORMATS = {
            logging.DEBUG:    "\033[0;37m[{levelname}]\033[0m {message}",
            logging.INFO:     "\033[0;34m[{levelname}]\033[0m {message}",
            logging.WARNING:  "\033[0;33m[{levelname}]\033[0m {message}",
            logging.ERROR:    "\033[0;31m[{levelname}]\033[0m {message}",
            logging.CRITICAL: "\033[0;31m[{levelname}]\033[0m {message}",
        }

        def format(self, record):
            formatter = logging.Formatter(ColoredFormatter.FORMATS[record.levelno], style="{")
            return formatter.format(record)

    class LevelFilter:
        def __init__(self, min_level, max_level):
            self.min_level = min_level
            self.max_level = max_level

        def filter(self, record):
            return self.min_level <= record.levelno <= self.max_level

    class CustomHandler(logging.StreamHandler):
        def __init__(self, stream, min_level=logging.DEBUG, max_level=logging.CRITICAL):
            logging.StreamHandler.__init__(self, stream)
            self.setFormatter(ColoredFormatter())
            self.addFilter(LevelFilter(min_level, max_level))

        def emit(self, record):
            logging.StreamHandler.emit(self, record)
            if record.levelno >= logging.CRITICAL:
                sys.exit("aborting")

    root_logger = logging.getLogger("")
    # Remove old handlers
    for handler in root_logger.handlers:
        root_logger.removeHandler(handler)
    root_logger.addHandler(CustomHandler(sys.stdout, max_level=logging.WARNING))
    root_logger.addHandler(CustomHandler(sys.stderr, min_level=logging.ERROR))
    root_logger.setLevel(level)


# Printing ---------------------------------------------------------------------

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
        logging.warning(
            f"Invalid choice '{choice}'. Please respond with 'yes' or 'no'."
        )
        return query_yes_no(text, default)

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
    if config.get().marking_mode == "exercise":
        # TODO: I'm not sure why I added the team_id here. Add it back in if
        # it's necessary, remove these lines otherwise.
        # team_id = team_dir.name.split("_")[0]
        # prefix = team_id + "_" + prefix
        file_name += config.get().tutor_name + "_"
        file_name += "_".join([f"ex{exercise}" for exercise in args.exercises])
    elif config.get().marking_mode == "random":
        file_name += config.get().tutor_name
    elif config.get().marking_mode == "static":
        # Remove trailing underscore.
        file_name = file_name[:-1]
    else:
        logging.critical(f"Unsupported marking mode {config.get().marking_mode}!")
    return file_name


def get_combined_feedback_file_name() -> str:
    return FEEDBACK_FILE_PREFIX + get_adam_sheet_name_string()


def get_marks_file_path():
    return (
        args.sheet_root_dir
        / f"points_{config.get().tutor_name.lower()}_{get_adam_sheet_name_string()}.json"
    )


# Miscellaneous ----------------------------------------------------------------

def is_hidden_file(name: str) -> bool:
    """
    Check if a given file name could be a hidden file. In particular a file that
    should not be sent to students as feedback.
    """
    return name.startswith(".") or is_macos_path(name)


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
    Move all content of source directory to destination directory.
    This does not complain if the dst directory already exists.
    """
    assert src.is_dir() and dst.is_dir()
    with tempfile.TemporaryDirectory() as temp_dir:
        shutil.copytree(src, temp_dir, dirs_exist_ok=True)
        shutil.rmtree(src)
        shutil.copytree(temp_dir, dst, dirs_exist_ok=True)


def verify_sheet_root_dir() -> None:
    """
    Ensure that the given sheet root directory is valid. Needed for multiple
    sub-commands such as 'collect', or 'send'.
    """
    if not args.sheet_root_dir.is_dir():
        logging.critical("The given sheet directory is not valid!")


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


def email_to_text(email: EmailMessage) -> None:
    to = email["To"]
    cc = email["CC"]
    subject = email["Subject"]
    content = ""
    attachments = []
    for part in email.walk():
        if part.is_attachment():
            attachments.append(part.get_filename())
        elif not part.is_multipart():
            content += part.get_content()
    lines = []
    lines.append(f"To: {to}")
    if cc:
        lines.append(f"CC: {cc}")
    if attachments:
        lines.append(f"Attachments: {', '.join(attachments)}")
    lines.append(f"Subject: {subject}")
    lines.append("Text:")
    lines.append(content)
    return "\n".join(lines)


def print_emails(emails: list[EmailMessage]) -> None:
    logging.info("Sending emails now would send the following emails:")
    for email in emails:
        print(email_to_text(email))
        print(f"===========\n")


def send_messages(emails: list[EmailMessage]) -> None:
    with smtplib.SMTP(config.get().smtp_url, config.get().smtp_port) as smtp:
        smtp.starttls()
        if config.get().smtp_user:
            password = getpass("Email password: ")
            smtp.login(config.get().smtp_user, password)
        for email in emails:
            logging.info(f"Sending email to {email['To']}")
            smtp.send_message(email)
        logging.info("Done sending emails.")


def get_team_email_subject() -> str:
    """
    Builds the email subject.
    """
    return f"Feedback {args.adam_sheet_name} | {config.get().lecture_title}"


def get_assistant_email_subject() -> str:
    """
    Builds the email subject.
    """
    return f"Marks for {args.adam_sheet_name} | {config.get().lecture_title}"


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


def get_team_email_content(name_list: list[str]) -> str:
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


def get_assistant_email_content() -> str:
    """
    Builds the body of the email.
    """
    return textwrap.dedent(
        f"""
    Dear assistant for {config.get().lecture_title}

    Please find my marks for {args.adam_sheet_name} in the attachment.

    Best,
    Your Tutors
    """
    )[
        1:
    ]  # Removes the leading newline.


def create_email_to_team(team_dir):
    team = args.team_dir_to_team[team_dir.name]
    team_first_names, _, team_emails = zip(*team)
    return construct_email(
        list(team_emails),
        config.get().feedback_email_cc,
        get_team_email_subject(),
        get_team_email_content(team_first_names),
        config.get().tutor_email,
        get_collected_feedback_file(team_dir),
    )


def create_email_to_assistent():
    return construct_email(
        [config.get().assistant_email],
        config.get().feedback_email_cc,
        get_assistant_email_subject(),
        get_assistant_email_content(),
        config.get().tutor_email,
        get_marks_file_path(),
    )


def send() -> None:
    """
    After the collection step finished successfully, send the feedback to the
    students via email. This currently only works if the tutor's email account
    is whitelisted for the smpt-ext.unibas.ch server, or if the tutor uses
    smtp.unibas.ch with an empty smpt_user.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()
    if config.get().marking_mode == "exercise":
        logging.critical(
            "Sending for marking mode 'exercise' is not implemented because "
            "collection is not yet figured out."
        )
    # Send emails.
    emails: list[EmailMessage] = []
    for team_dir in get_relevant_team_dirs():
        emails.append(create_email_to_team(team_dir))
    if config.get().assistant_email:
        emails.append(create_email_to_assistent())
    logging.info(f"Ready to send {len(emails)} email(s).")
    if args.dry_run:
        print_emails(emails)
    else:
        send_messages(emails)


# ============================ Collect Sub-Command =============================


def validate_marks_json() -> None:
    """
    Verify that all necessary marks are present in the MARK_FILE_NAME file and
    adhere to the granularity defined in the config file.
    """
    marks_json_file = get_marks_file_path()
    if not marks_json_file.is_file():
        logging.critical(
            f"Missing points file in directory '{args.sheet_root_dir}'!"
        )
    with open(marks_json_file, "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)
    relevant_teams = [
        relevant_team.name for relevant_team in get_relevant_team_dirs()
    ]
    marked_teams = list(marks.keys())
    if sorted(relevant_teams) != sorted(marked_teams):
        logging.critical(
            "There is no 1-to-1 mapping between team directories "
            "that need to be marked and entries in the "
            f"'{marks_json_file.name}' "
            "file! Make sure that it contains exactly one entry for every team "
            "directory that needs to be marked, and that directory name and "
            "key are the same."
        )
    if config.get().points_per == "exercise":
        marks_list = [
            mark
            for team_marks in marks.values()
            for mark in team_marks.values()
        ]
    else:
        marks_list = marks.values()
    if not all(marks_list):
        logging.critical(
            f"There are missing points in the '{marks_json_file.name}' file!"
        )
    if not all(
        (float(mark) / config.get().min_point_unit).is_integer() for mark in marks_list
    ):
        logging.critical(
            f"'{marks_json_file.name}' contains marks that are more fine-grained "
            "than allowed! You may only award points in "
            f"'{config.get().min_point_unit}' increments."
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
        logging.critical(f"Missing feedback directory for team {team_dir.name}!")
    content = list(feedback_dir.iterdir())
    if any(".todo" in file_or_dir.name for file_or_dir in content):
        logging.critical(
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
        if file.is_file() and not file.suffix in config.get().ignore_feedback_suffix
    ]
    # Ask for confirmation if the feedback directory contains hidden files that
    # are maybe not supposed to be part of the collected feedback.
    hidden_files = [f for f in feedback_files if is_hidden_file(f.name)]
    for hidden_file in hidden_files:
        include_anyway = query_yes_no(
            (
                "There seem to be hidden files in your feedback directory, "
                f"e.g. '{str(hidden_file)}'. Do you want to include them in "
                "your feedback anyway? (Consider adding ignored suffixes in "
                "your individual configuration file to avoid this prompt in "
                "the future.)"
            ),
            default=False,
        )
        if not include_anyway:
            feedback_files.remove(hidden_file)

    if not feedback_files:
        logging.critical(
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
        logging.warning(f"The feedback for {team_dir.name} contains no PDF file!")


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
    logging.info("Exporting .xopp files...")
    for team_dir in get_relevant_team_dirs():
        feedback_dir = team_dir / FEEDBACK_DIR_NAME
        xopp_files = [
            file for file in feedback_dir.rglob("*") if file.suffix == ".xopp"
        ]
        for xopp_file in xopp_files:
            dest = xopp_file.with_suffix(".pdf")
            subprocess.run(["xournalpp", "-p", dest, xopp_file])
    logging.info("Done exporting .xopp files.")


def print_marks() -> None:
    """
    Prints the marks so that they can be easily copy-pasted to the file where
    marks are collected.
    """
    # Read marks file.
    # Don't check whether the marks file exists because `validate_marks_json()`
    # would have already complained.
    with open(get_marks_file_path(), "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)

    # Print marks.
    logging.info("Start of copy-paste marks...")
    # We want all teams printed, not just the marked ones.
    for team_to_print in config.get().teams:
        for team_dir, team in args.team_dir_to_team.items():
            # Every team should only be the value of at most one entry in
            # `team_dir_to_team`.
            if team == team_to_print:
                for student in team:
                    full_name = f"{student[0]} {student[1]}"
                    output_str = f"{full_name:>35};"
                    if config.get().points_per == "exercise":
                        # The value `marks` assigns to the team_dir key is a
                        # dict with (exercise name, mark) pairs.
                        team_marks = marks.get(team_dir, {"null": ""})
                        _, exercise_marks = zip(*team_marks.items())
                        for mark in exercise_marks:
                            output_str += f"{mark:>3};"
                    else:
                        sheet_mark = marks.get(team_dir, "")
                        output_str += f"{sheet_mark:>3}"
                    print(output_str)
    logging.info("End of copy-paste marks.")


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
            logging.info(f"Could not write share archive. Aborting command.")
            return
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
                logging.critical(
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
            logging.info(f"Could not write collected feedback archives. Aborting command.")
            return
    if args.xopp:
        export_xopp_files()
    create_collected_feedback_directories()
    for team_dir in get_relevant_team_dirs():
        collect_feedback_files(team_dir)
    if config.get().marking_mode == "exercise":
        create_share_archive(overwrite)
    if config.get().use_marks_file:
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
        logging.critical(
            f"No share archives exist in {args.sheet_root_dir}. " + instructions
        )
    if len(list(share_archive_files)) == 1:
        logging.warning(
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
            logging.info(f"Could not write to '{combined_dir}'. Aborting combine command.")
            return

    combined_dir.mkdir()

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined

    # Create subdirectories for teams.
    for team_dir in get_relevant_team_dirs():
        combined_team_dir = combined_dir / team_dir.name
        combined_team_dir.mkdir()

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined
    #     ├── 12345_Muster-Meier-Mueller
    #     .

    teams_all = [team_dir.name for team_dir in get_relevant_team_dirs()]
    # Extract feedback files from share archives into their respective team
    # directories in the combined directory.
    for share_archive_file in args.sheet_root_dir.glob(
        SHARE_ARCHIVE_PREFIX + "*.zip"
    ):
        with ZipFile(share_archive_file, mode="r") as share_archive:
            # Check if this share archive is missing team archives for any team.
            teams_present = [
                pathlib.Path(team).stem for team in share_archive.namelist()
            ]
            teams_not_present = list(set(teams_all) - set(teams_present))
            for team_not_present in teams_not_present:
                logging.warning(
                    f"The shared archive {share_archive_file} contains no"
                    f" feedback for team {team_not_present}."
                )
            for team in teams_present:
                # Extract team_archive from share_archive.
                team_archive_file = share_archive.extract(
                    team + ".zip", combined_dir / team
                )
                with ZipFile(team_archive_file, mode="r") as team_archive:
                    team_archive.extractall(path=combined_dir / team)
                pathlib.Path(team_archive_file).unlink()

    """
    I think the step above already accomplishes what this step is supposed to
    accomplish. If no problems arise long-term (today is 2023-10-31), remove
    this block.

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined
    #     ├── 12345_Muster-Meier-Mueller
    #     .   ├── feedback_exercise_sheet_01_tutor1_ex1.pdf
    #     .   └── feedback_exercise_sheet_01_tutor2_ex2.zip

    # Extract zipped feedback in combined directory.
    for team_dir in combined_dir.iterdir():
        assert team_dir.is_dir()
        for feedback_file in team_dir.iterdir():
            assert feedback_file.is_file()
            # If the feedback is not an archive but a single pdf, move on.
            if feedback_file.suffix != ".zip":
                continue
            # Otherwise, extract feedback from feedback archive.
            with ZipFile(feedback_file, mode="r") as feedback_archive:
                feedback_archive.extractall(path=combined_dir / team)
            # Remove feedback archive.
            feedback_file.unlink()
    """

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined
    #     ├── 12345_Muster-Meier-Mueller
    #     .   ├── feedback_exercise_sheet_01_tutor1_ex1.pdf
    #     .   ├── feedback_exercise_sheet_01_tutor2_ex2.pdf
    #     .   └── feedback_exercise_sheet_01_tutor2_ex2_code.cc

    # Zip up feedback files.
    for team_dir in combined_dir.iterdir():
        feedback_files = list(team_dir.iterdir())
        combined_team_archive = team_dir / (
            get_combined_feedback_file_name() + ".zip"
        )
        with ZipFile(combined_team_archive, mode="w") as combined_zip:
            for feedback_file in feedback_files:
                combined_zip.write(feedback_file, arcname=feedback_file.name)
                feedback_file.unlink()

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined
    #     ├── 12345_Muster-Meier-Mueller
    #     .   └── feedback_exercise_sheet_01.zip


# ============================== Init Sub-Command ==============================


def extract_adam_zip() -> tuple[pathlib.Path, str]:
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
            # Should be the name of the exercise sheet, for example "Exercise Sheet 2".
            with ZipFile(args.adam_zip_path, mode="r") as zip_file:
                zip_content = zip_file.namelist()
                sheet_root_dir = pathlib.Path(temp_dir) / zip_content[0]
                filtered_extract(zip_file, pathlib.Path(temp_dir))
        else:
            # Assume the directory is an extracted ADAM zip.
            unzipped_path = pathlib.Path(args.adam_zip_path)
            unzipped_destination_path = (
                pathlib.Path(temp_dir) / unzipped_path.name
            )
            sheet_root_dir = pathlib.Path(
                shutil.copytree(unzipped_path, unzipped_destination_path)
            )
        # Store ADAM exercise sheet name to use as random seed.
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
    move_content_and_delete(sub_dirs[0], sheet_root_dir)
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
            for team in config.get().teams
            if any(submission_email in student for student in team)
        ]
        if len(teams) == 0:
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
    if config.get().marking_mode == "static":
        return config.get().classes[config.get().tutor_name]
    elif config.get().marking_mode == "random":
        # Here not all teams are assigned to a tutor, but only those that
        # submitted something. This is to ensure that submissions can be
        # distributed fairly among tutors.
        num_tutors = len(config.get().tutor_list)
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
        shuffled_tutors = config.get().tutor_list.copy()
        random.Random(seed).shuffle(shuffled_tutors)
        return chunks[shuffled_tutors.index(config.get().tutor_name)]
    elif config.get().marking_mode == "exercise":
        return config.get().teams
    else:
        logging.critical(f"Unsupported marking mode {config.get().marking_mode}!")
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
            logging.warning(
                f"There are multiple submissions for group '{team_dir.name}'!"
            )
        if len(team_submission_dirs) < 1:
            logging.warning(
                f"The submission of group '{team_dir.name}' is empty!"
            )
        for team_submission_dir in team_submission_dirs:
            move_content_and_delete(team_submission_dir, team_dir)


def unzip_internal_zips() -> None:
    """
    If multiple files are uploaded to ADAM, the submission becomes a single zip
    file. Here we extract this zip. I'm not sure if nested zip files are also
    extracted. Additionally we flatten the directory as long as a level only
    consists of a single directory.
    """
    for team_dir in get_all_team_dirs():
        if not team_dir.is_dir():
            continue
        for zip_file in team_dir.glob("**/*.zip"):
            with ZipFile(zip_file, mode="r") as zf:
                filtered_extract(zf, zip_file.parent)
            os.remove(zip_file)
        sub_dirs = list(team_dir.iterdir())
        while len(sub_dirs) == 1 and sub_dirs[0].is_dir():
            move_content_and_delete(sub_dirs[0], team_dir)
            sub_dirs = list(team_dir.iterdir())


def create_marks_file() -> None:
    """
    Write a json file to add the marks for all relevant teams and exercises.
    """
    exercise_dict: Union[str, dict[str, str]] = ""
    if config.get().points_per == "exercise":
        if config.get().marking_mode == "static" or config.get().marking_mode == "random":
            exercise_dict = {
                f"exercise_{i}": "" for i in range(1, args.num_exercises + 1)
            }
        elif config.get().marking_mode == "exercise":
            exercise_dict = {f"exercise_{i}": "" for i in args.exercises}
    else:
        exercise_dict = ""

    marks_dict = {}
    for team_dir in sorted(list(get_relevant_team_dirs())):
        marks_dict.update({team_dir.name: exercise_dict})

    with open(get_marks_file_path(), "w", encoding="utf-8") as marks_json:
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

    logging.info("Generating .xopp files...")
    for team_dir in get_relevant_team_dirs():
        pdf_paths = list(team_dir.glob("*.pdf"))
        if len(pdf_paths) != 1:
            logging.warning(
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
            logging.warning(
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
    logging.info("Done generating .xopp files.")


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
    if config.get().marking_mode == "exercise":
        info_dict.update({"exercises": args.exercises})
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "w", encoding="utf-8"
    ) as sheet_info_file:
        # Sorting the keys here is essential because the order of teams here
        # will influence the assignment returned by `get_relevant_teams()` in
        # case config.get().marking_mode == "random".
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
        team for team in config.get().teams if not team in adam_id_to_team.values()
    ]
    if missing_teams:
        logging.warning("There are no submissions for the following team(s):")
        for missing_team in missing_teams:
            print(f"* {team_to_string(missing_team)}")


def init() -> None:
    """
    Prepares the directory structure holding the submissions.
    """
    # Catch wrong combinations of marking_mode/points_per/-n/-e.
    # Not possible eariler because marking_mode and points_per are given by the
    # config file.
    if config.get().points_per == "exercise":
        if config.get().marking_mode == "exercise" and not args.exercises:
            logging.critical(
                "You must provide a list of exercise numbers to be marked with "
                "the '-e' flag, for example '-e 1 3 4'."
            )
        if (
            config.get().marking_mode == "random" or config.get().marking_mode == "static"
        ) and not args.num_exercises:
            logging.critical(
                "You must provide the number of exercises in the sheet with "
                "the '-n' flag, for example '-n 5'."
            )
    else:  # points per sheet
        if config.get().marking_mode == "exercise":
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

    if config.get().use_marks_file:
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
    configure_logging()
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
    # Subcommands ==============================================================
    subparsers = parser.add_subparsers(
        required=True,
        help="available sub-commands",
        dest="sub_command",
    )
    # Help command -------------------------------------------------------------
    parser_init = subparsers.add_parser(
        "help",
        help=(
            "print this help message; "
            "run e.g. 'python3 adam-script.py init -h' to print help of "
            "sub-command"
        ),
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
    parser_send.add_argument(
        "-d",
        "--dry_run",
        action="store_true",
        help="only print emails instead of sending them",
    )
    parser_send.set_defaults(func=send)

    args = parser.parse_args()

    if args.sub_command == "help":
        parser.print_help()
        sys.exit(0)

    # Process config files =====================================================
    config.load(args.config_shared, args.config_individual)

    # Execute subcommand =======================================================
    logging.info(f"Running command '{args.sub_command}'...")
    # This calls the function set as default in the parser.
    # For example, `func` is set to `init` if the subcommand is "init".
    args.func()
    logging.info(f"Command '{args.sub_command}' terminated successfully. 🎉")

# Only in Python 3.7+ are dicts order preserving, using older Pythons may cause
# the random assignments to not match up.
