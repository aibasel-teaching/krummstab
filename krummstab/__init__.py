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
import json
import logging
import mimetypes
import os
import pathlib
import shutil
import subprocess
import sys
import tempfile
import textwrap
from importlib import resources
import jsonschema

from zipfile import ZipFile

# For typing annotations.
from typing import Any, Optional, Union
from collections.abc import Iterator

# Needed for email stuff.
import smtplib
from email.message import EmailMessage
from getpass import getpass

from . import config
from . import schemas

Student = tuple[str, str, str]
Team = list[Student]

DEFAULT_SHARED_CONFIG_FILE = "config-shared.json"
DEFAULT_INDIVIDUAL_CONFIG_FILE = "config-individual.json"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_COLLECTED_DIR_NAME = "feedback_collected"
FEEDBACK_FILE_PREFIX = "feedback_"
SHEET_INFO_FILE_NAME = "sheet.json"
SUBMISSION_INFO_FILE_NAME = "submission.json"
SHARE_ARCHIVE_PREFIX = "share_archive"
COMBINED_DIR_NAME = "feedback_combined"

args = None

# Might be necessary to make colored output work on Windows.
os.system("")


# ============================= Utility Functions ==============================

# Logging ----------------------------------------------------------------------


def configure_logging(level=logging.INFO):
    class ColoredFormatter(logging.Formatter):
        FORMATS = {
            logging.DEBUG: "\033[0;37m[{levelname}]\033[0m {message}",
            logging.INFO: "\033[0;34m[{levelname}]\033[0m {message}",
            logging.WARNING: "\033[0;33m[{levelname}]\033[0m {message}",
            logging.ERROR: "\033[0;31m[{levelname}]\033[0m {message}",
            logging.CRITICAL: "\033[0;31m[{levelname}]\033[0m {message}",
        }

        def format(self, record):
            formatter = logging.Formatter(
                ColoredFormatter.FORMATS[record.levelno], style="{"
            )
            return formatter.format(record)

    class LevelFilter:
        def __init__(self, min_level, max_level):
            self.min_level = min_level
            self.max_level = max_level

        def filter(self, record):
            return self.min_level <= record.levelno <= self.max_level

    class CustomHandler(logging.StreamHandler):
        def __init__(
            self, stream, min_level=logging.DEBUG, max_level=logging.CRITICAL
        ):
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


def get_team_key(submission_info: dict) -> str:
    """
    Create a string representation of the team with the ADAM ID: team_id_LastName1-LastName2
    """
    team_id = submission_info.get("adam_id")
    team = submission_info.get("team")
    return team_id + "_" + team_to_string(team)


def get_adam_sheet_name_string() -> str:
    """
    Turn the sheet name given by ADAM into a string usable for file names.
    """
    return args.adam_sheet_name.replace(" ", "_").lower()


def get_feedback_file_name(_the_config: config.Config) -> str:
    file_name = FEEDBACK_FILE_PREFIX + get_adam_sheet_name_string() + "_"
    if _the_config.marking_mode == "exercise":
        # TODO: I'm not sure why I added the team_id here. Add it back in if
        # it's necessary, remove these lines otherwise.
        # team_id = team_dir.name.split("_")[0]
        # prefix = team_id + "_" + prefix
        file_name += _the_config.tutor_name + "_"
        file_name += "_".join([f"ex{exercise}" for exercise in args.exercises])
    elif _the_config.marking_mode == "static":
        # Remove trailing underscore.
        file_name = file_name[:-1]
    else:
        unsupported_marking_mode_error(_the_config)
    return file_name


def get_combined_feedback_file_name() -> str:
    return FEEDBACK_FILE_PREFIX + get_adam_sheet_name_string()


def get_marks_file_path(_the_config: config.Config):
    return (
        args.sheet_root_dir
        / f"points_{_the_config.tutor_name.lower()}_{get_adam_sheet_name_string()}.json"
    )


def get_individual_marks_file_path(_the_config: config.Config):
    marks_file_path = get_marks_file_path(_the_config)
    individual_marks_file_path = marks_file_path.with_name(
        marks_file_path.stem + "_individual" + marks_file_path.suffix
    )
    return individual_marks_file_path


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


def get_submission_info(path: pathlib.Path) -> dict:
    """
    Load the submission info of the team directory. In particular the
    team and if it is a relevant team.
    """
    try:
        with open(
                path / SUBMISSION_INFO_FILE_NAME, "r", encoding="utf-8"
        ) as submission_info_file:
            submission_info = json.load(submission_info_file)
            submission_info_schema = json.loads(
                resources.files(schemas).joinpath("submission-info-schema.json").read_text(encoding="utf-8"))
            jsonschema.validate(submission_info, submission_info_schema, jsonschema.Draft7Validator)
            return submission_info
    except FileNotFoundError:
        logging.critical("The submission.json file does not exist.")
    except NotADirectoryError:
        logging.critical(f"The path '{path}' is not a team directory.")
    except jsonschema.exceptions.ValidationError as error:
        logging.critical(f"The submission.json file does not have the right format: {error.message}")


def get_all_team_dirs() -> Iterator[pathlib.Path]:
    """
    Return all team directories within the sheet root directory. To find the team
    directories, the submission.json files are used. This excludes other
    directories that may be created in the sheet root directory, such as
    one containing combined feedback.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if team_dir.is_dir() and team_dir != args.sheet_root_dir / COMBINED_DIR_NAME:
            yield team_dir


def get_relevant_team_dirs() -> Iterator[pathlib.Path]:
    """
    Return the team directories of the teams whose submission has to be
    corrected by the tutor running the script.
    """
    for team_dir in get_all_team_dirs():
        submission_info = get_submission_info(team_dir)
        if submission_info.get("relevant"):
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
    if not collected_feedback_dir.is_dir():
        logging.critical(
            "The directory for collected feedback at"
            f" '{str(collected_feedback_dir)}' does not exist. You probably"
            " have to run the 'collect' command first."
        )
    collected_feedback_files = list(collected_feedback_dir.iterdir())
    assert (
        len(collected_feedback_files) == 1
        and collected_feedback_files[0].is_file()
        and collected_feedback_files[0].suffix in [".pdf", ".zip"]
    )
    return collected_feedback_files[0]


def get_combined_feedback_file(team_dir: pathlib.Path) -> pathlib.Path:
    """
    Given a team directory, return the combined feedback file. This is always a
    zip archive because in the usual case it will contain feedback from multiple
    tutors.
    """
    combined_feedback_dir = args.sheet_root_dir / COMBINED_DIR_NAME
    if not combined_feedback_dir.is_dir():
        logging.critical(
            "The directory for combined feedback at"
            f" '{str(combined_feedback_dir)}' does not exist. You probably"
            " have to run the 'combine' command first."
        )
    combined_feedback_files = list(
        (combined_feedback_dir / team_dir.name).iterdir()
    )
    assert (
        len(combined_feedback_files) == 1
        and combined_feedback_files[0].is_file()
        and combined_feedback_files[0].suffix == ".zip"
    )
    return combined_feedback_files[0]


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


def unsupported_marking_mode_error(_the_config: config.Config) -> None:
    """
    Throw an error if a marking mode is encountered that is not handled
    correctly. This is primarily used in the else case of if-elif-else branches
    on the marking_mode, which shouldn't be reached in regular operation anyway
    because the config settings should be validated first. But this function
    offers an easy way to find sections to consider if we ever want to add a new
    marking_mode.
    """
    logging.critical(f"Unsupported marking mode {_the_config.marking_mode}!")


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
    if cc:
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

    def format_line(left: str, right: str) -> str:
        return f"\033[0;34m{left}\033[0m {right}"

    lines.append(format_line("To:", to))
    if cc:
        lines.append(format_line("CC:", cc))
    if attachments:
        lines.append(format_line("Attachments:", ", ".join(attachments)))
    lines.append(format_line("Subject:", subject))
    lines.append(format_line("Start Body", ""))
    if content[-1] == "\n":
        content = content[:-1]
    lines.append(content)
    lines.append(format_line("End Body", ""))
    return "\n".join(lines)


def print_emails(emails: list[EmailMessage]) -> None:
    separator = "\n\033[0;33m" + 80 * "=" + "\033[0m\n"
    email_strings = [email_to_text(email) for email in emails]
    print(f"{separator}{separator.join(email_strings)}{separator}")


def send_messages(emails: list[EmailMessage], _the_config: config.Config) -> None:
    with smtplib.SMTP(_the_config.smtp_url, _the_config.smtp_port) as smtp:
        smtp.starttls()
        if _the_config.smtp_user:
            logging.warning(
                "The setting 'smtp_user' should probably be empty for the"
                " 'send' command to work, trying anyway."
            )
            password = getpass("Email password: ")
            smtp.login(_the_config.smtp_user, password)
        for email in emails:
            logging.info(f"Sending email to {email['To']}")
            # During testing, I didn't manage to trigger the exceptions below.
            # Additionally `refused_recipients` was always empty, even when the
            # documentation of smtplib states that it should be populated when
            # some but not all of the recipients are refused. Instead I always
            # get receive an email from the Outlook server containing the error
            # message.
            try:
                refused_recipients = smtp.send_message(email)
            except smtplib.SMTPRecipientsRefused:
                logging.warning(
                    f"Email to '{email['To']}' failed to deliver because all"
                    " recipients were refused."
                )
            except smtplib.SMTPSenderRefused:
                logging.critical(
                    "Email sender was refused, failed to deliver any emails."
                )
            except (
                smtplib.SMTPHeloError,
                smtplib.SMTPDataError,
                smtplib.SMTPNotSupportedError,
            ):
                logging.warning(
                    f"Email to '{email['To']}' failed to deliver because of"
                    " some weird error."
                )
            for refused_recipient, (
                smtp_error,
                error_message,
            ) in refused_recipients.items():
                logging.warning(
                    f"Email to '{refused_recipient}' failed to deliver because"
                    " the recipient was refused with the SMTP error code"
                    f" '{smtp_error}' and the message '{error_message}'."
                )
        logging.info("Done sending emails.")


def get_team_email_subject(_the_config: config.Config) -> str:
    """
    Builds the email subject.
    """
    return f"Feedback {args.adam_sheet_name} | {_the_config.lecture_title}"


def get_assistant_email_subject(_the_config: config.Config) -> str:
    """
    Builds the email subject.
    """
    return f"Marks for {args.adam_sheet_name} | {_the_config.lecture_title}"


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


def get_team_email_content(name_list: list[str], _the_config: config.Config) -> str:
    """
    Builds the body of the email that sends feedback to students.
    """
    return textwrap.dedent(
        f"""
    {get_email_greeting(name_list)}

    Please find feedback on your submission for {args.adam_sheet_name} in the attachment.
    If you have any questions, you can contact us in the exercise session or by replying to this email (reply to all).

    Best,
    {_the_config.email_signature}"""  # noqa
    )[
        1:
    ]  # Removes the leading newline.


def get_assistant_email_content(_the_config: config.Config) -> str:
    """
    Builds the body of the email that sends the points to the assistent.
    """
    return textwrap.dedent(
        f"""
    Dear assistant for {_the_config.lecture_title}

    Please find my marks for {args.adam_sheet_name} in the attachment.

    Best,
    {_the_config.email_signature}"""
    )[
        1:
    ]  # Removes the leading newline.


def get_assistant_email_attachment_path(_the_config: config.Config) -> pathlib.Path:
    """
    Sending the marks file to the assistant as is has the disadvantage that
    points are only listed per team. This makes it difficult for the assistant
    to figure out how many points each individual student has. So we use the
    individual marks file where points are listed per student.
    """
    return get_individual_marks_file_path(_the_config)


def create_email_to_team(team_dir, _the_config: config.Config):
    submission_info = get_submission_info(team_dir)
    team_first_names, _, team_emails = zip(*submission_info.get("team"))
    if _the_config.marking_mode == "exercise":
        feedback_file_path = get_combined_feedback_file(team_dir)
    elif _the_config.marking_mode == "static":
        feedback_file_path = get_collected_feedback_file(team_dir)
    else:
        unsupported_marking_mode_error(_the_config)
    return construct_email(
        list(team_emails),
        _the_config.feedback_email_cc,
        get_team_email_subject(_the_config),
        get_team_email_content(team_first_names, _the_config),
        _the_config.tutor_email,
        feedback_file_path,
    )


def create_email_to_assistant(_the_config: config.Config):
    return construct_email(
        [_the_config.assistant_email],
        _the_config.feedback_email_cc,
        get_assistant_email_subject(_the_config),
        get_assistant_email_content(_the_config),
        _the_config.tutor_email,
        get_assistant_email_attachment_path(_the_config),
    )


def send(_the_config: config.Config) -> None:
    """
    After the collection step finished successfully, send the feedback to the
    students via email. This currently only works if the tutor's email account
    is whitelisted for the smtp-ext.unibas.ch server, or if the tutor uses
    smtp.unibas.ch with an empty smtp_user.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()
    # Send emails.
    emails: list[EmailMessage] = []
    for team_dir in get_relevant_team_dirs():
        emails.append(create_email_to_team(team_dir, _the_config))
    # TODO: As of now the plan is to only send assistant emails if the marking
    # mode is "static" because there the assistant collects the points
    # centrally. In case of "exercise", we plan to distribute the point files
    # through the share_archives, so there is no need to send an email to the
    # assistent, but this may change in the future.
    if _the_config.marking_mode != "exercise" and _the_config.assistant_email:
        emails.append(create_email_to_assistant(_the_config))
    logging.info(f"Drafted {len(emails)} email(s).")
    if args.dry_run:
        logging.info("Sending emails now would send the following emails:")
        print_emails(emails)
        logging.info("No emails sent.")
    else:
        print_emails(emails)
        really_send = query_yes_no(
            (
                f"Do you really want to send the {len(emails)} email(s) "
                "printed above?"
            ),
            default=False,
        )
        if really_send:
            send_messages(emails, _the_config)
        else:
            logging.info("No emails sent.")


# ============================ Collect Sub-Command =============================


def validate_marks_json(_the_config: config.Config) -> None:
    """
    Verify that all necessary marks are present in the MARK_FILE_NAME file and
    adhere to the granularity defined in the config file.
    """
    marks_json_file = get_marks_file_path(_the_config)
    if not marks_json_file.is_file():
        logging.critical(
            f"Missing points file in directory '{args.sheet_root_dir}'!"
        )
    with open(marks_json_file, "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)
    relevant_teams = []
    for team_dir in get_relevant_team_dirs():
        submission_info = get_submission_info(team_dir)
        relevant_teams.append(get_team_key(submission_info))
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
    if _the_config.points_per == "exercise":
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
        (float(mark) / _the_config.min_point_unit).is_integer() for mark in marks_list
    ):
        logging.critical(
            f"'{marks_json_file.name}' contains marks that are more"
            " fine-grained than allowed! You may only award points in"
            f" '{_the_config.min_point_unit}' increments."
        )


def collect_feedback_files(team_dir: pathlib.Path, _the_config: config.Config) -> None:
    """
    Take the contents of a {team_dir}/feedback directory and collect the files
    that actually contain feedback (e.g., no .xopp files). If there are
    multiple, add them to a zip archive and save it to
    {team_dir}/feedback_collected. If there is only a single pdf, copy it to
    {team_dir}/feedback_collected.
    """
    feedback_dir = team_dir / FEEDBACK_DIR_NAME
    collected_feedback_dir = team_dir / FEEDBACK_COLLECTED_DIR_NAME
    collected_feedback_zip_name = get_feedback_file_name(_the_config) + ".zip"
    # Error handling.
    if not feedback_dir.exists():
        logging.critical(
            f"Missing feedback directory for team {team_dir.name}!"
        )
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
        if file.is_file() and file.suffix not in _the_config.ignore_feedback_suffix
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
        logging.warning(
            f"The feedback for {team_dir.name} contains no PDF file!"
        )


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


def print_marks(_the_config: config.Config) -> None:
    """
    Prints the marks so that they can be easily copy-pasted to the file where
    marks are collected.
    """
    # Read marks file.
    # Don't check whether the marks file exists because `validate_marks_json()`
    # would have already complained.
    with open(get_marks_file_path(_the_config), "r", encoding="utf-8") as marks_file:
        marks = json.load(marks_file)

    # Print marks.
    logging.info("Start of copy-paste marks...")
    # We want all teams printed, not just the marked ones.
    for team_to_print in _the_config.teams:
        for team_dir in get_all_team_dirs():
            submission_info = get_submission_info(team_dir)
            if submission_info.get("team") == team_to_print:
                key = get_team_key(submission_info)
                for student in submission_info.get("team"):
                    full_name = f"{student[0]} {student[1]}"
                    output_str = f"{full_name:>35};"
                    if _the_config.points_per == "exercise":
                        # The value `marks` assigned to the team_dir key is a
                        # dict with (exercise name, mark) pairs.
                        team_marks = marks.get(key, {"null": ""})
                        _, exercise_marks = zip(*team_marks.items())
                        for mark in exercise_marks:
                            output_str += f"{mark:>3};"
                    else:
                        sheet_mark = marks.get(key, "")
                        output_str += f"{sheet_mark:>3}"
                    print(output_str)
    logging.info("End of copy-paste marks.")


def create_individual_marks_file(_the_config: config.Config) -> None:
    """
    Write a json file to add the marks per student.
    """
    with open(get_marks_file_path(_the_config), "r", encoding="utf-8") as marks_file:
        team_marks = json.load(marks_file)
    student_marks = {}
    for team_dir in get_relevant_team_dirs():
        submission_info = get_submission_info(team_dir)
        team_key = get_team_key(submission_info)
        for first_name, last_name, email in submission_info.get("team"):
            student_key = email.lower()
            student_marks.update({student_key: team_marks.get(team_key)})
    file_content = {
        "tutor_name": _the_config.tutor_name,
        "adam_sheet_name": get_adam_sheet_name_string(),
        "marks": student_marks
    }
    if _the_config.points_per == "exercise" and _the_config.marking_mode == "exercise":
        file_content["exercises"] = args.exercises
    with open(get_individual_marks_file_path(_the_config), "w", encoding="utf-8") as file:
        json.dump(file_content, file, indent=4, ensure_ascii=False)


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
    # be written to sheet.json and then read in by a load_sheet_info().
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
        assert overwrite
        if overwrite is None:
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
            logging.critical("Aborting 'combine' without overwriting existing share archive.")
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


def collect(_the_config: config.Config) -> None:
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
            logging.critical("Aborting 'collect' without overwriting existing collected feedback.")
    if args.xopp:
        export_xopp_files()
    create_collected_feedback_directories()
    for team_dir in get_relevant_team_dirs():
        collect_feedback_files(team_dir, _the_config)
    if _the_config.marking_mode == "exercise":
        create_share_archive(overwrite)
    if _the_config.use_marks_file:
        validate_marks_json(_the_config)
        print_marks(_the_config)
        create_individual_marks_file(_the_config)


# ============================ Combine Sub-Command =============================


def combine(_the_config: config.Config) -> None:
    """
    Combine multiple share archives so that in the end we have one zip archive
    per team containing all feedback for that team.
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
            logging.info(
                f"Could not write to '{combined_dir}'. Aborting combine"
                " command."
            )
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
            # Should be the name of the exercise sheet,
            # for example "Exercise Sheet 2".
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
    move_content_and_delete(sub_dirs[0], sheet_root_dir)
    return sheet_root_dir, adam_sheet_name


def mark_irrelevant_team_dirs(_the_config: config.Config) -> None:
    """
    Indicate which team directories do not have to be marked by adding the
    `DO_NOT_MARK_PREFIX` to their directory name.
    """
    for team_dir in get_all_team_dirs():
        submission_info = get_submission_info(team_dir)
        if not submission_info.get("relevant"):
            shutil.move(
                team_dir, team_dir.with_name(DO_NOT_MARK_PREFIX + team_dir.name)
            )


def get_relevant_teams(_the_config: config.Config) -> list[Team]:
    """
    Get a list of teams that the tutor specified in the config has to mark.
    We rename the directories using the `DO_NOT_MARK_PREFIX` and thereafter only
    access relevant teams via `get_relevant_team_dirs()`.
    """
    if _the_config.marking_mode == "static":
        return _the_config.classes[_the_config.tutor_name]
    elif _the_config.marking_mode == "exercise":
        return _the_config.teams
    else:
        unsupported_marking_mode_error(_the_config)
        return []


def rename_team_dirs() -> None:
    """
    The team directories are renamed to: team_id_LastName1-LastName2
    The team ID can be helpful to identify a team on the ADAM web interface.
    """
    for team_dir in get_all_team_dirs():
        submission_info = get_submission_info(team_dir)
        team_key = get_team_key(submission_info)
        team_dir = pathlib.Path(
            shutil.move(team_dir, team_dir.with_name(team_key))
        )


def flatten_team_dirs() -> None:
    """
    There can be multiple directories within a "Team 00000" directory. This
    happens when multiple members of the team upload solutions. Sometimes, only
    one directory contains submitted files, in this case we remove the empty
    ones silently. In case multiple submissions exist, we put the files within
    them next to each other and print a warning.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if team_dir.is_dir():
            # Remove empty subdirectories.
            for team_submission_dir in team_dir.iterdir():
                if team_submission_dir.is_dir() and len(list(team_submission_dir.iterdir())) == 0:
                    team_submission_dir.rmdir()
            # Store the list of team submission directories in variable, because the
            # generator may include subdirectories of team submission directories
            # that have already been flattened.
            team_submission_dirs = [path for path in team_dir.iterdir() if not path.name == SUBMISSION_INFO_FILE_NAME]
            if len(team_submission_dirs) > 1:
                logging.warning(
                    f"There are multiple submissions for group '{team_dir.name}'!"
                )
            if len(team_submission_dirs) < 1:
                logging.warning(
                    f"The submission of group '{team_dir.name}' is empty!"
                )
            for team_submission_dir in team_submission_dirs:
                if team_submission_dir.is_dir():
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
        sub_dirs = [path for path in team_dir.iterdir() if not path.name == SUBMISSION_INFO_FILE_NAME]
        while len(sub_dirs) == 1 and sub_dirs[0].is_dir():
            move_content_and_delete(sub_dirs[0], team_dir)
            sub_dirs = [path for path in team_dir.iterdir() if not path.name == SUBMISSION_INFO_FILE_NAME]


def create_marks_file(_the_config: config.Config) -> None:
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
            exercise_dict = {f"exercise_{i}": "" for i in args.exercises}
    else:
        exercise_dict = ""

    marks_dict = {}
    for team_dir in sorted(get_relevant_team_dirs()):
        submission_info = get_submission_info(team_dir)
        team_key = get_team_key(submission_info)
        marks_dict.update({team_key: exercise_dict})

    with open(get_marks_file_path(_the_config), "w", encoding="utf-8") as marks_json:
        json.dump(marks_dict, marks_json, indent=4, ensure_ascii=False)


def create_feedback_directories(_the_config: config.Config) -> None:
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

        feedback_file_name = get_feedback_file_name(_the_config)
        dummy_pdf_name = feedback_file_name + ".pdf.todo"
        pathlib.Path(feedback_dir / dummy_pdf_name).touch(exist_ok=True)

        # Copy non-pdf submission files into feedback directory with added
        # prefix.
        for submission_file in team_dir.glob("*"):
            if submission_file.is_dir() or submission_file.suffix == ".pdf" \
                    or submission_file.name == SUBMISSION_INFO_FILE_NAME:
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
    from pypdf import PdfReader

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


def create_submission_info(_the_config: config.Config) -> None:
    """
    Write in each team directory a JSON file which contains the team,
    the ADAM ID of the team which ADAM sets anew with each exercise sheet,
    and if the tutor specified in the config has to mark this team. At
    the same time, the "Team " prefix is removed from directory names.
    """
    relevant_teams = get_relevant_teams(_the_config)
    adam_id_to_team = {}
    for team_dir in args.sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        team_id = team_dir.name.split(" ")[1]
        submission_dir = list(team_dir.iterdir())[0]
        submission_email = submission_dir.name.split("_")[-2]
        teams = [
            team
            for team in _the_config.teams
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
        is_relevant = False
        submission_info = {}
        if teams[0] in relevant_teams:
            is_relevant = True
        submission_info.update({"team": teams[0]})
        submission_info.update({"adam_id": team_id})
        submission_info.update({"relevant": is_relevant})
        with open(
                team_dir / SUBMISSION_INFO_FILE_NAME, "w", encoding="utf-8"
        ) as submission_info_file:
            json.dump(
                submission_info,
                submission_info_file,
                indent=4,
                ensure_ascii=False,
            )


def create_sheet_info_file(_the_config: config.Config) -> None:
    """
    Write information generated during the execution of the `init` command in a
    sheet info file. In particular the name of the exercise sheet as given by
    ADAM and which exercises should be marked if the marking mode is `exercise`.
    Later commands (e.g. `collect`, or `send`) are meant to load the information
    stored in this file into the `args` object and access it that way.
    """
    info_dict = {}
    info_dict.update({"adam_sheet_name": args.adam_sheet_name})
    if _the_config.marking_mode == "exercise":
        info_dict.update({"exercises": args.exercises})
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "w", encoding="utf-8"
    ) as sheet_info_file:
        json.dump(
            info_dict,
            sheet_info_file,
            indent=4,
            ensure_ascii=False,
            sort_keys=True,
        )
    # Immediately load the info back into args.
    load_sheet_info()


def print_missing_submissions(_the_config: config.Config) -> None:
    """
    Print all teams that are listed in the config file, but whose submission is
    not present in the zip downloaded from ADAM.
    """
    teams_who_submitted = []
    for team_dir in get_all_team_dirs():
        teams_who_submitted.append(get_submission_info(team_dir).get("team"))
    missing_teams = [
        team for team in _the_config.teams if team not in teams_who_submitted
    ]
    if missing_teams:
        logging.warning("There are no submissions for the following team(s):")
        for missing_team in missing_teams:
            print(f"* {team_to_string(missing_team)}")


def init(_the_config: config.Config) -> None:
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
    sheet_root_dir, adam_sheet_name = extract_adam_zip()
    add_to_args("sheet_root_dir", sheet_root_dir)
    add_to_args("adam_sheet_name", adam_sheet_name)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── Team 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    create_submission_info(_the_config)
    print_missing_submissions(_the_config)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345
    # .   ├── Muster_Hans_hans.muster@unibas.ch_000000
    # .   │   └── submission.pdf or submission.zip
    # .   └── submission.json
    rename_team_dirs()

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── Muster_Hans_hans.muster@unibas.ch_000000
    # .   │   └── submission.pdf or submission.zip
    # .   └── submission.json
    flatten_team_dirs()

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345_Muster-Meier-Mueller
    # .   ├── submission.pdf or submission.zip
    # .   └── submission.json
    unzip_internal_zips()

    # From here on, we need information about relevant teams.
    mark_irrelevant_team_dirs(_the_config)

    create_sheet_info_file(_the_config)

    if _the_config.use_marks_file:
        create_marks_file(_the_config)

    create_feedback_directories(_the_config)

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
    if args.xopp:
        generate_xopp_files()


# ============================= Config Processing ==============================


def add_to_args(key: str, value: Any) -> None:
    """
    This adds new entries into the args object created by argparse, we are in
    the process of facing out this hack. This function (and the title above)
    should be removed when it is not called anywhere anymore.
    """
    vars(args).update({key: value})


# =============================== Main Function ================================


def main():
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
            "run e.g. 'krummstab init -h' to print the help of a sub-command"
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

    global args
    args = parser.parse_args()

    if args.sub_command == "help":
        parser.print_help()
        sys.exit(0)

    # Process config files =====================================================
    _the_config = config.Config([args.config_shared, args.config_individual])

    # Execute subcommand =======================================================
    logging.info(f"Running command '{args.sub_command}'...")
    # This calls the function set as default in the parser.
    # For example, `func` is set to `init` if the subcommand is "init".
    args.func(_the_config)
    logging.info(f"Command '{args.sub_command}' terminated successfully. 🎉")


if __name__ == "__main__":
    main()
