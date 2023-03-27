#!/usr/bin/env python3
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
import textwrap
import typing

from collections.abc import Iterator  # For typing.
from zipfile import ZipFile

# Needed for email stuff.
import smtplib
import ssl
from email.message import EmailMessage
from getpass import getpass

Student = tuple[str, str, str]
Team = list[Student]
Data = dict[str, typing.Any]

DEFAULT_CONFIG_FILE = "config.json"
DO_NOT_MARK_PREFIX = "DO_NOT_MARK_"
FEEDBACK_DIR_NAME = "feedback"
FEEDBACK_ARCHIVE_NAME = "feedback.zip"
FEEDBACK_ARCHIVE_PATH = pathlib.Path(FEEDBACK_DIR_NAME, FEEDBACK_ARCHIVE_NAME)
SHEET_INFO_FILE_NAME = ".sheet_info"
MARKS_FILE_NAME = "points.json"

# Might be necessary to make colored output work on Windows.
os.system("")


# ============================= Utility Functions ==============================
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


def throw_error(text: str) -> None:
    """
    Print an error message and exit.
    """
    print("\033[0;31m[Error]\033[0m " + text)
    sys.exit(1)


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


def get_relevant_team_dirs() -> Iterator[pathlib.Path]:
    """
    Return the team directories of the teams whose submission has to be
    corrected by the tutor running the script.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if team_dir.is_dir() and not DO_NOT_MARK_PREFIX in team_dir.name:
            yield team_dir


def load_sheet_info() -> None:
    """
    Load the information stored in the sheet info file into the args object.
    """
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "r"
    ) as sheet_info_file:
        sheet_info = json.load(sheet_info_file)
    for key, value in sheet_info.items():
        add_to_args(key, value)


def team_to_string(team: Team) -> str:
    """
    Concatenate the last names of students to get a pretty-ish string
    representation of teams.
    """
    return "_".join(sorted([student[1].replace(" ", "-") for student in team]))


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
    # Only keep one name per entry."Hans Jakob" becomes "Hans"
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
            team_dir / FEEDBACK_ARCHIVE_PATH,
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
    with open(marks_json_file, "r") as marks_file:
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


def archive_feedback(team_dir: pathlib.Path) -> None:
    """
    Take the contents of a '{team_dir}/feedback' directory and add them to the
    file '{team_dir}/feedback/feedback.zip'.
    """
    feedback_dir = team_dir / FEEDBACK_DIR_NAME
    feedback_zip = team_dir / FEEDBACK_ARCHIVE_PATH
    # Error handling.
    if not feedback_dir.exists():
        throw_error(f"Missing feedback directory for team {team_dir.name}!")
    content = list(feedback_dir.iterdir())
    if any(".todo" in file_or_dir.name for file_or_dir in content):
        throw_error(
            f"Feedback for {team_dir.name} contains placeholder TODO file!"
        )

    if feedback_zip.is_file():
        throw_error(
            f"A zipped feedback file already exists, "
            f"run the 'collect' command with the flag '-r' "
            f"to overwrite existing feedback archives."
        )
    # Zip up content of the feedback directory. By default, files with
    # extensions .zip and .xopp are not added to the feedback archive.
    # .zip have to be ignored because the feedback_zip would contain itself
    # otherwise. (Constructing it in a tmp directory could fix this.)
    feedback_contains_pdf = False
    with ZipFile(feedback_zip, "w") as zip_file:
        to_zip = [
            file
            for file in feedback_dir.rglob("*")
            if file.is_file()
            and not file.suffix
            in [".zip", ".xopp"] + args.ignore_feedback_suffix
        ]
        if len(to_zip) <= 0:
            throw_error(
                f"Feedback archive for team {team_dir.name} is empty! "
                "Did you forget the '-x' flag to export .xopp files?"
            )
        for file_to_zip in to_zip:
            if file_to_zip.suffix == ".pdf":
                feedback_contains_pdf = True
            zip_file.write(
                file_to_zip, arcname=file_to_zip.relative_to(feedback_dir)
            )
    if not feedback_contains_pdf:
        print_warning(f"The feedback for {team_dir.name} contains no PDF file!")


def delete_feedback_archives() -> None:
    """
    Removes existing feedback archives. Does not care about non-existing
    ones.
    """
    for team_dir in get_relevant_team_dirs():
        feedback_zip = team_dir / FEEDBACK_ARCHIVE_PATH
        feedback_zip.unlink(missing_ok=True)


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
    with open(marks_json_file, "r") as marks_file:
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


def collect() -> None:
    """
    After marking is done, add feedback files to archives and print marks to be
    copy-pasted to shared point spreadsheet.
    """
    # Prepare.
    verify_sheet_root_dir()
    load_sheet_info()
    # Collect feedback.
    if args.replace:
        delete_feedback_archives()
    if args.xopp:
        export_xopp_files()
    for team_dir in get_relevant_team_dirs():
        archive_feedback(team_dir)
    if args.marking_mode == "exercise":
        throw_error(
            "Collecting for marking mode 'exercise' is not implemented!"
        )
        # TODO: Implement. We probably want to create a single archive with all
        # feedback here, so that it can be shared with the other tutors. In the
        # end, one of the tutors has to run a command that combines all
        # feedback.
    validate_marks_json()
    print_marks()


# ============================== Init Sub-Command ==============================
def extract_adam_zip() -> tuple[pathlib.Path, str]:
    """
    Unzips the given ADAM zip file and renames the directory to *target* if one
    is given. This is done stupidly right now, it would be better to extract to
    a temporary folder and then move to once to the right location.
    """
    if args.adam_zip_path.is_file():
        # Unzip to the directory within the zip file.
        # Should be the name of the exercise sheet, for example "Exercise Sheet 2".
        with ZipFile(args.adam_zip_path, mode="r") as zip_file:
            zip_content = zip_file.namelist()
            sheet_root_dir = pathlib.Path(zip_content[0])
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
    move_content_and_delete(sheet_root_dir / "Abgaben", sheet_root_dir)
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


def ensure_single_submission_per_team() -> None:
    """
    There can be multiple directories within a "Team 00000" directory. This
    (probably) happens when multiple members of the team upload solutions, but
    I think only the directory of the most recent submission remains non-empty,
    so we delete the empty ones and throw an error if multiple remain.
    """
    for team_dir in args.sheet_root_dir.iterdir():
        if not team_dir.is_dir():
            continue
        for team_submission_dir in team_dir.iterdir():
            if len(list(team_submission_dir.iterdir())) == 0:
                team_submission_dir.rmdir()
        if len(list(team_dir.iterdir())) > 1:
            # If there are multiple non-empty submissions, we have a problem.
            throw_error(
                f"There are multiple submissions for group '{team_dir.name}'!"
            )


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
        team = [
            team
            for team in args.teams
            if any(submission_email in student for student in team)
        ]
        if len(team) == 0:
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
        assert len(team) == 1
        adam_id_to_team.update({team_id: team[0]})
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
        shuffled_teams = [team for _, team in args.team_dir_to_team.items()]
        seed = int(
            hashlib.sha256(args.adam_sheet_name.encode("utf-8")).hexdigest(), 16
        )
        random.Random(seed).shuffle(shuffled_teams)
        chunks = [shuffled_teams[i::num_tutors] for i in range(num_tutors)]
        assert len(chunks) == num_tutors
        assert all(
            abs(len(this) - len(that)) <= 1
            for this in chunks
            for that in chunks
        )
        return chunks[args.tutor_list.index(args.tutor_name)]
    elif args.marking_mode == "exercise":
        return args.teams
    else:
        throw_error(f"Unsupported marking mode {args.marking_mode}!")
        return []


def rename_team_dirs(adam_id_to_team: dict[str, Team]) -> None:
    """
    The team directories are renamed to: team_id_LastName1-LastName2
    The team ID can be helpful to identify a team on the ADAM web interface.
    Additionally the directory structure is flattened.
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
        sub_dirs = list(team_dir.iterdir())
        assert len(sub_dirs) == 1
        move_content_and_delete(sub_dirs[0], team_dir)


def unzip_internal_zips() -> None:
    """
    If multiple files are uploaded to ADAM, the submission becomes a single zip
    file. Here we extract this zip. I'm not sure if nested zip files are also
    extracted. Additionally we flatten the directory by one level if the zip
    contains only a single directory. Doing so recursively would be nicer.
    """
    for team_dir in args.sheet_root_dir.iterdir():
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
    exercise_dict: typing.Union[str, dict[str, str]] = ""
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

    with open(args.sheet_root_dir / MARKS_FILE_NAME, "w") as marks_json:
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

        prefix = "feedback_"
        if args.marking_mode == "exercise":
            team_id = team_dir.name.split("_")[0]
            prefix = team_id + "_" + prefix
            prefix += args.tutor_name + "_"
            prefix += "ex"
            for exercise in args.exercises:
                prefix += str(exercise) + "_"
        elif args.marking_mode == "random":
            prefix += args.tutor_name + "_"
        elif args.marking_mode == "static":
            pass

        dummy_pdf_name = prefix[:-1] + ".pdf.todo"
        pathlib.Path(feedback_dir / dummy_pdf_name).touch(exist_ok=True)

        # Copy non-pdf submission files into feedback directory with added
        # prefix.
        for submission_file in team_dir.glob("*"):
            if submission_file.is_dir() or submission_file.suffix == ".pdf":
                continue
            feedback_file_name = prefix + submission_file.name
            shutil.copy(submission_file, feedback_dir / feedback_file_name)


def generate_xopp_files() -> None:
    """
    Generate xopp files in the feedback directories that point to the single PDF
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
                f"Skipping {team_dir.name}: No or multiple PDF files."
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
            print_warning(f"Skipping {team_dir.name}: xopp file exists.")
            continue
        xopp_file = open(xopp_path, "w")
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


def create_sheet_info_file(
    adam_id_to_team: dict[str, Team], adam_sheet_name: str
) -> None:
    """
    Write information generated during the execution of the 'init' command in a
    sheet info file. In particular a mapping from team directory names to teams
    and the name of the exercise sheet as given by ADAM. The latter is used as
    the seed to make random assignment of submissions to tutors consistent
    between tutors, but still vary from sheet to sheet. Later commands (e.g.
    'collect', or 'send') are meant to load the information stored in this file
    into the 'args' object and access it that way.
    """
    info_dict: dict[str, typing.Union[str, dict[str, Team]]] = {}
    # Build the dict from team directory names to teams.
    team_dir_to_team = {}
    for team_dir in args.sheet_root_dir.iterdir():
        if team_dir.is_file():
            continue
        # Get ADAM ID from directory name.
        adam_id_match = re.search(r"\d+", team_dir.name)
        assert adam_id_match
        adam_id = adam_id_match.group()
        team = adam_id_to_team[adam_id]
        team_dir_to_team.update({team_dir.name: team})
    info_dict.update({"team_dir_to_team": team_dir_to_team})
    info_dict.update({"adam_sheet_name": adam_sheet_name})
    with open(
        args.sheet_root_dir / SHEET_INFO_FILE_NAME, "w"
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
            print_info(f"    {team_to_string(missing_team)}", True)


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

    # Structure at this point:
    # <sheet_root_dir>
    # ├── Team 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    ensure_single_submission_per_team()
    adam_id_to_team = get_adam_id_to_team_dict()
    print_missing_submissions(adam_id_to_team)

    # Structure at this point:
    # <sheet_root_dir>
    # ├── 12345
    # .   └── Muster_Hans_hans.muster@unibas.ch_000000
    # .       └── submission.pdf or submission.zip
    rename_team_dirs(adam_id_to_team)

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
    create_sheet_info_file(adam_id_to_team, adam_sheet_name)
    # then rename the irrelevant team directories...
    mark_irrelevant_team_dirs()
    # and finally recreate the sheet info file to reflect the final team
    # directory names.
    create_sheet_info_file(adam_id_to_team, adam_sheet_name)

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
        throw_error(
            "There are duplicate student emails in the config file!"
        )
    teams.sort()


def process_static_config(data: Data) -> None:
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


def process_dynamic_config(data: Data) -> None:
    """
    Extract and check the config values necessary for the dynamic correction
    marking modes, i.e., random and exercise.
    """
    tutor_list = data["tutor_list"]
    assert type(tutor_list) is list
    assert all(type(tutor) is str for tutor in tutor_list)
    assert args.tutor_name in tutor_list
    add_to_args("tutor_list", sorted(tutor_list))

    teams = data["teams"]
    validate_teams(teams)
    add_to_args("teams", teams)


def process_general_config(data: Data) -> None:
    """
    Extract and check config values that are necessary in all marking modes.
    """
    # Individual settings
    tutor_name = data["your_name"]
    assert type(tutor_name) is str
    add_to_args("tutor_name", tutor_name)

    # Use `get` because this config setting is optional.
    ignore_feedback_suffix = data.get("ignore_feedback_suffix", [])
    assert type(ignore_feedback_suffix) is list
    assert all(
        type(suffix) is str and suffix[0] == "."
        for suffix in ignore_feedback_suffix
    )
    add_to_args("ignore_feedback_suffix", ignore_feedback_suffix)

    # General settings
    lecture_title = data["lecture_title"]
    assert lecture_title and type(lecture_title) is str
    add_to_args("lecture_title", lecture_title)

    marking_mode = data["marking_mode"]
    assert marking_mode in ["static", "random", "exercise"]
    add_to_args("marking_mode", marking_mode)

    max_team_size = data["max_team_size"]
    assert type(max_team_size) is int and max_team_size > 0
    add_to_args("max_team_size", max_team_size)

    points_per = data["points_per"]
    assert type(points_per) is str
    assert points_per in ["sheet", "exercise"]
    add_to_args("points_per", points_per)

    min_point_unit = data["min_point_unit"]
    assert type(min_point_unit) is float or type(min_point_unit) is int
    assert min_point_unit > 0
    add_to_args("min_point_unit", min_point_unit)

    # Email settings, currently all optional because not fully functional
    tutor_email = data.get("your_email", "")
    # assert type(tutor_email) is str and is_email(tutor_email)
    add_to_args("tutor_email", tutor_email)

    feedback_email_cc = data.get("feedback_email_cc", [])
    assert type(feedback_email_cc) is list
    # assert all(
    #    type(email) is str and is_email(email) for email in feedback_email_cc
    # )
    add_to_args("feedback_email_cc", feedback_email_cc)

    smtp_url = data.get("smtp_url", "smtp-ext.unibas.ch")
    assert type(smtp_url) is str
    add_to_args("smtp_url", smtp_url)

    smtp_port = data.get("smtp_port", 587)
    assert type(smtp_port) is int
    add_to_args("smtp_port", smtp_port)

    smtp_user = data.get("smtp_user", "")
    assert type(smtp_user) is str
    add_to_args("smtp_user", smtp_user)


def add_to_args(key: str, value: typing.Any) -> None:
    """
    Settings defined in the config file are parsed, checked, and then added to
    the args object. Similar with information that is calculated using things
    defined in the config file. After parsing, all necessary information should
    be contained in the args object.
    """
    vars(args).update({key: value})


# =============================== Main Function ================================
if __name__ == "__main__":
    """
    This script uses the following, potentially ambiguous, terminology:
    team     - set of students that submit together, as defined on ADAM
    class    - set of teams that are in the same exercise class
    relevant - a team whose submission has to be marked by the tutor running the
               script is considered 'relevant'
    to mark  - grading/correcting a sheet; giving feedback
    marks    - points awarded for a sheet or exercise
    """
    parser = argparse.ArgumentParser(description="")
    # Main command arguments ---------------------------------------------------
    parser.add_argument(
        "-c",
        "--config",
        default=DEFAULT_CONFIG_FILE,
        type=pathlib.Path,
        help=(
            "path to the json config file containing student/email list and"
            " other settings"
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
        help="collect feedback files",
    )
    parser_collect.add_argument(
        "-r",
        "--replace",
        action=argparse.BooleanOptionalAction,
        default=False,
        help="replace existing feedback archives",
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

    # Get settings from config.
    print_info("Processing config...")
    with open(args.config, "r") as config_file:
        data = json.load(config_file)
    # We currently plan to support the following marking modes.
    # static:   Every tutor corrects the submissions of the teams assigned to
    #           that tutor. These will usually be the teams in that tutors
    #           exercise class.
    # random:   Every tutor corrects some submissions which are assigned
    #           randomly with every sheet.
    # exercise: Every tutor corrects some exercise(s) on all sheets.
    process_general_config(data)

    if args.marking_mode == "static":
        process_static_config(data)
    else:
        process_dynamic_config(data)
    print_info("Processed config successfully.")

    # Call the function associated with the selected sub-command, e.g.,
    # init(args) or xopp(args).
    print_info(f"Running command '{args.sub_command}'...")
    args.func()
    print_info(
        f"Command '{args.sub_command}' terminated successfully."
        " \033[0;32m:)\033[0m"
    )

# TODO: Ignore/delete Apple store files. e.g. folder "__MACOSX"
# TODO: Before collecting, check whether all feedback consists of a single PDF
#       file. If yes, don't zip the feedback.

# Only in Python 3.7+ are dicts order preserving, using older Pythons may cause
# the random assignments to not match up.
