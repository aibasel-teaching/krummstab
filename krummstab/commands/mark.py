import logging
import pathlib
import subprocess
import sys

from .. import config, sheets, submissions, utils


def run_command_and_wait(command: list[str], dry_run: bool) -> None:
    """
    Executes a command and waits for it to finish. The dry_run flag is only
    useful for automated tests and is not meant to be set by users.
    """
    logging.info(f"Running {command}")
    try:
        timeout = 1.0 if dry_run else None
        subprocess.run(
            command, check=True, capture_output=True, timeout=timeout
        )
    except subprocess.TimeoutExpired:
        if dry_run:
            sys.exit(0)
        else:
            # The intention here is to pass the error on to be caught be the
            # handler below. I'm not sure whether that would actually happen,
            # but the subprocess should not really time out without 'dry_run'.
            raise
    except subprocess.SubprocessError as error:
        if error.stderr:
            logging.error(
                f"The marking command {command} failed with the following "
                "error output."
            )
            print(
                f"{utils.SEPARATOR_LINE}\n"
                f"{error.stderr.decode()}"
                f"{utils.SEPARATOR_LINE}"
            )
        else:
            logging.error(
                "The marking command {command} failed without any error output."
            )
        logging.critical("Aborting 'mark'.")

def get_command_with_file(command: list[str], file: pathlib.Path) -> list[str]:
    """
    Creates the complete command with the given program and file.
    """
    xopp_file = file
    pdf_file = file
    command = [arg.format(**locals()) for arg in command]
    return command


def mark_submission(
    submission: submissions.Submission,
    command: list[str],
    suffix_to_mark: str,
    dry_run: bool,
) -> None:
    feedback_dir = submission.get_feedback_dir()
    files_to_mark = [
        file
        for file in feedback_dir.rglob("*")
        if file.suffix == suffix_to_mark
    ]
    if not files_to_mark:
        logging.warning(
            f"No files to mark for team {submission.team.pretty_print()}."
        )
        return
    for file_to_mark in files_to_mark:
        run_command_and_wait(
            get_command_with_file(command, file_to_mark), dry_run
        )


def mark(_the_config: config.Config, args) -> None:
    """
    Mark all submissions at once with a specific program such as Xournal++.
    Runs the program specified in the config parameter 'marking_command'.
    """
    sheet = sheets.Sheet(args.sheet_root_dir)
    command = _the_config.marking_command
    # TODO: Remove this check when reworking config defaults (Issue #84).
    if not command:
        command = ["xournalpp", "{xopp_file}"]
    has_xopp = "{xopp_file}" in command
    has_pdf = "{pdf_file}" in command
    if has_xopp == has_pdf:
        logging.critical(
            "The config option 'marking_command' must contain either "
            "'{xopp_file}' or '{pdf_file}', but not both."
        )
    suffix_to_mark = ".xopp" if has_xopp else ".pdf"
    for submission in sheet.get_relevant_submissions():
        mark_submission(submission, command, suffix_to_mark, args.dry_run)
