import logging
import pathlib
import subprocess

from .. import config, sheets


def run_command_and_wait(command: str) -> None:
    """
    Executes a shell command and waits for it to finish.
    """
    logging.info(f"Running: {command}")
    process = subprocess.Popen(command, shell=True)
    process.wait()


def get_command_with_file(command: str, file: pathlib.Path) -> str:
    """
    Creates the complete shell command with the given program and file.
    """
    return f"{command} '{file}'"


def correct_xopp_files(command: str, sheet: sheets.Sheet) -> None:
    """
    Finds all .xopp files and opens them with Xournal++.
    """
    for submission in sheet.get_relevant_submissions():
        feedback_dir = submission.get_feedback_dir()
        xopp_files = [
            file for file in feedback_dir.rglob("*") if file.suffix == ".xopp"
        ]
        if xopp_files:
            for xopp_file in xopp_files:
                command_with_file = get_command_with_file(
                    command, xopp_file
                )
                run_command_and_wait(command_with_file)


def correct_pdf_files(command: str, sheet: sheets.Sheet):
    """
    Finds all feedback PDFs and opens them with the program specified in the
    config parameter 'command'.
    """
    for submission in sheet.get_relevant_submissions():
        feedback_dir = submission.get_feedback_dir()
        pdf_files = [
            file for file in feedback_dir.rglob("*") if file.suffix == ".pdf"
        ]
        if pdf_files:
            for pdf_file in pdf_files:
                command_with_file = get_command_with_file(
                    command, pdf_file
                )
                run_command_and_wait(command_with_file)


def correct(_the_config: config.Config, args) -> None:
    """
    Correct all submissions at once with a specific program such as
    Xournal++. Runs the program specified in the config parameter 'command'.
    """
    sheet = sheets.Sheet(args.sheet_root_dir)
    if len(_the_config.command) == 0:
        # default
        correct_xopp_files("xournalpp", sheet)
    elif "xournalpp" in _the_config.command:
        correct_xopp_files(_the_config.command, sheet)
    else:
        correct_pdf_files(_the_config.command, sheet)
