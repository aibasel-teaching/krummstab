import logging
import pathlib
import subprocess

from .. import config, sheets


def run_command_and_wait(command: list[str]) -> None:
    """
    Executes a command and waits for it to finish.
    """
    logging.info(f"Running: {command}")
    subprocess.run(command)


def get_command_with_file(command: list[str], file: pathlib.Path) -> list[str]:
    """
    Creates the complete command with the given program and file.
    """
    xopp_file = file
    pdf_file = file
    command = [arg.format(**locals()) for arg in command]
    return command


def correct_xopp_files(command: list[str], sheet: sheets.Sheet) -> None:
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


def correct_pdf_files(command: list[str], sheet: sheets.Sheet) -> None:
    """
    Finds all feedback PDFs and opens them with the program specified in the
    config parameter 'marking_command'.
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
    Correct all submissions at once with a specific program such as Xournal++.
    Runs the program specified in the config parameter 'marking_command'.
    """
    sheet = sheets.Sheet(args.sheet_root_dir)
    cmd = _the_config.marking_command
    if not cmd:
        # default
        correct_xopp_files(["xournalpp", "{xopp_file}"], sheet)
    else:
        has_xopp = "{xopp_file}" in cmd
        has_pdf = "{pdf_file}" in cmd
        if has_xopp == has_pdf:
            logging.critical("The config must contain either '{xopp_file}' or "
                             "'{pdf_file}' in 'marking_command', but not both.")
        elif has_xopp:
            correct_xopp_files(cmd, sheet)
        else:
            correct_pdf_files(cmd, sheet)
