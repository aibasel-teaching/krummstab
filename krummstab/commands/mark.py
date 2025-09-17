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
    submission_num: int,
    submissions_total: int,
    dry_run: bool,
) -> None:
    feedback_dir = submission.get_feedback_dir()
    files_to_mark = [
        file
        for file in feedback_dir.rglob("*")
        if file.suffix == suffix_to_mark
    ]
    if not files_to_mark:
        logging.warning(f"No files to mark for team {submission.team}.")
        return
    for file_to_mark in files_to_mark:
        command_with_file = get_command_with_file(command, file_to_mark)
        logging.info(
            f"({submission_num:{len(str(submissions_total))}d}/{submissions_total}) "
            f"Running {command_with_file}"
        )
        run_command_and_wait(command_with_file, dry_run)


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

    submissions_to_mark = list(sheet.get_relevant_submissions())
    # By default, only open teams who do not already have marks entered in the
    # marks file. The idea of this is that the tutor does not need to look at
    # submissions again if she/he has already entered all marks. The '--force'
    # flag exists to circumvent this default behavior.
    # TODO: Generalize this to marking_mode == "exercise" and points_per ==
    #       "exercise".
    if (
        _the_config.use_marks_file
        and _the_config.marking_mode == "static"
        and _the_config.points_per == "sheet"
        and not args.force
    ):
        marks_data = utils.read_json(sheet.get_marks_file_path(_the_config))
        submissions_to_mark = [
            submission
            for submission in submissions_to_mark
            if marks_data[submission.team.get_team_key()] == ""
        ]

    submissions_total = len(submissions_to_mark)
    logging.info(
        f"{submissions_total} out of "
        f"{len(list(sheet.get_relevant_submissions()))} "
        f"submission{'s'[:submissions_total^1]} to mark."
    )

    for submission_num, submission in enumerate(submissions_to_mark, start=1):
        mark_submission(
            submission,
            command,
            suffix_to_mark,
            submission_num,
            submissions_total,
            args.dry_run,
        )
