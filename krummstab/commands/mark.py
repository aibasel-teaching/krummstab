import logging
import pathlib
import subprocess
import sys
from collections.abc import Iterator

from .. import config, sheets, strings, submissions, utils


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
        logging.error(
            f"The command {error.cmd} failed. If you are able to mark the "
            "submission as expected anyway, you can safely ignore this."
        )
        if error.stderr:
            logging.warning(f"Output to stderr:")
            print(
                f"{strings.SEPARATOR_LINE}\n"
                f"{error.stderr.decode()}"
                f"{strings.SEPARATOR_LINE}"
            )
        else:
            logging.info(f"No output to stderr.")
        if error.stdout:
            logging.warning(f"Output to stdout:")
            print(
                f"{strings.SEPARATOR_LINE}\n"
                f"{error.stdout.decode()}"
                f"{strings.SEPARATOR_LINE}"
            )
        else:
            logging.info(f"No output to stdout.")


def get_unmarked_submissions(
    _the_config: config.Config,
    sheet: sheets.Sheet,
    submissions_to_mark: Iterator[submissions.Submission],
) -> list[submissions.Submission]:
    marks_data = utils.read_json(sheet.get_marks_file_path(_the_config))
    if _the_config.points_per == "sheet":
        unmarked_submissions = [
            submission
            for submission in submissions_to_mark
            if marks_data[submission.team.get_team_key()] == ""
        ]
    else:
        assert _the_config.points_per == "exercise"
        unmarked_submissions = [
            submission
            for submission in submissions_to_mark
            if any(
                points == ""
                for exercise, points in marks_data[
                    submission.team.get_team_key()
                ].items()
            )
        ]
    return unmarked_submissions


def get_files_to_mark(
    submission: submissions.Submission,
    suffix_to_mark: str,
) -> list[str]:
    feedback_dir = submission.get_feedback_dir()
    files_to_mark = [
        str(file) for file in feedback_dir.rglob(f"*{suffix_to_mark}")
    ]
    if not files_to_mark:
        logging.warning(f"No files to mark for team {submission.team}.")
    return files_to_mark


def mark(_the_config: config.Config, args) -> None:
    """
    Mark all submissions at once with a specific program such as Xournal++.
    Runs the program specified in the config parameter 'marking_command'.
    """
    sheet = sheets.Sheet(args.sheet_root_dir)
    command_template = _the_config.marking_command
    # TODO: Remove this check when reworking config defaults (Issue #84).
    if not command_template:
        command_template = ["xournalpp", "{xopp_file}"]
    has_xopp = "{xopp_file}" in command_template
    has_pdf = "{pdf_file}" in command_template
    has_all_pdfs = "{all_pdf_files}" in command_template
    num_keywords_found = sum([has_xopp, has_pdf, has_all_pdfs])
    if num_keywords_found != 1:
        logging.critical(
            "The config option 'marking_command' must contain exactly one of "
            "the strings '{xopp_file}', '{pdf_file}', '{all_pdf_files}'. Make "
            "sure to include the braces around the keyword."
        )
    suffix_to_mark = ".xopp" if has_xopp else ".pdf"

    submissions_to_mark = list(sheet.get_relevant_submissions())
    if not submissions_to_mark:
        logging.info("There are no submissions that you have to mark.")
        return
    # By default, only open teams who do not already have marks entered in the
    # marks file. The idea of this is that the tutor does not need to look at
    # submissions again if she/he has already entered all marks. The '--force'
    # flag exists to circumvent this default behavior.
    if _the_config.use_marks_file and not args.force:
        submissions_to_mark = get_unmarked_submissions(
            _the_config, sheet, submissions_to_mark
        )

    submissions_total = len(submissions_to_mark)
    if not submissions_to_mark:
        logging.info(
            "You have already entered points for all the submissions you have "
            "to mark. Use the '-f' flag to open them anyway."
        )
        return
    logging.info(
        f"{submissions_total} out of "
        f"{len(list(sheet.get_relevant_submissions()))} "
        f"submission{'s'[:submissions_total ^ 1]} to mark."
    )

    if has_all_pdfs:
        all_files_to_mark = []
        for submission in submissions_to_mark:
            all_files_to_mark += get_files_to_mark(submission, suffix_to_mark)
        # Fill in the command_template by replacing the '{all_pdf_files}' string
        # by the list of files to be marked represented as strings
        i = command_template.index("{all_pdf_files}")
        del command_template[i]
        command_template[i:i] = all_files_to_mark
        if not all_files_to_mark:
            logging.critical(
                f"None of the submissions to be marked contain any PDF files."
            )
            return
        logging.info(f"Running {command_template}")
        run_command_and_wait(command_template, args.dry_run)
    else:
        assert has_xopp ^ has_pdf
        # Construct commands.
        commands = []
        for submission in submissions_to_mark:
            for file_to_mark in get_files_to_mark(submission, suffix_to_mark):
                command = [
                    s.format(xopp_file=file_to_mark, pdf_file=file_to_mark)
                    for s in command_template
                ]
                commands.append(command)
        num_commands = len(commands)
        # Run commands.
        for i, command in enumerate(commands, start=1):
            logging.info(
                f"({i:{len(str(num_commands))}d}"
                f"/{num_commands} file{'s'[:num_commands ^ 1]}) "
                f"Running {command}"
            )
            run_command_and_wait(command, args.dry_run)
