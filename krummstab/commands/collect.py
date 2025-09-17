import gzip
import json
import logging
import pathlib
import shutil
import subprocess
from typing import Optional
from zipfile import ZipFile

from .. import config, sheets, submissions, strings, utils


def validate_marks_json(
    _the_config: config.Config, sheet: sheets.Sheet
) -> None:
    """
    Verify that all necessary marks are present in the MARK_FILE_NAME file and
    adhere to the granularity defined in the config file.
    """
    marks_json_file = sheet.get_marks_file_path(_the_config)
    if not marks_json_file.is_file():
        logging.critical(
            f"Missing points file in directory '{sheet.root_dir}'!"
        )
    marks = utils.read_json(marks_json_file)
    relevant_teams = []
    for submission in sheet.get_relevant_submissions():
        relevant_teams.append(submission.team.get_team_key())
    marked_teams = list(marks.keys())
    if sorted(relevant_teams) != sorted(marked_teams):
        logging.critical(
            "There is no 1-to-1 mapping between teams "
            "that need to be marked and entries in the "
            f"'{marks_json_file.name}' "
            "file! Make sure that it contains exactly one entry for every team "
            "that needs to be marked, and that the team keys "
            "have the correct format. The team key has to match the content of "
            "the submission.json file in the team directory and consist of the "
            "ADAM ID and the alphabetically sorted last names of all team "
            "members in the following format: ID_Last-Name1_Last-Name2"
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
        (float(mark) / _the_config.min_point_unit).is_integer()
        for mark in marks_list
        if mark != strings.PLAGIARISM
    ):
        logging.critical(
            f"'{marks_json_file.name}' contains marks that are more"
            " fine-grained than allowed! You may only award points in"
            f" '{_the_config.min_point_unit}' increments."
        )


def collect_feedback_files(
    submission: submissions.Submission,
    _the_config: config.Config,
    sheet: sheets.Sheet,
) -> None:
    """
    Take the contents of a {team_dir}/feedback directory and collect the files
    that actually contain feedback (e.g., no .xopp files). If there are
    multiple, add them to a zip archive and save it to
    {team_dir}/feedback_collected. If there is only a single pdf, copy it to
    {team_dir}/feedback_collected.
    """
    feedback_dir = submission.get_feedback_dir()
    collected_feedback_dir = submission.get_collected_feedback_dir()
    collected_feedback_zip_name = (
        sheet.get_feedback_file_name(_the_config) + ".zip"
    )
    # Error handling.
    if not feedback_dir.exists():
        logging.critical(
            f"Missing feedback directory for team {submission.root_dir.name}!"
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
        if file.is_file()
        and file.suffix not in _the_config.ignore_feedback_suffix
    ]
    # Ask for confirmation if the feedback directory contains hidden files that
    # are maybe not supposed to be part of the collected feedback.
    hidden_files = [f for f in feedback_files if utils.is_hidden_file(f.name)]
    for hidden_file in hidden_files:
        include_anyway = utils.query_yes_no(
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
            f"Feedback archive for team {submission.root_dir.name} is empty!"
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
            f"The feedback for {submission.root_dir.name} contains no PDF file!"
        )


def delete_collected_feedback_directories(sheet: sheets.Sheet) -> None:
    """
    Removes existing collected feedback directories. Does not care about
    non-existing ones.
    """
    for submission in sheet.get_relevant_submissions():
        collected_feedback_dir = submission.get_collected_feedback_dir()
        shutil.rmtree(collected_feedback_dir, ignore_errors=True)


def create_collected_feedback_directories(sheet: sheets.Sheet) -> None:
    """
    Create an empty directory in each relevant team directory. The collected
    feedback will be saved to these directories.
    """
    for submission in sheet.get_relevant_submissions():
        collected_feedback_dir = submission.get_collected_feedback_dir()
        assert not collected_feedback_dir.is_dir() or not any(
            collected_feedback_dir.iterdir()
        )
        collected_feedback_dir.mkdir(exist_ok=True)


def is_gzipped(filename: pathlib.Path) -> bool:
    """
    Checks if a file is gzipped.
    """
    try:
        with gzip.open(filename, "rb") as f:
            f.read(1)
        return True
    except OSError:
        return False


def export_xopp_files(sheet: sheets.Sheet) -> None:
    """
    Exports all xopp feedback files.
    """
    logging.info("Exporting .xopp files...")
    for submission in sheet.get_relevant_submissions():
        feedback_dir = submission.get_feedback_dir()
        xopp_files = [
            file for file in feedback_dir.rglob("*") if file.suffix == ".xopp"
        ]
        for xopp_file in xopp_files:
            if not is_gzipped(xopp_file):
                logging.critical(
                    f"File {xopp_file} has not been altered and saved by "
                    "Xournal++. It does not contain any feedback."
                )
            dest = xopp_file.with_suffix(".pdf")
            subprocess.run(["xournalpp", "-p", dest, xopp_file])
    logging.info("Done exporting .xopp files.")


def create_individual_marks_file(
    _the_config: config.Config, sheet: sheets.Sheet
) -> None:
    """
    Write a json file to add the marks per student.
    """
    team_marks = utils.read_json(sheet.get_marks_file_path(_the_config))
    student_marks = {}
    for submission in sheet.get_relevant_submissions():
        team_key = submission.team.get_team_key()
        for student in submission.team.members:
            student_key = student.email.lower()
            student_marks.update({student_key: team_marks.get(team_key)})
    file_content = {
        "tutor_name": _the_config.tutor_name,
        "adam_sheet_name": sheet.name,
        "marks": student_marks,
    }
    if (
        _the_config.points_per == "exercise"
        and _the_config.marking_mode == "exercise"
    ):
        file_content["exercises"] = sheet.exercises
    with open(
        sheet.get_individual_marks_file_path(_the_config),
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(file_content, file, indent=4, ensure_ascii=False)


def create_share_archive(
    overwrite: Optional[bool], sheet: sheets.Sheet
) -> None:
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
    # be written to sheet.json and then read in when initializing sheets.Sheet.
    assert sheet.exercises
    # Build share archive file name.
    share_archive_file = sheet.get_share_archive_file_path()
    if share_archive_file.is_file():
        # If the user has already chosen to overwrite when considering feedback
        # zips, then overwrite here too. Otherwise, ask here.
        # We should not be here if the user chose 'No' (making overwrite False)
        # before, so we catch this case.
        assert overwrite
        if overwrite is None:
            overwrite = utils.query_yes_no(
                (
                    "There already exists a share archive. Do you want to"
                    " overwrite it?"
                ),
                default=False,
            )
        if overwrite:
            share_archive_file.unlink(missing_ok=True)
        else:
            logging.critical(
                "Aborting 'combine' without overwriting existing share archive."
            )
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
        for submission in sheet.get_relevant_submissions():
            collected_feedback_file = submission.get_collected_feedback_path()
            sub_zip_name = f"{submission.root_dir.name}.zip"
            if collected_feedback_file.suffix == ".pdf":
                # Create a temporary zip file in the collected feedback
                # directory and add the single pdf.
                temp_zip_file = (
                    submission.get_collected_feedback_dir() / "temp_zip.zip"
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


def collect(_the_config: config.Config, args) -> None:
    """
    After marking is done, add feedback files to archives and print marks to be
    copy-pasted to shared point spreadsheet.
    """
    # Prepare.
    sheet = sheets.Sheet(args.sheet_root_dir)
    # Collect feedback.

    # Check if there is a collected feedback directory with files inside
    # already.
    collected_feedback_exists = any(
        (submission.get_collected_feedback_dir()).is_dir()
        and any((submission.get_collected_feedback_dir()).iterdir())
        for submission in sheet.get_relevant_submissions()
    )
    # Ask the user whether collected feedback should be overwritten in case it
    # exists already.
    overwrite = None
    if collected_feedback_exists:
        overwrite = utils.query_yes_no(
            (
                "There already exists collected feedback. Do you want to"
                " overwrite it?"
            ),
            default=False,
        )
        if overwrite:
            delete_collected_feedback_directories(sheet)
        else:
            logging.critical(
                "Aborting 'collect' without overwriting existing collected feedback."
            )
    if _the_config.xopp:
        export_xopp_files(sheet)
    create_collected_feedback_directories(sheet)
    for submission in sheet.get_relevant_submissions():
        collect_feedback_files(submission, _the_config, sheet)
    if _the_config.marking_mode == "exercise":
        create_share_archive(overwrite, sheet)
    if _the_config.use_marks_file:
        validate_marks_json(_the_config, sheet)
        create_individual_marks_file(_the_config, sheet)
