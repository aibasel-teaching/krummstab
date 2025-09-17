import logging
import pathlib
import shutil
from zipfile import ZipFile

from .. import config, sheets, utils


def combine(_the_config: config.Config, args) -> None:
    """
    Combine multiple share archives so that in the end we have one zip archive
    per team containing all feedback for that team.
    """
    # Prepare.
    sheet = sheets.Sheet(args.sheet_root_dir)

    share_archive_files = sheet.get_share_archive_files()
    instructions = (
        "Run `collect` to generate the share archive for your own feedback and"
        " save the share archives you received from the other tutors under"
        f" {sheet.root_dir}."
    )
    if len(list(share_archive_files)) == 0:
        logging.critical(
            f"No share archives exist in {sheet.root_dir}. " + instructions
        )
    if len(list(share_archive_files)) == 1:
        logging.warning(
            "Only a single share archive is being combined. " + instructions
        )

    # Create directory to store combined feedback in.
    combined_dir = sheet.get_combined_feedback_path()
    if combined_dir.exists() and combined_dir.is_dir():
        overwrite = utils.query_yes_no(
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
    for submission in sheet.get_relevant_submissions():
        combined_team_dir = combined_dir / submission.root_dir.name
        combined_team_dir.mkdir()

    # Structure at this point:
    # <sheet_root_dir>
    # └── feedback_combined
    #     ├── 12345_Muster-Meier-Mueller
    #     .

    teams_all = [
        submission.root_dir.name
        for submission in sheet.get_relevant_submissions()
    ]
    # Extract feedback files from share archives into their respective team
    # directories in the combined directory.
    for share_archive_file in sheet.get_share_archive_files():
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
            sheet.get_combined_feedback_file_name() + ".zip"
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
