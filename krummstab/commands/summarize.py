import json
import logging
from collections import defaultdict
from pathlib import Path
import xlsxwriter
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range_abs
from xlsxwriter.worksheet import Worksheet

from .. import config
from ..teams import *

BOLD = {'bold': True}
BORDER = {'border': 1}
BORDER_LEFT = {'left': 1}
BORDER_TOP_BOTTOM = {'top': 1, 'bottom': 1}
GRAY = {'bg_color': '#E2E2E2'}
GREEN = {'bg_color': '#9BE189'}
YELLOW = {'bg_color': '#FFE699'}
RED = {'bg_color': '#EE7868'}
PLAGIARISM_RED = {'bg_color': 'red'}
PERCENT = {'num_format': '0%'}
TEXT_WRAP = {'text_wrap': True}

PERCENTAGE_TO_PASS_CELL = '$B$4'
IMPROVE_AVG_RED_CELL = '$E$3'
IMPROVE_AVG_GREEN_CELL = '$E$1'

PLAGIARISM = "Plagiarism"


def add_legend(workbook: Workbook, worksheet: Worksheet):
    """
    Adds a legend that explains the conditional formatting.
    """
    worksheet.set_row(0, 28)
    worksheet.set_row(1, 28)
    worksheet.set_row(2, 28)
    worksheet.set_row(3, 28)
    worksheet.merge_range(
        0, 0, 0, 2, "Pass", workbook.add_format(GREEN | BORDER)
    )
    worksheet.merge_range(
        1, 0, 1, 2, "Fail", workbook.add_format(RED | BORDER)
    )
    worksheet.merge_range(
        2, 0, 2, 1, "2 x Plagiarism",
        workbook.add_format(PLAGIARISM_RED | BORDER)
    )
    worksheet.write(
        0, 3,
        "Does not need to improve average\nif percentage is lower "
        "than:",
        workbook.add_format(GREEN | TEXT_WRAP)
    )
    worksheet.write_number(
        0, 4,
        -0.05, workbook.add_format(GREEN | PERCENT)
    )
    worksheet.write(
        1, 3,
        "Should improve average\nif percentage is between:",
        workbook.add_format(YELLOW | TEXT_WRAP)
    )
    worksheet.write_formula(
        1, 4,
        f'TEXT({IMPROVE_AVG_GREEN_CELL},"0%")'
        f'& " to " & TEXT({IMPROVE_AVG_RED_CELL},"0%")',
        workbook.add_format(YELLOW | PERCENT)
    )
    worksheet.write(
        2, 3,
        "Has to improve average\nby at least the following percentage:",
        workbook.add_format(RED | TEXT_WRAP)
    )
    worksheet.write_number(
        2, 4,
        0.10, workbook.add_format(RED | PERCENT)
    )
    worksheet.write(
        3, 0, "Percentage of points\nrequired to pass:",
        workbook.add_format(TEXT_WRAP)
    )
    worksheet.write_number(
        3, 1, 0.5, workbook.add_format(PERCENT)
    )


def add_pass_or_fail_conditional_formatting(workbook: Workbook, worksheet: Worksheet, row: int,
                                            graded_sheet_names, all_sheet_names):
    """
    Colors the name red if not enough points can be collected.
    Colors the name green when enough points have been collected.
    """
    num_legend_rows = 4
    total_points_all_sheets_range = xl_range_abs(
        1 + num_legend_rows, 5, 1 + num_legend_rows, 5 + len(all_sheet_names) - 1
    )
    possible_points_range = xl_range_abs(
        1 + num_legend_rows, 5 + len(graded_sheet_names), 1 + num_legend_rows, 5 + len(all_sheet_names) - 1
    )
    worksheet.conditional_format(row, 0, row, 2, {
        'type': 'formula',
        'criteria': f"=(SUM(INDIRECT(ADDRESS(ROW(),3)), SUM({possible_points_range})))"
                    f" < (SUM({total_points_all_sheets_range}) * {PERCENTAGE_TO_PASS_CELL})",
        'format': workbook.add_format(RED)
    })
    worksheet.conditional_format(row, 0, row, 2, {
        'type': 'formula',
        'criteria': f"INDIRECT(ADDRESS(ROW(),3)) >= (SUM("
                    f"{total_points_all_sheets_range}) * {PERCENTAGE_TO_PASS_CELL})",
        'format': workbook.add_format(GREEN)
    })


def add_plagiarism_conditional_formatting(workbook: Workbook, worksheet: Worksheet,
                                          row: int, all_sheet_names):
    """
    Colors the name bright red if there are two plagiarisms.
    """
    student_marks_range = xl_range_abs(row, 5, row, 5 + len(all_sheet_names) - 1)
    worksheet.conditional_format(row, 0, row, 1, {
        'type': 'formula',
        'criteria': f'=COUNTIF({student_marks_range},"{PLAGIARISM}") >= 2',
        'format': workbook.add_format(PLAGIARISM_RED)
    })


def add_average_conditional_formatting(workbook: Workbook, worksheet: Worksheet, row: int):
    """
    Adds the conditional formatting for the current average and the percentage
    by which the average must be improved.
    """
    worksheet.conditional_format(row, 3, row, 4, {
        'type': 'formula',
        'criteria': f'=INDIRECT(ADDRESS(ROW(),5)) <= {IMPROVE_AVG_GREEN_CELL}',
        'format': workbook.add_format(GREEN)
    })
    worksheet.conditional_format(row, 3, row, 4, {
        'type': 'formula',
        'criteria': f'=INDIRECT(ADDRESS(ROW(),5)) <= {IMPROVE_AVG_RED_CELL}',
        'format': workbook.add_format(YELLOW)
    })
    worksheet.conditional_format(row, 3, row, 4, {
        'type': 'formula',
        'criteria': f'=INDIRECT(ADDRESS(ROW(),5)) > {IMPROVE_AVG_RED_CELL} ',
        'format': workbook.add_format(RED)
    })


def add_student_average(workbook: Workbook, worksheet: Worksheet, row: int, graded_sheet_names, all_sheet_names):
    """
    Calculates a student's current weighted average and the percentage
    by which the average must be improved to pass with the remaining
    exercise sheets.
    """
    num_legend_rows = 4
    student_marks_range = xl_range_abs(row, 5, row, 5 + len(all_sheet_names) - 1)
    total_points_all_sheets_range = xl_range_abs(
        1 + num_legend_rows, 5, 1 + num_legend_rows, 5 + len(all_sheet_names) - 1
    )
    possible_points_range = xl_range_abs(
        1 + num_legend_rows, 5 + len(graded_sheet_names), 1 + num_legend_rows, 5 + len(all_sheet_names) - 1
    )
    worksheet.write_formula(row, 3,
                            f'=IFERROR(SUMPRODUCT(ISNUMBER({student_marks_range})*1,{student_marks_range},'
                            f' {total_points_all_sheets_range}) / SUMPRODUCT(ISNUMBER({student_marks_range})*1,'
                            f' {total_points_all_sheets_range}),"")',
                            workbook.add_format(BORDER))
    worksheet.write_formula(row, 4,
                            f'=IFERROR((SUM('
                            f'{total_points_all_sheets_range})*{PERCENTAGE_TO_PASS_CELL}'
                            f' - (INDIRECT(ADDRESS(ROW(),3))))/(COLUMNS({possible_points_range})'
                            f'*(INDIRECT(ADDRESS(ROW(),4))))-1,"")',
                            workbook.add_format(BORDER | PERCENT))


def write_mark(worksheet, row, col, mark) -> None:
    """
    Writes a mark in the cell specified by row and column.
    """
    if mark is None:
        worksheet.write_blank(row, col, None)
    elif isinstance(mark, str) and not mark.replace(".", "", 1).isdigit():
        worksheet.write(row, col, mark)
    else:
        worksheet.write_number(row, col, float(mark))


def add_student_marks_worksheet_points_per_sheet(workbook: Workbook, worksheet: Worksheet,
                                                 _the_config: config.Config, row, email,
                                                 students_marks, all_sheet_names, graded_sheet_names,
                                                 points_per_sheet_cell_addresses):
    """
    Writes the points per sheet of a student in the worksheet for the sheets and calculates
    the sum to get the total points of the student.
    If points are given per exercise: Uses cell addresses from the points per
    exercise worksheet that contain the calculated points per sheet.
    """
    if _the_config.points_per == 'exercise':
        for col, cell_address in enumerate(points_per_sheet_cell_addresses[email], start=5):
            worksheet.write(row, col, f"={cell_address}")
    else:
        for col, sheet_name in enumerate(list(all_sheet_names)[:len(graded_sheet_names)], start=5):
            mark = students_marks[email].get(sheet_name)
            write_mark(worksheet, row, col, mark)
    student_marks_range = xl_range_abs(row, 5, row, 5 + len(all_sheet_names) - 1)
    worksheet.write_formula(row, 2, f"=SUM({student_marks_range})", workbook.add_format(BORDER))


def add_student_marks_worksheet_points_per_exercise(workbook: Workbook, worksheet: Worksheet,
                                                    row, email, students_marks, all_sheet_names,
                                                    graded_sheet_names) -> list[str]:
    """
    Writes the points per exercise of a student in the worksheet for the exercises, calculates
    the sum to get the points per sheet and handles plagiarism. Returns a list of
    cell addresses containing the calculated points per sheet.
    """
    col = 2
    points_per_sheet_cell_addresses = []
    for sheet_name in list(all_sheet_names)[:len(graded_sheet_names)]:
        student_marks_range = xl_range_abs(row, col + 1, row, col + len(graded_sheet_names[sheet_name]))
        worksheet.write_dynamic_array_formula(
            row, col, row, col,
            f'=IF(COUNTIF({student_marks_range},"{PLAGIARISM}") > 0,"{PLAGIARISM}",'
            f'IF(AND(NOT(ISNUMBER({student_marks_range}))),"",SUM({student_marks_range})))',
            workbook.add_format(BORDER_LEFT)
        )
        cell_address = xl_rowcol_to_cell(row, col)
        points_per_sheet_cell_addresses.append(f"'Points Per Exercise'!{cell_address}")
        col += 1
        for exercise in sorted(graded_sheet_names[sheet_name]):
            mark = students_marks[email][sheet_name].get(exercise)
            write_mark(worksheet, row, col, mark)
            col += 1
    return points_per_sheet_cell_addresses


def add_student_name(workbook, worksheet, row, name) -> None:
    """
    Writes the specified name of the student in the first two columns
    for the specified row. Changes the background color of the row
    alternately to gray for better readability.
    """
    first_name, last_name = name
    if row % 2 == 0:
        worksheet.write(row, 0, first_name, workbook.add_format(BORDER_TOP_BOTTOM | GRAY))
        worksheet.write(row, 1, last_name, workbook.add_format(BORDER_TOP_BOTTOM | GRAY))
        worksheet.set_row(row, None, workbook.add_format(GRAY))
    else:
        worksheet.write(row, 0, first_name, workbook.add_format(BORDER_TOP_BOTTOM))
        worksheet.write(row, 1, last_name, workbook.add_format(BORDER_TOP_BOTTOM))


def create_worksheet_points_per_exercise(workbook: Workbook, email_to_name, students_marks, all_sheet_names,
                                         graded_sheet_names) -> defaultdict[str, list[str]]:
    """
    Creates an Excel worksheet that contains the points per exercise. Returns a dict with the
    students' email addresses and the cell addresses that contain the calculated points per sheet.
    """
    worksheet = workbook.add_worksheet("Points Per Exercise")
    worksheet.write(0, 0, "First Name", workbook.add_format(BOLD))
    worksheet.write(0, 1, "Last Name", workbook.add_format(BOLD))
    col = 2
    for sheet_name in all_sheet_names:
        worksheet.write(0, col, sheet_name, workbook.add_format(BOLD | BORDER_LEFT))
        if sheet_name in graded_sheet_names:
            col += 1
            for exercise in sorted(graded_sheet_names[sheet_name]):
                worksheet.write(0, col, exercise.replace("exercise", "task"), workbook.add_format(BOLD))
                col += 1
        else:
            col += 1
    all_points_per_sheet_cell_addresses = defaultdict(list)
    student_start_row = 1
    sorted_emails = sorted(email_to_name.keys(), key=lambda e: (email_to_name[e][0], email_to_name[e][1]))
    for row, email in enumerate(sorted_emails, start=student_start_row):
        add_student_name(workbook, worksheet, row, email_to_name[email])
        points_per_sheet_cell_addresses = add_student_marks_worksheet_points_per_exercise(
            workbook, worksheet, row, email, students_marks, all_sheet_names, graded_sheet_names
        )
        all_points_per_sheet_cell_addresses[email] = points_per_sheet_cell_addresses
    worksheet.autofit()
    return all_points_per_sheet_cell_addresses


def create_worksheet_points_per_sheet(workbook: Workbook, _the_config: config.Config,
                                      email_to_name, students_marks,
                                      all_sheet_names, graded_sheet_names,
                                      points_per_sheet_cell_addresses) -> None:
    """
    Creates an Excel worksheet that contains the points per sheet,
    useful values calculated with functions, and conditional formatting.
    """
    worksheet = workbook.add_worksheet("Points Summary")
    worksheet.activate()
    num_legend_rows = 4
    worksheet.write(3 + num_legend_rows, 0, "First Name", workbook.add_format(BOLD))
    worksheet.write(3 + num_legend_rows, 1, "Last Name", workbook.add_format(BOLD))
    worksheet.write(3 + num_legend_rows, 2, "Total Points", workbook.add_format(BOLD))
    worksheet.write(3 + num_legend_rows, 3, "Current Average", workbook.add_format(BOLD))
    worksheet.write(3 + num_legend_rows, 4, "Improve", workbook.add_format(BOLD))
    worksheet.write(1 + num_legend_rows, 0, "Max Points", workbook.add_format(BOLD))
    worksheet.write(2 + num_legend_rows, 0, "Average", workbook.add_format(BOLD))

    worksheet.set_row(0 + num_legend_rows, cell_format=workbook.add_format(BORDER_TOP_BOTTOM))
    student_start_row = 4 + num_legend_rows
    student_end_row = student_start_row + len(email_to_name) - 1
    for col, sheet_name in enumerate(all_sheet_names, start=5):
        worksheet.write(0 + num_legend_rows, col, sheet_name, workbook.add_format(BOLD | BORDER_TOP_BOTTOM))
        max_points_value = _the_config.max_points_per_sheet.get(sheet_name)
        worksheet.write(1 + num_legend_rows, col, max_points_value)
        avg_range = xl_range_abs(student_start_row, col, student_end_row, col)
        worksheet.write_formula(2 + num_legend_rows, col, f'=IFERROR(AVERAGE({avg_range}),"")')

    sorted_emails = sorted(email_to_name.keys(), key=lambda e: (email_to_name[e][0], email_to_name[e][1]))
    for row, email in enumerate(sorted_emails, start=student_start_row):
        add_student_name(workbook, worksheet, row, email_to_name[email])
        add_student_marks_worksheet_points_per_sheet(workbook, worksheet, _the_config,
                                                     row, email, students_marks,
                                                     all_sheet_names, graded_sheet_names,
                                                     points_per_sheet_cell_addresses)
        add_student_average(workbook, worksheet, row, graded_sheet_names, all_sheet_names)
        add_average_conditional_formatting(workbook, worksheet, row)
        add_plagiarism_conditional_formatting(workbook, worksheet, row, all_sheet_names)
        add_pass_or_fail_conditional_formatting(workbook, worksheet, row, graded_sheet_names, all_sheet_names)
    add_legend(workbook, worksheet)
    worksheet.autofit()


def load_marks_files(marks_dir: Path, _the_config: config.Config):
    """
    Loads the data of the individual marks files in the specified directory.
    """
    marks_files = marks_dir.glob("points_*_individual.json")
    if _the_config.points_per == 'exercise':
        students_marks = defaultdict(lambda: defaultdict(dict))
        graded_sheet_names = defaultdict(list)
    else:
        students_marks = defaultdict(dict)
        graded_sheet_names = set()
    tutors = defaultdict(set)
    for file in marks_files:
        with open(file, 'r') as f:
            data = json.load(f)
            sheet_name = data["adam_sheet_name"]
            marks = data["marks"]
            tutors[sheet_name].add(data["tutor_name"])
            if _the_config.points_per == 'exercise':
                for email, exercises in marks.items():
                    for exercise, mark in exercises.items():
                        if students_marks[email][sheet_name].get(exercise):
                            logging.warning(f'{exercise} of sheet {sheet_name} is marked multiple times for {email}!')
                        students_marks[email][sheet_name][exercise] = mark
                        if exercise not in graded_sheet_names[sheet_name]:
                            graded_sheet_names[sheet_name].append(exercise)
            else:
                graded_sheet_names.add(sheet_name)
                for email, mark in marks.items():
                    if students_marks[email].get(sheet_name):
                        logging.warning(f'Sheet {sheet_name} is marked multiple times for {email}!')
                    students_marks[email][sheet_name] = mark
    for sheet_name, tutor_list in tutors.items():
        for tutor in _the_config.classes if _the_config.marking_mode == 'static' else _the_config.tutor_list:
            if tutor not in tutor_list:
                logging.warning(f'There is no file from tutor {tutor} for sheet {sheet_name}!')
    return students_marks, graded_sheet_names


def create_marks_summary_excel_file(_the_config: config.Config, marks_dir: Path) -> None:
    """
    Generates an Excel file that summarizes the students' marks. Uses a path
    to a directory containing the individual marks files.
    """
    email_to_name = create_email_to_name_dict(_the_config.teams)
    workbook = xlsxwriter.Workbook("Points_Summary_Report.xlsx")
    students_marks, graded_sheet_names = load_marks_files(marks_dir, _the_config)
    all_sheet_names = _the_config.max_points_per_sheet.keys()
    points_per_sheet_cell_addresses = None
    if _the_config.points_per == 'exercise':
        points_per_sheet_cell_addresses = create_worksheet_points_per_exercise(
            workbook, email_to_name, students_marks, all_sheet_names, graded_sheet_names
        )
    create_worksheet_points_per_sheet(
        workbook, _the_config, email_to_name, students_marks, all_sheet_names, graded_sheet_names,
        points_per_sheet_cell_addresses
    )
    workbook.close()


def summarize(_the_config: config.Config, args) -> None:
    """
    Generate an Excel file summarizing students' marks after the individual
    marks files have been collected in a directory.
    """
    if not args.marks_dir.is_dir():
        logging.critical("The given individual marks directory is not valid!")
    create_marks_summary_excel_file(
        _the_config, args.marks_dir
    )
