import json
import logging
from collections import defaultdict
from pathlib import Path
import xlsxwriter
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_range, xl_rowcol_to_cell
from xlsxwriter.worksheet import Worksheet

from . import config


BOLD = {'bold': True}
BORDER = {'border': 1}
BORDER_LEFT = {'left': 1}
BORDER_TOP_BOTTOM = {'top': 1, 'bottom': 1}
GRAY = {'bg_color': '#E2E2E2'}
GREEN = {'bg_color': '#9BE189'}
YELLOW = {'bg_color': '#FFFF00'}
RED = {'bg_color': '#EE7868'}
PLAGIARISM_RED = {'bg_color': 'red'}


def add_total_points_conditional_formatting(workbook: Workbook, worksheet: Worksheet,
                                            row: int, graded_sheet_names):
    graded_sheets_points_range = xl_range(1, 3, 1, 3 + len(graded_sheet_names) - 1)
    worksheet.conditional_format(row, 2, row, 2, {
        'type': 'formula',
        'criteria': f"=(INDIRECT(ADDRESS(ROW(),COLUMN()))/SUM({graded_sheets_points_range})) >= 0.65",
        'format': workbook.add_format(GREEN)
    })
    worksheet.conditional_format(row, 2, row, 2, {
        'type': 'formula',
        'criteria': f"=(INDIRECT(ADDRESS(ROW(),COLUMN()))/SUM({graded_sheets_points_range})) >= 0.5",
        'format': workbook.add_format(YELLOW)
    })
    worksheet.conditional_format(row, 2, row, 2, {
        'type': 'formula',
        'criteria': f"=(INDIRECT(ADDRESS(ROW(),COLUMN()))/SUM({graded_sheets_points_range})) < 0.5",
        'format': workbook.add_format(RED)
    })


def add_pass_or_fail_conditional_formatting(workbook: Workbook, worksheet: Worksheet, row: int,
                                            graded_sheet_names, all_sheet_names):
    total_points_all_sheets_range = xl_range(1, 3, 1, 3 + len(all_sheet_names) - 1)
    possible_points_range = xl_range(1, 3 + len(graded_sheet_names), 1, 3 + len(all_sheet_names) - 1)
    worksheet.conditional_format(row, 0, row, 0, {
        'type': 'formula',
        'criteria': f"=(SUM(INDIRECT(ADDRESS(ROW(),3)), SUM({possible_points_range})))"
                    f" < (SUM({total_points_all_sheets_range}) * 0.5)",
        'format': workbook.add_format(RED)
    })
    worksheet.conditional_format(row, 1, row, 1, {
        'type': 'formula',
        'criteria': f"=(SUM(INDIRECT(ADDRESS(ROW(),3)), SUM({possible_points_range})))"
                    f" < (SUM({total_points_all_sheets_range}) * 0.5)",
        'format': workbook.add_format(RED)
    })
    worksheet.conditional_format(row, 0, row, 0, {
        'type': 'formula',
        'criteria': f"INDIRECT(ADDRESS(ROW(),3)) >= (SUM({total_points_all_sheets_range}) * 0.5)",
        'format': workbook.add_format(GREEN)
    })
    worksheet.conditional_format(row, 1, row, 1, {
        'type': 'formula',
        'criteria': f"INDIRECT(ADDRESS(ROW(),3)) >= (SUM({total_points_all_sheets_range}) * 0.5)",
        'format': workbook.add_format(GREEN)
    })


def add_plagiarism_conditional_formatting(workbook: Workbook, worksheet: Worksheet,
                                          row: int, all_sheet_names):
    student_marks_range = xl_range(row, 3, row, 3 + len(all_sheet_names) - 1)
    worksheet.conditional_format(row, 0, row, 0, {
        'type': 'formula',
        'criteria': f'=COUNTIF({student_marks_range},"Plagiat") >= 2',
        'format': workbook.add_format(PLAGIARISM_RED)
    })
    worksheet.conditional_format(row, 1, row, 1, {
        'type': 'formula',
        'criteria': f'=COUNTIF({student_marks_range},"Plagiat") >= 2',
        'format': workbook.add_format(PLAGIARISM_RED)
    })


def write_mark(worksheet, row, col, mark) -> None:
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
    if _the_config.points_per == 'exercise':
        for col, cell_address in enumerate(points_per_sheet_cell_addresses[email], start=3):
            worksheet.write(row, col, f"={cell_address}")
    else:
        for col, sheet_name in enumerate(list(all_sheet_names)[:len(graded_sheet_names)], start=3):
            mark = students_marks[email].get(sheet_name)
            write_mark(worksheet, row, col, mark)
    student_marks_range = xl_range(row, 3, row, 3 + len(all_sheet_names) - 1)
    worksheet.write_formula(row, 2, f"=SUM({student_marks_range})", workbook.add_format(BORDER))


def add_student_marks_worksheet_points_per_exercise(workbook: Workbook, worksheet: Worksheet,
                                                    row, email, students_marks, all_sheet_names,
                                                    graded_sheet_names) -> list[str]:
    col = 2
    points_per_sheet_cell_addresses = []
    for sheet_name in list(all_sheet_names)[:len(graded_sheet_names)]:
        student_marks_range = xl_range(row, col + 1, row, col + len(graded_sheet_names[sheet_name]))
        worksheet.write_formula(
            row, col,
            f'=IF(COUNTIF({student_marks_range},"Plagiat") > 0,"Plagiat",SUM({student_marks_range}))',
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
    first_name, last_name = name
    if row % 2 == 0:
        worksheet.write(row, 0, first_name, workbook.add_format(BORDER_TOP_BOTTOM | GRAY))
        worksheet.write(row, 1, last_name, workbook.add_format(BORDER_TOP_BOTTOM | GRAY))
        worksheet.set_row(row, None, workbook.add_format(GRAY))
    else:
        worksheet.write(row, 0, first_name, workbook.add_format(BORDER_TOP_BOTTOM))
        worksheet.write(row, 1, last_name, workbook.add_format(BORDER_TOP_BOTTOM))


def create_worksheet_points_per_exercise(workbook: Workbook, email_to_name, students_marks, all_sheet_names,
                                         graded_sheet_names) -> (Worksheet, defaultdict[str, list[str]]):
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
    sorted_emails = sorted(students_marks.keys(), key=lambda e: email_to_name[e][0])
    for row, email in enumerate(sorted_emails, start=student_start_row):
        add_student_name(workbook, worksheet, row, email_to_name[email])
        points_per_sheet_cell_addresses = add_student_marks_worksheet_points_per_exercise(
            workbook, worksheet, row, email, students_marks, all_sheet_names, graded_sheet_names
        )
        all_points_per_sheet_cell_addresses[email] = points_per_sheet_cell_addresses
    return worksheet, all_points_per_sheet_cell_addresses


def create_worksheet_points_per_sheet(workbook: Workbook, _the_config: config.Config,
                                      email_to_name, students_marks,
                                      all_sheet_names, graded_sheet_names,
                                      points_per_sheet_cell_addresses) -> None:
    worksheet = workbook.add_worksheet("Points Summary")
    worksheet.activate()
    worksheet.write(3, 0, "First Name", workbook.add_format(BOLD))
    worksheet.write(3, 1, "Last Name", workbook.add_format(BOLD))
    worksheet.write(3, 2, "Total Points", workbook.add_format(BOLD))
    worksheet.write(1, 0, "Max Points", workbook.add_format(BOLD))
    worksheet.write(2, 0, "Average", workbook.add_format(BOLD))

    student_start_row = 4
    student_end_row = student_start_row + len(students_marks) - 1
    for col, sheet_name in enumerate(all_sheet_names, start=3):
        worksheet.write(0, col, sheet_name, workbook.add_format(BOLD))
        max_points_value = _the_config.max_points_per_sheet.get(sheet_name)
        worksheet.write(1, col, max_points_value)
        if sheet_name in graded_sheet_names:
            avg_range = xl_range(student_start_row, col, student_end_row, col)
            worksheet.write_formula(2, col, f"=AVERAGE({avg_range})")

    sorted_emails = sorted(students_marks.keys(), key=lambda e: email_to_name[e][0])
    for row, email in enumerate(sorted_emails, start=student_start_row):
        add_student_name(workbook, worksheet, row, email_to_name[email])
        add_student_marks_worksheet_points_per_sheet(workbook, worksheet, _the_config,
                                                     row, email, students_marks,
                                                     all_sheet_names, graded_sheet_names,
                                                     points_per_sheet_cell_addresses)
        add_plagiarism_conditional_formatting(workbook, worksheet, row, all_sheet_names)
        add_pass_or_fail_conditional_formatting(workbook, worksheet, row, graded_sheet_names, all_sheet_names)
        add_total_points_conditional_formatting(workbook, worksheet, row, graded_sheet_names)


def load_marks_files(marks_dir: Path, _the_config: config.Config):
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
                for email, mark in marks.items():
                    if students_marks[email][sheet_name]:
                        logging.warning(f'Sheet {sheet_name} is marked multiple times for {email}!')
                    students_marks[email][sheet_name] = mark
    for sheet_name, tutor_list in tutors.items():
        for tutor in _the_config.classes if _the_config.marking_mode == 'static' else _the_config.tutor_list:
            if tutor not in tutor_list:
                logging.warning(f'There is no file from tutor {tutor} for sheet {sheet_name}!')
    return students_marks, graded_sheet_names


def create_email_to_name_dict(_the_config: config.Config) -> dict[str, tuple[str, str]]:
    email_to_name = {}
    for team in _the_config.teams:
        for first_name, last_name, email in team:
            email_to_name[email] = (first_name, last_name)
    return email_to_name


def create_marks_summary_excel_file(_the_config: config.Config, marks_dir: Path) -> None:
    email_to_name = create_email_to_name_dict(_the_config)
    workbook = xlsxwriter.Workbook("Points_Summary_Report.xlsx")
    students_marks, graded_sheet_names = load_marks_files(marks_dir, _the_config)
    all_sheet_names = _the_config.max_points_per_sheet.keys()
    points_per_sheet_cell_addresses = None
    if _the_config.points_per == 'exercise':
        worksheet_exercise, points_per_sheet_cell_addresses = create_worksheet_points_per_exercise(
            workbook, email_to_name, students_marks, all_sheet_names, graded_sheet_names
        )
    create_worksheet_points_per_sheet(
        workbook, _the_config, email_to_name, students_marks, all_sheet_names, graded_sheet_names,
        points_per_sheet_cell_addresses
    )
    workbook.close()
