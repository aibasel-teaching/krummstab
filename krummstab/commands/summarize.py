import logging
from collections import defaultdict
from pathlib import Path
import xlsxwriter
from xlsxwriter import Workbook
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range, xl_range_abs
from xlsxwriter.worksheet import Worksheet
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import quote_sheetname, absolute_coordinate

from .. import config, utils
from ..teams import *

BOLD = {'bold': True}
BORDER = {'border': 1}
BORDER_LEFT = {'left': 1}
BORDER_RIGHT = {'right': 1}
BORDER_LEFT_RIGHT = {'left': 1, 'right': 1}
BORDER_TOP = {'top': 1}
BORDER_BOTTOM = {'bottom': 1}
BORDER_TOP_BOTTOM = {'top': 1, 'bottom': 1}
BORDER_FULL = {'top': 1, 'bottom': 1, 'left': 1, 'right': 1}
GRAY = {'bg_color': '#E2E2E2'}
GREEN = {'bg_color': '#9BE189'}
YELLOW = {'bg_color': '#FFE699'}
RED = {'bg_color': '#EE7868'}
PLAGIARISM_RED = {'bg_color': 'red'}
PERCENT = {'num_format': '0%'}
TEXT_WRAP = {'text_wrap': True}
CENTERED = {'align': 'center'}
RIGHT_ALIGNED = {'align': 'right'}
TWO_DIGIT_FLOAT = {'num_format': '0.00'}

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

# ---------------------
def convert_to_float_if_possible(value):
    try:
        return float(value)
    except ValueError:
        return value

def sum_numbers(values):
    total = 0
    for v in values:
        try:
            total += float(v)
        except ValueError:
            pass
    return total

def sort_items_by_score(email_scores_pair):
    scores = email_scores_pair[1].values()
    return sum_numbers(scores)

VAR_MIN_GOOD_IMPROVEMENT = "min_good_improvement"
VAR_MIN_BAD_IMPROVEMENT = "min_bad_improvement"
VAR_MARKS_BY_SHEET = "marks_by_sheet"
VAR_SHEET_WAS_MARKED = "was_marked"
VAR_MARKS_ALL_SHEETS = "marks_available_all_sheet"
VAR_MARKS_PAST_SHEETS = "marks_available_past_sheet"
VAR_MARKS_FUTURE_SHEETS = "marks_available_future_sheet"
VAR_MARKS_TO_PASS = "marks_to_pass"

class PointsSummarySheetBuilder:
    def __init__(self, _the_config, sheets, email_to_name, students_marks, graded_sheet_names):
        self.config = _the_config
        self.sheets = sheets
        self.email_to_name = email_to_name
        self.students_marks = students_marks
        self.graded_sheet_names = graded_sheet_names
        self.workbook = None
        self.worksheet = None

    def get_available_marks(self):
        return [self.config.max_points_per_sheet.get(sheet) for sheet in self.sheets]

    def get_was_marked_values(self):
        return [int(sheet in self.graded_sheet_names) for sheet in self.sheets]

    def write(self, row, col, value, format=None, row_abs=True, col_abs=True):
        self.worksheet.write(row, col, value, self.workbook.add_format(format))
        return xl_rowcol_to_cell(row, col, row_abs, col_abs)

    def write_formula(self, row, col, formula, format=None, row_abs=True, col_abs=True):
        return self.write(row, col, formula, format, row_abs, col_abs)

    def merge_range(self, row_start, col_start, row_end, col_end, value, format=None):
        self.worksheet.merge_range(row_start, col_start, row_end, col_end,
                                   value, self.workbook.add_format(format))

    def write_row(self, row, col, values, format=None, abs_ref=True):
        for current_col, value in enumerate(values, start=col):
            self.write(row, current_col, value, format)
        if abs_ref:
            return xl_range_abs(row, col, row, col + len(values) - 1)
        else:
            return xl_range(row, col, row, col + len(values) - 1)

    def define_name(self, var, range):
        ref =  f"{absolute_coordinate(range)}"
        self.workbook.define_name(var, ref)
#        ref =  f"{quote_sheetname(self.worksheet.title)}!{absolute_coordinate(range)}"
#        defined_name = DefinedName(var, attr_text=ref)
#        self.workbook.defined_names.add(defined_name)

    def add_conditional_format(self, range, formula, format):
        self.worksheet.conditional_format(range, {
            'type': 'formula',
            'criteria': formula,
            'format': self.workbook.add_format(format)
        })

    def write_sheet_name_row(self, row, col):
        return self.write_row(row, col, self.sheets, BOLD | BORDER_TOP_BOTTOM)

    def write_available_marks_row(self, row, col):
        self.merge_range(row, col, row, col + 1, "Available marks", BOLD | RIGHT_ALIGNED)
        ref = self.write_row(row, col + 2, self.get_available_marks())
        self.define_name(VAR_MARKS_BY_SHEET, ref)

    def write_was_marked_row(self, row, col):
        self.merge_range(row, col, row, col + 1, "Was marked", BOLD | RIGHT_ALIGNED)
        ref = self.write_row(row, col + 2, self.get_was_marked_values())
        self.define_name(VAR_SHEET_WAS_MARKED, ref)

    def write_color_key(self, row, col):
        self.write(row, 0, "Color Key", BOLD)
        self.merge_range(row + 1, col, row + 1, col + 1, "Passed", GREEN)
        self.merge_range(row + 2, col, row + 2, col + 1, "Failed", RED)
        self.merge_range(row + 3, col, row + 3, col + 1, "2x Plagiarism", PLAGIARISM_RED)

        self.merge_range(row + 1, col + 4, row + 1, col + 5,
                         "Does not need to improve\naverage if percentage is\nlower than:", GREEN | TEXT_WRAP)
        self.merge_range(row + 2, col + 4, row + 2, col + 5,
                         "Should improve average if\npercentage is between:", YELLOW | TEXT_WRAP)
        self.merge_range(row + 3, col + 4, row + 3, col + 5,
                         "Has to improve average by\nat least the following\npercentage:", RED | TEXT_WRAP)

        cellref_min_good_improvement = self.write(row + 1, col + 6, -0.05, GREEN | PERCENT)
        self.define_name(VAR_MIN_GOOD_IMPROVEMENT, cellref_min_good_improvement)
        cellref_min_bad_improvement = self.write(row + 3, col + 6, 0.1, RED | PERCENT)
        self.define_name(VAR_MIN_BAD_IMPROVEMENT, cellref_min_bad_improvement)
        self.write(row + 2, col + 6, 
                   f"=TEXT({VAR_MIN_GOOD_IMPROVEMENT},\"0%\")& \" to \" & TEXT({VAR_MIN_BAD_IMPROVEMENT},\"0%\")", YELLOW)

    def add_conditional_formatting_for_warnings(self, range, representative_student_improvement):
        self.add_conditional_format(range, f"={representative_student_improvement} <= {VAR_MIN_GOOD_IMPROVEMENT}", GREEN)
        self.add_conditional_format(range, f"={representative_student_improvement} <= {VAR_MIN_BAD_IMPROVEMENT}", YELLOW)
        self.add_conditional_format(range, f"={representative_student_improvement} > {VAR_MIN_BAD_IMPROVEMENT}", RED)

    def add_conditional_formatting_for_zebra_stripes(self, range):
        self.add_conditional_format(range, f"=ISEVEN(ROW())", GRAY)

    def write_student_marks_table(self):
        self.write_sheet_name_row(0, 7)
        self.write_available_marks_row(1, 5)
        self.write_was_marked_row(2, 5)

        summed_marks_format = BOLD | GRAY | CENTERED
        self.merge_range(0, 1, 0, 3, "Available Marks", summed_marks_format | BORDER_FULL)

        self.write(1, 1, "All Sheets", summed_marks_format | BORDER_TOP_BOTTOM)
        cellref_available_all_sheets = self.write_formula(
            2, 1, f"=SUM({VAR_MARKS_BY_SHEET})", summed_marks_format | BORDER_TOP_BOTTOM)
        self.define_name(VAR_MARKS_ALL_SHEETS, cellref_available_all_sheets)

        self.write(1, 2, "Past Sheets", summed_marks_format | BORDER_TOP_BOTTOM)
        cellref_available_past_sheets = self.write_formula(
            2, 2, f"=SUMPRODUCT({VAR_MARKS_BY_SHEET},{VAR_SHEET_WAS_MARKED})", summed_marks_format | BORDER_TOP_BOTTOM)
        self.define_name(VAR_MARKS_PAST_SHEETS, cellref_available_past_sheets)

        self.write(1, 3, "Future Sheets", summed_marks_format | BORDER_TOP_BOTTOM | BORDER_RIGHT)
        cellref_available_future_sheets = self.write_formula(
            2, 3, f"={VAR_MARKS_ALL_SHEETS}-{VAR_MARKS_PAST_SHEETS}", summed_marks_format | BORDER_TOP_BOTTOM | BORDER_RIGHT)
        self.define_name(VAR_MARKS_FUTURE_SHEETS, cellref_available_future_sheets)

        self.write(0, 0, "Marks required", summed_marks_format | BORDER_LEFT_RIGHT | BORDER_BOTTOM  | TEXT_WRAP)
        self.write(1, 0, "to pass", summed_marks_format | BORDER_LEFT_RIGHT | BORDER_BOTTOM  | TEXT_WRAP)
        cellref_marks_to_pass = self.write_formula(
            2, 0, f"={VAR_MARKS_ALL_SHEETS}*0.5", summed_marks_format | BORDER_FULL) # TODO turn 0.5 into a config value
        self.define_name(VAR_MARKS_TO_PASS, cellref_marks_to_pass)

        headers = ["First Name", "Last Name", "Total Marks", "Missing Marks", "Avg Marks", "Required Avg", "Improve"]
        self.write_row(4, 0, headers, BOLD | BORDER_TOP_BOTTOM)
        first_score_column =  len(headers)
        last_score_column = first_score_column + len(self.sheets) - 1
        self.write_sheet_name_row(4, first_score_column)
        sorted_marks = sorted(self.students_marks.items(), key=sort_items_by_score)
        first_student_row = 5
        last_student_row = first_student_row + len(sorted_marks) - 1
        for row, (email, student_marks) in enumerate(sorted_marks, start=first_student_row):
            student_score_values = [convert_to_float_if_possible(student_marks.get(sheet, "")) for sheet in self.sheets]
            rangeref_student_marks = self.write_row(row, first_score_column, student_score_values, abs_ref=False)
            # Name
            first_name, last_name = self.email_to_name.get(email, ("Unknown", "Unknown"))
            self.write(row, 0, first_name, BORDER_TOP_BOTTOM | BORDER_LEFT)
            self.write(row, 1, last_name, BORDER_TOP_BOTTOM | BORDER_RIGHT)
            # Total Marks
            cellref_student_marks_total = self.write_formula(
                row, 2, f"=SUMPRODUCT({rangeref_student_marks},{VAR_SHEET_WAS_MARKED})",
                BORDER_FULL, col_abs=True, row_abs=False)
            # Missing Marks
            cellref_student_marks_missing = self.write_formula(
                row, 3, f"={VAR_MARKS_TO_PASS} - {cellref_student_marks_total}",
                BORDER_FULL, col_abs=True, row_abs=False)
            # Avg Marks
            current_avg_formula = (
                f"{{=AVERAGE(IF(ISNUMBER({rangeref_student_marks}) * {VAR_SHEET_WAS_MARKED}, "
                + f"{rangeref_student_marks} / {VAR_MARKS_BY_SHEET}, "
                + f"\"Ignoring plagiarism and sheets not submitted for the average\"))}}")
            cellref_student_current_average = self.write_formula(
                row, 4, current_avg_formula, format=PERCENT | BORDER_FULL, col_abs=True, row_abs=False)
            # Required Avg
            cellref_student_required_average = self.write_formula(
                row, 5, f"={cellref_student_marks_missing}/{VAR_MARKS_FUTURE_SHEETS}",
                format=PERCENT | BORDER_FULL, col_abs=True, row_abs=False)
            # Improve
            self.write_formula(
                row, 6, f"={cellref_student_required_average} - {cellref_student_current_average}",
                format=PERCENT | BORDER_FULL)

        self.merge_range(3, 5, 3, 6, "Average achieved", BOLD | RIGHT_ALIGNED)
        for col in range(first_score_column, last_score_column + 1):
            rangeref_score_column = xl_range(first_student_row, col, last_student_row, col)
            self.write_formula(3, col, f"=IFERROR(AVERAGE({rangeref_score_column}),"")", TWO_DIGIT_FLOAT)

        # Conditional Formatting
        representative_student_range = (
            xl_rowcol_to_cell(first_student_row, first_score_column, row_abs=False, col_abs=True)
            + ":" +
            xl_rowcol_to_cell(first_student_row, last_score_column, row_abs=False, col_abs=True)
        )
        rangeref_students_name_columns = xl_range(first_student_row, 0, last_student_row, 1)
        self.add_conditional_format(
            rangeref_students_name_columns,  f"=COUNTIF({representative_student_range},\"Plagiarism\") >= 2", PLAGIARISM_RED)
        representative_student_total_marks = xl_rowcol_to_cell(
            first_student_row, 2, row_abs=False, col_abs=True)
        self.add_conditional_format(
            rangeref_students_name_columns, f"={representative_student_total_marks} + {VAR_MARKS_FUTURE_SHEETS} < {VAR_MARKS_TO_PASS}", RED)
        self.add_conditional_format(rangeref_students_name_columns, f"={representative_student_total_marks} >= {VAR_MARKS_TO_PASS}", GREEN)

        self.write_color_key(last_student_row + 2, 0)

        rangeref_students_averge_columns = xl_range(first_student_row, 4, last_student_row, 6)
        representative_student_improvement = xl_rowcol_to_cell(
            first_student_row, 6, row_abs=False, col_abs=True)
        self.add_conditional_formatting_for_warnings(rangeref_students_averge_columns, representative_student_improvement)


        rangeref_all_students = xl_range(first_student_row, 0, last_student_row, last_score_column)
        self.add_conditional_formatting_for_zebra_stripes(rangeref_all_students)



    def add_summary_table_to_workbook(self, workbook: Workbook):
        self.workbook = workbook
        self.worksheet = workbook.add_worksheet("Points Summary")
        self.worksheet.activate()
        self.write_student_marks_table()
        self.worksheet.autofit()
        self.merge_range(3, 0, 3, 4, "If all formulas show up as 0, press Shift+Ctrl+F9 to recalculate them.", BOLD)




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
        data = utils.read_json(file)
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
    sheets = _the_config.max_points_per_sheet.keys()
    email_to_name = create_email_to_name_dict(_the_config.teams)
    students_marks, graded_sheet_names = load_marks_files(marks_dir, _the_config)

    builder = PointsSummarySheetBuilder(_the_config, sheets, email_to_name, students_marks, graded_sheet_names)

#    workbook = openpyxl.Workbook()
    workbook = xlsxwriter.Workbook("Points_Summary_Report.xlsx")
    builder.add_summary_table_to_workbook(workbook)

#    if _the_config.points_per == 'exercise':
#        points_per_sheet_cell_addresses = create_worksheet_points_per_exercise(
#            workbook, email_to_name, students_marks, all_sheet_names, graded_sheet_names
#        )
    workbook.close()

#    workbook.save("Points_Summary_Report.xlsx")

#    pyxl_worksheet = pyxl_workbook.active
    # make sure sheetnames and cell references are quoted correctly
#    pyxl_worksheet["A1"] = 1
#    pyxl_worksheet["A2"] = 12
#    pyxl_worksheet["A3"] = 14
#    pyxl_worksheet["A5"] = 16
#    pyxl_worksheet["B2"] = "=SUM(global_range)"



def summarize(_the_config: config.Config, args) -> None:
    """
    Generate an Excel file summarizing students' marks after the individual
    marks files have been collected in a directory.
    """
    if not args.marks_dir.is_dir():
        logging.critical("The given individual marks directory is not valid!")
    create_marks_summary_excel_file(_the_config, args.marks_dir)
