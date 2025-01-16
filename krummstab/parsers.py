import argparse
import pathlib

from .commands import *


DEFAULT_SHARED_CONFIG_FILE = "config-shared.json"
DEFAULT_INDIVIDUAL_CONFIG_FILE = "config-individual.json"


def add_parsers():
    parser = argparse.ArgumentParser(description="")
    add_main_command_parser(parser)
    subparsers = add_subcommand_parser(parser)
    add_help_command_parser(subparsers)
    add_init_command_parser(subparsers)
    add_collect_command_parser(subparsers)
    add_combine_command_parser(subparsers)
    add_send_command_parser(subparsers)
    add_summarize_command_parser(subparsers)
    return parser


def add_main_command_parser(parser):
    parser.add_argument(
        "-s",
        "--config-shared",
        default=DEFAULT_SHARED_CONFIG_FILE,
        type=pathlib.Path,
        help=(
            "path to the json config file containing shared settings such as "
            "the student/email list"
        ),
    )
    parser.add_argument(
        "-i",
        "--config-individual",
        default=DEFAULT_INDIVIDUAL_CONFIG_FILE,
        type=pathlib.Path,
        help=(
            "path to the json config file containing individual settings such "
            "as tutor name and email configuration"
        ),
    )


def add_subcommand_parser(parser):
    subparsers = parser.add_subparsers(
        required=True,
        help="available sub-commands",
        dest="sub_command",
    )
    return subparsers


def add_help_command_parser(subparsers):
    parser_init = subparsers.add_parser(
        "help",
        help=(
            "print this help message; "
            "run e.g. 'krummstab init -h' to print the help of a sub-command"
        ),
    )


def add_init_command_parser(subparsers):
    parser_init = subparsers.add_parser(
        "init",
        help="unpack zip file from ADAM and prepare directory structure",
    )
    parser_init.add_argument(
        "adam_zip_path",
        type=pathlib.Path,
        help="path to the zip file downloaded from ADAM",
    )
    parser_init.add_argument(
        "-t",
        "--target",
        type=pathlib.Path,
        required=False,
        help="path to the directory that will contain the submissions",
    )
    group = parser_init.add_mutually_exclusive_group(required=False)
    group.add_argument(
        "-n",
        "--num-exercises",
        dest="num_exercises",
        type=int,
        help="the number of exercises in the sheet",
    )
    group.add_argument(
        "-e",
        "--exercises",
        dest="exercises",
        nargs="+",
        type=int,
        help="the exercises you have to mark",
    )
    parser_init.set_defaults(func=init)


def add_collect_command_parser(subparsers):
    parser_collect = subparsers.add_parser(
        "collect",
        help="collect feedback files after marking is done",
    )
    parser_collect.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_collect.set_defaults(func=collect)


def add_combine_command_parser(subparsers):
    parser_combine = subparsers.add_parser(
        "combine",
        help=(
            "combine multiple share archives, only necessary if tutors mark per"
            " exercise and have to integrate their individual feedback into a"
            " single ZIP file to send to the students"
        ),
    )
    parser_combine.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_combine.set_defaults(func=combine)


def add_send_command_parser(subparsers):
    parser_send = subparsers.add_parser(
        "send",
        help="send feedback via email",
    )
    parser_send.add_argument(
        "sheet_root_dir",
        type=pathlib.Path,
        help="path to the sheet's directory",
    )
    parser_send.add_argument(
        "-d",
        "--dry_run",
        action="store_true",
        help="only print emails instead of sending them",
    )
    parser_send.set_defaults(func=send)


def add_summarize_command_parser(subparsers):
    parser_summarize = subparsers.add_parser(
        "summarize",
        help="summarize individual marks files into Excel report",
    )
    parser_summarize.add_argument(
        "marks_dir",
        type=pathlib.Path,
        help="path to the directory with all individual marks files",
    )
    parser_summarize.set_defaults(func=summarize)
