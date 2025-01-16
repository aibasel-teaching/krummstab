#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
This script uses the following, potentially ambiguous, terminology:
team     - set of students that submit together, as defined on ADAM
class    - set of teams that are in the same exercise class
relevant - a team whose submission has to be marked by the tutor running the
           script is considered 'relevant'
to mark  - grading/correcting a sheet; giving feedback
marks    - points awarded for a sheet or exercise
"""
import logging
import os
import sys

from . import config, parsers, utils


# Might be necessary to make colored output work on Windows.
os.system("")


def main():
    utils.configure_logging()

    parser = parsers.add_parsers()
    args = parser.parse_args()

    if args.sub_command == "help":
        parser.print_help()
        sys.exit(0)

    # Process config files
    _the_config = config.Config([args.config_shared, args.config_individual])

    # Execute subcommand. This calls the function set as default in the parser.
    # For example, `args.func` is set to `init` if the subcommand is "init".
    logging.info(f"Running command '{args.sub_command}'...")
    args.func(_the_config, args)
    logging.info(f"Command '{args.sub_command}' terminated successfully. 🎉")


if __name__ == "__main__":
    main()
