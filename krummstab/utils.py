import json
import logging
import pathlib
import shutil
import sys
import tempfile
from zipfile import ZipFile
import jsonschema

from . import strings

# Logging ----------------------------------------------------------------------


def configure_logging(level=logging.INFO):
    class ColoredFormatter(logging.Formatter):
        FORMATS = {
            logging.DEBUG: "\033[0;37m[{levelname}]\033[0m {message}",
            logging.INFO: "\033[0;34m[{levelname}]\033[0m {message}",
            logging.WARNING: "\033[0;33m[{levelname}]\033[0m {message}",
            logging.ERROR: "\033[0;31m[{levelname}]\033[0m {message}",
            logging.CRITICAL: "\033[0;31m[{levelname}]\033[0m {message}",
        }

        def format(self, record):
            formatter = logging.Formatter(
                ColoredFormatter.FORMATS[record.levelno], style="{"
            )
            return formatter.format(record)

    class LevelFilter:
        def __init__(self, min_level, max_level):
            self.min_level = min_level
            self.max_level = max_level

        def filter(self, record):
            return self.min_level <= record.levelno <= self.max_level

    class CustomHandler(logging.StreamHandler):
        def __init__(
            self, stream, min_level=logging.DEBUG, max_level=logging.CRITICAL
        ):
            logging.StreamHandler.__init__(self, stream)
            self.setFormatter(ColoredFormatter())
            self.addFilter(LevelFilter(min_level, max_level))

        def emit(self, record):
            logging.StreamHandler.emit(self, record)
            if record.levelno >= logging.CRITICAL:
                sys.exit("aborting")

    root_logger = logging.getLogger("")
    # Remove old handlers
    for handler in root_logger.handlers:
        root_logger.removeHandler(handler)
    root_logger.addHandler(CustomHandler(sys.stdout, max_level=logging.WARNING))
    root_logger.addHandler(CustomHandler(sys.stderr, min_level=logging.ERROR))
    root_logger.setLevel(level)


# Printing ---------------------------------------------------------------------


def query_yes_no(text: str, default: bool = True) -> bool:
    """
    Ask the user a yes/no question and return answer.
    """
    valid = {"yes": True, "y": True, "ye": True, "no": False, "n": False}
    options = "[Y/n]" if default else "[y/N]"
    print("\033[0;35m[Query]\033[0m " + text + f" {options}")
    choice = input().lower()
    if choice == "":
        return default
    elif choice in valid:
        return valid[choice]
    else:
        logging.warning(
            f"Invalid choice '{choice}'. Please respond with 'yes' or 'no'."
        )
        return query_yes_no(text, default)


# JSON parsing -------------------------------------------------------


def validate_json(
    data: dict,
    schema: dict,
    source: str = "file",
    schema_version=jsonschema.Draft7Validator,
) -> None:
    """
    Validates a JSON object against a given schema.
    """
    try:
        jsonschema.validate(data, schema, schema_version)
    except jsonschema.exceptions.ValidationError as error:
        logging.critical(
            f"Validation error: {source} does not have the right format: "
            f"{error.message}"
        )


def read_json(source: str | pathlib.Path, source_name: str = "file") -> dict:
    """
    Reads a JSON file and returns its contents.
    """
    data = {}
    try:
        if isinstance(source, pathlib.Path):
            source_name = source
            json_str = source.read_text(encoding="utf-8")
        else:
            json_str = source
        data = json.loads(json_str)
    except json.decoder.JSONDecodeError as error:
        logging.critical(f"Wrong JSON format in {source_name}: {error}")
    return data


# Miscellaneous ----------------------------------------------------------------


def is_hidden_file(name: str) -> bool:
    """
    Check if a given file name could be a hidden file. In particular a file that
    should not be sent to students as feedback.
    """
    return name.startswith(".") or is_macos_path(name)


def is_macos_path(path: str) -> bool:
    """
    Check if the given path is non-essential file created by MacOS.
    """
    return "__MACOSX" in path or ".DS_Store" in path


def filtered_extract(zip_file: ZipFile, dest: pathlib.Path) -> None:
    """
    Extract all files except for MACOS helper files.
    """
    zip_content = zip_file.namelist()
    for file_path in zip_content:
        if is_macos_path(file_path):
            continue
        zip_file.extract(file_path, dest)


def move_content_and_delete(src: pathlib.Path, dst: pathlib.Path) -> None:
    """
    Move all content of source directory to destination directory.
    This does not complain if the dst directory already exists.
    """
    assert src.is_dir() and dst.is_dir()
    with tempfile.TemporaryDirectory() as temp_dir:
        shutil.copytree(src, temp_dir, dirs_exist_ok=True)
        shutil.rmtree(src)
        shutil.copytree(temp_dir, dst, dirs_exist_ok=True)
