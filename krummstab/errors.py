import logging
import pathlib


def unsupported_marking_mode_error(marking_mode: str) -> None:
    """
    Throw this error if a marking mode is encountered that is not handled
    correctly. This is primarily used in the else case of if-elif-else branches
    on the marking_mode, which shouldn't be reached in regular operation anyway
    because the config settings should be validated first. But this function
    offers an easy way to find sections to consider if we ever want to add a new
    marking_mode.
    """
    logging.critical(f"Unsupported marking mode {marking_mode}!")


def unexpected_zip_structure(adam_zip_path: pathlib.Path) -> None:
    """
    Throw this error if the structure of the zip file provided to `init` does
    not match our assumptions. This is expected to happen if ADAM changes their
    output format.
    """
    logging.critical(
        f"The zip file at {adam_zip_path} has an unexpected structure. The "
        "file at the given location is possibly not a submission zip provided "
        "by ADAM. If it is, contact your assistant as ADAM may have changed "
        "their output format in which case Krummstab needs to be adapted."
    )
