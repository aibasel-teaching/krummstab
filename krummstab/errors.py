import logging


def unsupported_marking_mode_error(marking_mode) -> None:
    """
    Throw an error if a marking mode is encountered that is not handled
    correctly. This is primarily used in the else case of if-elif-else branches
    on the marking_mode, which shouldn't be reached in regular operation anyway
    because the config settings should be validated first. But this function
    offers an easy way to find sections to consider if we ever want to add a new
    marking_mode.
    """
    logging.critical(f"Unsupported marking mode {marking_mode}!")
