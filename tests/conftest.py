# Inspiration on how to add an option to pytest from here:
# https://stackoverflow.com/a/42145604
def pytest_addoption(parser):
    parser.addoption("--skip-mark-test", action="store_true", default=False)


def pytest_generate_tests(metafunc):
    # This is called for every test. Only get/set command line arguments
    # if the argument is specified in the list of test "fixturenames".
    option_value = metafunc.config.option.skip_mark_test
    if "skip_mark_test" in metafunc.fixturenames and option_value is not None:
        metafunc.parametrize("skip_mark_test", [option_value])
