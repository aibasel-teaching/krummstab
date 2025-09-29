# One-time setup
#    sudo apt install python3-build twine
# or in a venv:
#    pip install build twine

# Publishing a new version:
#
# 1. Update the version tag in this file.
# 2. Remove the `dist/` and the `krummstab.egg-info` directories
# 3. Run the following steps (needs `pip install build twine`):
#
#     $ python3 -m build
#     $ python3 -m twine upload dist/*
#
# 4. Enter the API token

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="krummstab",
    version="1.1.1",
    description="Efficiently give feedback on ADAM submissions at University of Basel",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Remo Christen",
    author_email="remo.christen@unibas.ch",
    url="https://github.com/aibasel-teaching/krummstab",
    license="GPL3+",
    classifiers=[
        "Environment :: Console",
        "Intended Audience :: Education",
        "License :: OSI Approved :: GNU General Public License v3 or later (GPLv3+)",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3.10",
        "Topic :: Education",
    ],
    packages=find_packages(),
    python_requires=">=3.10",
    install_requires=[
        "jsonschema==4.23.0",
        "openpyxl==3.1.5",
        "pypdf==3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "krummstab = krummstab:main",
        ]
    },
    include_package_data=True,
    package_data={
        "krummstab": [
            "schemas/config-schema.json",
            "schemas/submission-info-schema.json",
        ],
    },
)
