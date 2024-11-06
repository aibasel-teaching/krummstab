# Krummstab Feedback Script

The purpose of this script is to automate some of the menial steps involved in
marking ADAM submissions.

The system is made up of three components: the central
[krummstab](https://pypi.org/project/krummstab) PyPI project, and two JSON
configuration files, `config-shared.json` and `config-individual.json`.

The shared config file contains [general settings](#general-settings) that need
to be adapted to the course that is being taught, but should remain static
thereafter. Additionally, it lists all students and their
[team assignment](#teams). This part of the file is subject to change during the
semester as students drop the course or teams are reassigned. It is important
that all tutors have an identical copy of the shared config file, meaning that
whenever a tutor makes changes to the file, she or he should share the new
version with the others, for example via the Discord server or uploading it to
ADAM. (Assistants can find tips for creating this file
[here](#setting-up-shared-config-file).)

The individual config file contains [personal settings](#individual-settings)
that are only relevant to each tutor. These only need to be set once at the
beginning of the course.

Depending on the general settings of the shared config file, different command
line options may be mandatory. The `help` option provides information about the
script, its subcommands (currently `init`, `collect`, `combine` and `send`), and
their parameters. Once you have completed the one-time setup below, you'll be
able to access the help via:
```
krummstab -h
krummstab <subcommand> -h
```
In the following I will go over the recommended workflow using the settings of
the Foundations of Artificial Intelligence lecture from the spring semester 2023
as an example.


## Requirements

- `Python 3.10+`: I only tested the script with 3.10 and I think it makes use of
some new-ish language features, so I cannot guarantee that everything works as
expected with older Python versions.


## One-Time Setup

> 📝 I'm assuming a Linux environment in the following. In case you are using
> macOS, I hope that the following instructions work without major differences.
> In case you are using Windows, I recommend trying to install a Windows
> Subsystem for Linux (WSL), which should allow you to follow these
> instructions exactly. Alternatively you can try to install the necessary
> software natively, but I don't offer support here.

To get started, create an empty directory where you want to do your marking, in
this example the directory will be called `ki-fs23-marking`:
```
mkdir ki-fs23-marking
```
Navigate to this directory, set up a virtual Python environment, and activate
it:
```
cd ki-fs23-marking
python3 -m venv .venv
source .venv/bin/activate
```
Then you can install Krummstab in this environment:
```
pip install krummstab
```
To test the installation, you can print the help string:
```
krummstab -h
```

With the script installed, we continue with the config files. You should have
gotten a `config-shared.json` file from the teaching assistant, copy this file
into the directory you just created, in our example `ki-fs23-marking`. Similarly
you can copy the `config-individual.json` file from the `tests` directory of
this repository. Replace the example entries in the individual configurations
with your own information; The parameters are explained
[here](#individual-settings). Make sure that the string you enter in the field
`tutor_name` in your individual config exactly matches the existing entry in the
`tutor_list` field of the shared config.

In general, it is important that the all configurations, besides the individual
ones you just adjusted, are exactly the same across all tutors, as otherwise
submissions may be assigned to multiple or no tutors. If you think that
something should be changed in the shared settings, please let the teaching
assistant and the other tutors know, so that the configurations remain in sync.
This may in particular be necessary if teams change throughout the semester.

In order to work with the script, you will have to call the `krummstab` command
from a command line whose working directory is the one which contains the two
config files. If you'd like to keep the config files somewhere else, you'll have
to provide the paths to the files with the `-s path/to/shared` and `-i
path/to/individual` flags whenever you call `krummstab`.


## Marking a Sheet

While the steps above are only necessary for the initial setup, the following
procedure applies to every exercise sheet. The first step is always to activate
the virtual environment in which we have installed Krummstab. You do this by
navigating to the marking directory and using the source command.
```
cd ki-fs23-marking
source .venv/bin/activate
```
If you forget this step you'll get an error saying that the `krummstab` command
could not be found.

### init
First, download the submissions from ADAM and save the zip file in the marking
directory. (It's important that you only download the submissions after the
ADAM deadline has passed, so that all tutors have the same, complete pool of
submissions.) Our example directory `ki-fs23-marking`, with `Sheet 1.zip` being
the file downloaded from ADAM, should look like this:
```
.
├── .venv
├── config-individual.json
├── config-shared.json
└── Sheet 1.zip
```
We can now finally make the script do something useful by running:
```
krummstab init -n 4 -t sheet01 "Sheet 1.zip"
```
This will unzip the submissions and prepare them for marking. The flag `-n`
expects the number of exercises in the sheet, `-t` is optional and takes the
name of the directory the submissions should be extracted to. By default it's
the name of the zip file, but I'm choosing to rename it in order to get rid of
the whitespace in the directory name. The directory should now look something
like this:
```
.
├── .venv
├── config-individual.json
├── config-shared.json
├── sheet01
│   ├── 12345_Muster_Müller
│   │   ├── feedback
│   │   │   └── feedback_tutor-name.pdf.todo
│   │   └── Sheet1_MaxMuster_MayaMueller.pdf
│   .
│   ├── DO_NOT_MARK_12346_Meier_Meyer
│   │   └── submission_exercise_sheet1.pdf
│   .
│   └── points.json
└── Sheet 1.zip
```
As you may have guessed, the submissions you need to mark are those without the
`DO_NOT_MARK_` prefix. Those directories contain the files submitted by the
respective team, as well as a directory called `feedback`, which in turn
contains an empty placeholder PDF file and copies of submitted files that are
not PDFs (e.g. source files).

The idea is that you can give feedback to non-PDFs by adding your comments to
these copies directly, and delete the ones you don't need to comment on. For the
PDF feedback you can use whichever tool you like, and overwrite the `.pdf.todo`
placeholder with the resulting output. If this tool adds files to the feedback
directory that you do not want to send to the students, you can add their
endings to the config file under the `ignore_feedback_suffix` key. Marking with
Xournal++ is supported by default: Simply add the flag `-x` to the `init`
command above to automatically create the relevant `.xopp` files.

While writing the feedback, you can keep track of the points the teams get in
the file `points.json`.

### collect
Once you have marked all the teams assigned to you and added their points to
the `points.json` file, you can run the next command, where `sheet01` is the
path to the directory created by the `init` command:
```
krummstab collect sheet01
```
This will create a zip archive in every feedback directory containing the
feedback for that team. Additionally, a semicolon-separated list of all points
is printed. This can be useful in case you have to paste the points into a
shared spreadsheet. The names are there to be able to double-check that the rows
match up.

In case you need make changes to the markings and rerun the collection step, use
the `-r` flag to overwrite existing feedback archives. If you are using
Xournal++, you can also use the `-x` flag here to automatically export the
`.xopp` files before collecting the feedback.

### combine
This command is only relevant for the `exercise` marking mode.
`TODO: Document this.`

### send
For the `static` marking mode, it is possible to directly send the
feedback to the students via e-mail. For this to work you have to be in the
university network, which likely means you'll have to connect to the university
VPN. You may find the `--dry_run` option useful, instead of sending the e-mails
directly, it only prints them so that you can double-check that everything looks
as expected.


## Config File Details

### Setting Up Shared Config File
- References for the structure of the file can be found in the `tests`
  directory.
- You should be able to get the list of students as an Excel file from ADAM:
  course page > tab 'Content' ('Inhalt') > exercise page > tab 'Submissions and
  Grades' ('Abgaben und Noten') > 'Grades View' ('Notenübersicht') > button
  'Export (Excel)' at top of page. You can then run the shell script
  `scripts/xlsx-to-config.sh` with the downloaded file as input to get a student
  list in JSON format as a starting point for the shared config file.

### Individual Settings
- `tutor_name`: ID of the tutor, this must match with either an element of
  `tutor_list` (for `exercise`) or a key in `teams` (for `static`)
- `tutor_email`: tutor's email address, feedback will be sent via this address
- `feedback_email_cc`: list of email addresses that will be CC'd with every
  feedback email, for example the addresses of all tutors
- `smtp_url`: the URL of the SMTP server, `smtp.unibas.ch` by default (you may
  use `smtp-ext.unibas.ch` if your email address is white-listed; this is
  usually not the case and you would likely know if it is)
- `smtp_port`: SMTP port to connect to, `25` by default (use `587` for an
  `smtp-ext` setup)
- `smtp_user`: SMTP user, empty by default (use your short unibas account name
  for an `smtp-ext` setup)
- `ignore_feedback_suffix`: a list of extensions that should be ignored by the
  `collect` sub-command; this is useful if the tools you use for marking create
  files in the feedback folders that you don't want to send to the students

### General Settings
- `lecture_title`: lecture name to be printed in feedback emails
- `marking_mode`
    - `static`: student teams are assigned to a tutor who will mark all their
      submissions
    - `exercise`: with every sheet, tutors distribute the exercises and only
      correct those, but for all submissions
- `points_per`
    - `exercise`: tutors keep track how many points teams got for every exercise
    - `sheet`: tutors only keep track of the total number of points per sheet
- `min_point_unit`: a float denoting the smallest allowed point fraction, for
  example `0.5`, or `1`
- `tutor_list`: list to identify tutors, for example a list of first names
- `max_team_size`: integer denoting the maximum number of members a team may
  have

### Teams
- `teams`: depending on the `marking_mode` teams are structured slightly
  differently
    - `exercise`: list of teams, each consisting of a list of students,
      where each student entry is a list of the form `[ "first_name",
      "last_name", "email@unibas.ch" ]`
    - `static`: similar to before, but teams are not just listed, but assigned
      to a tutor; this is done via a dictionary where some ID for the tutors
      (e.g. first names) are the keys, and the values are the list of teams
      assigned to each tutor


# Development

We added some tests that use the `pytest` framework. Simply install `pytest` via
`pip3 install pytest` (or `pip`, not sure what the difference is), and run the
command `pytest`. Currently it tests the `init` and `collect` steps for the
modes `static` and `exercise`, the `combine` step for the mode `exercise`, and
the `send` step for the mode `static`.
