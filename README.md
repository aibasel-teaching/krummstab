<img src="./logo/banner.svg" alt="krummstab-logo">

# Krummstab Feedback Script

The purpose of this script is to automate some of the tedious steps involved in
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
script, its subcommands (currently `init`, `collect`, `combine`, `mark`, `send`
and `summarize`), and their parameters. Once you have completed the one-time
setup below, you'll be able to access the help via:
```
krummstab -h
krummstab <subcommand> -h
```
In the following I will go over the recommended workflow using the settings of
the Foundations of Artificial Intelligence lecture from the spring semester 2023
as an example.


## Requirements

- Python 3.10+: Our `pytest` setup tests Python versions 3.10, 3.11, and 3.12
  on an Ubuntu machine. Older Python versions will likely cause problems.
- [Xournal++](https://github.com/xournalpp/xournalpp) (optional): The script
  includes some convenient functionality when using Xournal++ to write feedback,
  but you can also do so by any other means.


## One-Time Setup

> 📝 We're assuming a Linux environment in the following. In case you are using
> macOS, we hope that the following instructions work without major differences.
> In case you are using Windows, we recommend using Krummstab on a native Python
> installation, but we don't provide instructions below. Installing inside a
> Windows Subsystem for Linux (WSL) is possible but not compatible with some
> features such as automatically opening all files that need to be marked and
> sending feedback per email.

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
to provide the paths to the files with the `-s <path to shared>` and `-i
<path to individual>` flags whenever you call `krummstab`.


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
First, download the submissions from ADAM and save the ZIP file in the marking
directory. (It's important that you only download the submissions after the ADAM
deadline has passed, so that all tutors have the same, complete pool of
submissions.) Our example directory `ki-fs23-marking`, with `Sheet 1.zip` being
the file downloaded from ADAM, should look like this:
```
.
├── .venv
├── config-individual.json
├── config-shared.json
└── Sheet 1.zip
```
Before running the command we have to look at the shared config to figure out if
we have to add information about exercises to the `init` command. In case we
only record points per sheet (`points_per: sheet`) this is not needed,
but if we record points per individual exercise (`points_per: exercise`), then
we have to provide information about the exercises on the sheet by either
providing the total number of exercises via `-n <num exercises>` (in case of
`marking_mode: sheet`) or via `-e <list of exercise numbers>` (in case of
`marking_mode: exercise`).

Assuming our configuration sets `points_per: exercise` and `marking_mode:
sheet`, we can now finally make the script do something useful by running:
```
krummstab init -n 5 -t sheet01 "Sheet 1.zip"
```
This will unzip the submissions and prepare them for marking. The flag `-t` is
optional and takes the name of the directory the submissions should be extracted
to. By default it's the name of the ZIP file, but I'm choosing to rename it in
order to get rid of the whitespace in the directory name. The directory should
now look something like this:
```
.
├── .venv
├── config-individual.json
├── config-shared.json
├── sheet01
│   ├── 12345_Muster_Müller
│   │   ├── feedback
│   │   │   └── feedback_tutor-name.pdf
│   │   └── Sheet1_MaxMuster_MayaMueller.pdf
│   .
│   ├── DO_NOT_MARK_12346_Meier_Meyer
│   │   └── submission_exercise_sheet1.pdf
│   .
│   └── points_*.json
└── Sheet 1.zip
```
As you may have guessed, the submissions you need to mark are those without the
`DO_NOT_MARK_` prefix. Those directories contain the files submitted by the
respective team, as well as a directory called `feedback`, which in turn
contains copies of the submitted files.

The idea is that you can give feedback by adding your comments to these copies
directly, and delete the ones you don't need to comment on. It is possible to
add a `--plain` or `-p` flag to the `init` command that prevents non-PDF files
from being copied into the feedback directories. When marking by exercise, this
is useful for the tutors that do not have to mark the programming exercises.

For the PDF feedback you can use whichever tool you like. If this tool adds
files to the feedback directory that you do not want to send to the students,
you can add their endings to the config file under the `ignore_feedback_suffix`
key. Marking with Xournal++ is supported by default: Simply set the value of the
`xopp` key in the config file to `true` to automatically create the relevant
`.xopp` files. If you like, you can use the `mark` command for easier feedback
writing, which is explained next.

While writing the feedback, you can keep track of the points the teams get in
the file `points_*.json`. In the case of plagiarism, write `Plagiarism` in the
place for the points.

### mark

This command allows you to mark all submissions at once with a specific
program such as Xournal++. It opens all relevant PDF feedback files or `.xopp`
files one after the other with the program that you can specify with the config
parameter `marking_command`. This parameter is a list of strings, starting with
the program command, with the following elements being arguments. Arguments that
would be separated by a space on the command line should be separate strings in
the list. One of these arguments has to be either `{xopp_file}`, `{pdf_file}`,
or `{all_pdf_files}`. In the case of `{xopp_file}` and `{pdf_file}`, `mark`
executes the command for each file to be marked while replacing the placeholder
argument with the path to that file. In the case of `{all_pdf_files}`, `mark`
executes the command one time while replacing the placeholder argument with the
list of all files to be marked. The default Xournal++ command would for example
look as follows in the config:
```json
"marking_command": ["xournalpp", "{xopp_file}"],
```
and is equivalent to running:
```shell
xournalpp <path to a file to be marked>
```
on the command line, one by one for each file to be marked. `mark` waits for the
process of the current marking command to finish before starting the next one.
By default, `mark` only opens the files of those teams for which points are
missing in the `points_*.json` file, use the `-f` flag to force `mark` to open
all files.

To run `mark` you need to provide the path to the directory created by the
`init` command which is `sheet01` in our running example:
```
krummstab mark sheet01
```
> 📝 On a native Windows installation you may have to add the parent directory
> of the executable you would like to use for marking to the `PATH` environment
> variable. If you do so for Xournal++ and set `xopp: true` in your individual
> config, then `mark` should work without changes to the `marking_command`
> option.

### collect
Once you have marked all the teams assigned to you and added their points to
the `points_*.json` file, you can run the next command, where `sheet01` is the
path to the directory created by the `init` command:
```
krummstab collect sheet01
```
This will create a ZIP archive in every feedback directory containing the
feedback for that team. A JSON file containing the individual points per student
is also generated.

In case you need make changes to the markings and rerun the collection step, use
the `-r` flag to overwrite existing feedback archives. When using Xournal++
(that is, when the `xopp` key is set to true), the `.xopp` files will be
exported automatically before collecting the feedback.

### combine
This command is only relevant for the `exercise` marking mode.
`TODO: Document this.`

### send
For the `static` marking mode, it is possible to directly send the feedback to
the students via e-mail. For this to work you have to be in the university
network, which likely means you'll have to connect to the university VPN. You
may find the `--dry-run` option useful, instead of sending the e-mails directly,
it only prints them so that you can double-check that everything looks as
expected.

### summarize
This command generates an Excel file that summarizes the students' marks. It
needs a path to a directory containing the individual marks JSON files:
```
krummstab summarize <path to a directory containing individual marks files>
```
If you use LibreOffice, it is possible that the formulas are not calculated
immediately. To calculate them, use the Recalculate Hard command in LibreOffice.
To access this command
- From the menu bar: Data > Calculate > Recalculate Hard
- From the keyboard: Ctrl + Shift + F9

To avoid having to do so manually, you can configure LibreOffice to always
recalculate upon opening a file. You can do so by setting the field "Excel 2007
and newer" under Tools > Options > LibreOffice Calc > Formula > Recalculation on
File Load to "Always recalculate" (or "Prompt user").


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
- `xopp`: if you use Xournal++ for marking, set the value to `true`; the
  relevant `xopp` files are then automatically created with the `init`
  subcommand and exported with the `collect` subcommand before the feedback is
  collected
- `ignore_feedback_suffix`: a list of extensions that should be ignored by the
  `collect` sub-command; this is useful if the tools you use for marking create
  files in the feedback folders that you don't want to send to the students
- `marking_command`: a list of strings that the `mark` subcommand should use,
  starting with program command, with the following elements being arguments;
  one argument has to be either `{xopp_file}` or `{pdf_file}`, which will be
  automatically replaced with file paths later; contains `xournalpp` with
  `{xopp_file}` by default

### General Settings
- `lecture_title`: lecture name to be printed in feedback emails
- `marking_mode`
    - `static`: student teams are assigned to a tutor who will mark all their
      submissions
    - `exercise`: with every sheet, tutors distribute the exercises and only
      mark those, but for all submissions
- `points_per`
    - `exercise`: tutors keep track how many points teams got for every exercise
    - `sheet`: tutors only keep track of the total number of points per sheet
- `min_point_unit`: a float denoting the smallest allowed point fraction, for
  example `0.5`, or `1`
- `tutor_list`: list to identify tutors, for example a list of first names
- `max_points_per_sheet`: a dictionary with all exercise sheet names as keys
  and their maximum possible points as values
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


## Handling Special Cases

There may be situations that require manual changes. This section provides
instructions for handling these special cases.

It is important to note that the teams in the shared config file are only used
for the `init` and `summarize` commands. After the `init` command, there is a
file with the name `submission.json` in each team folder that contains
information about the submission, including the team. The information in these
files is used for the other commands.

### Manually Adding Late Submissions
If you have already executed the `init` command and have already started to mark
the sheet, but there is a late submission that needs to be added, the
following steps are necessary:
1. Creating a new folder for the submission: In the directory generated by the
   `init` command, create a new directory for the submission. The folder name
   does not matter, let's say it's called `late_submission`.
2. Adding a `submission.json` file: Add a file with the name `submission.json`
   to `late_submission`. The internal structure of this file needs particular
   attributes:
   - `team`: list of students of the team, where each student entry is a list of
     the form `[ "first_name", "last_name", "email@stud.unibas.ch" ]`
   - `adam_id`: a string of numbers usually generated by ADAM, you can choose
     this arbitrarily
   - `relevant`: set the value to `true` to specify that you will mark this
     submission

     The structure should look similar to the following:
     ```json
     {
         "team": [
             [
                 "first_name1",
                 "last_name1",
                 "email1@stud.unibas.ch"
             ],
             [
                 "first_name2",
                 "last_name2",
                 "email2@stud.unibas.ch"
             ]
         ],
         "adam_id": "11910",
         "relevant": true
     }
     ```
3. Creating a feedback folder: Create a new subfolder in `late_submission` with
   the name `feedback`. When you have marked the team, you can add your feedback
   files here. You can add the original submitted files to `late_submission`,
   but this is not mandatory.
4. Modifying the `points_*.json` file: Add the team key with the points that the
   team gets to the `points_*.json` file. The team key consists of the ADAM ID
   you chose in step 1 and the alphabetically sorted last names of all team
   members in the following format: `<ADAM ID>_<Last-Name-1>_<Last-Name-2>`,
   that is, last names are capitalized and spaces within the last names are
   replaced by hyphens (`-`).

After completing these steps, the new submission will be processed as usual by
future calls to Krummstab, in particular by the `collect` and `send` commands.

### Multiple Submissions from the Same Team
There can be multiple submissions for the same team in ADAM. This can happen in
two ways, either

- team members submit separately without forming a team on ADAM, or
- two submissions with different file names are uploaded for the same ADAM team.

In the first case, Krummstab will create a separate submission directory for
each team member. To resolve this, you can create a new submission directory
according to the instructions [here](#manually-adding-late-submissions). In the
second case you will have to figure out which submission to mark, but for
Krummstab it only matters that the files you want to send as feedback are in the
`feedback` folder of the team's submission directory. To avoid this situation,
students should use the same file name when uploading an updated submission and
of course refrain from uploading two separate submissions.

### Only Sending Feedback to Some Teams
The `send` command only sends feedback to teams that are marked as being
`relevant` in the their respective `submission.json` file. To avoid sending
feedback to team Hans and Hanna Muster, manually edit the file
`00000_Muster_Muster/submission.json` and change `"relevant": true` to
`"relevant": false`. The marks for the team you entered in the points file the
will be sent to the assistant anyway, so you may want to remove the
corresponding entry in `points_*.json` or give 0 marks
explicitly. You can still later send the feedback through Krummstab by reverting
the changes above and marking the teams you already sent feedback to earlier as
"not relevant".

# Development

To set up for development, you have to
- clone this repository
- create a virtual environment at its root directory (`python3 -m venv .venv`)
- enter the virtual environment (`source .venv/bin/activate`)
- install the local version of Krummstab using `pip` (`pip install -e .`)
After this initial setup you simply have to make sure that you are in this same
virtual environment when making changes to Krummstab. If that is the case, all
changes to the source files will be directly reflected in any call to
`krummstab`.

There are some tests written in the `pytest` framework. This requires `pytest`,
which can be installed via `pip install pytest` for example, and
[Xournal++](https://github.com/xournalpp/xournalpp), which can be installed via
`sudo apt install xournalpp` on Ubuntu. Tests can then be executed by running
`pytest` in the root directory of the project (while in the virtual
environment).

All code should be formatted according to [Black](https://github.com/psf/black).
The only change from the default configuration is that we limit lines to 80
instead of 88 characters.
