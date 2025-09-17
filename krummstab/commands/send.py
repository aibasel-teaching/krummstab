import logging
import mimetypes
import os
import pathlib
import smtplib
import textwrap

from email.message import EmailMessage
from getpass import getpass

from .. import config, errors, sheets, strings, utils


def add_attachment(mail: EmailMessage, path: pathlib.Path) -> None:
    """
    Add a file as attachment to an email.
    This is copied from Patrick's/Silvan's script, not entirely sure what's
    going on here.
    """
    assert path.exists()
    ctype, encoding = mimetypes.guess_type(path)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(path, "rb") as fp:
        mail.add_attachment(
            fp.read(),
            maintype=maintype,
            subtype=subtype,
            filename=os.path.basename(path),
        )


def construct_email(
    receivers: list[str],
    cc: list[str],
    subject: str,
    content: str,
    sender: str,
    attachment: pathlib.Path,
) -> EmailMessage:
    """
    Construct an email object.
    """
    assert isinstance(receivers, list)
    assert subject and content and sender and attachment
    mail = EmailMessage()
    mail.set_content(content)
    mail["Subject"] = subject
    mail["From"] = sender
    mail["To"] = ", ".join(receivers)
    if cc:
        mail["Cc"] = ", ".join(cc)
    add_attachment(mail, attachment)
    return mail


def email_to_text(email: EmailMessage) -> None:
    to = email["To"]
    cc = email["CC"]
    subject = email["Subject"]
    content = ""
    attachments = []
    for part in email.walk():
        if part.is_attachment():
            attachments.append(part.get_filename())
        elif not part.is_multipart():
            content += part.get_content()
    lines = []

    def format_line(left: str, right: str) -> str:
        return f"\033[0;34m{left}\033[0m {right}"

    lines.append(format_line("To:", to))
    if cc:
        lines.append(format_line("CC:", cc))
    if attachments:
        lines.append(format_line("Attachments:", ", ".join(attachments)))
    lines.append(format_line("Subject:", subject))
    lines.append(format_line("Start Body", ""))
    if content[-1] == "\n":
        content = content[:-1]
    lines.append(content)
    lines.append(format_line("End Body", ""))
    return "\n".join(lines)


def print_emails(emails: list[EmailMessage]) -> None:
    email_strings = [email_to_text(email) for email in emails]
    print(
        f"{strings.SEPARATOR_LINE}"
        f"{strings.SEPARATOR_LINE.join(email_strings)}"
        f"{strings.SEPARATOR_LINE}"
    )


def send_messages(
    emails: list[EmailMessage], _the_config: config.Config
) -> None:
    with smtplib.SMTP(_the_config.smtp_url, _the_config.smtp_port) as smtp:
        smtp.starttls()
        if _the_config.smtp_user:
            logging.warning(
                "The setting 'smtp_user' should probably be empty for the"
                " 'send' command to work, trying anyway."
            )
            password = getpass("Email password: ")
            smtp.login(_the_config.smtp_user, password)
        for email in emails:
            logging.info(f"Sending email to {email['To']}")
            # During testing, I didn't manage to trigger the exceptions below.
            # Additionally `refused_recipients` was always empty, even when the
            # documentation of smtplib states that it should be populated when
            # some but not all of the recipients are refused. Instead I always
            # get receive an email from the Outlook server containing the error
            # message.
            try:
                refused_recipients = smtp.send_message(email)
            except smtplib.SMTPRecipientsRefused:
                logging.warning(
                    f"Email to '{email['To']}' failed to deliver because all"
                    " recipients were refused."
                )
            except smtplib.SMTPSenderRefused:
                logging.critical(
                    "Email sender was refused, failed to deliver any emails."
                )
            except (
                smtplib.SMTPHeloError,
                smtplib.SMTPDataError,
                smtplib.SMTPNotSupportedError,
            ):
                logging.warning(
                    f"Email to '{email['To']}' failed to deliver because of"
                    " some weird error."
                )
            for refused_recipient, (
                smtp_error,
                error_message,
            ) in refused_recipients.items():
                logging.warning(
                    f"Email to '{refused_recipient}' failed to deliver because"
                    " the recipient was refused with the SMTP error code"
                    f" '{smtp_error}' and the message '{error_message}'."
                )
        logging.info("Done sending emails.")


def get_team_email_subject(
    _the_config: config.Config, sheet: sheets.Sheet
) -> str:
    """
    Builds the email subject.
    """
    return f"Feedback {sheet.name} | {_the_config.lecture_title}"


def get_assistant_email_subject(
    _the_config: config.Config, sheet: sheets.Sheet
) -> str:
    """
    Builds the email subject.
    """
    return f"Marks for {sheet.name} | {_the_config.lecture_title}"


def get_email_greeting(name_list: list[str]) -> str:
    """
    Builds the first line of the email.
    """
    # Only keep one name per entry, "Hans Jakob" becomes "Hans".
    name_list = [name.split(" ")[0] for name in name_list]
    name_list.sort()
    assert len(name_list) > 0
    if len(name_list) == 1:
        names = name_list[0]
    elif len(name_list) == 2:
        names = name_list[0] + " and " + name_list[1]
    else:
        names = ", ".join(name_list[:-1]) + ", and " + name_list[-1]
    return "Dear " + names + ","


def get_team_email_content(
    name_list: list[str], _the_config: config.Config, sheet: sheets.Sheet
) -> str:
    """
    Builds the body of the email that sends feedback to students.
    """
    return textwrap.dedent(
        f"""
    {get_email_greeting(name_list)}

    Please find feedback on your submission for {sheet.name} in the attachment.
    If you have any questions, you can contact us in the exercise session or by replying to this email (reply to all).

    Best,
    {_the_config.email_signature}"""  # noqa
    )[
        1:
    ]  # Removes the leading newline.


def get_assistant_email_content(
    _the_config: config.Config, sheet: sheets.Sheet
) -> str:
    """
    Builds the body of the email that sends the points to the assistant.
    """
    return textwrap.dedent(
        f"""
    Dear assistant for {_the_config.lecture_title},

    Please find my marks for {sheet.name} in the attachment.

    Best,
    {_the_config.email_signature}"""
    )[
        1:
    ]  # Removes the leading newline.


def get_assistant_email_attachment_path(
    _the_config: config.Config, sheet: sheets.Sheet
) -> pathlib.Path:
    """
    Instead of sending the regular marks file where points are listed per team
    to the assistant, we send the individual marks file that lists points per
    student. The idea is that the assistent can collect these files and use the
    `summarize` command to generate an overview.
    """
    return sheet.get_individual_marks_file_path(_the_config)


def create_email_to_team(
    submission, _the_config: config.Config, sheet: sheets.Sheet
):
    team_first_names = submission.team.get_first_names()
    team_emails = submission.team.get_emails()
    if _the_config.marking_mode == "exercise":
        feedback_file_path = submission.get_combined_feedback_file()
    elif _the_config.marking_mode == "static":
        feedback_file_path = submission.get_collected_feedback_path()
    else:
        errors.unsupported_marking_mode_error(_the_config.marking_mode)
    return construct_email(
        list(team_emails),
        _the_config.feedback_email_cc,
        get_team_email_subject(_the_config, sheet),
        get_team_email_content(team_first_names, _the_config, sheet),
        _the_config.tutor_email,
        feedback_file_path,
    )


def create_email_to_assistant(_the_config: config.Config, sheet: sheets.Sheet):
    return construct_email(
        [_the_config.assistant_email],
        _the_config.feedback_email_cc,
        get_assistant_email_subject(_the_config, sheet),
        get_assistant_email_content(_the_config, sheet),
        _the_config.tutor_email,
        get_assistant_email_attachment_path(_the_config, sheet),
    )


def send(_the_config: config.Config, args) -> None:
    """
    After the collection step finished successfully, send the feedback to the
    students via email. This currently only works if the tutor's email account
    is whitelisted for the smtp-ext.unibas.ch server, or if the tutor uses
    smtp.unibas.ch with an empty smtp_user.
    """
    # Prepare.
    sheet = sheets.Sheet(args.sheet_root_dir)
    # Send emails.
    emails: list[EmailMessage] = []
    for submission in sheet.get_relevant_submissions():
        emails.append(create_email_to_team(submission, _the_config, sheet))
    # TODO: As of now the plan is to only send assistant emails if the marking
    # mode is "static" because there the assistant collects the points
    # centrally. In case of "exercise", we plan to distribute the point files
    # through the share_archives, so there is no need to send an email to the
    # assistant, but this may change in the future.
    if _the_config.marking_mode != "exercise" and _the_config.assistant_email:
        emails.append(create_email_to_assistant(_the_config, sheet))
    logging.info(f"Drafted {len(emails)} email(s).")
    if args.dry_run:
        logging.info("Sending emails now would send the following emails:")
        print_emails(emails)
        logging.info("No emails sent.")
    else:
        print_emails(emails)
        really_send = utils.query_yes_no(
            (
                f"Do you really want to send the {len(emails)} email(s) "
                "printed above?"
            ),
            default=False,
        )
        if really_send:
            send_messages(emails, _the_config)
        else:
            logging.info("No emails sent.")
