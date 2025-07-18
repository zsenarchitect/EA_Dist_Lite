"""This module is for sending email. It is a wrapper for the Emailer app."""

import time
import EXE
import DATA_FILE
import IMAGE
import USER
import ENVIRONMENT
import TIME
import SPEAK

if ENVIRONMENT.IS_REVIT_ENVIRONMENT:
    from REVIT import REVIT_APPLICATION


def email(
    receiver_email_list,
    body,
    subject=ENVIRONMENT.PLUGIN_NAME + " Auto Email",
    body_folder_link_list=None,
    body_image_link_list=None,
    attachment_list=None,
):
    """Send email using the Emailer app.

    Args:
        receiver_email_list (list): List of email addresses.
        body (str): Body of the email.
        subject (str, optional): Subject of the email. Defaults to " Auto Email".
        body_folder_link_list (list, optional): List of folder links to be included in the email body. Defaults to None.
        body_image_link_list (list, optional): List of image links to be included in the email body. Defaults to None.
        attachment_list (list, optional): List of file paths to be attached to the email. Defaults to None.
    """
    
    if not body:
        print("Missing body of the email.....")
        return

    if not receiver_email_list:
        print("missing email receivers....")
        return

    if isinstance(receiver_email_list, str):
        print("Prefer list but ok.")
        print (receiver_email_list)
        receiver_email_list = receiver_email_list.rstrip().split(";")

    body = body.replace("\n", "<br>")

    data = {}
    data["receiver_email_list"] = receiver_email_list
    data["subject"] = subject
    data["body"] = body
    data["body_folder_link_list"] = body_folder_link_list
    data["body_image_link_list"] = body_image_link_list
    data["attachment_list"] = attachment_list
    data["logo_image_path"] = IMAGE.get_image_path_by_name("logo.png")
    DATA_FILE.set_data(data, "email_data")

    EXE.try_open_app("Emailer")
    SPEAK.speak(
        "enni-ed tab email is sent out. Subject line: {}".format(
            subject.lower().replace("ennead", "enni-ed ")
        )
    )


def email_error(
    traceback, tool_name, error_from_user, subject_line=ENVIRONMENT.PLUGIN_NAME + " Auto Email Error Log"
):
    """Send automated email when an error occurs.

    Args:
        traceback (str): Traceback of the error.
        tool_name (str): Name of the tool that caused the error.
        error_from_user (str): Error message from the user.
        subject_line (str, optional): Subject of the email. Defaults to " Auto Email Error Log".
    """
    t = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(int(time.time())))
    computer_name = ENVIRONMENT.get_computer_name()
    additional_note = ""
    
    try:
        if ENVIRONMENT.IS_REVIT_ENVIRONMENT:
            from pyrevit import versionmgr
            pyrvt_ver = versionmgr.get_pyrevit_version()
            nice_version = 'v{}'.format(pyrvt_ver.get_formatted())
            app_uptime = TIME.get_revit_uptime()
            

            app = REVIT_APPLICATION.get_app()

            try:
                doc_name = app.ActiveUIDocument.Document.Title
            except:
                try:
                    doc_name = REVIT_APPLICATION.get_doc().Title
                except:
                    doc_name = "N/A"

            additional_note = "pyRevit Version: {}\n\nRevit Version Build: {}\nRevit Version Name: {}\nDoc name:{}\n\nRevit UpTime: {}".format(
                nice_version,
                app.VersionBuild,
                app.VersionName,
                doc_name,
                app_uptime,
            )
        elif ENVIRONMENT.IS_RHINO_ENVIRONMENT:
            import rhinoscriptsyntax as rs
            import scriptcontext as sc

            additional_note = (
                "File in trouble:{}\nCommand history before disaster:\n{}".format(
                    sc.doc.Path or None, rs.CommandHistory()
                )
            )

        else:
            additional_note = ""
    except Exception as e:
        print(e)
        additional_note = str(e)
        
    body = "{}\nError happens on {}'s machine [{}] when running [{}].\n\nDetail below:\n{}\n\n{}".format(
        t, error_from_user, computer_name, tool_name, traceback, additional_note
    )

    developer_emails = []
    try:
        if ENVIRONMENT.IS_REVIT_ENVIRONMENT:
            developer_emails = USER.get_revit_developer_emails()
            if not developer_emails:  # Check if list is empty
                developer_emails = ["szhang@ennead.com"]
            
            # Check Revit uptime for notification
            if "h" in app_uptime and 50 < int(app_uptime.split("h")[0]):
                email_to_self(
                    subject="I am tired...Revit running non-stop for {}".format(app_uptime),
                    body="Hello,\nI have been running for {}.\nLet me rest and clear cache!\n\nDid you know that restarting your Revit regularly can improve performance?\nBest regard,\nYour poor Revit.".format(
                        app_uptime
                    ),
                )
        elif ENVIRONMENT.IS_RHINO_ENVIRONMENT:
            developer_emails = USER.get_rhino_developer_emails()

        if USER.IS_DEVELOPER:
            developer_emails = [USER.get_company_email_address()]
            
        # Ensure developer_emails is always a list
        if not isinstance(developer_emails, list):
            developer_emails = [developer_emails] if developer_emails else ["szhang@ennead.com"]
    except Exception as e:
        print("Error getting developer emails: {}".format(e))
        developer_emails = ["szhang@ennead.com"]  # Fallback to default

    email(
        receiver_email_list=developer_emails,
        body=body,
        subject=subject_line,
        body_folder_link_list=None,
        body_image_link_list=None,
        attachment_list=None,
    )


def email_to_self(
    subject=ENVIRONMENT.PLUGIN_NAME + " Auto Email to Self",
    body=None,
    body_folder_link_list=None,
    body_image_link_list=None,
    attachment_list=None,
):
    """Send email to self.

    Args:
        subject (str, optional): Subject of the email. Defaults to " Auto Email to Self".
        body (str, optional): Body of the email. Defaults to None.
        body_folder_link_list (list, optional): List of folder links to be included in the email body. Defaults to None.
        body_image_link_list (list, optional): List of image links to be included in the email body. Defaults to None.
        attachment_list (list, optional): List of file paths to be attached to the email. Defaults to None
    """
    email(
        receiver_email_list=[USER.get_company_email_address()],
        subject=subject,
        body=body,
        body_folder_link_list=body_folder_link_list,
        body_image_link_list=body_image_link_list,
        attachment_list=attachment_list,
    )


def unit_test():
    email_to_self(
        subject="Test Email for compiler",
        body="Happy Howdy. This is a quick email test to see if the base communication still working",
    )


if __name__ == "__main__":
    unit_test()