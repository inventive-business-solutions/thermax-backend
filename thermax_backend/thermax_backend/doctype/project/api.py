import frappe
import smtplib

from email.message import EmailMessage

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


@frappe.whitelist()
def trigger_approver_notification_mail(
    approvar_email, creator_email, project_oc_number, project_name, sent_by, subject
):
    user = frappe.get_doc("User", approvar_email)
    project_creator = frappe.get_doc("User", creator_email)
    template = frappe.render_template(
        "/templates/approver_assignment.html",
        {
            "first_name": user.first_name,
            "last_name": user.last_name,
            "project_creator_first_name": project_creator.first_name,
            "project_creator_last_name": project_creator.last_name,
            "email": approvar_email,
            "project_oc_number": project_oc_number,
            "project_name": project_name,
            "sent_by": sent_by,
        },
    )
    frappe.sendmail(
        recipients=approvar_email,
        cc=creator_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Notification mail sent to approver successfully"


@frappe.whitelist()
def send_custom_mail(receiver_email):
    """
    Send Custom Mail
    """
    # SMTP server configuration
    smtp_server = "192.168.255.200"
    smtp_port = 25
    default_email = "noreply.enimax@thermaxglobal.com"

    # Email details
    sender_email = default_email
    subject = "Test Email from Python"
    body = "This is a test email sent from Python using the provided SMTP server."

    try:
        # Create the email
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = receiver_email
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))

        # Connect to the SMTP server and send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.sendmail(sender_email, receiver_email, message.as_string())

        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")
