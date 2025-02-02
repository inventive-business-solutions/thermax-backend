import frappe


@frappe.whitelist()
def trigger_next_reset_password(email, reset_link, sent_by):
    user = frappe.get_doc("User", email)
    template = frappe.render_template(
        "/templates/nextauth_reset_password.html",
        {
            "first_name": user.first_name,
            "last_name": user.last_name,
            "link": reset_link,
            "sent_by": sent_by,
        },
    )
    frappe.sendmail(
        recipients=email,
        subject="Password Reset",
        message=template,
        now=True,
    )
    return "Password reset link has been sent to your email."


@frappe.whitelist()
def trigger_email_verification_mail(email, division_name, verification_link, sent_by):
    user = frappe.get_doc("User", email)
    template = frappe.render_template(
        "/templates/email_verification_template.html",
        {
            "first_name": user.first_name,
            "last_name": user.last_name,
            "division_name": division_name,
            "verification_link": verification_link,
            "sent_by": sent_by,
        },
    )
    frappe.sendmail(
        recipients=email,
        subject="Verify your email account",
        message=template,
        now=True,
    )
    return "Email verification link has been sent to your email."


@frappe.whitelist()
def trigger_send_credentials(
    recipient_email, cc_email, password, division_name, is_superuser, sent_by, subject
):
    user = frappe.get_doc("User", recipient_email)
    template = frappe.render_template(
        "/templates/send_credentials_template.html",
        {
            "first_name": user.first_name,
            "last_name": user.last_name,
            "email": recipient_email,
            "password": password,
            "division_name": division_name,
            "is_superuser": bool(is_superuser),
            "sent_by": sent_by,
        },
    )
    frappe.sendmail(
        recipients=recipient_email,
        cc=cc_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Account credentials has been sent to your email."


@frappe.whitelist()
def trigger_delete_user(
    recipient_email, cc_email, subject, division_name, is_superuser, sent_by
):
    user = frappe.get_doc("User", recipient_email)
    template = frappe.render_template(
        "/templates/delete_superuser_template.html",
        {
            "first_name": user.first_name,
            "last_name": user.last_name,
            "division_name": division_name,
            "is_superuser": bool(is_superuser),
            "sent_by": sent_by,
        },
    )
    frappe.sendmail(
        recipients=recipient_email,
        cc=cc_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "User has been deleted successfully."


@frappe.whitelist()
def get_user_by_role():
    user = frappe.get_all(
        "Thermax Extended User", filters={"is_superuser": True}, fields=["*"]
    )
    return user
