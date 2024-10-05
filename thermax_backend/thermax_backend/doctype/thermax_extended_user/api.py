import frappe

@frappe.whitelist()
def trigger_next_reset_password(email, reset_link, sent_by):
    user = frappe.get_doc("User", email)
    template = frappe.render_template('/templates/nextauth_reset_password.html', {
        "first_name": user.first_name,
        "last_name": user.last_name,
        "link": reset_link,
        "sent_by": sent_by
    })
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
    template = frappe.render_template('/templates/email_verification_template.html', {
        "first_name": user.first_name,
        "last_name": user.last_name,
        "division_name": division_name,
        "verification_link": verification_link,
        "sent_by": sent_by
    })
    frappe.sendmail(
        recipients=email,
        subject="Verify your email account",
        message=template,
        now=True,
    )
    return "Email verification link has been sent to your email."

@frappe.whitelist()
def trigger_send_credentials(email, password, sent_by):
    user = frappe.get_doc("User", email)
    template = frappe.render_template('/templates/send_credentials_template.html', {
        "first_name": user.first_name,
        "last_name": user.last_name,
        "email": email,
        "password": password,
        "sent_by": sent_by
    })
    frappe.sendmail(
        recipients=email,
        subject="Welcome to EniMax",
        message=template,
        now=True,
    )
    return "Account credentials has been sent to your email."