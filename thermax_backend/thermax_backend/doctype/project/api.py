import frappe

@frappe.whitelist()
def trigger_approver_notification_mail(email, project_owner, project_oc_number, project_name, sent_by, subject):
    user = frappe.get_doc("User", email)
    template = frappe.render_template('/templates/approver_assignment.html', {
        "first_name": user.first_name,
        "last_name": user.last_name,
        "email": email,
        "project_owner": project_owner,
        "project_oc_number": project_oc_number,
        "project_name": project_name,
        "sent_by": sent_by
    })
    frappe.sendmail(
        recipients=email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Notification mail sent to approver successfully"