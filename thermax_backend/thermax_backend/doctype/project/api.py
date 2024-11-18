import frappe

@frappe.whitelist()
def trigger_approver_notification_mail(approvar_email, creator_email, project_oc_number, project_name, sent_by, subject):
    user = frappe.get_doc("User", approvar_email)
    project_creator = frappe.get_doc("User", creator_email)
    template = frappe.render_template('/templates/approver_assignment.html', {
        "first_name": user.first_name,
        "last_name": user.last_name,
        "project_creator_first_name": project_creator.first_name,
        "project_creator_last_name": project_creator.last_name,
        "email": approvar_email,
        "project_oc_number": project_oc_number,
        "project_name": project_name,
        "sent_by": sent_by
    })
    frappe.sendmail(
        recipients=approvar_email,
        cc=creator_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Notification mail sent to approver successfully"