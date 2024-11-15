import frappe

@frappe.whitelist()
def trigger_review_submission_mail(approver_email, project_owner_email, project_oc_number, project_name, subject):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template('/templates/db_review_submission.html', {
        "approver_first_name": approver.first_name,
        "approver_last_name": approver.last_name,
        "project_oc_number": project_oc_number,
        "project_name": project_name,
        "sent_by": f"{project_owner.first_name} {project_owner.last_name}",
    })
    frappe.sendmail(
        recipients=approver_email,  
        cc=project_owner_email,      
        subject=subject,
        message=template,
        now=True,
    )
    return "Submit for review notification mail sent successfully"

@frappe.whitelist()
def trigger_review_resubmission_mail(approver_email, project_owner_email, project_oc_number, project_name, feedback_description, subject, attachments):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template('/templates/db_review_resubmission.html', {
        "owner_first_name": project_owner.first_name,
        "owner_last_name": project_owner.last_name,
        "project_oc_number": project_oc_number,
        "project_name": project_name,
        "feedback_description": feedback_description,
        "approvar_name": f"{approver.first_name} {approver.last_name}",
    })
    frappe.sendmail(
        recipients=project_owner_email,  
        cc=approver_email,      
        subject=subject,
        message=template,
        now=True,
        attachments=attachments
    )
    return "Resubmit for review notification mail sent successfully"

@frappe.whitelist()
def trigger_review_approval_mail(approver_email, project_owner_email, project_oc_number, project_name,  subject):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template('/templates/db_review_approval.html', {
        "owner_first_name": project_owner.first_name,
        "owner_last_name": project_owner.last_name,
        "project_oc_number": project_oc_number,
        "project_name": project_name,
        "approvar_name": f"{approver.first_name} {approver.last_name}",
    })
    frappe.sendmail(
        recipients=project_owner_email,  
        cc=approver_email,      
        subject=subject,
        message=template,
        now=True,
    )
    return "Approval notification mail sent successfully"