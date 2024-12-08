import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime

@frappe.whiltelist()
def trigger_review_submission_mail(
    approver_email, project_owner_email, project_oc_number, project_name, subject
):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template(
        "/templates/db_review_submission.html",
        {
            "approver_first_name": approver.first_name,
            "approver_last_name": approver.last_name,
            "project_oc_number": project_oc_number,
            "project_name": project_name,
            "sent_by": f"{project_owner.first_name} {project_owner.last_name}",
        },
    )
    frappe.sendmail(
        recipients=approver_email,
        cc=project_owner_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Submit for review notification mail sent successfully"

@frappe.whitelist()
def trigger_review_resubmission_mail(
    approver_email,
    project_owner_email,
    project_oc_number,
    project_name,
    feedback_description,
    subject,
    attachments,
):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template(
        "/templates/db_review_resubmission.html",
        {
            "owner_first_name": project_owner.first_name,
            "owner_last_name": project_owner.last_name,
            "project_oc_number": project_oc_number,
            "project_name": project_name,
            "feedback_description": feedback_description,
            "approvar_name": f"{approver.first_name} {approver.last_name}",
        },
    )
    frappe.sendmail(
        recipients=project_owner_email,
        cc=approver_email,
        subject=subject,
        message=template,
        now=True,
        attachments=attachments,
    )
    return "Resubmit for review notification mail sent successfully"


@frappe.whitelist()
def trigger_review_approval_mail(
    approver_email, project_owner_email, project_oc_number, project_name, subject
):
    approver = frappe.get_doc("User", approver_email)
    project_owner = frappe.get_doc("User", project_owner_email)
    template = frappe.render_template(
        "/templates/db_review_approval.html",
        {
            "owner_first_name": project_owner.first_name,
            "owner_last_name": project_owner.last_name,
            "project_oc_number": project_oc_number,
            "project_name": project_name,
            "approvar_name": f"{approver.first_name} {approver.last_name}",
        },
    )
    frappe.sendmail(
        recipients=project_owner_email,
        cc=approver_email,
        subject=subject,
        message=template,
        now=True,
    )
    return "Approval notification mail sent successfully"


@frappe.whitelist()
def get_local_isolator_excel(): 
    payload = frappe.local.form_dict
    metadata = payload.get("metadata")
    project = payload.get("project")
    # division = payload.get("division")
    document_revision = payload.get("documentRevisions")
    project_id  = payload.get("project_id")
    status = payload.get("status")
    description = payload.get("description")
    electrical_load_list_data = payload.get("electrical_load_list_data")

    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "local_isolator_specification_template.xlsx.xlsx"
    )

    template_workbook = load_workbook(template_path)

    cover_sheet = template_workbook["COVER"]
    revision_sheet = template_workbook["REVISION"]
    isolator_sheet = template_workbook["ISOLATOR"]
    bom_list_sheet = template_workbook["BOM LIST"]

    # COVER

    division_name = metadata.get("division_name").upper()  # Get the division name and convert to uppercase
    # cover_sheet["A4"] = "411 026" just to ensure
    if division_name == "WWS SPG":
        cover_sheet["A3"] = "Water & Waste Solution".upper()  # Replace with desired text
        cover_sheet["A4"] = "411 026"
    elif division_name == "Enviro".upper():
        cover_sheet["A4"] = "411 026"
    else:
        cover_sheet["A3"] = division_name.upper()  # Otherwise, use the original division name

    
    cover_sheet["D6"] = "LOCAL ISOLATOR SPECIFICATION"
    cover_sheet["D7"] = project.get("client_name").upper()
    cover_sheet["D8"] = project.get("consultant_name").upper()
    cover_sheet["D9"] = project.get("project_name").upper()
    cover_sheet["D10"] = project.get("project_oc_number").upper()
    cover_sheet["D11"] = "LOCAL ISOLATOR SPECIFICATION"


    cover_sheet["B36"] = "0" # revision number (index or length - 1)
    cover_sheet["D36"] = "Not Released"
    cover_sheet["E36"] = "SP"
    cover_sheet["F36"] = "JS"
    cover_sheet["G36"] = "RBB"




    # REVISION

    start_row = 6
    

    if len(document_revision) > 1: 
        for idx, revision in enumerate(document_revision) :
            modified_revision_date = revision.get("modified")

            if modified_revision_date:
                modified_revision_date = "date1"
            else:
                modified_revision_date = "date2"
                revision_sheet[f"B{start_row + idx}"] = revision.get("idx")
                revision_sheet[f"D{start_row + idx}"] = modified_revision_date
                revision_sheet[f"E{start_row + idx}"] = revision.get("status")
        

    # revision_sheet["B6"] = "R0"
    # revision_sheet["D6"] = "DATE"
    # revision_sheet["E6"] = "Status"


    # ISOLATOR 


    isolator_sheet["E3"] = "type"
    isolator_sheet["E4"] = "enclosure"
    isolator_sheet["E5"] = "material"
    isolator_sheet["E6"] = "quantity"
    isolator_sheet["E7"] = "color shade"
    isolator_sheet["E8"] = "cable entry"

    local_push_button_status = "All"
    isolator_sheet["B20"] = f"{local_push_button_status} Local Push Button station shall have canopy on top."


    # BOM LIST

    for i in range(0,8):

        bom_list_sheet["A3"] = f"{i}-Sr."
        bom_list_sheet["B3"] = f"{i}-Motor TAg"
        bom_list_sheet["C3"] = f"{i}-Desctiption"
        bom_list_sheet["D3"] = f"{i}-kw rating"
        bom_list_sheet["E3"] = f"{i}-canopy required"
        bom_list_sheet["F3"] = f"{i}-part code"
        bom_list_sheet["G3"] = f"{i}-remark"
