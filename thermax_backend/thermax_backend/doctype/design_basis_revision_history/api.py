import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime

from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.cover_sheet import (
    create_cover_sheet,
)
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.design_basis_sheet import (
    get_design_basis_sheet,
)
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.enviro_ipg_db_excel import (
    get_enviro_ipg_db_excel,
)
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.heating_db_excel import (
    get_heating_db_excel,
)
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.spg_db_excel import (
    get_spg_db_excel,
)


@frappe.whitelist()
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
def get_design_basis_excel():
    # Retrieve the payload from the request
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")
    revision_data = frappe.get_doc(
        "Design Basis Revision History", revision_id
    ).as_dict()
    project_id = revision_data.get("project_id")
    project_data = frappe.get_doc("Project", project_id).as_dict()

    division_name = project_data.get("division")

    template_path = ""

    if division_name == "Heating":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "heating_design_basis_template.xlsx"
        )
    elif division_name == "WWS SPG":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "spg_design_basis_template.xlsx"
        )
    elif division_name == "Enviro" or division_name == "WWS IPG":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "enviro_ipg_design_basis_template.xlsx"
        )
    else:
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "heating_load_list_template.xlsx"
        )

    template_workbook = load_workbook(template_path)
    cover_sheet = template_workbook["COVER"]
    design_basis_sheet = template_workbook["Design Basis"]
    mcc_sheet = template_workbook["MCC"]
    pcc_sheet = template_workbook["PCC"]
    mcc_cum_plc_sheet = template_workbook["MCC CUM PLC"]

    make_of_components_data = frappe.db.get_list(
        "Design Basis Make of Component", {"revision_id": revision_id}, "*"
    )
    make_of_components_data = make_of_components_data[0]

    cover_sheet = create_cover_sheet(
        cover_sheet, project_data, revision_data, division_name
    )
    design_basis_sheet = get_design_basis_sheet(
        design_basis_sheet=design_basis_sheet,
        project_id=project_id,
        revision_id=revision_id,
        division_name=division_name,
        make_of_components_data=make_of_components_data,
    )

    if division_name == "Enviro" or division_name == "WWS IPG":
        template_workbook = get_enviro_ipg_db_excel(
            template_workbook=template_workbook,
            mcc_sheet=mcc_sheet,
            pcc_sheet=pcc_sheet,
            mcc_cum_plc_sheet=mcc_cum_plc_sheet,
            project_data=project_data,
            make_of_components_data=make_of_components_data,
            revision_id=revision_id,
        )
    elif division_name == "Heating":
        template_workbook = get_heating_db_excel(
            template_workbook=template_workbook,
            mcc_sheet=mcc_sheet,
            pcc_sheet=pcc_sheet,
            mcc_cum_plc_sheet=mcc_cum_plc_sheet,
            project_data=project_data,
            make_of_components_data=make_of_components_data,
            revision_id=revision_id,
        )
    elif division_name == "WWS SPG":
        template_workbook = get_spg_db_excel(
            template_workbook=template_workbook,
            mcc_sheet=mcc_sheet,
            pcc_sheet=pcc_sheet,
            mcc_cum_plc_sheet=mcc_cum_plc_sheet,
            project_data=project_data,
            make_of_components_data=make_of_components_data,
            revision_id=revision_id,
        )
    else:
        template_workbook = get_heating_db_excel(
            template_workbook=template_workbook,
            mcc_sheet=mcc_sheet,
            pcc_sheet=pcc_sheet,
            mcc_cum_plc_sheet=mcc_cum_plc_sheet,
            project_data=project_data,
            make_of_components_data=make_of_components_data,
            revision_id=revision_id,
        )

    template_workbook.remove(mcc_sheet)
    template_workbook.remove(pcc_sheet)
    template_workbook.remove(mcc_cum_plc_sheet)

    # Load the workbook from the template path
    # template_workbook.save("design_basis.xlsx")

    # Create a BytesIO stream to save the workbook
    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    # Prepare the response for file download
    frappe.local.response.filename = "generated_design_basis.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
