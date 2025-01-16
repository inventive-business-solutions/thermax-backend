import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime

# @frappe.whiltelist()
# def trigger_review_submission_mail(
#     approver_email, project_owner_email, project_oc_number, project_name, subject
# ):
#     approver = frappe.get_doc("User", approver_email)
#     project_owner = frappe.get_doc("User", project_owner_email)
#     template = frappe.render_template(
#         "/templates/db_review_submission.html",
#         {
#             "approver_first_name": approver.first_name,
#             "approver_last_name": approver.last_name,
#             "project_oc_number": project_oc_number,
#             "project_name": project_name,
#             "sent_by": f"{project_owner.first_name} {project_owner.last_name}",
#         },
#     )
#     frappe.sendmail(
#         recipients=approver_email,
#         cc=project_owner_email,
#         subject=subject,
#         message=template,
#         now=True,
#     )
#     return "Submit for review notification mail sent successfully"

# @frappe.whitelist()
# def trigger_review_resubmission_mail(
#     approver_email,
#     project_owner_email,
#     project_oc_number,
#     project_name,
#     feedback_description,
#     subject,
#     attachments,
# ):
#     approver = frappe.get_doc("User", approver_email)
#     project_owner = frappe.get_doc("User", project_owner_email)
#     template = frappe.render_template(
#         "/templates/db_review_resubmission.html",
#         {
#             "owner_first_name": project_owner.first_name,
#             "owner_last_name": project_owner.last_name,
#             "project_oc_number": project_oc_number,
#             "project_name": project_name,
#             "feedback_description": feedback_description,
#             "approvar_name": f"{approver.first_name} {approver.last_name}",
#         },
#     )
#     frappe.sendmail(
#         recipients=project_owner_email,
#         cc=approver_email,
#         subject=subject,
#         message=template,
#         now=True,
#         attachments=attachments,
#     )
#     return "Resubmit for review notification mail sent successfully"


# @frappe.whitelist()
# def trigger_review_approval_mail(
#     approver_email, project_owner_email, project_oc_number, project_name, subject
# ):
#     approver = frappe.get_doc("User", approver_email)
#     project_owner = frappe.get_doc("User", project_owner_email)
#     template = frappe.render_template(
#         "/templates/db_review_approval.html",
#         {
#             "owner_first_name": project_owner.first_name,
#             "owner_last_name": project_owner.last_name,
#             "project_oc_number": project_oc_number,
#             "project_name": project_name,
#             "approvar_name": f"{approver.first_name} {approver.last_name}",
#         },
#     )
#     frappe.sendmail(
#         recipients=project_owner_email,
#         cc=approver_email,
#         subject=subject,
#         message=template,
#         now=True,
#     )
#     return "Approval notification mail sent successfully"


@frappe.whitelist()
def get_lpbs_excel(): 
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", revision_id
    ).as_dict()
    project_id = design_basis_revision_data.get("project_id")

    # Loading the workbook 
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "local_isolator_specification_template.xlsx"
    )
    template_workbook = load_workbook(template_path)

    # project data for cover page
    project_data = frappe.get_doc("Project", project_id).as_dict()

    project_description = design_basis_revision_data.get("description")
    project_status = design_basis_revision_data.get("status")
    owner = design_basis_revision_data.get("owner")

    division_name = project_data.get("division")
    project_name = project_data.get("project_name")
    project_oc_number = project_data.get("project_oc_number")
    approver = project_data.get("approver")
    client_name = project_data.get("client_name")
    consultant_name = project_data.get("consultant_name")
    modified = project_data.get("modified")

    # loading the sheets 

    cover_sheet = template_workbook["COVER"]
    instruction_name_sheet = template_workbook["INSTRUCTION PAGE"]
    specificaiton_sheet = template_workbook["SPECIFICATION"]
    lpbs_safe_sheet = template_workbook[" LPBS LIST SAFE AREA"]
    lpbs_safe_sheet = template_workbook["LPBS LIST HAZARDOUS AREA "]
    isolator_sheet = template_workbook["ISOLATOR"]
    isolator_safe_area_sheet = template_workbook["ISOLATOR  LIST SAFE AREA"]
    isolator_hazard_area_sheet = template_workbook["ISOLATOR LIST HAZARDOUS AREA"]


    # cover page sheet populating

    prepped_by_initial = frappe.db.get_value(
        "Thermax Extended User", owner, "name_initial"
    )
    checked_by_initial = frappe.db.get_value(
        "Thermax Extended User", approver, "name_initial"
    )
    super_user_initial = frappe.db.get_value(
        "Thermax Extended User",
        {"is_superuser": 1, "division": division_name},
        "name_initial",
    )

    revision_date = modified.strftime("%d-%m-%Y")
    revision_data_with_pid = frappe.db.get_list("Design Basis Revision History", {"project_id": project_id}, "*")


    cover_sheet["A3"] = division_name.upper()
    cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = "LOCAL ISOLATOR SPECIFICAITON LIST"

    index = 33

    for i in range(len(revision_data_with_pid) - 1, -1, -1):
        cover_sheet[f"B{index}"] = f"R{len(revision_data_with_pid) - i - 1}"
        cover_sheet[f"C{index}"] = revision_date
        if (len(revision_data_with_pid) - i - 1) == 0:
            project_description = "Issued for Approval"
        cover_sheet[f"D{index}"] = project_description
        cover_sheet[f"E{index}"] = prepped_by_initial
        cover_sheet[f"F{index}"] = checked_by_initial
        cover_sheet[f"G{index}"] = super_user_initial
        index = index - 1

    match division_name:
        case "Heating":
            cover_sheet["A4"] = "PUNE - 411 019"
        case "WWS SPG":
            cover_sheet["A3"] = "WATER & WASTE SOLUTION"
            cover_sheet["A4"] = "PUNE - 411 026"
        case "WWS IPG":
            cover_sheet["A3"] = "WATER & WASTE SOLUTION"
            cover_sheet["A4"] = "PUNE - 411 026"
        case "Enviro":
            cover_sheet["A4"] = "PUNE - 411 026"
        case _:
            cover_sheet["A4"] = "PUNE - 411 026"


    # ISOLATOR SHEET 

    def num_to_string(value):
        if value == 1 or value == "1":
            return "Applicable"
        return "Not Applicable"


    def na_to_string(value):
        if "NA" in value or value is None:
            return "Not Applicable"
        return value
    # Fetch the Design Basis revision data (then isolator data form that)

    common_config_data = frappe.db.get_list(
        "Common Configuration 2", {"revision_id": revision_id}, "*"
    )
    common_config_data = common_config_data[0]

    cc_is_field_motor_isolator_selected = common_config_data.get("is_field_motor_isolator_selected") 
    cc_is_safe_area_isolator_selected = common_config_data.get("is_safe_area_isolator_selected") 
    cc_safe_field_motor_type = common_config_data.get("safe_field_motor_type") 
    cc_hazardous_field_motor_type = common_config_data.get("hazardous_field_motor_type") 
    cc_safe_field_motor_enclosure = common_config_data.get("safe_field_motor_enclosure") 
    cc_hazardous_field_motor_enclosure = common_config_data.get("hazardous_field_motor_enclosure") 
    cc_safe_field_motor_material = common_config_data.get("safe_field_motor_material") 
    cc_hazardous_field_motor_material = common_config_data.get("hazardous_field_motor_material") 
    cc_safe_field_motor_thickness = common_config_data.get("safe_field_motor_thickness") 
    cc_hazardous_field_motor_thickness = common_config_data.get("hazardous_field_motor_thickness") 
    cc_safe_field_motor_qty = common_config_data.get("safe_field_motor_qty") 
    cc_hazardous_field_motor_qty = common_config_data.get("hazardous_field_motor_qty") 
    cc_safe_field_motor_isolator_color_shade = common_config_data.get("safe_field_motor_isolator_color_shade") 
    cc_hazardous_field_motor_isolator_color_shade = common_config_data.get("hazardous_field_motor_isolator_color_shade") 
    cc_safe_field_motor_cable_entry = common_config_data.get("safe_field_motor_cable_entry") 
    cc_hazardous_field_motor_cable_entry = common_config_data.get("hazardous_field_motor_cable_entry") 
    cc_safe_field_motor_canopy = common_config_data.get("safe_field_motor_canopy") 
    cc_hazardous_field_motor_canopy = common_config_data.get("hazardous_field_motor_canopy") 
    cc_safe_field_motor_canopy_type = common_config_data.get("safe_field_motor_canopy_type") 
    cc_hazardous_field_motor_canopy_type = common_config_data.get("hazardous_field_motor_canopy_type") 

    if int(cc_is_field_motor_isolator_selected) == 0 or int(cc_is_safe_area_isolator_selected) == 0:
        cc_safe_field_motor_type = "Not Applicable"
        cc_safe_field_motor_enclosure = "Not Applicable"
        cc_safe_field_motor_material = "Not Applicable"
        cc_safe_field_motor_qty = "Not Applicable"
        cc_safe_field_motor_isolator_color_shade = "Not Applicable"
        cc_safe_field_motor_cable_entry = "Not Applicable"
        cc_safe_field_motor_canopy = "Not Applicable"
        cc_safe_field_motor_canopy_type = "Not Applicable"

        cc_hazardous_field_motor_type = "Not Applicable"
        cc_hazardous_field_motor_enclosure = "Not Applicable"
        cc_hazardous_field_motor_material = "Not Applicable"
        cc_hazardous_field_motor_qty = "Not Applicable"
        cc_hazardous_field_motor_isolator_color_shade = "Not Applicable"
        cc_hazardous_field_motor_cable_entry = "Not Applicable"
        cc_hazardous_field_motor_canopy = "Not Applicable"
        cc_hazardous_field_motor_canopy_type = "Not Applicable"

    
    isolator_sheet["C3"] = cc_safe_field_motor_type
    isolator_sheet["C4"] = na_to_string(cc_safe_field_motor_enclosure)

    if cc_safe_field_motor_material == "CRCA" or cc_safe_field_motor_material == "SS 316" or cc_safe_field_motor_material == "SS 306":
        cc_safe_field_motor_material = f"{cc_safe_field_motor_material}, {cc_safe_field_motor_thickness} mm"
        cc_safe_field_motor_cable_entry = f"{cc_safe_field_motor_cable_entry}, 3 mm"
    elif cc_safe_field_motor_material == "NA":
        cc_safe_field_motor_material = "Not Applicable"

    isolator_sheet["C5"] = cc_safe_field_motor_material
    isolator_sheet["C6"] = na_to_string(cc_safe_field_motor_qty)
    isolator_sheet["C7"] = na_to_string(cc_safe_field_motor_isolator_color_shade)
    isolator_sheet["C8"] = cc_safe_field_motor_cable_entry
    isolator_sheet["C9"] = na_to_string(cc_safe_field_motor_canopy)
    isolator_sheet["C10"] = na_to_string(cc_safe_field_motor_canopy_type)


    isolator_sheet["D3"] = cc_hazardous_field_motor_type
    isolator_sheet["D4"] = na_to_string(cc_hazardous_field_motor_enclosure)
    isolator_sheet["D5"] = na_to_string(cc_hazardous_field_motor_material)
    isolator_sheet["D6"] = na_to_string(cc_hazardous_field_motor_qty)
    isolator_sheet["D7"] = na_to_string(cc_hazardous_field_motor_isolator_color_shade)
    isolator_sheet["D8"] = cc_hazardous_field_motor_cable_entry
    isolator_sheet["D9"] = na_to_string(cc_hazardous_field_motor_canopy)
    isolator_sheet["D10"] = na_to_string(cc_hazardous_field_motor_canopy_type)



    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")