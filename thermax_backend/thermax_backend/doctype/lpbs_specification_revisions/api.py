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
def get_lpbs_specification_excel(): 
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    lpbs_specifications_revision_data = frappe.get_doc(
        "LPBS Specification Revisions",
        revision_id,
        "*"
    ).as_dict()
    
    project_id = lpbs_specifications_revision_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", {"project_id": project_id}
    ).as_dict()

    # Loading the workbook 
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "lpbs_specification_template.xlsx"
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
    instruction_sheet = template_workbook["INSTRUCTION PAGE"]
    specification_sheet = template_workbook["SPECIFICATION"]
    lpbs_safe_sheet = template_workbook[" LPBS LIST SAFE AREA"]
    lpbs_hazard_sheet = template_workbook["LPBS LIST HAZARDOUS AREA "]
    selection_sheet = template_workbook["SELECTION CRITERIA"]


    # # cover page sheet populating

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
    static_document_list_data = frappe.get_doc("Static Document List", {"project_id":project_id}, "*").as_dict()

    lpbs_specifications_and_list = static_document_list_data.get("lpbs_specifications_and_list")
    is_safe_lpbs_selected = lpbs_specifications_revision_data.get("is_safe_lpbs_selected")
    is_hazardous_lpbs_selected = lpbs_specifications_revision_data.get("is_hazardous_lpbs_selected")


    cover_sheet["A3"] = division_name.upper()
    # cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = lpbs_specifications_and_list

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


    # # ISOLATOR SHEET 

    def num_to_string(value):
        if value == 1 or value == "1":
            return "Applicable"
        return "Not Applicable"


    def na_to_string(value):
        if "NA" in value or value is None:
            return "Not Applicable"
        return value
    # Fetch the Design Basis revision data (then isolator data form that)

    lpbs_revision_data = frappe.get_doc(
        "LPBS Specification Revisions", revision_id, "*"
    ).as_dict()

    lpbs_specification = lpbs_revision_data.get("lpbs_specification_data")
    lpbs_specification_data = lpbs_specification[0]
    config_data = lpbs_specification[1]
    lpbs_specification_motor_details = lpbs_revision_data.get("lpbs_specifications_motor_details")

    safe_lpbs_type = lpbs_specification_data.get("safe_lpbs_type")
    safe_lpbs_ip_protection = lpbs_specification_data.get("safe_lpbs_ip_protection")
    safe_lpbs_moc = lpbs_specification_data.get("safe_lpbs_moc")
    safe_lpbs_thickness = lpbs_specification_data.get("safe_lpbs_thickness")
    safe_lpbs_quantity = lpbs_specification_data.get("safe_lpbs_quantity")
    safe_lpbs_color_shade = lpbs_specification_data.get("safe_lpbs_color_shade")
    safe_lpbs_cable_entry = "Bottom" # safe cabel entry
    safe_lpbs_canopy = lpbs_specification_data.get("safe_lpbs_canopy")
    safe_lpbs_canopy_type = lpbs_specification_data.get("safe_lpbs_canopy_type")

    if (
        safe_lpbs_moc == "CRCA"
        or safe_lpbs_moc == "SS 316"
        or safe_lpbs_moc == "SS 306"
    ):
        safe_lpbs_moc = (
            f"{safe_lpbs_moc}, {safe_lpbs_thickness}"
        )
        safe_lpbs_cable_entry = f"{safe_lpbs_cable_entry}, 3 mm"
    elif safe_lpbs_moc == "NA":
        safe_lpbs_moc = "Not Applicable"

    if int(is_safe_lpbs_selected) == 0:
        safe_lpbs_type = "Not Applicable"
        safe_lpbs_ip_protection = "Not Applicable"
        safe_lpbs_moc = "Not Applicable"
        safe_lpbs_quantity = "Not Applicable"
        safe_lpbs_color_shade = "Not Applicable"
        safe_lpbs_cable_entry = "Not Applicable"
        safe_lpbs_canopy = "Not Applicable"
        safe_lpbs_canopy_type = "Not Applicable"

    specification_sheet["C3"] = safe_lpbs_type
    specification_sheet["C4"] = safe_lpbs_ip_protection
    specification_sheet["C5"] = safe_lpbs_moc
    specification_sheet["C6"] = safe_lpbs_quantity
    specification_sheet["C7"] = safe_lpbs_color_shade
    specification_sheet["C8"] = safe_lpbs_cable_entry
    specification_sheet["C9"] = safe_lpbs_canopy
    specification_sheet["C10"] = safe_lpbs_canopy_type


    hazardous_lpbs_type = lpbs_specification_data.get("hazardous_lpbs_type")
    hazardous_ip_protection = lpbs_specification_data.get("hazardous_ip_protection")
    hazardous_lpbs_moc = lpbs_specification_data.get("hazardous_lpbs_moc")
    hazardous_lpbs_thickness = lpbs_specification_data.get("hazardous_lpbs_thickness")
    hazardous_lpbs_qty = lpbs_specification_data.get("hazardous_lpbs_qty")
    hazardous_lpbs_color_shade = lpbs_specification_data.get("hazardous_lpbs_color_shade")
    hazardous_lpbs_cable_entry = "Bottom" #hazardous cable entry
    hazardous_lpbs_canopy = lpbs_specification_data.get("hazardous_lpbs_canopy")
    hazardous_lpbs_canopy_type = lpbs_specification_data.get("hazardous_lpbs_canopy_type")

    if (
        hazardous_lpbs_moc == "CRCA"
        or hazardous_lpbs_moc == "SS 316"
        or hazardous_lpbs_moc == "SS 306"
    ):
        hazardous_lpbs_moc = (
            f"{hazardous_lpbs_moc}, {hazardous_lpbs_thickness}"
        )
        hazardous_lpbs_cable_entry = f"{hazardous_lpbs_cable_entry}, 3 mm"
    elif hazardous_lpbs_moc == "NA":
        hazardous_lpbs_moc = "Not Applicable"

        
    if int(is_hazardous_lpbs_selected) == 0:
        hazardous_lpbs_type = "Not Applicable"
        hazardous_ip_protection = "Not Applicable"
        hazardous_lpbs_moc = "Not Applicable"
        hazardous_lpbs_qty = "Not Applicable"
        hazardous_lpbs_color_shade = "Not Applicable"
        hazardous_lpbs_cable_entry = "Not Applicable"
        hazardous_lpbs_canopy = "Not Applicable"
        hazardous_lpbs_canopy_type = "Not Applicable"

    specification_sheet["D3"] = hazardous_lpbs_type
    specification_sheet["D4"] = hazardous_ip_protection
    specification_sheet["D5"] = hazardous_lpbs_moc
    specification_sheet["D6"] = hazardous_lpbs_qty
    specification_sheet["D7"] = hazardous_lpbs_color_shade
    specification_sheet["D8"] = hazardous_lpbs_cable_entry
    specification_sheet["D9"] = hazardous_lpbs_canopy
    specification_sheet["D10"] = hazardous_lpbs_canopy_type

    # Push Button Color
    specification_sheet["C13"] = lpbs_specification_data.get("lpbs_push_button_start_color")
    specification_sheet["C14"] = lpbs_specification_data.get("lpbs_forward_push_button_start_color")
    specification_sheet["C15"] = lpbs_specification_data.get("lpbs_reverse_push_button_start_color")
    specification_sheet["C16"] = lpbs_specification_data.get("lpbs_push_button_ess_color")
    specification_sheet["C17"] = lpbs_specification_data.get("lpbs_speed_increase_color")
    specification_sheet["C18"] = lpbs_specification_data.get("lpbs_speed_decrease_color")
    specification_sheet["C19"] = lpbs_specification_data.get("lpbs_indication_lamp_start_color")
    specification_sheet["C20"] = lpbs_specification_data.get("lpbs_indication_lamp_stop_color")

    id = 22

    keys_with_yes = [key for key, value in config_data.items() if value == 'yes']

    for j in range(len(keys_with_yes)):
        specification_sheet[f"{id}"] = keys_with_yes[j]
        id = id + 1


    # motor details sheet 
    safe_motor_details = []
    hazard_motor_details = []

    for i in range(len(lpbs_specification_motor_details)):
        if lpbs_specification_motor_details[i].get("area") == "Safe":
            safe_motor_details.append(lpbs_specification_motor_details[i])
        else:
            hazard_motor_details.append(lpbs_specification_motor_details[i])

    index = 3

    for i in range(len(safe_motor_details)):
        lpbs_safe_sheet[f"A{index}"] = i
        lpbs_safe_sheet[f"B{index}"] = safe_motor_details[i].get("tag_number")
        lpbs_safe_sheet[f"C{index}"] = safe_motor_details[i].get("service_description")
        lpbs_safe_sheet[f"D{index}"] = safe_motor_details[i].get("working_kw")
        lpbs_safe_sheet[f"E{index}"] = safe_motor_details[i].get("lpbs_type")
        lpbs_safe_motor_location = safe_motor_details[i].get("motor_location")
        lpbs_safe_sheet[f"G{index}"] = lpbs_safe_motor_location

        canopy_required = ""
        
        if safe_lpbs_canopy == "All":
            canopy_required = "Yes"
        else: 
            if "OUT" in safe_lpbs_canopy and "OUT" in lpbs_safe_motor_location :
                canopy_required = "Yes"
            else:
                canopy_required = "No"

        lpbs_safe_sheet[f"F{index}"] = canopy_required
        lpbs_safe_sheet[f"H{index}"] = safe_motor_details[i].get("gland_size")

    type_count = {motor_type["lpbs_type"]: sum(1 for t in safe_motor_details if t["lpbs_type"] == motor_type["lpbs_type"]) for motor_type in safe_motor_details}
    a = 3
    for key, value in type_count.items():
        lpbs_safe_sheet[f"M{a}"] = key
        lpbs_safe_sheet[f"N{a}"] = value
        a = a + 1

    index = 3
    for i in  range(len(hazard_motor_details)):
        lpbs_hazard_sheet[f"A{index}"] = i
        lpbs_hazard_sheet[f"B{index}"] = hazard_motor_details[i].get("tag_number")
        lpbs_hazard_sheet[f"C{index}"] = hazard_motor_details[i].get("service_description")
        lpbs_hazard_sheet[f"D{index}"] = hazard_motor_details[i].get("working_kw")
        lpbs_hazard_sheet[f"E{index}"] = hazard_motor_details[i].get("lpbs_type")
        hazard_lpbs_motor_location = hazard_motor_details[i].get("motor_location")
        lpbs_hazard_sheet[f"K{index}"] = hazard_lpbs_motor_location
        canopy_required = ""
        
        if hazardous_lpbs_canopy == "All":
            canopy_required = "Yes"
        else: 
            if "OUT" in hazardous_lpbs_canopy and "OUT" in hazard_lpbs_motor_location :
                canopy_required = "Yes"
            else:
                canopy_required = "No"

        lpbs_hazard_sheet[f"F{index}"] = canopy_required
        lpbs_hazard_sheet[f"G{index}"] = hazard_motor_details[i].get("standard")
        lpbs_hazard_sheet[f"H{index}"] = hazard_motor_details[i].get("zone")
        lpbs_hazard_sheet[f"I{index}"] = hazard_motor_details[i].get("gas_group")
        lpbs_hazard_sheet[f"J{index}"] = hazard_motor_details[i].get("temperature_class")
        lpbs_hazard_sheet[f"L{index}"] = hazard_motor_details[i].get("gland_size")

    if int(is_safe_lpbs_selected) == 0:
        template_workbook.remove(lpbs_safe_sheet)
    
    if int(is_hazardous_lpbs_selected) == 0:
        template_workbook.remove(lpbs_hazard_sheet)

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")