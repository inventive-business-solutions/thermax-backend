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

const_revision_id = "st486uu99i"


@frappe.whitelist()
def get_local_isolator_excel(): 
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
    instruction_sheet = template_workbook["INSTRUCTION PAGE"]
    specification_sheet = template_workbook["SPECIFICATION"]
    lpbs_safe_sheet = template_workbook[" LPBS LIST SAFE AREA"]
    lpbs_hazard_sheet = template_workbook["LPBS LIST HAZARDOUS AREA "]
    selection_sheet = template_workbook["SELECTION CRITERIA"]


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
    static_document_list_data = frappe.get_doc("Static Document List", {"project_id":project_id}, "*").as_dict()

    lpbs_specifications_and_list = static_document_list_data.get("lpbs_specifications_and_list")


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

    lpbs_revision_data = frappe.get_doc(
        "LPBS Specification Revisions", const_revision_id, "*"
    ).as_dict()

    lpbs_data = lpbs_revision_data.get("lpbs_specification_data")
    safe_lpbs_data = {}
    hazard_lpbs_data = {}

    for data in lpbs_data:
        if data["area"] == "Safe":
            safe_lpbs_data = data
        else:
            hazard_lpbs_data = data


    
    specification_sheet["C3"] = safe_lpbs_data.get("lpbs_type")
    specification_sheet["C4"] = safe_lpbs_data.get("lpbs_enclosure")

    safe_lpbs_material = safe_lpbs_data.get("lpbs_material")
    # safe_fmi_enclosure_thickness = safe_lpbs_data.get("fmi_enclosure_thickness")
    safe_lpbs_cable_entry = safe_lpbs_data.get("ifm_cable_entry")

    if (
        safe_lpbs_material == "CRCA"
        or safe_lpbs_material == "SS 316"
        or safe_lpbs_material == "SS 306"
    ):
        # safe_lpbs_material = (
        #     f"{safe_lpbs_material}, {safe_fmi_enclosure_thickness} mm"
        # )
        safe_lpbs_cable_entry = f"{safe_lpbs_cable_entry}, 3 mm"
    elif safe_lpbs_material == "NA":
        safe_lpbs_material = "Not Applicable"


    specification_sheet["C5"] = safe_lpbs_material
    specification_sheet["C6"] = safe_lpbs_data.get("fmi_qty")
    specification_sheet["C7"] = safe_lpbs_data.get("ifm_isolator_color_shade")
    specification_sheet["C8"] = safe_lpbs_cable_entry
    specification_sheet["C9"] = safe_lpbs_data.get("canopy")
    specification_sheet["C10"] = safe_lpbs_data.get("canopy_type")


    hazard_fmi_enclouser_moc = hazard_lpbs_data.get("fmi_enclouser_moc")
    hazard_fmi_enclosure_thickness = hazard_lpbs_data.get("fmi_enclosure_thickness")
    hazard_ifm_cable_entry = hazard_lpbs_data.get("ifm_cable_entry")

    if (
        hazard_fmi_enclouser_moc == "CRCA"
        or hazard_fmi_enclouser_moc == "SS 316"
        or hazard_fmi_enclouser_moc == "SS 306"
    ):
        # hazard_fmi_enclouser_moc = (
        #     f"{hazard_fmi_enclouser_moc}, {hazard_fmi_enclosure_thickness} mm"
        # )
        hazard_ifm_cable_entry = f"{hazard_ifm_cable_entry}, 3 mm"
    elif hazard_fmi_enclouser_moc == "NA":
        hazard_fmi_enclouser_moc = "Not Applicable"

    specification_sheet["D3"] = hazard_lpbs_data.get("fmi_type")
    specification_sheet["D4"] = hazard_lpbs_data.get("fmi_ip_protection")
    specification_sheet["D5"] = hazard_fmi_enclouser_moc
    specification_sheet["D6"] = hazard_lpbs_data.get("fmi_qty")
    specification_sheet["D7"] = hazard_lpbs_data.get("ifm_isolator_color_shade")
    specification_sheet["D8"] = hazard_ifm_cable_entry
    specification_sheet["D9"] = hazard_lpbs_data.get("canopy")
    specification_sheet["D10"] = hazard_lpbs_data.get("canopy_type")
    
    local_isolator_motor_details_data = lpbs_revision_data.get("local_isolator_motor_details_data")
    safe_motor_details = []
    hazard_motor_details = []

    for i in range(len(local_isolator_motor_details_data)):
        if local_isolator_motor_details_data[i].get("area") == "Safe":
            safe_motor_details.append(local_isolator_motor_details_data[i])
        else:
            hazard_motor_details.append(local_isolator_motor_details_data[i])

    index = 3

    for i in range(len(safe_motor_details)):
        # area_data = local_isolator_motor_details_data[i].get("area")
        # if area_data == "Safe":
        lpbs_safe_sheet[f"A{index}"] = i + 1
        lpbs_safe_sheet[f"B{index}"] = safe_motor_details[i].get("tag_number")
        lpbs_safe_sheet[f"C{index}"] = safe_motor_details[i].get("service_description")
        lpbs_safe_sheet[f"D{index}"] = safe_motor_details[i].get("working_kw")
        lpbs_safe_sheet[f"E{index}"] = ""
        motor_location = safe_motor_details[i].get("motor_location")
        area = safe_motor_details[i].get("area")

        lpbs_safe_sheet[f"G{index}"] = motor_location

        if area == "Safe":
            canopy = safe_lpbs_data.get("canopy")
        else: 
            canopy = hazard_lpbs_data.get("canopy")


        canopy_required = ""
        if canopy == "All":
            canopy_required = "Yes"
        else: 
            if canopy == "OUTDOOR" and motor_location == "OUTDOOR":
                canopy_required = "Yes"
            else:
                canopy_required = "No"
            

        lpbs_safe_sheet[f"F{index}"] = canopy_required
        lpbs_safe_sheet[f"H{index}"] = safe_motor_details[i].get("gland_size")
        index = index + 1

    lpbs_safe_sheet[f"C{index + 5}"] = "Total Quantity"
    lpbs_safe_sheet[f"D{index + 5}"] = int(len(safe_motor_details))
    lpbs_safe_sheet[f"F{index + 5}"] = "Nos"

    for i in range(len(hazard_motor_details)):
        # area_data = local_isolator_motor_details_data[i].get("area")
        # if area_data == "Hazardous":
        lpbs_hazard_sheet[f"A{index}"] = i + 1
        lpbs_hazard_sheet[f"B{index}"] = hazard_motor_details[i].get("tag_number")
        lpbs_hazard_sheet[f"C{index}"] = hazard_motor_details[i].get("service_description")
        lpbs_hazard_sheet[f"D{index}"] = hazard_motor_details[i].get("working_kw")
        lpbs_hazard_sheet[f"E{index}"] = ""
        motor_location = hazard_motor_details[i].get("motor_location")
        area = hazard_motor_details[i].get("area")

        lpbs_hazard_sheet[f"K{index}"] = motor_location

        if area == "Safe":
            canopy = safe_lpbs_data.get("canopy")
        else: 
            canopy = hazard_lpbs_data.get("canopy")


        canopy_required = ""
        if canopy == "All":
            canopy_required = "Yes"
        else: 
            if canopy == "OUTDOOR" and motor_location == "OUTDOOR":
                canopy_required = "Yes"
            else:
                canopy_required = "No"
            

        lpbs_hazard_sheet[f"F{index}"] = canopy_required
        lpbs_hazard_sheet[f"G{index}"] = hazard_motor_details[i].get("standard")
        lpbs_hazard_sheet[f"H{index}"] = hazard_motor_details[i].get("zone")
        lpbs_hazard_sheet[f"I{index}"] = hazard_motor_details[i].get("gas_group")
        lpbs_hazard_sheet[f"J{index}"] = hazard_motor_details[i].get("temprature_class")
        lpbs_hazard_sheet[f"L{index}"] = hazard_motor_details[i].get("gland_size")
        index = index + 1

    lpbs_hazard_sheet[f"C{index + 5}"] = "Total Quantity"
    lpbs_hazard_sheet[f"D{index + 5}"] = int(len(hazard_motor_details))
    lpbs_hazard_sheet[f"F{index + 5}"] = "Nos"

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")