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

# const_revision_id = "st486uu99i"


@frappe.whitelist()
def get_local_isolator_excel(): 
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    local_isolator_revisions_data = frappe.get_doc(
        "Local Isolator Revisions", revision_id, "*"
    ).as_dict()

    project_id = local_isolator_revisions_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", {"project_id":project_id}, "*"
    ).as_dict()

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

    # # loading the sheets 

    cover_sheet = template_workbook["COVER"]
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
    static_document_list_data = frappe.get_doc("Static Document List", {"project_id":project_id}, "*").as_dict()

    local_isolator_specifications_and_list = static_document_list_data.get("local_isolator_specifications_and_list")


    cover_sheet["A3"] = division_name.upper()
    # cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = local_isolator_specifications_and_list

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

    local_isolator_data = local_isolator_revisions_data.get("local_isolator_data")
    
    is_safe_area_isolator_selected = local_isolator_revisions_data.get("is_safe_area_isolator_selected")
    is_hazardous_area_isolator_selected = local_isolator_revisions_data.get("is_hazardous_area_isolator_selected")


    safe_isolator_data = {}
    hazard_isolator_data = {}

    for data in local_isolator_data:
        if data["area"] == "Safe":
            safe_isolator_data = data
        else:
            hazard_isolator_data = data


    safe_fmi_type = safe_isolator_data.get("fmi_type")
    safe_fmi_ip_protection = safe_isolator_data.get("fmi_ip_protection")
    safe_fmi_enclouser_moc = safe_isolator_data.get("fmi_enclouser_moc")
    safe_fmi_enclosure_thickness = safe_isolator_data.get("fmi_enclosure_thickness")
    safe_ifm_cable_entry = safe_isolator_data.get("ifm_cable_entry")
    safe_fmi_qty = safe_isolator_data.get("fmi_qty")
    safe_ifm_isolator_color_shade = safe_isolator_data.get("ifm_isolator_color_shade")
    safe_canopy = safe_isolator_data.get("canopy")
    safe_canopy_type = safe_isolator_data.get("canopy_type")

    if (
        safe_fmi_enclouser_moc == "CRCA"
        or safe_fmi_enclouser_moc == "SS 316"
        or safe_fmi_enclouser_moc == "SS 306"
    ):
        safe_fmi_enclouser_moc = (
            f"{safe_fmi_enclouser_moc}, {safe_fmi_enclosure_thickness}"
        )
        safe_ifm_cable_entry = f"{safe_ifm_cable_entry}, 3 mm"
    elif safe_fmi_enclouser_moc == "NA":
        safe_fmi_enclouser_moc = "Not Applicable"

    
    if int(is_safe_area_isolator_selected) == 0:
        safe_fmi_type = "Not Applicable"
        safe_fmi_ip_protection = "Not Applicable"
        safe_fmi_enclouser_moc = "Not Applicable"
        safe_fmi_enclosure_thickness = "Not Applicable"
        safe_ifm_cable_entry = "Not Applicable"
        safe_fmi_qty = "Not Applicable"
        safe_ifm_isolator_color_shade = "Not Applicable"
        safe_canopy = "Not Applicable"
        safe_canopy_type = "Not Applicable"

    isolator_sheet["C3"] = safe_fmi_type
    isolator_sheet["C4"] = safe_fmi_ip_protection
    isolator_sheet["C5"] = safe_fmi_enclouser_moc
    isolator_sheet["C6"] = safe_fmi_qty
    isolator_sheet["C7"] = safe_ifm_isolator_color_shade
    isolator_sheet["C8"] = safe_ifm_cable_entry
    isolator_sheet["C9"] = safe_canopy
    isolator_sheet["C10"] = safe_canopy_type

    hazard_fmi_type = hazard_isolator_data.get("fmi_type")
    hazard_fmi_ip_protection = hazard_isolator_data.get("fmi_ip_protection")
    hazard_fmi_enclouser_moc = hazard_isolator_data.get("fmi_enclouser_moc")
    hazard_fmi_enclosure_thickness = hazard_isolator_data.get("fmi_enclosure_thickness")
    hazard_ifm_cable_entry = hazard_isolator_data.get("ifm_cable_entry") 
    hazard_fmi_qty = hazard_isolator_data.get("fmi_qty")
    hazard_ifm_isolator_color_shade = hazard_isolator_data.get("ifm_isolator_color_shade")
    hazard_canopy = hazard_isolator_data.get("canopy")
    hazard_canopy_type = hazard_isolator_data.get("canopy_type")

    if (
        hazard_fmi_enclouser_moc == "CRCA"
        or hazard_fmi_enclouser_moc == "SS 316"
        or hazard_fmi_enclouser_moc == "SS 306"
    ):
        hazard_fmi_enclouser_moc = (
            f"{hazard_fmi_enclouser_moc}, {hazard_fmi_enclosure_thickness}"
        )
        hazard_ifm_cable_entry = f"{hazard_ifm_cable_entry}, 3 mm"
    elif hazard_fmi_enclouser_moc == "NA":
        hazard_fmi_enclouser_moc = "Not Applicable"

    
    if int(is_hazardous_area_isolator_selected) == 0:
        hazard_fmi_type = "Not Applicable"
        hazard_fmi_ip_protection = "Not Applicable"
        hazard_fmi_enclouser_moc = "Not Applicable"
        hazard_fmi_enclosure_thickness = "Not Applicable"
        hazard_ifm_cable_entry = "Not Applicable"
        hazard_fmi_qty = "Not Applicable"
        hazard_ifm_isolator_color_shade = "Not Applicable"
        hazard_canopy = "Not Applicable"
        hazard_canopy_type = "Not Applicable"

    isolator_sheet["D3"] = hazard_fmi_type
    isolator_sheet["D4"] = hazard_fmi_ip_protection
    isolator_sheet["D5"] = hazard_fmi_enclouser_moc
    isolator_sheet["D6"] = hazard_fmi_qty
    isolator_sheet["D7"] = hazard_ifm_isolator_color_shade
    isolator_sheet["D8"] = hazard_ifm_cable_entry
    isolator_sheet["D9"] = hazard_canopy
    isolator_sheet["D10"] = hazard_canopy_type
    
    local_isolator_motor_details_data = local_isolator_revisions_data.get("local_isolator_motor_details_data")
    safe_motor_details = []
    hazard_motor_details = []

    for i in range(len(local_isolator_motor_details_data)):
        if local_isolator_motor_details_data[i].get("local_isolator") == "Yes":
            if local_isolator_motor_details_data[i].get("area") == "Safe" or local_isolator_motor_details_data[i].get("area") == "NA":
                safe_motor_details.append(local_isolator_motor_details_data[i])
            else:
                hazard_motor_details.append(local_isolator_motor_details_data[i])

    index = 3

    for i in range(len(safe_motor_details)):
        # area_data = local_isolator_motor_details_data[i].get("area")
        # if area_data == "Safe":
        isolator_safe_area_sheet[f"A{index}"] = i + 1
        isolator_safe_area_sheet[f"B{index}"] = safe_motor_details[i].get("tag_number")
        isolator_safe_area_sheet[f"C{index}"] = safe_motor_details[i].get("service_description")
        isolator_safe_area_sheet[f"D{index}"] = safe_motor_details[i].get("working_kw")
        isolator_safe_area_sheet[f"E{index}"] = ""
        motor_location = safe_motor_details[i].get("motor_location")
        area = safe_motor_details[i].get("area")

        isolator_safe_area_sheet[f"G{index}"] = motor_location

        if area == "Safe":
            canopy = safe_isolator_data.get("canopy")
        else: 
            canopy = hazard_isolator_data.get("canopy")


        canopy_required = ""
        if canopy == "All":
            canopy_required = "Yes"
        else: 
            if "OUT" in canopy and "OUT" in motor_location:
                canopy_required = "Yes"
            else:
                canopy_required = "No"
            

        isolator_safe_area_sheet[f"F{index}"] = canopy_required
        isolator_safe_area_sheet[f"H{index}"] = safe_motor_details[i].get("gland_size")
        index = index + 1

    isolator_safe_area_sheet[f"C{index + 5}"] = "Total Quantity"
    isolator_safe_area_sheet[f"D{index + 5}"] = int(len(safe_motor_details))
    isolator_safe_area_sheet[f"E{index + 5}"] = "Nos"

    index = 3

    for i in range(len(hazard_motor_details)):
        # area_data = local_isolator_motor_details_data[i].get("area")
        # if area_data == "Hazardous":
        isolator_hazard_area_sheet[f"A{index}"] = i + 1
        isolator_hazard_area_sheet[f"B{index}"] = hazard_motor_details[i].get("tag_number")
        isolator_hazard_area_sheet[f"C{index}"] = hazard_motor_details[i].get("service_description")
        isolator_hazard_area_sheet[f"D{index}"] = hazard_motor_details[i].get("working_kw")
        isolator_hazard_area_sheet[f"E{index}"] = ""
        motor_location = hazard_motor_details[i].get("motor_location")
        area = hazard_motor_details[i].get("area")

        isolator_hazard_area_sheet[f"K{index}"] = motor_location

        if area == "Safe":
            canopy = safe_isolator_data.get("canopy")
        else: 
            canopy = hazard_isolator_data.get("canopy")


        canopy_required = ""
        if canopy == "All":
            canopy_required = "Yes"
        else: 
            if "OUT" in canopy  and "OUT" in motor_location:
                canopy_required = "Yes"
            else:
                canopy_required = "No"
            

        isolator_hazard_area_sheet[f"F{index}"] = canopy_required
        isolator_hazard_area_sheet[f"G{index}"] = hazard_motor_details[i].get("standard")
        isolator_hazard_area_sheet[f"H{index}"] = hazard_motor_details[i].get("zone")
        isolator_hazard_area_sheet[f"I{index}"] = hazard_motor_details[i].get("gas_group")
        isolator_hazard_area_sheet[f"J{index}"] = hazard_motor_details[i].get("temprature_class")
        isolator_hazard_area_sheet[f"L{index}"] = hazard_motor_details[i].get("gland_size")
        index = index + 1

    isolator_hazard_area_sheet[f"C{index + 5}"] = "Total Quantity"
    isolator_hazard_area_sheet[f"D{index + 5}"] = int(len(hazard_motor_details))
    isolator_hazard_area_sheet[f"E{index + 5}"] = "Nos"

    if len(safe_motor_details) < 1:
        template_workbook.remove(isolator_safe_area_sheet)
    
    # if len(hazard_motor_details) < 1:
    #     template_workbook.remove(isolator_hazard_area_sheet)

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")