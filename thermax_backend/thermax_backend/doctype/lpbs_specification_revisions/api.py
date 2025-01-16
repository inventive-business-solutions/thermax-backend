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
    specification_sheet = template_workbook["SPECIFICATION"]
    lpbs_safe_sheet = template_workbook[" LPBS LIST SAFE AREA"]
    lpbs_hazard_sheet = template_workbook["LPBS LIST HAZARDOUS AREA "]
    selection_criteria_sheet = template_workbook["SELECTION CRITERIA"]
    solator_hazard_area_sheet = template_workbook["ISOLATOR LIST HAZARDOUS AREA"]


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

    is_local_push_button_station_selected = common_config_data.get("is_local_push_button_station_selected") 
    selector_switch_applicable = common_config_data.get("selector_switch_applicable") 
    selector_switch_lockable = common_config_data.get("selector_switch_lockable") 
    running_open = common_config_data.get("running_open") 
    stopped_closed = common_config_data.get("stopped_closed") 
    trip = common_config_data.get("trip") 
    safe_lpbs_type = common_config_data.get("safe_lpbs_type") 
    hazardous_lpbs_type = common_config_data.get("hazardous_lpbs_type") 
    safe_lpbs_enclosure = common_config_data.get("safe_lpbs_enclosure") 
    hazardous_lpbs_enclosure = common_config_data.get("hazardous_lpbs_enclosure") 
    safe_lpbs_thickness = common_config_data.get("safe_lpbs_thickness") 
    hazardous_lpbs_thickness = common_config_data.get("hazardous_lpbs_thickness") 
    safe_lpbs_material = common_config_data.get("safe_lpbs_material") 
    hazardous_lpbs_material = common_config_data.get("hazardous_lpbs_material") 
    safe_lpbs_qty = common_config_data.get("safe_lpbs_qty") 
    hazardous_lpbs_qty = common_config_data.get("hazardous_lpbs_qty") 
    safe_lpbs_color_shade = common_config_data.get("safe_lpbs_color_shade") 
    hazardous_lpbs_color_shade = common_config_data.get("hazardous_lpbs_color_shade") 
    safe_lpbs_canopy = common_config_data.get("safe_lpbs_canopy") 
    hazardous_lpbs_canopy = common_config_data.get("hazardous_lpbs_canopy") 
    safe_lpbs_canopy_type = common_config_data.get("safe_lpbs_canopy_type") 
    hazardous_lpbs_canopy_type = common_config_data.get("hazardous_lpbs_canopy_type") 
    lpbs_push_button_start_color = common_config_data.get("lpbs_push_button_start_color") 
    lpbs_indication_lamp_start_color = common_config_data.get("lpbs_indication_lamp_start_color") 
    lpbs_indication_lamp_stop_color = common_config_data.get("lpbs_indication_lamp_stop_color") 
    lpbs_speed_increase = common_config_data.get("lpbs_speed_increase") 
    lpbs_speed_decrease = common_config_data.get("lpbs_speed_decrease") 

    if cc_is_local_push_button_station_selected == 0 or cc_is_local_push_button_station_selected == "0":
        cc_lpbs_push_button_start_color = "Not Applicable"
        cc_forward_push_button_start = "Not Applicable"
        cc_reverse_push_button_start = "Not Applicable"
        cc_push_button_ess = "Not Applicable"
        cc_lpbs_speed_increase = "Not Applicable"
        cc_lpbs_speed_decrease = "Not Applicable"
        cc_lpbs_indication_lamp_start_color = "Not Applicable"
        cc_lpbs_indication_lamp_stop_color = "Not Applicable"
        cc_safe_lpbs_type = "Not Applicable"
        cc_safe_lpbs_enclosure = "Not Applicable"
        cc_safe_lpbs_material = "Not Applicable"
        cc_safe_lpbs_qty = "Not Applicable"
        cc_safe_lpbs_color_shade = "Not Applicable"
        cc_safe_lpbs_canopy = "Not Applicable"
        cc_safe_lpbs_canopy_type = "Not Applicable"

    specification_sheet["C160"] = na_to_string(cc_lpbs_push_button_start_color)
    specification_sheet["C161"] = na_to_string(cc_forward_push_button_start)
    specification_sheet["C162"] = na_to_string(cc_reverse_push_button_start)
    specification_sheet["C163"] = na_to_string(cc_push_button_ess)
    specification_sheet["C164"] = na_to_string(cc_lpbs_speed_increase)
    specification_sheet["C165"] = na_to_string(cc_lpbs_speed_decrease)
    specification_sheet["C166"] = na_to_string(cc_lpbs_indication_lamp_start_color)
    specification_sheet["C167"] = na_to_string(cc_lpbs_indication_lamp_stop_color)

    specification_sheet["C169"] = na_to_string(cc_safe_lpbs_type)
    specification_sheet["C170"] = na_to_string(cc_safe_lpbs_enclosure)
    specification_sheet["C171"] = na_to_string(cc_safe_lpbs_material)
    specification_sheet["C172"] = na_to_string(cc_safe_lpbs_qty)
    specification_sheet["C173"] = na_to_string(cc_safe_lpbs_color_shade)
    specification_sheet["C174"] = na_to_string(cc_safe_lpbs_canopy)
    specification_sheet["C175"] = na_to_string(cc_safe_lpbs_canopy_type)

    specification_sheet["D169"] = na_to_string(cc_hazardous_lpbs_type)
    specification_sheet["D170"] = na_to_string(cc_hazardous_lpbs_enclosure)
    specification_sheet["D171"] = na_to_string(cc_hazardous_lpbs_material)
    specification_sheet["D172"] = na_to_string(cc_hazardous_lpbs_qty)
    specification_sheet["D173"] = na_to_string(cc_hazardous_lpbs_color_shade)
    specification_sheet["D174"] = na_to_string(cc_hazardous_lpbs_canopy)
    specification_sheet["D175"] = na_to_string(cc_hazardous_lpbs_canopy_type)

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")