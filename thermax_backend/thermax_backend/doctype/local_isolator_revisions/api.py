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
    # project_oc_number = project_data.get("project_oc_number")
    # approver = project_data.get("approver")
    # client_name = project_data.get("client_name")
    # consultant_name = project_data.get("consultant_name")
    # modified = project_data.get("modified")

    # loading the sheets 

    cover_sheet = template_workbook["COVER"]
    # isolator_sheet = template_workbook["ISOLATOR"]
    # isolator_safe_area_sheet = template_workbook["ISOLATOR LIST SAFE AREA"]
    # isolator_hazard_area_sheet = template_workbook["ISOLATOR HAZARDOUS AREA"]


    # cover page sheet populating

    # prepped_by_initial = frappe.db.get_value(
    #     "Thermax Extended User", owner, "name_initial"
    # )
    # checked_by_initial = frappe.db.get_value(
    #     "Thermax Extended User", approver, "name_initial"
    # )
    # super_user_initial = frappe.db.get_value(
    #     "Thermax Extended User",
    #     {"is_superuser": 1, "division": division_name},
    #     "name_initial",
    # )

    # revision_date = modified.strftime("%d-%m-%Y")
    # revision_data_with_pid = frappe.db.get_list("Design Basis Revision History", {"project_id": project_id}, "*")


    cover_sheet["A3"] = division_name.upper()
    cover_sheet["D6"] = project_name.upper()
    # cover_sheet["D7"] = client_name.upper()
    # cover_sheet["D8"] = consultant_name.upper()
    # cover_sheet["D9"] = project_name.upper()
    # cover_sheet["D10"] = project_oc_number.upper()

    # index = 33

    # for i in range(len(revision_data_with_pid) - 1, -1, -1):
    #     cover_sheet[f"C{index}"] = revision_date
    #     cover_sheet[f"D{index}"] = project_description
    #     cover_sheet[f"E{index}"] = prepped_by_initial
    #     cover_sheet[f"F{index}"] = checked_by_initial
    #     cover_sheet[f"G{index}"] = super_user_initial
    #     index = index - 1

    # match division_name:
    #     case "Heating":
    #         cover_sheet["A4"] = "PUNE - 411 019"
    #     case "WWS SPG":
    #         cover_sheet["A3"] = "WATER & WASTE SOLUTION"
    #         cover_sheet["A4"] = "PUNE - 411 026"
    #     case "WWS IPG":
    #         cover_sheet["A3"] = "WATER & WASTE SOLUTION"
    #         cover_sheet["A4"] = "PUNE - 411 026"
    #     case "Enviro":
    #         cover_sheet["A4"] = "PUNE - 411 026"
    #     case _:
    #         cover_sheet["A4"] = "PUNE - 411 026"


    # ISOLATOR SHEET 
    # Fetch the Design Basis revision data (then isolator data form that)

    # cc_2 = frappe.db.get_list(
    #     "Common Configuration 2", {"revision_id": revision_id}, "*"
    # )

    # is_field_motor_isolator_selected = cc_2.get("is_field_motor_isolator_selected")
    # is_safe_area_isolator_selected = cc_2.get("is_safe_area_isolator_selected")
    # is_local_push_button_station_selected = cc_2.get("is_local_push_button_station_selected")
    # selector_switch_applicable = cc_2.get("selector_switch_applicable")
    # selector_switch_lockable = cc_2.get("selector_switch_lockable")
    # running_open = cc_2.get("running_open")
    # stopped_closed = cc_2.get("stopped_closed")
    # trip = cc_2.get("trip")
    # safe_field_motor_type = cc_2.get("safe_field_motor_type")
    # hazardous_field_motor_type = cc_2.get("hazardous_field_motor_type")
    # safe_field_motor_enclosure = cc_2.get("safe_field_motor_enclosure")
    # hazardous_field_motor_enclosure = cc_2.get("hazardous_field_motor_enclosure")
    # safe_field_motor_material = cc_2.get("safe_field_motor_material")
    # hazardous_field_motor_material = cc_2.get("hazardous_field_motor_material")
    # safe_field_motor_thickness = cc_2.get("safe_field_motor_thickness")
    # hazardous_field_motor_thickness = cc_2.get("hazardous_field_motor_thickness")
    # safe_field_motor_qty = cc_2.get("safe_field_motor_qty")
    # hazardous_field_motor_qty = cc_2.get("hazardous_field_motor_qty")
    # safe_field_motor_isolator_color_shade = cc_2.get("safe_field_motor_isolator_color_shade")
    # hazardous_field_motor_isolator_color_shade = cc_2.get("hazardous_field_motor_isolator_color_shade")
    # safe_field_motor_cable_entry = cc_2.get("safe_field_motor_cable_entry")
    # hazardous_field_motor_cable_entry = cc_2.get("hazardous_field_motor_cable_entry")
    # safe_field_motor_canopy = cc_2.get("safe_field_motor_canopy")
    # hazardous_field_motor_canopy = cc_2.get("hazardous_field_motor_canopy")
    # safe_field_motor_canopy_type = cc_2.get("safe_field_motor_canopy_type")
    # hazardous_field_motor_canopy_type = cc_2.get("hazardous_field_motor_canopy_type")
    # safe_lpbs_type = cc_2.get("safe_lpbs_type")
    # hazardous_lpbs_type = cc_2.get("hazardous_lpbs_type")
    # safe_lpbs_enclosure = cc_2.get("safe_lpbs_enclosure")
    # hazardous_lpbs_enclosure = cc_2.get("hazardous_lpbs_enclosure")
    # safe_lpbs_thickness = cc_2.get("safe_lpbs_thickness")
    # hazardous_lpbs_thickness = cc_2.get("hazardous_lpbs_thickness")
    # safe_lpbs_material = cc_2.get("safe_lpbs_material")
    # hazardous_lpbs_material = cc_2.get("hazardous_lpbs_material")
    # safe_lpbs_qty = cc_2.get("safe_lpbs_qty")
    # hazardous_lpbs_qty = cc_2.get("hazardous_lpbs_qty")
    # safe_lpbs_color_shade = cc_2.get("safe_lpbs_color_shade")
    # hazardous_lpbs_color_shade = cc_2.get("hazardous_lpbs_color_shade")
    # safe_lpbs_canopy = cc_2.get("safe_lpbs_canopy")
    # hazardous_lpbs_canopy = cc_2.get("hazardous_lpbs_canopy")
    # safe_lpbs_canopy_type = cc_2.get("safe_lpbs_canopy_type")
    # hazardous_lpbs_canopy_type = cc_2.get("hazardous_lpbs_canopy_type")
    # lpbs_push_button_start_color = cc_2.get("lpbs_push_button_start_color")
    # lpbs_indication_lamp_start_color = cc_2.get("lpbs_indication_lamp_start_color")
    # lpbs_indication_lamp_stop_color = cc_2.get("lpbs_indication_lamp_stop_color")
    # lpbs_speed_increase = cc_2.get("lpbs_speed_increase")
    # lpbs_speed_decrease = cc_2.get("lpbs_speed_decrease")
    # apfc_relay = cc_2.get("apfc_relay")
    # power_bus_main_busbar_selection = cc_2.get("power_bus_main_busbar_selection")
    # power_bus_heat_pvc_sleeve = cc_2.get("power_bus_heat_pvc_sleeve")
    # power_bus_material = cc_2.get("power_bus_material")
    # power_bus_current_density = cc_2.get("power_bus_current_density")
    # power_bus_rating_of_busbar = cc_2.get("power_bus_rating_of_busbar")
    # control_bus_main_busbar_selection = cc_2.get("control_bus_main_busbar_selection")
    # control_bus_heat_pvc_sleeve = cc_2.get("control_bus_heat_pvc_sleeve")
    # control_bus_material = cc_2.get("control_bus_material")
    # control_bus_current_density = cc_2.get("control_bus_current_density")
    # control_bus_rating_of_busbar = cc_2.get("control_bus_rating_of_busbar")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")