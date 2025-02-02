import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime

from thermax_backend.thermax_backend.doctype.lpbs_specification_revisions.hazardous_lpbs_excel import (
    create_hazardous_area_lpbs_excel,
)
from thermax_backend.thermax_backend.doctype.lpbs_specification_revisions.safe_lpbs_excel import (
    create_safe_area_lpbs_excel,
)


@frappe.whitelist()
def get_lpbs_specification_excel():
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    lpbs_specifications_revision_data = frappe.get_doc(
        "LPBS Specification Revisions", revision_id
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
    revision_data_with_pid = frappe.db.get_list(
        "Design Basis Revision History", {"project_id": project_id}, "*"
    )
    static_document_list_data = frappe.get_doc(
        "Static Document List", {"project_id": project_id}
    ).as_dict()

    lpbs_specifications_and_list = static_document_list_data.get(
        "lpbs_specifications_and_list"
    )
    is_safe_lpbs_selected = lpbs_specifications_revision_data.get(
        "is_safe_lpbs_selected"
    )
    is_hazardous_lpbs_selected = lpbs_specifications_revision_data.get(
        "is_hazardous_lpbs_selected"
    )

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

    # # INSTRUCTION PAGE

    instruction_sheet["A1"] = (
        f"{project_oc_number.upper()} -INSTRUCTIONS TO LOCAL PUSH BUTTON STATIONS VENDORS"
    )

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
        "LPBS Specification Revisions", revision_id
    ).as_dict()

    lpbs_specification = lpbs_revision_data.get("lpbs_specification_data")
    lpbs_specification_data = lpbs_specification[0]
    config_data = lpbs_specification[1]
    lpbs_specification_motor_details = lpbs_revision_data.get(
        "lpbs_specifications_motor_details"
    )

    safe_lpbs_type = lpbs_specification_data.get("safe_lpbs_type")
    safe_lpbs_ip_protection = lpbs_specification_data.get("safe_lpbs_ip_protection")
    safe_lpbs_moc = lpbs_specification_data.get("safe_lpbs_moc")
    safe_lpbs_thickness = lpbs_specification_data.get("safe_lpbs_thickness")
    safe_lpbs_quantity = lpbs_specification_data.get("safe_lpbs_quantity")
    safe_lpbs_color_shade = lpbs_specification_data.get("safe_lpbs_color_shade")
    safe_lpbs_cable_entry = "Bottom"  # safe cabel entry
    safe_lpbs_canopy = lpbs_specification_data.get("safe_lpbs_canopy")
    safe_lpbs_canopy_type = lpbs_specification_data.get("safe_lpbs_canopy_type")

    if safe_lpbs_moc in ["CRCA", "SS 306", "SS 316"]:
        safe_lpbs_moc = f"{safe_lpbs_moc}, {safe_lpbs_thickness}"
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
    hazardous_lpbs_color_shade = lpbs_specification_data.get(
        "hazardous_lpbs_color_shade"
    )
    hazardous_lpbs_cable_entry = "Bottom"  # hazardous cable entry
    hazardous_lpbs_canopy = lpbs_specification_data.get("hazardous_lpbs_canopy")
    hazardous_lpbs_canopy_type = lpbs_specification_data.get(
        "hazardous_lpbs_canopy_type"
    )

    if hazardous_lpbs_moc in ["CRCA", "SS 306", "SS 316"]:
        hazardous_lpbs_moc = f"{hazardous_lpbs_moc}, {hazardous_lpbs_thickness}"
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
    specification_sheet["C13"] = lpbs_specification_data.get(
        "lpbs_push_button_start_color"
    )
    specification_sheet["C14"] = lpbs_specification_data.get(
        "lpbs_forward_push_button_start_color"
    )
    specification_sheet["C15"] = lpbs_specification_data.get(
        "lpbs_reverse_push_button_start_color"
    )
    specification_sheet["C16"] = lpbs_specification_data.get(
        "lpbs_push_button_ess_color"
    )
    specification_sheet["C17"] = lpbs_specification_data.get(
        "lpbs_speed_increase_color"
    )
    specification_sheet["C18"] = lpbs_specification_data.get(
        "lpbs_speed_decrease_color"
    )
    specification_sheet["C19"] = lpbs_specification_data.get(
        "lpbs_indication_lamp_start_color"
    )
    specification_sheet["C20"] = lpbs_specification_data.get(
        "lpbs_indication_lamp_stop_color"
    )

    id = 22

    def handle_label(value):
        switcher = {
            "lpbs_start_push_button": "Start Push Button",
            "off_indication_lamp_push_button": "OFF Indication Lamp",
            "on_indication_lamp_push_button": "ON Indication Lamp",
            "analog_ammeter_push_button": "Analog Ammeter",
            "analog_rpm_push_button": "Analog RPM",
            "emergency_stop_push_button": "Emergency Stop Button",
            "forward_start_push_button": "Forward Start Push Button",
            "reverse_start_push_button": "Reverse Start Push Button",
            "speed_decrease_push_button": "Speed Increase Push Button",
            "speed_increase_push_button": "Speed Decrease Push Button",
        }
        return switcher.get(value, "Invalid option")

    keys_with_yes = [key for key, value in config_data.items() if value == "Yes"]

    for j in range(len(keys_with_yes)):
        value = handle_label(keys_with_yes[j])
        specification_sheet[f"B{id}"] = value
        id = id + 1

    # motor details sheet
    safe_motor_details = []
    hazard_motor_details = []

    for i in range(len(lpbs_specification_motor_details)):
        if lpbs_specification_motor_details[i].get("area") == "Safe":
            safe_motor_details.append(lpbs_specification_motor_details[i])
        else:
            hazard_motor_details.append(lpbs_specification_motor_details[i])

    # SAFE AREA LPBS SHEET
    lpbs_safe_sheet = create_safe_area_lpbs_excel(
        lpbs_safe_sheet=lpbs_safe_sheet,
        safe_motor_details=safe_motor_details,
        safe_lpbs_canopy=safe_lpbs_canopy,
    )

    # HAZARDOUS AREA LPBS SHEET
    lpbs_hazard_sheet = create_hazardous_area_lpbs_excel(
        lpbs_hazard_sheet=lpbs_hazard_sheet,
        hazard_motor_details=hazard_motor_details,
        hazardous_lpbs_canopy=hazardous_lpbs_canopy,
    )

    if is_safe_lpbs_selected == 0:
        template_workbook.remove(lpbs_safe_sheet)

    if is_hazardous_lpbs_selected == 0:
        template_workbook.remove(lpbs_hazard_sheet)

    # template_workbook.save("lpbs_specification.xlsx")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "lpbs_specification.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
