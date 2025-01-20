import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime

# revision_id = "st486uu99i"


@frappe.whitelist()
def get_motor_specification_excel():
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    motor_spec_revision_data = frappe.get_doc(
        "Motor Specification Revisions", revision_id, "*"
    ).as_dict()

    project_id = motor_spec_revision_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", {"project_id": project_id}
    ).as_dict()

    project_revision_id = design_basis_revision_data.get("revision_id")

    # # Loading the workbook
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "motor_specification_template.xlsx"
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
    safe_area_motor_list_sheet = template_workbook["SAFE AREA MOTOR LIST  "]
    safe_area_motor_bom_sheet = template_workbook[" SAFE AREA MOTOR BOM"]
    hazardous_area_motor_list_sheet = template_workbook[" HAZARDOUS AREA MOTOR LIST "]
    hazardous_area_motor_bom_sheet = template_workbook[" HAZARDOUS AREA MOTOR BOM  "]

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
    revision_data_with_pid = frappe.db.get_list(
        "Design Basis Revision History", {"project_id": project_id}, "*"
    )
    static_document_list_data = frappe.get_doc(
        "Static Document List", {"project_id": project_id}, "*"
    ).as_dict()

    motor_specification_static_document = static_document_list_data.get(
        "motor_specification"
    )

    cover_sheet["A3"] = division_name.upper()
    # cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = motor_specification_static_document

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

    # SPECIFICATION SHEET

    project_info_data =  frappe.db.get_list(
        "Motor Specification Revisions", {"project_id": project_id}, "*"
    )
    project_info_data = project_info_data[0]

    motor_specification_data = motor_spec_revision_data.get("motor_specification_data")

    specification_sheet["D4"] = (
        f'{project_info_data.get("standard")}, {project_info_data.get("zone")}, {project_info_data.get("gas_group")}, {project_info_data.get("temperature_class")}'
    )
    specification_sheet["D5"] = project_info_data.get("ambient_temperature_max")
    specification_sheet["D6"] = project_info_data.get("ambient_temperature_min")
    specification_sheet["D7"] = project_info_data.get(
        "electrical_design_temperature"
    )
    specification_sheet["D8"] = project_info_data.get("max_humidity")
    specification_sheet["D9"] = project_info_data.get("min_humidity")
    specification_sheet["D10"] = project_info_data.get("avg_humidity")
    specification_sheet["D11"] = project_info_data.get("performance_humidity")
    specification_sheet["D12"] = project_info_data.get("altitude")
    specification_sheet["D13"] = project_info_data.get("seismic_zone")

    specification_sheet["E4"] = (
        f'{project_info_data.get("standard")}, {project_info_data.get("zone")}, {project_info_data.get("gas_group")}, {project_info_data.get("temperature_class")}'
    )
    specification_sheet["E5"] = project_info_data.get("ambient_temperature_max")
    specification_sheet["E6"] = project_info_data.get("ambient_temperature_min")
    specification_sheet["E7"] = project_info_data.get(
        "electrical_design_temperature"
    )
    specification_sheet["E8"] = project_info_data.get("max_humidity")
    specification_sheet["E9"] = project_info_data.get("min_humidity")
    specification_sheet["E10"] = project_info_data.get("avg_humidity")
    specification_sheet["E11"] = project_info_data.get("performance_humidity")
    specification_sheet["E12"] = project_info_data.get("altitude")
    specification_sheet["E13"] = project_info_data.get("seismic_zone")


    # ELECTRICAL PARAMETERS 
    specification_sheet["D15"] = project_info_data.get("main_supply_lv")
    specification_sheet["D16"] = project_info_data.get("main_supply_lv_phase")
    specification_sheet["D17"] = project_info_data.get("frequency")
    specification_sheet["D18"] = project_info_data.get("fault_level")
    specification_sheet["D19"] = "50 KA for 0.25 Sec. for motors"


    specification_sheet["E15"] = project_info_data.get("main_supply_lv")
    specification_sheet["E16"] = project_info_data.get("main_supply_lv_phase")
    specification_sheet["E17"] = project_info_data.get("frequency")
    specification_sheet["E18"] = project_info_data.get("fault_level")
    specification_sheet["E19"] = "50 KA for 0.25 Sec. for motors"


    cc_1 = frappe.get_doc("Common Configuration 1", {"revision_id": project_revision_id}, "*").as_dict()
    cc_2 = frappe.get_doc("Common Configuration 2", {"revision_id": project_revision_id}, "*").as_dict()
    cc_3 = frappe.get_doc("Common Configuration 3", {"revision_id": project_revision_id}, "*").as_dict()

    common_config_data = cc_1 | cc_2 | cc_3

    specification_sheet["D21"] = common_config_data.get("dm_standard")
    specification_sheet["D22"] = "Low Voltage Squirrel Cage Induction Motor"
    specification_sheet["D23"] = "Copper"

    specification_sheet["D21"] = common_config_data.get("dm_standard")
    specification_sheet["D22"] = "Low Voltage Squirrel Cage Induction Motor"
    specification_sheet["D23"] = "Copper"


    # Motor Parameters

    # specification_sheet["D41"] = 
    

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
