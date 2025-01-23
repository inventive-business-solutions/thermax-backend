import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
from collections import defaultdict
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

    project_revision_id = design_basis_revision_data.get("name")

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


    # INSTRUCTION SHEET

    instruction_sheet["A1"] = f"{project_oc_number}  -INSTRUCTIONS TO LOCAL PUSH BUTTON STATIONS VENDORS"

    # SPECIFICATION SHEET

    project_info_data =  frappe.db.get_list(
        "Project Information", {"project_id": project_id}, "*"
    )
    project_info_data = project_info_data[0]

    motor_specification_data = motor_spec_revision_data.get("motor_specification_data")
    motor_specification_data = motor_specification_data[0]

    specification_sheet["C4"] = "Not Applicable"
    specification_sheet["C5"] = project_info_data.get("ambient_temperature_max")
    specification_sheet["C6"] = project_info_data.get("ambient_temperature_min")
    specification_sheet["C7"] = project_info_data.get(
        "electrical_design_temperature"
    )
    specification_sheet["C8"] = project_info_data.get("max_humidity")
    specification_sheet["C9"] = project_info_data.get("min_humidity")
    specification_sheet["C10"] = project_info_data.get("avg_humidity")
    specification_sheet["C11"] = project_info_data.get("performance_humidity")
    specification_sheet["C12"] = project_info_data.get("altitude")
    specification_sheet["C13"] = project_info_data.get("seismic_zone")

    hazard_area_classification_data = (
        f"{motor_specification_data.get("standard")}, {motor_specification_data.get("zone")}, {motor_specification_data.get("gas_group")}, {motor_spec_revision_data.get("temperature_class")}"
    )

    # if "NA" in hazard_area_classification_data or "None" in hazard_area_classification_data:
    #     hazard_area_classification_data = "Not Applicable"

    # specification_sheet["D4"] = hazard_area_classification_data
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


    # ELECTRICAL PARAMETERS 
    specification_sheet["C15"] = project_info_data.get("main_supply_lv")
    specification_sheet["C16"] = project_info_data.get("main_supply_lv_phase")
    specification_sheet["C17"] = project_info_data.get("frequency")
    specification_sheet["C18"] = project_info_data.get("fault_level")
    specification_sheet["C19"] = "50 KA for 0.25 Sec. for motors"


    specification_sheet["D15"] = project_info_data.get("main_supply_lv")
    specification_sheet["D16"] = project_info_data.get("main_supply_lv_phase")
    specification_sheet["D17"] = project_info_data.get("frequency")
    specification_sheet["D18"] = project_info_data.get("fault_level")
    specification_sheet["D19"] = "50 KA for 0.25 Sec. for motors"


    cc_1 = frappe.db.get_list("Common Configuration 1", {"revision_id": project_revision_id}, "*")
    cc_1 = cc_1[0]
    cc_2 = frappe.db.get_list("Common Configuration 2", {"revision_id": project_revision_id}, "*")
    cc_2 = cc_2[0]
    cc_3 = frappe.db.get_list("Common Configuration 3", {"revision_id": project_revision_id}, "*")
    cc_3 = cc_3[0]

    common_config_data = cc_1 | cc_2 | cc_3

    specification_sheet["C21"] = common_config_data.get("dm_standard")
    specification_sheet["C22"] = "Low Voltage Squirrel Cage Induction Motor"
    specification_sheet["C23"] = "Copper"

    specification_sheet["D21"] = common_config_data.get("dm_standard")
    specification_sheet["D22"] = "Low Voltage Squirrel Cage Induction Motor"
    specification_sheet["D23"] = "Copper"


    # Motor Parameters

    motor_parameters_data = frappe.get_doc(
        "Design Basis Motor Parameters",
        {"revision_id": project_revision_id},
        "*"
    ).as_dict()

    specification_sheet["C24"] = motor_parameters_data.get("safe_area_enclosure_ip_rating")
    specification_sheet["C25"] = motor_parameters_data.get("safe_area_duty")
    

    specification_sheet["D24"] = motor_parameters_data.get("hazardous_area_enclosure_ip_rating")
    specification_sheet["D25"] = motor_parameters_data.get("hazardous_area_duty")


    specification_sheet["C30"] = motor_parameters_data.get("safe_area_insulation_class")
    specification_sheet["C31"] = motor_parameters_data.get("safe_area_temperature_rise")
    specification_sheet["C33"] = motor_parameters_data.get("safe_area_starts_hour_permissible")
    specification_sheet["C34"] = motor_parameters_data.get("safe_area_service_factor")
    specification_sheet["C35"] = motor_parameters_data.get("safe_area_cooling_type")
    specification_sheet["C41"] = motor_parameters_data.get("safe_area_body_material")
    specification_sheet["C42"] = motor_parameters_data.get("safe_area_terminal_box_ip_rating")
    specification_sheet["C45"] = motor_parameters_data.get("safe_area_paint_type_and_shade")

    specification_sheet["D30"] = motor_parameters_data.get("hazardous_area_insulation_class")
    specification_sheet["D31"] = motor_parameters_data.get("hazardous_area_temperature_rise")
    specification_sheet["D33"] = motor_parameters_data.get("hazardous_area_starts_hour_permissible")
    specification_sheet["D34"] = motor_parameters_data.get("hazardous_area_service_factor")
    specification_sheet["D35"] = motor_parameters_data.get("hazardous_area_cooling_type")
    specification_sheet["D41"] = motor_parameters_data.get("hazardous_area_body_material")
    specification_sheet["D42"] = motor_parameters_data.get("hazardous_area_terminal_box_ip_rating")
    specification_sheet["D45"] = motor_parameters_data.get("hazardous_area_paint_type_and_shade")

    motor_details_data = motor_spec_revision_data.get("motor_details_data")
    safe_data = []
    hazard_data = []

    for data in motor_details_data:
        if data.get("area") == "Safe":
            safe_data.append(data)
        else: 
            hazard_data.append(data)

    # SAFE AREA MOTOR LIST 

    index = 3

    for data in safe_data:

        motor_rating_data = data.get("working_kw")
        kw_data = "W"
        
        if float(motor_rating_data) == 0:
            kw_data = "S"
            motor_rating_data = data.get("standby_kw")

        
        safe_area_motor_list_sheet[f"A{index}"] = index - 2
        safe_area_motor_list_sheet[f"B{index}"] = data.get("tag_number")
        safe_area_motor_list_sheet[f"C{index}"] = data.get("service_description")
        safe_area_motor_list_sheet[f"D{index}"] = kw_data
        safe_area_motor_list_sheet[f"E{index}"] = motor_rating_data
        safe_area_motor_list_sheet[f"F{index}"] = data.get("rpm")
        safe_area_motor_list_sheet[f"G{index}"] = data.get("type_of_mounting")
        safe_area_motor_list_sheet[f"H{index}"] = data.get("motor_frame_size")
        safe_area_motor_list_sheet[f"I{index}"] = data.get("motor_gd2")
        safe_area_motor_list_sheet[f"J{index}"] = data.get("gd2_of_driven_equipment")
        safe_area_motor_list_sheet[f"K{index}"] = data.get("bkw")
        safe_area_motor_list_sheet[f"L{index}"] = data.get("type_of_couplings")
        safe_area_motor_list_sheet[f"M{index}"] = data.get("motor_location")
        safe_area_motor_list_sheet[f"N{index}"] = data.get("supply_voltage")
        safe_area_motor_list_sheet[f"O{index}"] = 50
        safe_area_motor_list_sheet[f"P{index}"] = data.get("starter_type")
        safe_area_motor_list_sheet[f"Q{index}"] = data.get("cable_size")
        safe_area_motor_list_sheet[f"R{index}"] = data.get("space_heater")
        roller_bearing = "No"
        if data.get("type_of_bearing") == "Roller":
            roller_bearing = "Yes"
        
        safe_area_motor_list_sheet[f"S{index}"] = roller_bearing

        insulated_bearing = "No"
        if "nsulat" in data.get("type_of_bearing"):
            insulated_bearing = "Yes"

        safe_area_motor_list_sheet[f"T{index}"] = insulated_bearing
        safe_area_motor_list_sheet[f"U{index}"] = data.get("thermistor")
        safe_area_motor_list_sheet[f"V{index}"] = data.get("bearing_rtd")
        safe_area_motor_list_sheet[f"W{index}"] = data.get("winding_rtd")
        safe_area_motor_list_sheet[f"X{index}"] = data.get("efficiency")
        safe_area_motor_list_sheet[f"Y{index}"] = data.get("motor_rated_current")
        safe_area_motor_list_sheet[f"Z{index}"] = data.get("power_factor")
        # safe_area_motor_list_sheet[f"AA{index}"] = 
        safe_area_motor_list_sheet[f"AB{index}"] = data.get("make")
        safe_area_motor_list_sheet[f"AC{index}"] = data.get("part_code")
        safe_area_motor_list_sheet[f"AD{index}"] = data.get("remark")
        index = index + 1

    # safe_area_motor_bom_sheet
    count_dict = defaultdict(int)
    # Iterate through each motor detail in the correct structure
    for motor in safe_data:
        key = f"{motor['make']} Make LT Motor: {motor['working_kw']} kW, {motor['rpm']} POLE, {motor['type_of_mounting']}, {motor['efficiency']}"
        count_dict[key] += 1

    index = 3

    for key, count in count_dict.items():
        safe_area_motor_bom_sheet[f"A{index}"] = index - 2
        # safe_area_motor_bom_sheet[f"B{index}"] = index - 3
        safe_area_motor_bom_sheet[f"C{index}"] = key
        safe_area_motor_bom_sheet[f"D{index}"] = "NOS"
        safe_area_motor_bom_sheet[f"E{index}"] = count
        # safe_area_motor_bom_sheet[f"F{index}"] = 
        index += 1

    index = 3
    for data in hazard_data:
        motor_rating_data = data.get("working_kw")
        kw_data = "W"
        
        if float(motor_rating_data) == 0:
            kw_data = "S"
            motor_rating_data = data.get("standby_kw")

        hazardous_area_motor_list_sheet[f"A{index}"] = index - 2
        hazardous_area_motor_list_sheet[f"B{index}"] = data.get("tag_number")
        hazardous_area_motor_list_sheet[f"C{index}"] = data.get("service_description")
        hazardous_area_motor_list_sheet[f"D{index}"] = kw_data
        hazardous_area_motor_list_sheet[f"E{index}"] = motor_rating_data
        hazardous_area_motor_list_sheet[f"F{index}"] = data.get("rpm")
        hazardous_area_motor_list_sheet[f"G{index}"] = data.get("type_of_mounting")
        hazardous_area_motor_list_sheet[f"H{index}"] = data.get("motor_frame_size")
        hazardous_area_motor_list_sheet[f"I{index}"] = data.get("motor_gd2")
        hazardous_area_motor_list_sheet[f"J{index}"] = data.get("gd2_of_driven_equipment")
        hazardous_area_motor_list_sheet[f"K{index}"] = data.get("bkw")
        hazardous_area_motor_list_sheet[f"L{index}"] = data.get("type_of_couplings")
        hazardous_area_motor_list_sheet[f"M{index}"] = data.get("motor_location")
        hazardous_area_motor_list_sheet[f"N{index}"] = data.get("supply_voltage")
        hazardous_area_motor_list_sheet[f"O{index}"] = 50
        hazardous_area_motor_list_sheet[f"P{index}"] = data.get("starter_type")
        hazardous_area_motor_list_sheet[f"Q{index}"] = data.get("cable_size")
        hazardous_area_motor_list_sheet[f"R{index}"] = data.get("space_heater")
        roller_bearing = "No"
        if data.get("type_of_bearing") == "Roller":
            roller_bearing = "Yes"
        
        hazardous_area_motor_list_sheet[f"S{index}"] = roller_bearing

        insulated_bearing = "No"
        if "nsulat" in data.get("type_of_bearing"):
            insulated_bearing = "Yes"

        hazardous_area_motor_list_sheet[f"T{index}"] = insulated_bearing
        hazardous_area_motor_list_sheet[f"U{index}"] = data.get("thermistor")
        hazardous_area_motor_list_sheet[f"V{index}"] = data.get("bearing_rtd")
        hazardous_area_motor_list_sheet[f"W{index}"] = data.get("winding_rtd")
        hazardous_area_motor_list_sheet[f"X{index}"] = data.get("efficiency")
        hazardous_area_motor_list_sheet[f"Y{index}"] = data.get("motor_rated_current")
        hazardous_area_motor_list_sheet[f"Z{index}"] = data.get("power_factor")
        # safe_area_motor_list_sheet[f"AA{index}"] = 
        hazardous_area_motor_list_sheet[f"AB{index}"] = data.get("make")
        hazardous_area_motor_list_sheet[f"AC{index}"] = data.get("part_code")
        hazardous_area_motor_list_sheet[f"AD{index}"] = data.get("remark")
        index = index + 1

    count_dict = defaultdict(int)
    # Iterate through each motor detail in the correct structure
    for motor in hazard_data:
        key = f"{motor['make']} Make LT Motor: {motor['working_kw']} kW, {motor['rpm']} POLE, {motor['type_of_mounting']}, {motor['efficiency']}"
        count_dict[key] += 1

    index = 3

    for key, count in count_dict.items():
        hazardous_area_motor_bom_sheet[f"A{index}"] = index - 2
        hazardous_area_motor_bom_sheet[f"C{index}"] = key
        hazardous_area_motor_bom_sheet[f"D{index}"] = "NOS"
        hazardous_area_motor_bom_sheet[f"E{index}"] = count
        index += 1

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
