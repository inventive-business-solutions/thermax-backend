import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime


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
    revision_id = payload["revision_id"]

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", revision_id
    ).as_dict()
    project_id = design_basis_revision_data.get("project_id")
    project_data = frappe.get_doc("Project", project_id).as_dict()

    # Define the path to the template
    template_path = frappe.get_app_path(
        "thermax_backend", "templates", "design_basis_template.xlsx"
    )

    template_workbook = load_workbook(template_path)

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

    ########################################################################################################################

    # Loading the Sheets of templates

    cover_sheet = template_workbook["Cover"]  # template_workbook["COVER"]
    design_basis_sheet = template_workbook[
        "Design Basis"
    ]  # template_workbook["Design Basis"]
    mcc_sheet = template_workbook["MCC"]
    pcc_sheet = template_workbook["PCC"]
    mcc_cum_plc_sheet = template_workbook["MCC CUM PLC"]

    prepped_by_initial = frappe.db.get_value(
        "Thermax Extended User", owner, "name_initial"
    )
    checked_by_initial = frappe.db.get_value(
        "Thermax Extended User", owner, "name_initial"
    )
    super_user_initial = frappe.db.get_value(
        "Thermax Extended User",
        {"is_superuser": 1, "division": division_name},
        "name_initial",
    )

    revision_date = modified.strftime("%d-%m-%Y")

    # COVER SHEET ################################################################################################################################

    cover_sheet["A3"] = division_name.upper()
    cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()

    cover_sheet["C29"] = revision_date
    cover_sheet["D29"] = project_description
    cover_sheet["E29"] = prepped_by_initial
    cover_sheet["F29"] = checked_by_initial
    cover_sheet["G29"] = super_user_initial

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

    # DESIGN BASIS SHEET #
    project_info_data = frappe.get_doc("Project Information", project_id)

    main_supply_lv = project_info_data.get("main_supply_lv")
    main_supply_lv_variation = project_info_data.get("main_supply_lv_variation")
    main_supply_lv_phase = project_info_data.get("main_supply_lv_phase")
    lv_data = f"{main_supply_lv}, {main_supply_lv_variation}, {main_supply_lv_phase}"

    if main_supply_lv == "NA":
        lv_data = "Not Applicable"

    main_supply_mv = project_info_data.get("main_supply_mv")
    main_supply_mv_variation = project_info_data.get("main_supply_mv_variation")
    main_supply_mv_phase = project_info_data.get("main_supply_mv_phase")
    mv_data = f"{main_supply_mv}, {main_supply_mv_variation}%, {main_supply_mv_phase}"

    if main_supply_mv == "NA":
        mv_data = "Not Applicable"

    control_supply = project_info_data.get("control_supply")
    control_supply_variation = project_info_data.get("control_supply_variation")
    control_supply_phase = project_info_data.get("control_supply_phase")
    control_supply_data = (
        f"{control_supply}, {control_supply_variation}%, {control_supply_phase}"
    )
    if control_supply_variation == "NA":
        control_supply_data = control_supply

    utility_supply = project_info_data.get("utility_supply")
    utility_supply_variation = project_info_data.get("utility_supply_variation")
    utility_supply_phase = project_info_data.get("utility_supply_phase")
    utility_supply_data = (
        f"{utility_supply}, {utility_supply_variation}%, {utility_supply_phase}"
    )
    if utility_supply_variation == "NA":
        utility_supply_data = utility_supply

    project_info_freq = project_info_data.get("frequency")
    preojct_info_freq_var = project_info_data.get("frequency_variation")
    project_info_frequency_data = f"{project_info_freq} Hz , {preojct_info_freq_var}%"

    project_info_fault = project_info_data.get("fault_level")
    project_info_sec = project_info_data.get("sec")
    fault_data = f"{project_info_fault} kA, {project_info_sec} Sec"

    ambient_temperature_max = project_info_data.get("ambient_temperature_max")
    ambient_temperature_min = project_info_data.get("ambient_temperature_min")
    electrical_design_temperature = project_info_data.get(
        "electrical_design_temperature"
    )
    seismic_zone = project_info_data.get("seismic_zone")
    min_humidity = project_info_data.get("min_humidity")
    max_humidity = project_info_data.get("max_humidity")
    avg_humidity = project_info_data.get("avg_humidity")
    performance_humidity = project_info_data.get("performance_humidity")
    altitude = project_info_data.get("altitude")

    general_info_data = frappe.db.get_list(
        "Design Basis General Info", {"revision_id": revision_id}, "*"
    )
    general_info_data = general_info_data[0]
    battery_limit = general_info_data.get("battery_limit")

    design_basis_sheet["C4"] = mv_data
    design_basis_sheet["C5"] = lv_data
    design_basis_sheet["C6"] = control_supply_data
    design_basis_sheet["C7"] = utility_supply_data
    design_basis_sheet["C8"] = project_info_frequency_data
    design_basis_sheet["C9"] = fault_data
    design_basis_sheet["C10"] = f"{ambient_temperature_max} Deg. C"
    design_basis_sheet["C11"] = f"{ambient_temperature_min} Deg. C"
    design_basis_sheet["C12"] = f"{electrical_design_temperature} Deg. C"
    design_basis_sheet["C13"] = int(seismic_zone)
    design_basis_sheet["C14"] = f"{max_humidity}%"
    design_basis_sheet["C15"] = f"{min_humidity}%"
    design_basis_sheet["C16"] = f"{avg_humidity}%"
    design_basis_sheet["C17"] = f"{performance_humidity}%"
    design_basis_sheet["C18"] = f"{altitude} meters"

    main_packages_data = frappe.db.get_list(
        "Project Main Package",
        fields=["*"],
        filters={"revision_id": revision_id},
        order_by="creation asc",
    )
    main_packages_data = main_packages_data[0]

    # sub_package_data = main_packages_data["sub_packages"]
    # safe_sub_package = []
    # hazardous_sub_package = []

    # for sub_package in sub_package_data:
    #     if sub_package["area_of_classification"] == "Safe Area":

    # for main_package in main_packages_data:
    #     # Get all Sub Package records
    #     sub_packages = frappe.get_doc(
    #         "Project Main Package", main_package["name"]
    #     ).sub_packages
    #     main_package["sub_packages"] = sub_packages

    main_package_name = ""
    # if len(main_packages_data) > 1:
    main_package_name = main_packages_data.get("main_package_name")

    area_classification_data = frappe.db.get_value(
        "Project Main Package",
        {"revision_id": revision_id},
        ["standard", "zone", "gas_group", "temperature_class"],
    )

    default_values = {
        "standard": "IS",  # Replace with your actual default value
        "zone": "Zone 2",  # Replace with your actual default value
        "gas_group": "IIA/IIB",  # Replace with your actual default value
        "temperature_class": "T3",  # Replace with your actual default value
    }

    if area_classification_data is None:
        area_classification_data = [
            default_values[field] for field in default_values.keys()
        ]
    else:
        area_classification_data = [
            value if value is not None else default_values[field]
            for value, field in zip(area_classification_data, default_values.keys())
        ]

    # Safeguard against missing indices in area_classification_data
    standard = area_classification_data[0] if len(area_classification_data) > 0 else ""
    classification_1 = (
        area_classification_data[1] if len(area_classification_data) > 1 else ""
    )
    gas_group = area_classification_data[2] if len(area_classification_data) > 2 else ""
    temperature_class = (
        area_classification_data[3] if len(area_classification_data) > 3 else ""
    )

    design_basis_sheet["C20"] = main_package_name
    design_basis_sheet["C21"] = main_package_name
    design_basis_sheet["C22"] = main_package_name
    design_basis_sheet["C23"] = (
        f"Standard-{standard}, {classification_1}, Gas Group-{gas_group}, Temperature Class-{temperature_class}"
    )
    design_basis_sheet["C24"] = battery_limit

    # MOTOR PARAMETERS

    motor_parameters_data = frappe.db.get_list(
        "Design Basis Motor Parameters", {"revision_id": revision_id}, "*"
    )
    motor_parameters_data = motor_parameters_data[0]

    safe_area_efficiency_level = motor_parameters_data.get("safe_area_efficiency_level")
    safe_area_insulation_class = motor_parameters_data.get("safe_area_insulation_class")
    safe_area_temperature_rise = motor_parameters_data.get("safe_area_temperature_rise")
    safe_area_enclosure_ip_rating = motor_parameters_data.get(
        "safe_area_enclosure_ip_rating"
    )
    safe_area_max_temperature = motor_parameters_data.get("safe_area_max_temperature")
    safe_area_min_temperature = motor_parameters_data.get("safe_area_min_temperature")
    safe_area_altitude = motor_parameters_data.get("safe_area_altitude")
    safe_area_terminal_box_ip_rating = motor_parameters_data.get(
        "safe_area_terminal_box_ip_rating"
    )
    safe_area_thermister = motor_parameters_data.get("safe_area_thermister")
    safe_area_space_heater = motor_parameters_data.get("safe_area_space_heater")
    safe_area_certification = motor_parameters_data.get("safe_area_certification")
    safe_area_bearing_rtd = motor_parameters_data.get("safe_area_bearing_rtd")
    safe_area_winding_rtd = motor_parameters_data.get("safe_area_winding_rtd")
    safe_area_bearing_type = motor_parameters_data.get("safe_area_bearing_type")
    safe_area_duty = motor_parameters_data.get("safe_area_duty")
    safe_area_service_factor = motor_parameters_data.get("safe_area_service_factor")
    safe_area_cooling_type = motor_parameters_data.get("safe_area_cooling_type")
    safe_area_body_material = motor_parameters_data.get("safe_area_body_material")
    safe_area_terminal_box_material = motor_parameters_data.get(
        "safe_area_terminal_box_material"
    )
    safe_area_paint_type_and_shade = motor_parameters_data.get(
        "safe_area_paint_type_and_shade"
    )
    safe_area_starts_hour_permissible = motor_parameters_data.get(
        "safe_area_starts_hour_permissible"
    )

    is_hazardous_area_present = motor_parameters_data.get("is_hazardous_area_present")
    hazardous_area_efficiency_level = motor_parameters_data.get(
        "hazardous_area_efficiency_level"
    )
    hazardous_area_insulation_class = motor_parameters_data.get(
        "hazardous_area_insulation_class"
    )
    hazardous_area_temperature_rise = motor_parameters_data.get(
        "hazardous_area_temperature_rise"
    )
    hazardous_area_enclosure_ip_rating = motor_parameters_data.get(
        "hazardous_area_enclosure_ip_rating"
    )
    hazardous_area_max_temperature = motor_parameters_data.get(
        "hazardous_area_max_temperature"
    )
    hazardous_area_min_temperature = motor_parameters_data.get(
        "hazardous_area_min_temperature"
    )
    hazardous_area_altitude = motor_parameters_data.get("hazardous_area_altitude")
    hazardous_area_terminal_box_ip_rating = motor_parameters_data.get(
        "hazardous_area_terminal_box_ip_rating"
    )
    hazardous_area_thermister = motor_parameters_data.get("hazardous_area_thermister")
    hazardous_area_space_heater = motor_parameters_data.get(
        "hazardous_area_space_heater"
    )
    hazardous_area_certification = motor_parameters_data.get(
        "hazardous_area_certification"
    )
    hazardous_area_bearing_rtd = motor_parameters_data.get("hazardous_area_bearing_rtd")
    hazardous_area_winding_rtd = motor_parameters_data.get("hazardous_area_winding_rtd")
    hazardous_area_bearing_type = motor_parameters_data.get(
        "hazardous_area_bearing_type"
    )
    hazardous_area_duty = motor_parameters_data.get("hazardous_area_duty")
    hazardous_area_service_factor = motor_parameters_data.get(
        "hazardous_area_service_factor"
    )
    hazardous_area_cooling_type = motor_parameters_data.get(
        "hazardous_area_cooling_type"
    )
    hazardous_area_body_material = motor_parameters_data.get(
        "hazardous_area_body_material"
    )
    hazardous_area_terminal_box_material = motor_parameters_data.get(
        "hazardous_area_terminal_box_material"
    )
    hazardous_area_paint_type_and_shade = motor_parameters_data.get(
        "hazardous_area_paint_type_and_shade"
    )
    hazardous_area_starts_hour_permissible = motor_parameters_data.get(
        "hazardous_area_starts_hour_permissible"
    )

    design_basis_sheet["C27"] = safe_area_efficiency_level
    design_basis_sheet["C28"] = safe_area_insulation_class
    design_basis_sheet["C29"] = safe_area_temperature_rise
    design_basis_sheet["C30"] = safe_area_enclosure_ip_rating
    design_basis_sheet["C31"] = f"{safe_area_max_temperature} Deg. C"
    design_basis_sheet["C32"] = f"{safe_area_min_temperature} Deg. C"
    design_basis_sheet["C33"] = f"{safe_area_altitude} meters"
    design_basis_sheet["C34"] = f"{safe_area_terminal_box_ip_rating} kW & Above"
    design_basis_sheet["C35"] = f"{safe_area_thermister} kW & Above"
    design_basis_sheet["C36"] = f"{safe_area_space_heater} kW & Above"
    design_basis_sheet["C37"] = f"{safe_area_certification} kW & Above"
    design_basis_sheet["C38"] = f"{safe_area_bearing_rtd} kW & Above"
    design_basis_sheet["C39"] = f"{safe_area_winding_rtd} kW & Above"
    design_basis_sheet["C40"] = safe_area_bearing_type
    design_basis_sheet["C41"] = safe_area_duty
    design_basis_sheet["C42"] = (
        int(safe_area_service_factor) if safe_area_service_factor else ""
    )
    design_basis_sheet["C43"] = safe_area_cooling_type
    design_basis_sheet["C44"] = safe_area_body_material
    design_basis_sheet["C45"] = safe_area_terminal_box_material
    design_basis_sheet["C46"] = safe_area_paint_type_and_shade
    design_basis_sheet["C47"] = safe_area_starts_hour_permissible

    design_basis_sheet["D27"] = hazardous_area_efficiency_level
    design_basis_sheet["D28"] = hazardous_area_insulation_class
    design_basis_sheet["D29"] = hazardous_area_temperature_rise
    design_basis_sheet["D30"] = hazardous_area_enclosure_ip_rating
    design_basis_sheet["D31"] = f"{hazardous_area_max_temperature} Deg. C"
    design_basis_sheet["D32"] = f"{hazardous_area_min_temperature} Deg. C"
    design_basis_sheet["D33"] = f"{hazardous_area_altitude} meters"
    design_basis_sheet["D34"] = f"{hazardous_area_terminal_box_ip_rating} kW & Above"
    design_basis_sheet["D35"] = f"{hazardous_area_thermister} kW & Above"
    design_basis_sheet["D36"] = f"{hazardous_area_space_heater} kW & Above"
    design_basis_sheet["D37"] = f"{hazardous_area_certification} kW & Above"
    design_basis_sheet["D38"] = f"{hazardous_area_bearing_rtd} kW & Above"
    design_basis_sheet["D39"] = f"{hazardous_area_winding_rtd} kW & Above"
    design_basis_sheet["D40"] = hazardous_area_bearing_type
    design_basis_sheet["D41"] = hazardous_area_duty
    design_basis_sheet["D42"] = (
        int(hazardous_area_service_factor) if hazardous_area_service_factor else ""
    )
    design_basis_sheet["D43"] = hazardous_area_cooling_type
    design_basis_sheet["D44"] = hazardous_area_body_material
    design_basis_sheet["D45"] = hazardous_area_terminal_box_material
    design_basis_sheet["D46"] = hazardous_area_paint_type_and_shade
    design_basis_sheet["D47"] = hazardous_area_starts_hour_permissible

    # MAKE OF COMPONENTS

    make_of_components_data = frappe.db.get_list(
        "Design Basis Make of Component", {"revision_id": revision_id}, "*"
    )
    make_of_components_data = make_of_components_data[0]

    def handle_make_of_component(component):
        component = (
            component.replace('"', "").replace("[", "").replace("]", "")
            if component
            else "NA"
        )
        if (component) == "NA":
            return "Not Applicable"
        else:
            return component

    motor = make_of_components_data.get("motor")
    cable = make_of_components_data.get("cable")
    lv_switchgear = make_of_components_data.get("lv_switchgear")
    panel_enclosure = make_of_components_data.get("panel_enclosure")
    variable_frequency_speed_drive_vfd_vsd = make_of_components_data.get(
        "variable_frequency_speed_drive_vfd_vsd"
    )
    soft_starter = make_of_components_data.get("soft_starter")
    plc = make_of_components_data.get("plc")
    # gland_make = make_of_components_data.get("gland_make")

    preferred_motor = make_of_components_data.get("preferred_motor")
    preferred_cable = make_of_components_data.get("preferred_cable")
    preferred_lv_switchgear = make_of_components_data.get("preferred_lv_switchgear")
    preferred_panel_enclosure = make_of_components_data.get("preferred_panel_enclosure")
    preferred_vfdvsd = make_of_components_data.get("preferred_vfdvsd")
    preferred_soft_starter = make_of_components_data.get("preferred_soft_starter")
    preferred_plc = make_of_components_data.get("preferred_plc")
    # preferred_gland_make = make_of_components_data.get("preferred_gland_make") ,

    design_basis_sheet["C50"] = handle_make_of_component(motor)
    design_basis_sheet["C51"] = handle_make_of_component(cable)
    design_basis_sheet["C52"] = handle_make_of_component(lv_switchgear)
    design_basis_sheet["C53"] = handle_make_of_component(panel_enclosure)
    design_basis_sheet["C54"] = handle_make_of_component(
        variable_frequency_speed_drive_vfd_vsd
    )
    design_basis_sheet["C55"] = handle_make_of_component(soft_starter)
    design_basis_sheet["C56"] = handle_make_of_component(plc)

    # design_basis_sheet["D50"] = preferred_motor
    # design_basis_sheet["D51"] = preferred_cable
    # design_basis_sheet["D52"] = preferred_lv_switchgear
    # design_basis_sheet["D53"] = preferred_panel_enclosure
    # design_basis_sheet["D54"] = preferred_vfdvsd
    # design_basis_sheet["D55"] = preferred_soft_starter
    # design_basis_sheet["D56"] = preferred_plc

    # COMMON CONFIGURATION

    cc_1 = frappe.db.get_list(
        "Common Configuration 1", {"revision_id": revision_id}, "*"
    )
    cc_1 = cc_1[0]
    cc_2 = frappe.db.get_list(
        "Common Configuration 2", {"revision_id": revision_id}, "*"
    )
    cc_2 = cc_2[0]
    cc_3 = frappe.db.get_list(
        "Common Configuration 3", {"revision_id": revision_id}, "*"
    )
    cc_3 = cc_3[0]

    common_config_data = cc_1 | cc_2 | cc_3

    dol_starter = common_config_data.get("dol_starter")
    star_delta_starter = common_config_data.get("star_delta_starter")
    mcc_switchgear_type = common_config_data.get("mcc_switchgear_type")
    switchgear_combination = common_config_data.get("switchgear_combination")

    design_basis_sheet["C59"] = f"{dol_starter} kW & Above"
    design_basis_sheet["C60"] = f"{star_delta_starter} kW & Above"
    design_basis_sheet["C61"] = mcc_switchgear_type
    design_basis_sheet["C62"] = switchgear_combination

    is_control_transformer_applicable = common_config_data.get(
        "is_control_transformer_applicable"
    )
    control_transformer_primary_voltage = common_config_data.get(
        "control_transformer_primary_voltage"
    )
    control_transformer_secondary_voltage_copy = common_config_data.get(
        "control_transformer_secondary_voltage_copy"
    )
    control_transformer_coating = common_config_data.get("control_transformer_coating")
    control_transformer_quantity = common_config_data.get(
        "control_transformer_quantity"
    )
    control_transformer_configuration = common_config_data.get(
        "control_transformer_configuration"
    )
    control_transformer_type = common_config_data.get("control_transformer_type")

    design_basis_sheet["C64"] = control_transformer_primary_voltage
    design_basis_sheet["C65"] = control_transformer_secondary_voltage_copy
    design_basis_sheet["C66"] = control_transformer_coating
    design_basis_sheet["C67"] = control_transformer_quantity
    design_basis_sheet["C68"] = control_transformer_configuration
    design_basis_sheet["C69"] = control_transformer_type

    ammeter = common_config_data.get("ammeter")
    ammeter_configuration = common_config_data.get("ammeter_configuration")
    analog_meters = common_config_data.get("analog_meters")
    digital_meters = common_config_data.get("digital_meters")
    communication_protocol = common_config_data.get("communication_protocol")

    design_basis_sheet["C71"] = f"{ammeter} kW & Above"
    design_basis_sheet["C72"] = ammeter_configuration
    design_basis_sheet["C73"] = analog_meters
    design_basis_sheet["C74"] = digital_meters
    design_basis_sheet["C75"] = communication_protocol

    current_transformer = common_config_data.get("current_transformer")
    current_transformer_coating = common_config_data.get("current_transformer_coating")
    current_transformer_quantity = common_config_data.get(
        "current_transformer_quantity"
    )
    current_transformer_configuration = common_config_data.get(
        "current_transformer_configuration"
    )

    design_basis_sheet["C77"] = f"{current_transformer} kW & Above"
    design_basis_sheet["C78"] = current_transformer_coating
    design_basis_sheet["C79"] = current_transformer_quantity
    design_basis_sheet["C80"] = current_transformer_configuration

    pole = common_config_data.get("pole")
    design_basis_sheet["C82"] = pole

    supply_feeder_standard = common_config_data.get("supply_feeder_standard")
    dm_standard = common_config_data.get("dm_standard")

    # design_basis_sheet["84"] = dm_standard

    power_wiring_color = common_config_data.get("power_wiring_color")
    power_wiring_size = common_config_data.get("power_wiring_size")
    control_wiring_color = common_config_data.get("control_wiring_color")
    control_wiring_size = common_config_data.get("control_wiring_size")
    vdc_24_wiring_color = common_config_data.get("vdc_24_wiring_color")
    vdc_24_wiring_size = common_config_data.get("vdc_24_wiring_size")
    analog_signal_wiring_color = common_config_data.get("analog_signal_wiring_color")
    analog_signal_wiring_size = common_config_data.get("analog_signal_wiring_size")
    ct_wiring_color = common_config_data.get("ct_wiring_color")
    ct_wiring_size = common_config_data.get("ct_wiring_size")
    rtd_thermocouple_wiring_color = common_config_data.get(
        "rtd_thermocouple_wiring_color"
    )
    rtd_thermocouple_wiring_size = common_config_data.get(
        "rtd_thermocouple_wiring_size"
    )
    air_clearance_between_phase_to_phase_bus = common_config_data.get(
        "air_clearance_between_phase_to_phase_bus"
    )
    air_clearance_between_phase_to_neutral_bus = common_config_data.get(
        "air_clearance_between_phase_to_neutral_bus"
    )
    cable_insulation_pvc = common_config_data.get("cable_insulation_pvc")
    general_note_internal_wiring = common_config_data.get(
        "general_note_internal_wiring"
    )
    device_identification_of_components = common_config_data.get(
        "device_identification_of_components"
    )

    design_basis_sheet["C86"] = power_wiring_color
    design_basis_sheet["C87"] = power_wiring_size
    design_basis_sheet["C88"] = control_wiring_color
    design_basis_sheet["C89"] = control_wiring_size
    design_basis_sheet["C90"] = vdc_24_wiring_color
    design_basis_sheet["C91"] = vdc_24_wiring_size
    design_basis_sheet["C92"] = analog_signal_wiring_color
    design_basis_sheet["C93"] = analog_signal_wiring_size
    design_basis_sheet["C94"] = ct_wiring_color
    design_basis_sheet["C95"] = ct_wiring_size
    design_basis_sheet["C96"] = rtd_thermocouple_wiring_color
    design_basis_sheet["C97"] = rtd_thermocouple_wiring_size
    design_basis_sheet["C98"] = air_clearance_between_phase_to_phase_bus
    design_basis_sheet["C99"] = air_clearance_between_phase_to_neutral_bus
    design_basis_sheet["C100"] = cable_insulation_pvc
    design_basis_sheet["C101"] = device_identification_of_components
    design_basis_sheet["C102"] = general_note_internal_wiring

    power_terminal_clipon = common_config_data.get("power_terminal_clipon")
    power_terminal_busbar_type = common_config_data.get("power_terminal_busbar_type")
    control_terminal = common_config_data.get("control_terminal")
    spare_terminal = common_config_data.get("spare_terminal")

    design_basis_sheet["C104"] = power_terminal_clipon
    design_basis_sheet["C105"] = power_terminal_busbar_type
    design_basis_sheet["C106"] = control_terminal
    design_basis_sheet["C107"] = spare_terminal

    common_requirement = common_config_data.get("common_requirement")

    push_button_start = common_config_data.get("push_button_start")
    push_button_stop = common_config_data.get("push_button_stop")
    push_button_ess = common_config_data.get("push_button_ess")
    forward_push_button_start = common_config_data.get("forward_push_button_start")
    reverse_push_button_start = common_config_data.get("reverse_push_button_start")
    potentiometer = common_config_data.get("potentiometer")
    speed_increase_pb = common_config_data.get("speed_increase_pb")
    speed_decrease_pb = common_config_data.get("speed_decrease_pb")
    alarm_acknowledge_and_lamp_test = common_config_data.get(
        "alarm_acknowledge_and_lamp_test"
    )
    test_dropdown = common_config_data.get("test_dropdown")
    reset_dropdown = common_config_data.get("reset_dropdown")

    design_basis_sheet["C109"] = push_button_start
    design_basis_sheet["C110"] = push_button_stop
    design_basis_sheet["C111"] = push_button_ess
    design_basis_sheet["C112"] = forward_push_button_start
    design_basis_sheet["C113"] = reverse_push_button_start
    design_basis_sheet["C114"] = potentiometer
    design_basis_sheet["C115"] = speed_increase_pb
    design_basis_sheet["C116"] = speed_decrease_pb
    design_basis_sheet["C117"] = alarm_acknowledge_and_lamp_test
    design_basis_sheet["C118"] = test_dropdown
    design_basis_sheet["C119"] = reset_dropdown

    selector_switch_applicable = common_config_data.get("selector_switch_applicable")
    design_basis_sheet["C121"] = selector_switch_applicable

    running_open = common_config_data.get("running_open")
    stopped_closed = common_config_data.get("stopped_closed")
    trip = common_config_data.get("trip")

    design_basis_sheet["C123"] = running_open
    design_basis_sheet["C124"] = stopped_closed
    design_basis_sheet["C125"] = trip

    is_push_button_speed_selected = common_config_data.get(
        "is_push_button_speed_selected"
    )
    lamp_test_push_button = common_config_data.get("lamp_test_push_button")
    is_field_motor_isolator_selected = common_config_data.get(
        "is_field_motor_isolator_selected"
    )
    is_safe_area_isolator_selected = common_config_data.get(
        "is_safe_area_isolator_selected"
    )
    is_local_push_button_station_selected = common_config_data.get(
        "is_local_push_button_station_selected"
    )
    selector_switch_lockable = common_config_data.get("selector_switch_lockable")

    safe_field_motor_type = common_config_data.get("safe_field_motor_type")
    safe_field_motor_enclosure = common_config_data.get("safe_field_motor_enclosure")
    safe_field_motor_material = common_config_data.get("safe_field_motor_material")
    safe_field_motor_thickness = common_config_data.get("safe_field_motor_thickness")
    safe_field_motor_qty = common_config_data.get("safe_field_motor_qty")
    safe_field_motor_isolator_color_shade = common_config_data.get(
        "safe_field_motor_isolator_color_shade"
    )
    safe_field_motor_cable_entry = common_config_data.get(
        "safe_field_motor_cable_entry"
    )
    safe_field_motor_canopy = common_config_data.get("safe_field_motor_canopy")
    # safe_field_motor_canopy_type = common_config_data.get("safe_field_motor_canopy_type")

    hazardous_field_motor_type = common_config_data.get("hazardous_field_motor_type")
    hazardous_field_motor_enclosure = common_config_data.get(
        "hazardous_field_motor_enclosure"
    )
    hazardous_field_motor_material = common_config_data.get(
        "hazardous_field_motor_material"
    )
    # hazardous_field_motor_thickness = common_config_data.get("hazardous_field_motor_thickness")
    hazardous_field_motor_qty = common_config_data.get("hazardous_field_motor_qty")
    hazardous_field_motor_isolator_color_shade = common_config_data.get(
        "hazardous_field_motor_isolator_color_shade"
    )
    hazardous_field_motor_cable_entry = common_config_data.get(
        "hazardous_field_motor_cable_entry"
    )
    hazardous_field_motor_canopy = common_config_data.get(
        "hazardous_field_motor_canopy"
    )
    # hazardous_field_motor_canopy_type = common_config_data.get("hazardous_field_motor_canopy_type")

    design_basis_sheet["C128"] = safe_field_motor_type
    design_basis_sheet["C129"] = safe_field_motor_enclosure
    design_basis_sheet["C130"] = safe_field_motor_material
    design_basis_sheet["C131"] = safe_field_motor_qty
    design_basis_sheet["C132"] = safe_field_motor_isolator_color_shade
    design_basis_sheet["C133"] = safe_field_motor_cable_entry
    design_basis_sheet["C134"] = safe_field_motor_canopy

    design_basis_sheet["D128"] = hazardous_field_motor_type
    design_basis_sheet["D129"] = hazardous_field_motor_enclosure
    design_basis_sheet["D130"] = hazardous_field_motor_material
    design_basis_sheet["D131"] = hazardous_field_motor_qty
    design_basis_sheet["D132"] = hazardous_field_motor_isolator_color_shade
    design_basis_sheet["D133"] = hazardous_field_motor_cable_entry
    design_basis_sheet["D134"] = hazardous_field_motor_canopy

    lpbs_push_button_start_color = common_config_data.get(
        "lpbs_push_button_start_color"
    )
    lpbs_forward_push_button_start = common_config_data.get(
        "lpbs_forward_push_button_start"
    )
    lpbs_reverse_push_button_start = common_config_data.get(
        "lpbs_reverse_push_button_start"
    )
    lpbs_push_button_ess = common_config_data.get("lpbs_push_button_ess")
    lpbs_speed_increase = common_config_data.get("lpbs_speed_increase")
    lpbs_speed_decrease = common_config_data.get("lpbs_speed_decrease")
    lpbs_indication_lamp_start_color = common_config_data.get(
        "lpbs_indication_lamp_start_color"
    )  # ON
    lpbs_indication_lamp_stop_color = common_config_data.get(
        "lpbs_indication_lamp_stop_color"
    )  # OFF

    design_basis_sheet["C136"] = lpbs_push_button_start_color
    design_basis_sheet["C137"] = lpbs_forward_push_button_start
    design_basis_sheet["C138"] = lpbs_reverse_push_button_start
    design_basis_sheet["C139"] = lpbs_push_button_ess
    design_basis_sheet["C140"] = lpbs_speed_increase
    design_basis_sheet["C141"] = lpbs_speed_decrease
    design_basis_sheet["C142"] = lpbs_indication_lamp_start_color
    design_basis_sheet["C143"] = lpbs_indication_lamp_stop_color

    safe_lpbs_type = common_config_data.get("safe_lpbs_type")
    safe_lpbs_enclosure = common_config_data.get("safe_lpbs_enclosure")
    safe_lpbs_material = common_config_data.get("safe_lpbs_material")
    safe_lpbs_qty = common_config_data.get("safe_lpbs_qty")
    safe_lpbs_color_shade = common_config_data.get("safe_lpbs_color_shade")
    safe_lpbs_canopy = common_config_data.get("safe_lpbs_canopy")
    safe_lpbs_canopy_type = common_config_data.get("safe_lpbs_canopy_type")
    safe_lpbs_thickness = common_config_data.get("safe_lpbs_thickness")

    design_basis_sheet["C145"] = safe_lpbs_type
    design_basis_sheet["C146"] = safe_lpbs_enclosure
    design_basis_sheet["C147"] = safe_lpbs_material
    design_basis_sheet["C148"] = safe_lpbs_qty
    design_basis_sheet["C149"] = safe_lpbs_color_shade
    design_basis_sheet["C150"] = safe_lpbs_canopy
    design_basis_sheet["C151"] = safe_lpbs_canopy_type

    hazardous_lpbs_type = common_config_data.get("hazardous_lpbs_type")
    hazardous_lpbs_enclosure = common_config_data.get("hazardous_lpbs_enclosure")
    hazardous_lpbs_material = common_config_data.get("hazardous_lpbs_material")
    hazardous_lpbs_qty = common_config_data.get("hazardous_lpbs_qty")
    hazardous_lpbs_color_shade = common_config_data.get("hazardous_lpbs_color_shade")
    hazardous_lpbs_canopy = common_config_data.get("hazardous_lpbs_canopy")
    hazardous_lpbs_canopy_type = common_config_data.get("hazardous_lpbs_canopy_type")
    hazardous_lpbs_thickness = common_config_data.get("hazardous_lpbs_thickness")

    design_basis_sheet["D145"] = hazardous_lpbs_type
    design_basis_sheet["D146"] = hazardous_lpbs_enclosure
    design_basis_sheet["D147"] = hazardous_lpbs_material
    design_basis_sheet["D148"] = hazardous_lpbs_qty
    design_basis_sheet["D149"] = hazardous_lpbs_color_shade
    design_basis_sheet["D150"] = hazardous_lpbs_canopy
    design_basis_sheet["D151"] = hazardous_lpbs_canopy_type

    apfc_relay = common_config_data.get("apfc_relay")
    # design_basis_sheet["D153"] = f"{apfc_relay} Stage"

    power_bus_main_busbar_selection = common_config_data.get(
        "power_bus_main_busbar_selection"
    )
    power_bus_heat_pvc_sleeve = common_config_data.get("power_bus_heat_pvc_sleeve")
    power_bus_material = common_config_data.get("power_bus_material")
    power_bus_current_density = common_config_data.get("power_bus_current_density")
    power_bus_rating_of_busbar = common_config_data.get("power_bus_rating_of_busbar")

    design_basis_sheet["C155"] = power_bus_main_busbar_selection
    design_basis_sheet["C156"] = power_bus_heat_pvc_sleeve
    design_basis_sheet["C157"] = power_bus_material
    design_basis_sheet["C158"] = power_bus_current_density
    design_basis_sheet["C159"] = power_bus_rating_of_busbar

    control_bus_main_busbar_selection = common_config_data.get(
        "control_bus_main_busbar_selection"
    )
    control_bus_heat_pvc_sleeve = common_config_data.get("control_bus_heat_pvc_sleeve")
    control_bus_material = common_config_data.get("control_bus_material")
    control_bus_current_density = common_config_data.get("control_bus_current_density")
    control_bus_rating_of_busbar = common_config_data.get(
        "control_bus_rating_of_busbar"
    )

    design_basis_sheet["C161"] = control_bus_main_busbar_selection
    design_basis_sheet["C162"] = control_bus_heat_pvc_sleeve
    design_basis_sheet["C163"] = control_bus_material
    design_basis_sheet["C164"] = control_bus_current_density
    design_basis_sheet["C165"] = control_bus_rating_of_busbar

    earth_bus_main_busbar_selection = common_config_data.get(
        "earth_bus_main_busbar_selection"
    )
    earth_bus_busbar_position = common_config_data.get("earth_bus_busbar_position")
    earth_bus_material = common_config_data.get("earth_bus_material")
    earth_bus_current_density = common_config_data.get("earth_bus_current_density")
    earth_bus_rating_of_busbar = common_config_data.get("earth_bus_rating_of_busbar")
    door_earthing = common_config_data.get("door_earthing")
    instrument_earth = common_config_data.get("instrument_earth")
    general_note_busbar_and_insulation_materials = common_config_data.get(
        "general_note_busbar_and_insulation_materials"
    )

    design_basis_sheet["C167"] = earth_bus_main_busbar_selection
    design_basis_sheet["C168"] = earth_bus_busbar_position
    design_basis_sheet["C169"] = earth_bus_material
    design_basis_sheet["C170"] = earth_bus_current_density
    design_basis_sheet["C171"] = earth_bus_rating_of_busbar
    design_basis_sheet["C172"] = door_earthing
    design_basis_sheet["C173"] = instrument_earth
    design_basis_sheet["C174"] = general_note_busbar_and_insulation_materials

    ferrule = common_config_data.get("ferrule")
    ferrule_note = common_config_data.get("ferrule_note")

    design_basis_sheet["C176"] = ferrule
    design_basis_sheet["C177"] = ferrule_note
    design_basis_sheet["C178"] = device_identification_of_components

    cooling_fans = common_config_data.get("cooling_fans")
    louvers_and_filters = common_config_data.get("louvers_and_filters")

    design_basis_sheet["C180"] = cooling_fans
    design_basis_sheet["C181"] = louvers_and_filters

    metering_for_feeders = common_config_data.get("metering_for_feeders")
    alarm_annunciator = common_config_data.get("alarm_annunciator")
    control_transformer = common_config_data.get("control_transformer")
    commissioning_spare = common_config_data.get("commissioning_spare")
    two_year_operational_spare = common_config_data.get("two_year_operational_spare")

    # CABLE TRAY LAYOUT

    cable_tray_data = frappe.db.get_list(
        "Cable Tray Layout", {"revision_id": revision_id}, "*"
    )
    cable_tray_data = cable_tray_data[0]

    number_of_cores = cable_tray_data.get("number_of_cores")
    specific_requirement = cable_tray_data.get("specific_requirement")
    type_of_insulation = cable_tray_data.get("type_of_insulation")
    color_scheme = cable_tray_data.get("color_scheme")
    motor_voltage_drop_during_starting = cable_tray_data.get(
        "motor_voltage_drop_during_starting"
    )
    motor_voltage_drop_during_running = cable_tray_data.get(
        "motor_voltage_drop_during_running"
    )
    voltage_grade = cable_tray_data.get("voltage_grade")
    copper_conductor = cable_tray_data.get("copper_conductor")
    aluminium_conductor = cable_tray_data.get("aluminium_conductor")
    touching_factor_air = cable_tray_data.get("touching_factor_air")
    ambient_temp_factor_air = cable_tray_data.get("ambient_temp_factor_air")
    derating_factor_air = cable_tray_data.get("derating_factor_air")
    touching_factor_burid = cable_tray_data.get("touching_factor_burid")
    ambient_temp_factor_burid = cable_tray_data.get("ambient_temp_factor_burid")
    derating_factor_burid = cable_tray_data.get("derating_factor_burid")
    cable_installation = cable_tray_data.get("cable_installation")

    design_basis_sheet["C183"] = number_of_cores
    design_basis_sheet["C184"] = specific_requirement
    design_basis_sheet["C185"] = type_of_insulation
    design_basis_sheet["C186"] = color_scheme
    design_basis_sheet["C187"] = f"{motor_voltage_drop_during_starting} %"
    design_basis_sheet["C188"] = f"{motor_voltage_drop_during_running} %"
    design_basis_sheet["C189"] = voltage_grade
    design_basis_sheet["C190"] = f"{copper_conductor} Sq. mm"
    design_basis_sheet["C191"] = f"{aluminium_conductor} Sq. mm"
    design_basis_sheet["C192"] = touching_factor_air
    design_basis_sheet["C193"] = ambient_temp_factor_air
    design_basis_sheet["C194"] = derating_factor_air
    design_basis_sheet["C195"] = touching_factor_burid
    design_basis_sheet["C196"] = ambient_temp_factor_burid
    design_basis_sheet["C197"] = derating_factor_burid

    gland_make = cable_tray_data.get("gland_make")
    moc = cable_tray_data.get("moc")
    type_of_gland = cable_tray_data.get("type_of_gland")

    design_basis_sheet["C199"] = gland_make
    design_basis_sheet["C200"] = moc
    design_basis_sheet["C201"] = type_of_gland

    cable_tray_cover = cable_tray_data.get("cable_tray_cover")
    future_space_on_trays = cable_tray_data.get("future_space_on_trays")
    cable_placement = cable_tray_data.get("cable_placement")
    orientation = cable_tray_data.get("orientation")
    vertical_distance = cable_tray_data.get("vertical_distance")
    horizontal_distance = cable_tray_data.get("horizontal_distance")
    cable_tray_moc = cable_tray_data.get("cable_tray_moc")

    design_basis_sheet["C205"] = cable_tray_cover
    design_basis_sheet["C206"] = f"{future_space_on_trays} %"
    design_basis_sheet["C207"] = cable_placement
    design_basis_sheet["C208"] = orientation
    design_basis_sheet["C209"] = f"{vertical_distance} mm"
    design_basis_sheet["C210"] = f"{horizontal_distance} mm"
    design_basis_sheet["C211"] = cable_tray_moc

    cable_tray_moc_input = cable_tray_data.get("cable_tray_moc_input")
    is_dry_area_selected = cable_tray_data.get("is_dry_area_selected")
    is_wet_area_selected = cable_tray_data.get("is_wet_area_selected")

    # EARTHING LAYOUT

    earthing_layout_data = frappe.db.get_list(
        "Layout Earthing", {"revision_id": revision_id}, "*"
    )
    earthing_layout_data = earthing_layout_data[0]

    earthing_system = earthing_layout_data.get("earthing_system")
    earth_strip = earthing_layout_data.get("earth_strip")
    earth_pit = earthing_layout_data.get("earth_pit")
    soil_resistivity = earthing_layout_data.get("soil_resistivity")

    design_basis_sheet["C213"] = earthing_system
    design_basis_sheet["C214"] = earth_strip
    design_basis_sheet["C215"] = earth_pit
    design_basis_sheet["C216"] = f"{soil_resistivity} ohm"

    project_panel_data = frappe.db.get_list(
        "Project Panel Data", {"revision_id": revision_id}, "*"
    )

    for project_panel in project_panel_data:
        if project_panel.get("panel_main_type") == "MCC":

            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"revision_id": revision_id}, "*"
            )
            mcc_panel_data = mcc_panel_data[0]

            panel_id = mcc_panel_data.get("panel_id")
            incomer_ampere = mcc_panel_data.get("incomer_ampere")
            incomer_pole = mcc_panel_data.get("incomer_pole")
            incomer_type = mcc_panel_data.get("incomer_type")
            incomer_above_ampere = mcc_panel_data.get("incomer_above_ampere")
            incomer_above_pole = mcc_panel_data.get("incomer_above_pole")
            incomer_above_type = mcc_panel_data.get("incomer_above_type")
            is_under_or_over_voltage_selected = mcc_panel_data.get(
                "is_under_or_over_voltage_selected"
            )
            is_lsig_selected = mcc_panel_data.get("is_lsig_selected")
            is_lsi_selected = mcc_panel_data.get("is_lsi_selected")
            is_neural_link_with_disconnect_facility_selected = mcc_panel_data.get(
                "is_neural_link_with_disconnect_facility_selected"
            )
            is_led_type_lamp_selected = mcc_panel_data.get("is_led_type_lamp_selected")

            is_indication_on_selected = mcc_panel_data.get("is_indication_on_selected")
            is_indication_off_selected = mcc_panel_data.get(
                "is_indication_off_selected"
            )
            is_indication_trip_selected = mcc_panel_data.get(
                "is_indication_trip_selected"
            )
            led_type_on_input = mcc_panel_data.get("led_type_on_input")
            led_type_off_input = mcc_panel_data.get("led_type_off_input")
            led_type_trip_input = mcc_panel_data.get("led_type_trip_input")
            is_blue_cb_spring_charge_selected = mcc_panel_data.get(
                "is_blue_cb_spring_charge_selected"
            )
            is_red_cb_in_service = mcc_panel_data.get("is_red_cb_in_service")
            is_white_healthy_trip_circuit_selected = mcc_panel_data.get(
                "is_white_healthy_trip_circuit_selected"
            )
            alarm_annunciator = mcc_panel_data.get("alarm_annunciator")
            # led_type_other_input = mcc_panel_data.get("led_type_other_input")
            # is_other_selected = mcc_panel_data.get("is_other_selected")

            mcc_sheet["C6"] = led_type_on_input
            mcc_sheet["C7"] = led_type_off_input
            mcc_sheet["C8"] = led_type_trip_input
            mcc_sheet["C9"] = is_blue_cb_spring_charge_selected
            mcc_sheet["C10"] = is_red_cb_in_service
            mcc_sheet["C11"] = is_white_healthy_trip_circuit_selected
            mcc_sheet["C12"] = alarm_annunciator

            mi_analog = mcc_panel_data.get("mi_analog")
            mi_digital = mcc_panel_data.get("mi_digital")
            mi_communication_protocol = mcc_panel_data.get("mi_communication_protocol")

            mcc_sheet["C14"] = mi_analog
            mcc_sheet["C15"] = mi_digital
            mcc_sheet["C16"] = mi_communication_protocol

            mcc_current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
            )
            mcc_current_transformer_number = mcc_panel_data.get(
                "current_transformer_number"
            )
            mcc_current_transformer_configuration = mcc_panel_data.get(
                "current_transformer_configuration"
            )

            mcc_sheet["C18"] = mcc_current_transformer_coating
            mcc_sheet["C19"] = mcc_current_transformer_number
            mcc_sheet["C20"] = mcc_current_transformer_configuration

            mcc_control_transformer_coating = mcc_panel_data.get(
                "control_transformer_coating"
            )
            mcc_control_transformer_configuration = mcc_panel_data.get(
                "control_transformer_configuration"
            )

            is_power_and_bus_separation_section_selected = mcc_panel_data.get(
                "is_power_and_bus_separation_section_selected"
            )
            is_both_side_extension_section_selected = mcc_panel_data.get(
                "is_both_side_extension_section_selected"
            )

            ga_moc_material = mcc_panel_data.get("ga_moc_material")
            ga_moc_thickness_door = mcc_panel_data.get("ga_moc_thickness_door")
            ga_moc_thickness_covers = mcc_panel_data.get(
                "ga_moc_thickness_covers"
            )  # top and side thickness
            ga_gland_plate_thickness = mcc_panel_data.get("ga_gland_plate_thickness")
            ga_panel_mounting_frame = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            marshalling_section_text_area = mcc_panel_data.get(
                "marshalling_section_text_area"
            )
            is_cable_alley_section_selected = mcc_panel_data.get(
                "is_cable_alley_section_selected"
            )
            ga_power_and_control_busbar_separation = mcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            ga_enclosure_protection_degree = mcc_panel_data.get(
                "ga_enclosure_protection_degree"
            )
            ga_cable_entry_position = mcc_panel_data.get("ga_cable_entry_position")
            ga_busbar_chamber_position = mcc_panel_data.get(
                "ga_busbar_chamber_position"
            )

            mcc_sheet["C22"] = ga_moc_material
            mcc_sheet["C23"] = ""
            mcc_sheet["C24"] = ga_moc_thickness_door
            mcc_sheet["C25"] = ga_moc_thickness_covers
            mcc_sheet["C26"] = ga_gland_plate_thickness
            mcc_sheet["C27"] = ""
            mcc_sheet["C28"] = ""
            mcc_sheet["C29"] = ga_panel_mounting_frame
            mcc_sheet["C30"] = ga_panel_mounting_height
            mcc_sheet["C31"] = marshalling_section_text_area
            mcc_sheet["C32"] = is_cable_alley_section_selected
            mcc_sheet["C33"] = ga_power_and_control_busbar_separation
            mcc_sheet["C34"] = is_both_side_extension_section_selected
            mcc_sheet["C35"] = ga_busbar_chamber_position
            mcc_sheet["C36"] = is_both_side_extension_section_selected
            mcc_sheet["C37"] = ga_enclosure_protection_degree
            mcc_sheet["C38"] = ga_cable_entry_position

            general_requirments_for_construction = mcc_panel_data.get(
                "general_requirments_for_construction"
            )
            ppc_interior_and_exterior_paint_shade = mcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = mcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            ppc_minimum_coating_thickness = mcc_panel_data.get(
                "ppc_minimum_coating_thickness"
            )
            ppc_pretreatment_panel_standard = mcc_panel_data.get(
                "ppc_pretreatment_panel_standard"
            )

            ga_gland_plate_3mm_drill_type = mcc_panel_data.get(
                "ga_gland_plate_3mm_drill_type"
            )
            ga_mcc_compartmental = mcc_panel_data.get("ga_mcc_compartmental")
            ga_mcc_construction_front_type = mcc_panel_data.get(
                "ga_mcc_construction_front_type"
            )
            incoming_drawout_type = mcc_panel_data.get("incoming_drawout_type")
            outgoing_drawout_type = mcc_panel_data.get("outgoing_drawout_type")
            ga_mcc_construction_type = mcc_panel_data.get("ga_mcc_construction_type")
            busbar_material_of_construction = mcc_panel_data.get(
                "busbar_material_of_construction"
            )
            ga_current_density = mcc_panel_data.get("ga_current_density")
            is_marshalling_section_selected = mcc_panel_data.get(
                "is_marshalling_section_selected"
            )

            mcc_sheet["C40"] = "As per OEM Standard"
            mcc_sheet["C41"] = ppc_interior_and_exterior_paint_shade
            mcc_sheet["C42"] = ppc_component_mounting_plate_paint_shade
            mcc_sheet["C43"] = ppc_minimum_coating_thickness
            mcc_sheet["C44"] = "Black"
            mcc_sheet["C45"] = ppc_pretreatment_panel_standard
            mcc_sheet["C46"] = general_requirments_for_construction

            vfd_auto_manual_selection = mcc_panel_data.get("vfd_auto_manual_selection")
            two_year_operational_spare = mcc_panel_data.get(
                "two_year_operational_spare"
            )
            commissioning_spare = mcc_panel_data.get("commissioning_spare")

            mcc_sheet["C48"] = vfd_auto_manual_selection
            mcc_sheet["C50"] = commissioning_spare
            mcc_sheet["C51"] = two_year_operational_spare

            ppc_painting_standards = mcc_panel_data.get("ppc_painting_standards")
            ppc_base_frame_paint_shade = mcc_panel_data.get(
                "ppc_base_frame_paint_shade"
            )
            is_punching_details_for_boiler_selected = mcc_panel_data.get(
                "is_punching_details_for_boiler_selected"
            )
            boiler_model = mcc_panel_data.get("boiler_model")
            boiler_fuel = mcc_panel_data.get("boiler_fuel")
            boiler_year = mcc_panel_data.get("boiler_year")
            boiler_power_supply_vac = mcc_panel_data.get("boiler_power_supply_vac")
            boiler_power_supply_phase = mcc_panel_data.get("boiler_power_supply_phase")
            boiler_power_supply_frequency = mcc_panel_data.get(
                "boiler_power_supply_frequency"
            )
            boiler_control_supply_vac = mcc_panel_data.get("boiler_control_supply_vac")
            boiler_control_supply_phase = mcc_panel_data.get(
                "boiler_control_supply_phase"
            )
            boiler_control_supply_frequency = mcc_panel_data.get(
                "boiler_control_supply_frequency"
            )
            boiler_evaporation = mcc_panel_data.get("boiler_evaporation")
            boiler_output = mcc_panel_data.get("boiler_output")
            boiler_connected_load = mcc_panel_data.get("boiler_connected_load")
            boiler_design_pressure = mcc_panel_data.get("boiler_design_pressure")

            mcc_sheet["C54"] = boiler_model
            mcc_sheet["C55"] = boiler_fuel
            mcc_sheet["C56"] = boiler_year
            mcc_sheet["C57"] = (
                f"{boiler_power_supply_vac}, {boiler_power_supply_phase}, {boiler_power_supply_frequency}"
            )
            mcc_sheet["C58"] = (
                f"{boiler_control_supply_vac}, {boiler_control_supply_phase}, {boiler_control_supply_frequency}"
            )
            mcc_sheet["C59"] = f"{boiler_evaporation} kg/Hr"
            mcc_sheet["C60"] = f"{boiler_output} MW"
            mcc_sheet["C61"] = f"{boiler_connected_load} kW"
            mcc_sheet["C62"] = f"{boiler_design_pressure} kg/cm2(g)/Bar"

            is_punching_details_for_heater_selected = mcc_panel_data.get(
                "is_punching_details_for_heater_selected"
            )
            heater_model = mcc_panel_data.get("heater_model")
            heater_fuel = mcc_panel_data.get("heater_fuel")
            heater_year = mcc_panel_data.get("heater_year")
            heater_power_supply_vac = mcc_panel_data.get("heater_power_supply_vac")
            heater_power_supply_phase = mcc_panel_data.get("heater_power_supply_phase")
            heater_power_supply_frequency = mcc_panel_data.get(
                "heater_power_supply_frequency"
            )
            heater_control_supply_vac = mcc_panel_data.get("heater_control_supply_vac")
            heater_control_supply_phase = mcc_panel_data.get(
                "heater_control_supply_phase"
            )
            heater_control_supply_frequency = mcc_panel_data.get(
                "heater_control_supply_frequency"
            )
            heater_evaporation = mcc_panel_data.get("heater_evaporation")
            heater_output = mcc_panel_data.get("heater_output")
            heater_connected_load = mcc_panel_data.get("heater_connected_load")
            heater_temperature = mcc_panel_data.get("heater_temperature")

            mcc_sheet["C54"] = heater_model
            mcc_sheet["C55"] = heater_fuel
            mcc_sheet["C56"] = heater_year
            mcc_sheet["C57"] = (
                f"{heater_power_supply_vac}, {heater_power_supply_phase}, {heater_power_supply_frequency}"
            )
            mcc_sheet["C58"] = (
                f"{heater_control_supply_vac}, {heater_control_supply_phase}, {heater_control_supply_frequency}"
            )
            mcc_sheet["C59"] = f"{heater_evaporation} Kcl/Hr"
            mcc_sheet["C60"] = f"{heater_output} MW"
            mcc_sheet["C61"] = f"{heater_connected_load} kW"
            mcc_sheet["C62"] = f"{heater_temperature} Deg. C"

            is_spg_applicable = mcc_panel_data.get("is_spg_applicable")
            spg_name_plate_unit_name = mcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = mcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = mcc_panel_data.get(
                "spg_name_plate_manufacturing_year"
            )
            spg_name_plate_weight = mcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = mcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = mcc_panel_data.get("spg_name_plate_part_code")
            # special_note = mcc_panel_data.get("special_note")

            mcc_sheet["C74"] = spg_name_plate_unit_name
            mcc_sheet["C75"] = spg_name_plate_capacity
            mcc_sheet["C76"] = spg_name_plate_manufacturing_year
            mcc_sheet["C77"] = spg_name_plate_weight
            mcc_sheet["C78"] = spg_name_plate_oc_number
            mcc_sheet["C79"] = spg_name_plate_part_code

        elif project_panel.get("panel_main_type") == "PCC":

            pcc_panel_data = frappe.db.get_list(
                "PCC Panel", {"revision_id": revision_id}, "*"
            )
            pcc_panel_data = pcc_panel_data[0]

            panel_id = pcc_panel_data.get("panel_id")
            incomer_ampere = pcc_panel_data.get("incomer_ampere")
            incomer_pole = pcc_panel_data.get("incomer_pole")
            incomer_type = pcc_panel_data.get("incomer_type")
            incomer_above_ampere = pcc_panel_data.get("incomer_above_ampere")
            incomer_above_pole = pcc_panel_data.get("incomer_above_pole")
            incomer_above_type = pcc_panel_data.get("incomer_above_type")
            is_under_or_over_voltage_selected = pcc_panel_data.get(
                "is_under_or_over_voltage_selected"
            )
            is_lsig_selected = pcc_panel_data.get("is_lsig_selected")
            is_lsi_selected = pcc_panel_data.get("is_lsi_selected")
            is_neural_link_with_disconnect_facility_selected = pcc_panel_data.get(
                "is_neural_link_with_disconnect_facility_selected"
            )
            is_led_type_lamp_selected = pcc_panel_data.get("is_led_type_lamp_selected")
            is_indication_on_selected = pcc_panel_data.get("is_indication_on_selected")
            led_type_on_input = pcc_panel_data.get("led_type_on_input")
            is_indication_off_selected = pcc_panel_data.get(
                "is_indication_off_selected"
            )
            led_type_off_input = pcc_panel_data.get("led_type_off_input")
            is_indication_trip_selected = pcc_panel_data.get(
                "is_indication_trip_selected"
            )
            led_type_trip_input = pcc_panel_data.get("led_type_trip_input")
            is_blue_cb_spring_charge_selected = pcc_panel_data.get(
                "is_blue_cb_spring_charge_selected"
            )
            is_red_cb_in_service = pcc_panel_data.get("is_red_cb_in_service")
            is_white_healthy_trip_circuit_selected = pcc_panel_data.get(
                "is_white_healthy_trip_circuit_selected"
            )
            is_other_selected = pcc_panel_data.get("is_other_selected")
            control_transformer_coating = pcc_panel_data.get(
                "control_transformer_coating"
            )
            control_transformer_configuration = pcc_panel_data.get(
                "control_transformer_configuration"
            )
            current_transformer_coating = pcc_panel_data.get(
                "current_transformer_coating"
            )
            current_transformer_number = pcc_panel_data.get(
                "current_transformer_number"
            )
            current_transformer_configuration = pcc_panel_data.get(
                "current_transformer_configuration"
            )
            alarm_annunciator = pcc_panel_data.get("alarm_annunciator")
            led_type_other_input = pcc_panel_data.get("led_type_other_input")
            mi_analog = pcc_panel_data.get("mi_analog")
            mi_digital = pcc_panel_data.get("mi_digital")
            mi_communication_protocol = pcc_panel_data.get("mi_communication_protocol")
            ga_moc_material = pcc_panel_data.get("ga_moc_material")
            ga_moc_thickness_door = pcc_panel_data.get("ga_moc_thickness_door")
            ga_moc_thickness_covers = pcc_panel_data.get("ga_moc_thickness_covers")
            ga_pcc_compartmental = pcc_panel_data.get("ga_pcc_compartmental")
            ga_pcc_construction_front_type = pcc_panel_data.get(
                "ga_pcc_construction_front_type"
            )
            ga_pcc_construction_type = pcc_panel_data.get("ga_pcc_construction_type")
            incoming_drawout_type = pcc_panel_data.get("incoming_drawout_type")
            outgoing_drawout_type = pcc_panel_data.get("outgoing_drawout_type")
            busbar_material_of_construction = pcc_panel_data.get(
                "busbar_material_of_construction"
            )
            ga_current_density = pcc_panel_data.get("ga_current_density")
            ga_panel_mounting_frame = pcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = pcc_panel_data.get("ga_panel_mounting_height")
            is_marshalling_section_selected = pcc_panel_data.get(
                "is_marshalling_section_selected"
            )
            marshalling_section_text_area = pcc_panel_data.get(
                "marshalling_section_text_area"
            )
            is_cable_alley_section_selected = pcc_panel_data.get(
                "is_cable_alley_section_selected"
            )
            is_power_and_bus_separation_section_selected = pcc_panel_data.get(
                "is_power_and_bus_separation_section_selected"
            )
            is_both_side_extension_section_selected = pcc_panel_data.get(
                "is_both_side_extension_section_selected"
            )
            ga_gland_plate_3mm_drill_type = pcc_panel_data.get(
                "ga_gland_plate_3mm_drill_type"
            )
            ga_gland_plate_3mm_attachment_type = pcc_panel_data.get(
                "ga_gland_plate_3mm_attachment_type"
            )
            ga_busbar_chamber_position = pcc_panel_data.get(
                "ga_busbar_chamber_position"
            )
            ga_power_and_control_busbar_separation = pcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            ga_enclosure_protection_degree = pcc_panel_data.get(
                "ga_enclosure_protection_degree"
            )
            ga_cable_entry_position = pcc_panel_data.get("ga_cable_entry_position")
            general_requirments_for_construction = pcc_panel_data.get(
                "general_requirments_for_construction"
            )
            ppc_painting_standards = pcc_panel_data.get("ppc_painting_standards")
            ppc_interior_and_exterior_paint_shade = pcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = pcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            ppc_base_frame_paint_shade = pcc_panel_data.get(
                "ppc_base_frame_paint_shade"
            )
            ppc_minimum_coating_thickness = pcc_panel_data.get(
                "ppc_minimum_coating_thickness"
            )
            ppc_pretreatment_panel_standard = pcc_panel_data.get(
                "ppc_pretreatment_panel_standard"
            )
            commissioning_spare = pcc_panel_data.get("commissioning_spare")
            two_year_operational_spare = pcc_panel_data.get(
                "two_year_operational_spare"
            )
            is_punching_details_for_boiler_selected = pcc_panel_data.get(
                "is_punching_details_for_boiler_selected"
            )
            boiler_model = pcc_panel_data.get("boiler_model")
            boiler_fuel = pcc_panel_data.get("boiler_fuel")
            boiler_year = pcc_panel_data.get("boiler_year")
            boiler_power_supply_vac = pcc_panel_data.get("boiler_power_supply_vac")
            boiler_power_supply_phase = pcc_panel_data.get("boiler_power_supply_phase")
            boiler_power_supply_frequency = pcc_panel_data.get(
                "boiler_power_supply_frequency"
            )
            boiler_control_supply_vac = pcc_panel_data.get("boiler_control_supply_vac")
            boiler_control_supply_phase = pcc_panel_data.get(
                "boiler_control_supply_phase"
            )
            boiler_control_supply_frequency = pcc_panel_data.get(
                "boiler_control_supply_frequency"
            )
            boiler_evaporation = pcc_panel_data.get("boiler_evaporation")
            boiler_output = pcc_panel_data.get("boiler_output")
            boiler_connected_load = pcc_panel_data.get("boiler_connected_load")
            boiler_design_pressure = pcc_panel_data.get("boiler_design_pressure")
            is_punching_details_for_heater_selected = pcc_panel_data.get(
                "is_punching_details_for_heater_selected"
            )
            heater_model = pcc_panel_data.get("heater_model")
            heater_fuel = pcc_panel_data.get("heater_fuel")
            heater_year = pcc_panel_data.get("heater_year")
            heater_power_supply_vac = pcc_panel_data.get("heater_power_supply_vac")
            heater_power_supply_phase = pcc_panel_data.get("heater_power_supply_phase")
            heater_power_supply_frequency = pcc_panel_data.get(
                "heater_power_supply_frequency"
            )
            heater_control_supply_vac = pcc_panel_data.get("heater_control_supply_vac")
            heater_control_supply_phase = pcc_panel_data.get(
                "heater_control_supply_phase"
            )
            heater_control_supply_frequency = pcc_panel_data.get(
                "heater_control_supply_frequency"
            )
            heater_evaporation = pcc_panel_data.get("heater_evaporation")
            heater_output = pcc_panel_data.get("heater_output")
            heater_connected_load = pcc_panel_data.get("heater_connected_load")
            heater_temperature = pcc_panel_data.get("heater_temperature")
            is_spg_applicable = pcc_panel_data.get("is_spg_applicable")
            spg_name_plate_unit_name = pcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = pcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = pcc_panel_data.get(
                "spg_name_plate_manufacturing_year"
            )
            spg_name_plate_weight = pcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = pcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = pcc_panel_data.get("spg_name_plate_part_code")
            special_note = pcc_panel_data.get("special_note")

            pcc_sheet["C6"] = led_type_on_input
            pcc_sheet["C7"] = led_type_off_input
            pcc_sheet["C8"] = led_type_trip_input
            pcc_sheet["C9"] = is_blue_cb_spring_charge_selected
            pcc_sheet["C10"] = is_red_cb_in_service
            pcc_sheet["C11"] = is_white_healthy_trip_circuit_selected
            pcc_sheet["C12"] = alarm_annunciator

            pcc_sheet["C14"] = mi_analog
            pcc_sheet["C15"] = mi_digital
            pcc_sheet["C16"] = mi_communication_protocol
            pcc_sheet["C18"] = mcc_current_transformer_coating
            pcc_sheet["C19"] = mcc_current_transformer_number
            pcc_sheet["C20"] = mcc_current_transformer_configuration

            pcc_sheet["C22"] = ga_moc_material
            pcc_sheet["C23"] = ""
            pcc_sheet["C24"] = ga_moc_thickness_door
            pcc_sheet["C25"] = ga_moc_thickness_covers
            pcc_sheet["C26"] = ga_gland_plate_thickness
            pcc_sheet["C27"] = ""
            pcc_sheet["C28"] = ""
            pcc_sheet["C29"] = ga_panel_mounting_frame
            pcc_sheet["C30"] = ga_panel_mounting_height
            pcc_sheet["C31"] = marshalling_section_text_area
            pcc_sheet["C32"] = is_cable_alley_section_selected
            pcc_sheet["C33"] = ga_power_and_control_busbar_separation
            pcc_sheet["C34"] = is_both_side_extension_section_selected
            pcc_sheet["C35"] = ga_busbar_chamber_position
            pcc_sheet["C36"] = is_both_side_extension_section_selected
            pcc_sheet["C37"] = ga_enclosure_protection_degree
            pcc_sheet["C38"] = ga_cable_entry_position

            pcc_sheet["C40"] = "As per OEM Standard"
            pcc_sheet["C41"] = ppc_interior_and_exterior_paint_shade
            pcc_sheet["C42"] = ppc_component_mounting_plate_paint_shade
            pcc_sheet["C43"] = ppc_minimum_coating_thickness
            pcc_sheet["C44"] = "Black"
            pcc_sheet["C45"] = ppc_pretreatment_panel_standard
            pcc_sheet["C46"] = general_requirments_for_construction

            pcc_sheet["C48"] = vfd_auto_manual_selection
            # pcc_sheet["C50"] = commissioning_spare
            # pcc_sheet["C51"] = two_year_operational_spare

            pcc_sheet["C54"] = boiler_model
            pcc_sheet["C55"] = boiler_fuel
            pcc_sheet["C56"] = boiler_year
            pcc_sheet["C57"] = (
                f"{boiler_power_supply_vac}, {boiler_power_supply_phase}, {boiler_power_supply_frequency}"
            )
            pcc_sheet["C58"] = (
                f"{boiler_control_supply_vac}, {boiler_control_supply_phase}, {boiler_control_supply_frequency}"
            )
            pcc_sheet["C59"] = f"{boiler_evaporation} kg/Hr"
            pcc_sheet["C60"] = f"{boiler_output} MW"
            # pcc_sheet["C61"] = f"{boiler_connected_load} kW"
            pcc_sheet["C62"] = f"{boiler_design_pressure} kg/cm2(g)/Bar"

            pcc_sheet["C54"] = heater_model
            pcc_sheet["C55"] = heater_fuel
            pcc_sheet["C56"] = heater_year
            pcc_sheet["C57"] = (
                f"{heater_power_supply_vac}, {heater_power_supply_phase}, {heater_power_supply_frequency}"
            )
            pcc_sheet["C58"] = (
                f"{heater_control_supply_vac}, {heater_control_supply_phase}, {heater_control_supply_frequency}"
            )
            pcc_sheet["C59"] = f"{heater_evaporation} Kcl/Hr"
            pcc_sheet["C60"] = f"{heater_output} MW"
            # pcc_sheet["C61"] = f"{heater_connected_load} kW"
            pcc_sheet["C62"] = f"{heater_temperature} Deg. C"

            pcc_sheet["C74"] = spg_name_plate_unit_name
            pcc_sheet["C75"] = spg_name_plate_capacity
            pcc_sheet["C76"] = spg_name_plate_manufacturing_year
            pcc_sheet["C77"] = spg_name_plate_weight
            pcc_sheet["C78"] = spg_name_plate_oc_number
            pcc_sheet["C79"] = spg_name_plate_part_code

        elif project_panel.get("panel_main_type") == "MCC cum PCC":

            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"revision_id": revision_id}, "*"
            )
            mcc_panel_data = mcc_panel_data[0]

            plc_1 = frappe.db.get_list(
                "Panel PLC 1 - 3", {"revision_id": revision_id}, "*"
            )
            plc_1 = plc_1[0] if plc_1 else {}

            plc_2 = frappe.db.get_list(
                "Panel PLC 2 - 3", {"revision_id": revision_id}, "*"
            )
            plc_2 = plc_2[0] if plc_2 else {}

            plc_3 = frappe.db.get_list(
                "Panel PLC 3 - 3", {"revision_id": revision_id}, "*"
            )
            plc_3 = plc_3[0] if plc_3 else {}

            plc_panel_data = plc_1 | plc_2 | plc_3

            panel_id = mcc_panel_data.get("panel_id")
            incomer_ampere = mcc_panel_data.get("incomer_ampere")
            incomer_pole = mcc_panel_data.get("incomer_pole")
            incomer_type = mcc_panel_data.get("incomer_type")
            incomer_above_ampere = mcc_panel_data.get("incomer_above_ampere")
            incomer_above_pole = mcc_panel_data.get("incomer_above_pole")
            incomer_above_type = mcc_panel_data.get("incomer_above_type")
            is_under_or_over_voltage_selected = mcc_panel_data.get(
                "is_under_or_over_voltage_selected"
            )
            is_lsig_selected = mcc_panel_data.get("is_lsig_selected")
            is_lsi_selected = mcc_panel_data.get("is_lsi_selected")
            is_neural_link_with_disconnect_facility_selected = mcc_panel_data.get(
                "is_neural_link_with_disconnect_facility_selected"
            )
            is_led_type_lamp_selected = mcc_panel_data.get("is_led_type_lamp_selected")

            is_indication_on_selected = mcc_panel_data.get("is_indication_on_selected")
            is_indication_off_selected = mcc_panel_data.get(
                "is_indication_off_selected"
            )
            is_indication_trip_selected = mcc_panel_data.get(
                "is_indication_trip_selected"
            )
            led_type_on_input = mcc_panel_data.get("led_type_on_input")
            led_type_off_input = mcc_panel_data.get("led_type_off_input")
            led_type_trip_input = mcc_panel_data.get("led_type_trip_input")
            is_blue_cb_spring_charge_selected = mcc_panel_data.get(
                "is_blue_cb_spring_charge_selected"
            )
            is_red_cb_in_service = mcc_panel_data.get("is_red_cb_in_service")
            is_white_healthy_trip_circuit_selected = mcc_panel_data.get(
                "is_white_healthy_trip_circuit_selected"
            )
            alarm_annunciator = mcc_panel_data.get("alarm_annunciator")
            # led_type_other_input = mcc_panel_data.get("led_type_other_input")
            # is_other_selected = mcc_panel_data.get("is_other_selected")

            mcc_sheet["C6"] = led_type_on_input
            mcc_sheet["C7"] = led_type_off_input
            mcc_sheet["C8"] = led_type_trip_input
            mcc_sheet["C9"] = is_blue_cb_spring_charge_selected
            mcc_sheet["C10"] = is_red_cb_in_service
            mcc_sheet["C11"] = is_white_healthy_trip_circuit_selected
            mcc_sheet["C12"] = alarm_annunciator

            mi_analog = mcc_panel_data.get("mi_analog")
            mi_digital = mcc_panel_data.get("mi_digital")
            mi_communication_protocol = mcc_panel_data.get("mi_communication_protocol")

            mcc_sheet["C14"] = mi_analog
            mcc_sheet["C15"] = mi_digital
            mcc_sheet["C16"] = mi_communication_protocol

            mcc_current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
            )
            mcc_current_transformer_number = mcc_panel_data.get(
                "current_transformer_number"
            )
            mcc_current_transformer_configuration = mcc_panel_data.get(
                "current_transformer_configuration"
            )

            mcc_sheet["C18"] = mcc_current_transformer_coating
            mcc_sheet["C19"] = mcc_current_transformer_number
            mcc_sheet["C20"] = mcc_current_transformer_configuration

            mcc_control_transformer_coating = mcc_panel_data.get(
                "control_transformer_coating"
            )
            mcc_control_transformer_configuration = mcc_panel_data.get(
                "control_transformer_configuration"
            )

            is_power_and_bus_separation_section_selected = mcc_panel_data.get(
                "is_power_and_bus_separation_section_selected"
            )
            is_both_side_extension_section_selected = mcc_panel_data.get(
                "is_both_side_extension_section_selected"
            )

            ga_moc_material = mcc_panel_data.get("ga_moc_material")
            ga_moc_thickness_door = mcc_panel_data.get("ga_moc_thickness_door")
            ga_moc_thickness_covers = mcc_panel_data.get(
                "ga_moc_thickness_covers"
            )  # top and side thickness
            ga_gland_plate_thickness = mcc_panel_data.get("ga_gland_plate_thickness")
            ga_panel_mounting_frame = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            marshalling_section_text_area = mcc_panel_data.get(
                "marshalling_section_text_area"
            )
            is_cable_alley_section_selected = mcc_panel_data.get(
                "is_cable_alley_section_selected"
            )
            ga_power_and_control_busbar_separation = mcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            ga_enclosure_protection_degree = mcc_panel_data.get(
                "ga_enclosure_protection_degree"
            )
            ga_cable_entry_position = mcc_panel_data.get("ga_cable_entry_position")
            ga_busbar_chamber_position = mcc_panel_data.get(
                "ga_busbar_chamber_position"
            )

            mcc_sheet["C22"] = ga_moc_material
            mcc_sheet["C23"] = ""
            mcc_sheet["C24"] = ga_moc_thickness_door
            mcc_sheet["C25"] = ga_moc_thickness_covers
            mcc_sheet["C26"] = ga_gland_plate_thickness
            mcc_sheet["C27"] = ""
            mcc_sheet["C28"] = ""
            mcc_sheet["C29"] = ga_panel_mounting_frame
            mcc_sheet["C30"] = ga_panel_mounting_height
            mcc_sheet["C31"] = marshalling_section_text_area
            mcc_sheet["C32"] = is_cable_alley_section_selected
            mcc_sheet["C33"] = ga_power_and_control_busbar_separation
            mcc_sheet["C34"] = is_both_side_extension_section_selected
            mcc_sheet["C35"] = ga_busbar_chamber_position
            mcc_sheet["C36"] = is_both_side_extension_section_selected
            mcc_sheet["C37"] = ga_enclosure_protection_degree
            mcc_sheet["C38"] = ga_cable_entry_position

            general_requirments_for_construction = mcc_panel_data.get(
                "general_requirments_for_construction"
            )
            ppc_interior_and_exterior_paint_shade = mcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = mcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            ppc_minimum_coating_thickness = mcc_panel_data.get(
                "ppc_minimum_coating_thickness"
            )
            ppc_pretreatment_panel_standard = mcc_panel_data.get(
                "ppc_pretreatment_panel_standard"
            )

            ga_gland_plate_3mm_drill_type = mcc_panel_data.get(
                "ga_gland_plate_3mm_drill_type"
            )
            ga_mcc_compartmental = mcc_panel_data.get("ga_mcc_compartmental")
            ga_mcc_construction_front_type = mcc_panel_data.get(
                "ga_mcc_construction_front_type"
            )
            incoming_drawout_type = mcc_panel_data.get("incoming_drawout_type")
            outgoing_drawout_type = mcc_panel_data.get("outgoing_drawout_type")
            ga_mcc_construction_type = mcc_panel_data.get("ga_mcc_construction_type")
            busbar_material_of_construction = mcc_panel_data.get(
                "busbar_material_of_construction"
            )
            ga_current_density = mcc_panel_data.get("ga_current_density")
            is_marshalling_section_selected = mcc_panel_data.get(
                "is_marshalling_section_selected"
            )

            mcc_sheet["C40"] = "As per OEM Standard"
            mcc_sheet["C41"] = ppc_interior_and_exterior_paint_shade
            mcc_sheet["C42"] = ppc_component_mounting_plate_paint_shade
            mcc_sheet["C43"] = ppc_minimum_coating_thickness
            mcc_sheet["C44"] = "Black"
            mcc_sheet["C45"] = ppc_pretreatment_panel_standard
            mcc_sheet["C46"] = general_requirments_for_construction

            vfd_auto_manual_selection = mcc_panel_data.get("vfd_auto_manual_selection")
            two_year_operational_spare = mcc_panel_data.get(
                "two_year_operational_spare"
            )
            commissioning_spare = mcc_panel_data.get("commissioning_spare")

            mcc_sheet["C48"] = vfd_auto_manual_selection
            mcc_sheet["C50"] = commissioning_spare
            mcc_sheet["C51"] = two_year_operational_spare

            ppc_painting_standards = mcc_panel_data.get("ppc_painting_standards")
            ppc_base_frame_paint_shade = mcc_panel_data.get(
                "ppc_base_frame_paint_shade"
            )
            is_punching_details_for_boiler_selected = mcc_panel_data.get(
                "is_punching_details_for_boiler_selected"
            )
            boiler_model = mcc_panel_data.get("boiler_model")
            boiler_fuel = mcc_panel_data.get("boiler_fuel")
            boiler_year = mcc_panel_data.get("boiler_year")
            boiler_power_supply_vac = mcc_panel_data.get("boiler_power_supply_vac")
            boiler_power_supply_phase = mcc_panel_data.get("boiler_power_supply_phase")
            boiler_power_supply_frequency = mcc_panel_data.get(
                "boiler_power_supply_frequency"
            )
            boiler_control_supply_vac = mcc_panel_data.get("boiler_control_supply_vac")
            boiler_control_supply_phase = mcc_panel_data.get(
                "boiler_control_supply_phase"
            )
            boiler_control_supply_frequency = mcc_panel_data.get(
                "boiler_control_supply_frequency"
            )
            boiler_evaporation = mcc_panel_data.get("boiler_evaporation")
            boiler_output = mcc_panel_data.get("boiler_output")
            boiler_connected_load = mcc_panel_data.get("boiler_connected_load")
            boiler_design_pressure = mcc_panel_data.get("boiler_design_pressure")

            mcc_sheet["C54"] = boiler_model
            mcc_sheet["C55"] = boiler_fuel
            mcc_sheet["C56"] = boiler_year
            mcc_sheet["C57"] = (
                f"{boiler_power_supply_vac}, {boiler_power_supply_phase}, {boiler_power_supply_frequency}"
            )
            mcc_sheet["C58"] = (
                f"{boiler_control_supply_vac}, {boiler_control_supply_phase}, {boiler_control_supply_frequency}"
            )
            mcc_sheet["C59"] = f"{boiler_evaporation} kg/Hr"
            mcc_sheet["C60"] = f"{boiler_output} MW"
            mcc_sheet["C61"] = f"{boiler_connected_load} kW"
            mcc_sheet["C62"] = f"{boiler_design_pressure} kg/cm2(g)/Bar"

            is_punching_details_for_heater_selected = mcc_panel_data.get(
                "is_punching_details_for_heater_selected"
            )
            heater_model = mcc_panel_data.get("heater_model")
            heater_fuel = mcc_panel_data.get("heater_fuel")
            heater_year = mcc_panel_data.get("heater_year")
            heater_power_supply_vac = mcc_panel_data.get("heater_power_supply_vac")
            heater_power_supply_phase = mcc_panel_data.get("heater_power_supply_phase")
            heater_power_supply_frequency = mcc_panel_data.get(
                "heater_power_supply_frequency"
            )
            heater_control_supply_vac = mcc_panel_data.get("heater_control_supply_vac")
            heater_control_supply_phase = mcc_panel_data.get(
                "heater_control_supply_phase"
            )
            heater_control_supply_frequency = mcc_panel_data.get(
                "heater_control_supply_frequency"
            )
            heater_evaporation = mcc_panel_data.get("heater_evaporation")
            heater_output = mcc_panel_data.get("heater_output")
            heater_connected_load = mcc_panel_data.get("heater_connected_load")
            heater_temperature = mcc_panel_data.get("heater_temperature")

            mcc_sheet["C54"] = heater_model
            mcc_sheet["C55"] = heater_fuel
            mcc_sheet["C56"] = heater_year
            mcc_sheet["C57"] = (
                f"{heater_power_supply_vac}, {heater_power_supply_phase}, {heater_power_supply_frequency}"
            )
            mcc_sheet["C58"] = (
                f"{heater_control_supply_vac}, {heater_control_supply_phase}, {heater_control_supply_frequency}"
            )
            mcc_sheet["C59"] = f"{heater_evaporation} Kcl/Hr"
            mcc_sheet["C60"] = f"{heater_output} MW"
            mcc_sheet["C61"] = f"{heater_connected_load} kW"
            mcc_sheet["C62"] = f"{heater_temperature} Deg. C"

            is_spg_applicable = mcc_panel_data.get("is_spg_applicable")
            spg_name_plate_unit_name = mcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = mcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = mcc_panel_data.get(
                "spg_name_plate_manufacturing_year"
            )
            spg_name_plate_weight = mcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = mcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = mcc_panel_data.get("spg_name_plate_part_code")
            # special_note = mcc_panel_data.get("special_note")

            mcc_sheet["C74"] = spg_name_plate_unit_name
            mcc_sheet["C75"] = spg_name_plate_capacity
            mcc_sheet["C76"] = spg_name_plate_manufacturing_year
            mcc_sheet["C77"] = spg_name_plate_weight
            mcc_sheet["C78"] = spg_name_plate_oc_number
            mcc_sheet["C79"] = spg_name_plate_part_code

            # PLC fields

    ###############################################################################################################

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
