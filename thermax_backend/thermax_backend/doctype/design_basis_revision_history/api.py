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


def num_to_string(value):
    if value == 1 or value == "1":
        return "Applicable"
    return "Not Applicable"


def na_to_string(value):
    if value == "NA" or value is None:
        return "Not Applicable"
    return value


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
        "Thermax Extended User", approver, "name_initial"
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

    cover_sheet["C33"] = revision_date
    cover_sheet["D33"] = project_description
    cover_sheet["E33"] = prepped_by_initial
    cover_sheet["F33"] = checked_by_initial
    cover_sheet["G33"] = super_user_initial

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
    lv_data = f"{main_supply_lv}, {main_supply_lv_variation}%, {main_supply_lv_phase}"

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
    main_package_name = []
    safe_sub_package = []
    hazardous_sub_package = []
    for main_package in main_packages_data:
        main_package_name.append(main_package.get("main_package_name"))
        sub_packages = frappe.get_doc(
            "Project Main Package", main_package.get("name")
        ).as_dict()

        sub_package_data = sub_packages.get("sub_packages", [])

        for sub_package in sub_package_data:
            if sub_package["area_of_classification"] == "Safe Area":
                safe_sub_package.append(sub_package["sub_package_name"])
            else:
                hazardous_sub_package.append(sub_package["sub_package_name"])

    main_package_name = ", ".join(main_package_name)
    safe_sub_package = ", ".join(safe_sub_package)
    hazardous_sub_package = ", ".join(hazardous_sub_package)

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
    design_basis_sheet["C21"] = safe_sub_package
    design_basis_sheet["C22"] = hazardous_sub_package
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

    if safe_area_bearing_rtd == "NA":
        safe_area_bearing_rtd = "Not Applicable"

    if safe_area_winding_rtd == "NA":
        safe_area_winding_rtd = "Not Applicable"

    if hazardous_area_bearing_rtd == "NA":
        hazardous_area_bearing_rtd = "Not Applicable"

    if hazardous_area_winding_rtd == "NA":
        hazardous_area_winding_rtd = "Not Applicable"

    design_basis_sheet["C27"] = safe_area_efficiency_level
    design_basis_sheet["C28"] = safe_area_insulation_class
    design_basis_sheet["C29"] = safe_area_temperature_rise
    design_basis_sheet["C30"] = safe_area_enclosure_ip_rating
    design_basis_sheet["C31"] = f"{safe_area_max_temperature} Deg. C"
    design_basis_sheet["C32"] = f"{safe_area_min_temperature} Deg. C"
    design_basis_sheet["C33"] = f"{safe_area_altitude} meters"
    design_basis_sheet["C34"] = f"{safe_area_terminal_box_ip_rating}"
    design_basis_sheet["C35"] = f"{safe_area_thermister} kW & Above"
    design_basis_sheet["C36"] = f"{safe_area_space_heater} kW & Above"
    design_basis_sheet["C37"] = f"{safe_area_certification}"
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
    design_basis_sheet["D34"] = f"{hazardous_area_terminal_box_ip_rating}"
    design_basis_sheet["D35"] = f"{hazardous_area_thermister} kW & Above"
    design_basis_sheet["D36"] = f"{hazardous_area_space_heater} kW & Above"
    design_basis_sheet["D37"] = f"{hazardous_area_certification}"
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

    design_basis_sheet["C50"] = handle_make_of_component(motor)
    design_basis_sheet["C51"] = handle_make_of_component(cable)
    design_basis_sheet["C52"] = handle_make_of_component(lv_switchgear)
    design_basis_sheet["C53"] = handle_make_of_component(panel_enclosure)
    design_basis_sheet["C54"] = handle_make_of_component(
        variable_frequency_speed_drive_vfd_vsd
    )
    design_basis_sheet["C55"] = handle_make_of_component(soft_starter)
    design_basis_sheet["C56"] = handle_make_of_component(plc)

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

    cc_dol_starter = common_config_data.get("dol_starter")
    cc_star_delta_starter = common_config_data.get("star_delta_starter")
    cc_ammeter = common_config_data.get("ammeter")
    cc_ammeter_configuration = common_config_data.get("ammeter_configuration")
    cc_mcc_switchgear_type = common_config_data.get("mcc_switchgear_type")
    cc_switchgear_combination = common_config_data.get("switchgear_combination")
    cc_is_control_transformer_applicable = common_config_data.get(
        "is_control_transformer_applicable"
    )
    cc_control_transformer_primary_voltage = common_config_data.get(
        "control_transformer_primary_voltage"
    )
    cc_control_transformer_secondary_voltage_copy = common_config_data.get(
        "control_transformer_secondary_voltage_copy"
    )
    cc_control_transformer_coating = common_config_data.get(
        "control_transformer_coating"
    )
    cc_control_transformer_quantity = common_config_data.get(
        "control_transformer_quantity"
    )
    cc_control_transformer_configuration = common_config_data.get(
        "control_transformer_configuration"
    )
    cc_digital_meters = common_config_data.get("digital_meters")
    cc_analog_meters = common_config_data.get("analog_meters")
    cc_communication_protocol = common_config_data.get("communication_protocol")
    cc_pole = common_config_data.get("pole")
    cc_supply_feeder_standard = common_config_data.get("supply_feeder_standard")
    cc_dm_standard = common_config_data.get("dm_standard")
    cc_power_wiring_color = common_config_data.get("power_wiring_color")
    cc_power_wiring_size = common_config_data.get("power_wiring_size")
    cc_control_wiring_color = common_config_data.get("control_wiring_color")
    cc_control_wiring_size = common_config_data.get("control_wiring_size")
    cc_vdc_24_wiring_color = common_config_data.get("vdc_24_wiring_color")
    cc_vdc_24_wiring_size = common_config_data.get("vdc_24_wiring_size")
    cc_analog_signal_wiring_color = common_config_data.get("analog_signal_wiring_color")
    cc_analog_signal_wiring_size = common_config_data.get("analog_signal_wiring_size")
    cc_ct_wiring_color = common_config_data.get("ct_wiring_color")
    cc_ct_wiring_size = common_config_data.get("ct_wiring_size")
    cc_rtd_thermocouple_wiring_color = common_config_data.get(
        "rtd_thermocouple_wiring_color"
    )
    cc_rtd_thermocouple_wiring_size = common_config_data.get(
        "rtd_thermocouple_wiring_size"
    )
    cc_cable_insulation_pvc = common_config_data.get("cable_insulation_pvc")
    cc_air_clearance_between_phase_to_phase_bus = common_config_data.get(
        "air_clearance_between_phase_to_phase_bus"
    )
    cc_air_clearance_between_phase_to_neutral_bus = common_config_data.get(
        "air_clearance_between_phase_to_neutral_bus"
    )
    cc_ferrule = common_config_data.get("ferrule")
    cc_ferrule_note = common_config_data.get("ferrule_note")
    cc_device_identification_of_components = common_config_data.get(
        "device_identification_of_components"
    )
    cc_general_note_internal_wiring = common_config_data.get(
        "general_note_internal_wiring"
    )
    cc_common_requirement = common_config_data.get("common_requirement")
    cc_power_terminal_clipon = common_config_data.get("power_terminal_clipon")
    cc_power_terminal_busbar_type = common_config_data.get("power_terminal_busbar_type")
    cc_control_terminal = common_config_data.get("control_terminal")
    cc_spare_terminal = common_config_data.get("spare_terminal")
    cc_forward_push_button_start = common_config_data.get("forward_push_button_start")
    cc_reverse_push_button_start = common_config_data.get("reverse_push_button_start")
    cc_push_button_start = common_config_data.get("push_button_start")
    cc_push_button_stop = common_config_data.get("push_button_stop")
    cc_push_button_ess = common_config_data.get("push_button_ess")
    cc_potentiometer = common_config_data.get("potentiometer")
    cc_is_push_button_speed_selected = common_config_data.get(
        "is_push_button_speed_selected"
    )
    cc_speed_increase_pb = common_config_data.get("speed_increase_pb")
    cc_speed_decrease_pb = common_config_data.get("speed_decrease_pb")
    cc_alarm_acknowledge_and_lamp_test = common_config_data.get(
        "alarm_acknowledge_and_lamp_test"
    )
    cc_lamp_test_push_button = common_config_data.get("lamp_test_push_button")
    cc_test_dropdown = common_config_data.get("test_dropdown")
    cc_reset_dropdown = common_config_data.get("reset_dropdown")
    cc_is_field_motor_isolator_selected = common_config_data.get(
        "is_field_motor_isolator_selected"
    )
    cc_is_safe_area_isolator_selected = common_config_data.get(
        "is_safe_area_isolator_selected"
    )
    cc_is_local_push_button_station_selected = common_config_data.get(
        "is_local_push_button_station_selected"
    )
    cc_selector_switch_applicable = common_config_data.get("selector_switch_applicable")
    cc_selector_switch_lockable = common_config_data.get("selector_switch_lockable")
    cc_running_open = common_config_data.get("running_open")
    cc_stopped_closed = common_config_data.get("stopped_closed")
    cc_trip = common_config_data.get("trip")
    cc_safe_field_motor_type = common_config_data.get("safe_field_motor_type")
    cc_hazardous_field_motor_type = common_config_data.get("hazardous_field_motor_type")
    cc_safe_field_motor_enclosure = common_config_data.get("safe_field_motor_enclosure")
    cc_hazardous_field_motor_enclosure = common_config_data.get(
        "hazardous_field_motor_enclosure"
    )
    cc_safe_field_motor_material = common_config_data.get("safe_field_motor_material")
    cc_hazardous_field_motor_material = common_config_data.get(
        "hazardous_field_motor_material"
    )
    cc_safe_field_motor_thickness = common_config_data.get("safe_field_motor_thickness")
    cc_hazardous_field_motor_thickness = common_config_data.get(
        "hazardous_field_motor_thickness"
    )
    cc_safe_field_motor_qty = common_config_data.get("safe_field_motor_qty")
    cc_hazardous_field_motor_qty = common_config_data.get("hazardous_field_motor_qty")
    cc_safe_field_motor_isolator_color_shade = common_config_data.get(
        "safe_field_motor_isolator_color_shade"
    )
    cc_hazardous_field_motor_isolator_color_shade = common_config_data.get(
        "hazardous_field_motor_isolator_color_shade"
    )
    cc_safe_field_motor_cable_entry = common_config_data.get(
        "safe_field_motor_cable_entry"
    )
    cc_hazardous_field_motor_cable_entry = common_config_data.get(
        "hazardous_field_motor_cable_entry"
    )
    cc_safe_field_motor_canopy = common_config_data.get("safe_field_motor_canopy")
    cc_hazardous_field_motor_canopy = common_config_data.get(
        "hazardous_field_motor_canopy"
    )
    cc_safe_field_motor_canopy_type = common_config_data.get(
        "safe_field_motor_canopy_type"
    )
    cc_hazardous_field_motor_canopy_type = common_config_data.get(
        "hazardous_field_motor_canopy_type"
    )
    cc_safe_lpbs_type = common_config_data.get("safe_lpbs_type")
    cc_hazardous_lpbs_type = common_config_data.get("hazardous_lpbs_type")
    cc_safe_lpbs_enclosure = common_config_data.get("safe_lpbs_enclosure")
    cc_hazardous_lpbs_enclosure = common_config_data.get("hazardous_lpbs_enclosure")
    cc_safe_lpbs_thickness = common_config_data.get("safe_lpbs_thickness")
    cc_hazardous_lpbs_thickness = common_config_data.get("hazardous_lpbs_thickness")
    cc_safe_lpbs_material = common_config_data.get("safe_lpbs_material")
    cc_hazardous_lpbs_material = common_config_data.get("hazardous_lpbs_material")
    cc_safe_lpbs_qty = common_config_data.get("safe_lpbs_qty")
    cc_hazardous_lpbs_qty = common_config_data.get("hazardous_lpbs_qty")
    cc_safe_lpbs_color_shade = common_config_data.get("safe_lpbs_color_shade")
    cc_hazardous_lpbs_color_shade = common_config_data.get("hazardous_lpbs_color_shade")
    cc_safe_lpbs_canopy = common_config_data.get("safe_lpbs_canopy")
    cc_hazardous_lpbs_canopy = common_config_data.get("hazardous_lpbs_canopy")
    cc_safe_lpbs_canopy_type = common_config_data.get("safe_lpbs_canopy_type")
    cc_hazardous_lpbs_canopy_type = common_config_data.get("hazardous_lpbs_canopy_type")
    cc_lpbs_push_button_start_color = common_config_data.get(
        "lpbs_push_button_start_color"
    )
    cc_lpbs_indication_lamp_start_color = common_config_data.get(
        "lpbs_indication_lamp_start_color"
    )
    cc_lpbs_indication_lamp_stop_color = common_config_data.get(
        "lpbs_indication_lamp_stop_color"
    )
    cc_lpbs_speed_increase = common_config_data.get("lpbs_speed_increase")
    cc_lpbs_speed_decrease = common_config_data.get("lpbs_speed_decrease")
    cc_apfc_relay = common_config_data.get("apfc_relay")
    cc_power_bus_main_busbar_selection = common_config_data.get(
        "power_bus_main_busbar_selection"
    )
    cc_power_bus_heat_pvc_sleeve = common_config_data.get("power_bus_heat_pvc_sleeve")
    cc_power_bus_material = common_config_data.get("power_bus_material")
    cc_power_bus_current_density = common_config_data.get("power_bus_current_density")
    cc_power_bus_rating_of_busbar = common_config_data.get("power_bus_rating_of_busbar")
    cc_control_bus_main_busbar_selection = common_config_data.get(
        "control_bus_main_busbar_selection"
    )
    cc_control_bus_heat_pvc_sleeve = common_config_data.get(
        "control_bus_heat_pvc_sleeve"
    )
    cc_control_bus_material = common_config_data.get("control_bus_material")
    cc_control_bus_current_density = common_config_data.get(
        "control_bus_current_density"
    )
    cc_control_bus_rating_of_busbar = common_config_data.get(
        "control_bus_rating_of_busbar"
    )
    cc_earth_bus_main_busbar_selection = common_config_data.get(
        "earth_bus_main_busbar_selection"
    )
    cc_earth_bus_busbar_position = common_config_data.get("earth_bus_busbar_position")
    cc_earth_bus_material = common_config_data.get("earth_bus_material")
    cc_earth_bus_current_density = common_config_data.get("earth_bus_current_density")
    cc_earth_bus_rating_of_busbar = common_config_data.get("earth_bus_rating_of_busbar")
    cc_door_earthing = common_config_data.get("door_earthing")
    cc_instrument_earth = common_config_data.get("instrument_earth")
    cc_general_note_busbar_and_insulation_materials = common_config_data.get(
        "general_note_busbar_and_insulation_materials"
    )
    cc_metering_for_feeders = common_config_data.get("metering_for_feeders")
    cc_cooling_fans = common_config_data.get("cooling_fans")
    cc_louvers_and_filters = common_config_data.get("louvers_and_filters")
    cc_alarm_annunciator = common_config_data.get("alarm_annunciator")
    cc_control_transformer = common_config_data.get("control_transformer")
    cc_commissioning_spare = common_config_data.get("commissioning_spare")
    cc_two_year_operational_spare = common_config_data.get("two_year_operational_spare")
    cc_current_transformer = common_config_data.get("current_transformer")
    cc_current_transformer_coating = common_config_data.get(
        "current_transformer_coating"
    )
    cc_current_transformer_quantity = common_config_data.get(
        "current_transformer_quantity"
    )
    cc_current_transformer_configuration = common_config_data.get(
        "current_transformer_configuration"
    )
    cc_control_transformer_type = common_config_data.get("control_transformer_type")

    design_basis_sheet["C59"] = na_to_string(cc_dm_standard)
    design_basis_sheet["C61"] = na_to_string(cc_control_transformer_primary_voltage)
    design_basis_sheet["C62"] = na_to_string(
        cc_control_transformer_secondary_voltage_copy
    )
    design_basis_sheet["C63"] = na_to_string(cc_control_transformer_coating)
    design_basis_sheet["C64"] = na_to_string(cc_control_transformer_quantity)
    design_basis_sheet["C65"] = na_to_string(cc_control_transformer_configuration)
    design_basis_sheet["C66"] = na_to_string(cc_control_transformer_type)

    design_basis_sheet["C68"] = cc_apfc_relay

    design_basis_sheet["C70"] = cc_power_bus_main_busbar_selection
    design_basis_sheet["C71"] = cc_power_bus_heat_pvc_sleeve
    design_basis_sheet["C72"] = cc_power_bus_material
    design_basis_sheet["C73"] = cc_power_bus_current_density
    design_basis_sheet["C74"] = cc_power_bus_rating_of_busbar

    design_basis_sheet["C76"] = cc_control_bus_main_busbar_selection
    design_basis_sheet["C77"] = cc_control_bus_heat_pvc_sleeve
    design_basis_sheet["C78"] = cc_control_bus_material
    design_basis_sheet["C79"] = cc_control_bus_current_density
    design_basis_sheet["C80"] = cc_control_bus_rating_of_busbar

    design_basis_sheet["C82"] = cc_earth_bus_main_busbar_selection
    design_basis_sheet["C83"] = cc_earth_bus_busbar_position
    design_basis_sheet["C84"] = cc_earth_bus_material
    design_basis_sheet["C85"] = cc_earth_bus_current_density
    design_basis_sheet["C86"] = cc_earth_bus_rating_of_busbar
    design_basis_sheet["C87"] = cc_door_earthing
    design_basis_sheet["C88"] = cc_instrument_earth
    design_basis_sheet["C89"] = cc_general_note_busbar_and_insulation_materials

    design_basis_sheet["C91"] = f"{cc_dol_starter} kW & Above"
    design_basis_sheet["C92"] = f"{cc_star_delta_starter} kW & Above"
    design_basis_sheet["C93"] = cc_mcc_switchgear_type
    design_basis_sheet["C94"] = cc_switchgear_combination

    design_basis_sheet["C96"] = f"{cc_ammeter} kW & Above"
    design_basis_sheet["C97"] = cc_ammeter_configuration
    design_basis_sheet["C98"] = (
        cc_analog_meters.replace("[", "")
        .replace("]", "")
        .replace(",", ", ")
        .replace('"', "")
    )
    design_basis_sheet["C99"] = (
        cc_digital_meters.replace("[", "")
        .replace("]", "")
        .replace(",", ", ")
        .replace('"', "")
    )
    design_basis_sheet["C100"] = cc_communication_protocol

    design_basis_sheet["C102"] = f"{cc_current_transformer} kW & Above"
    design_basis_sheet["C103"] = cc_current_transformer_coating
    design_basis_sheet["C104"] = cc_current_transformer_quantity
    design_basis_sheet["C105"] = cc_current_transformer_configuration

    design_basis_sheet["C107"] = cc_pole

    design_basis_sheet["C109"] = cc_power_wiring_color
    design_basis_sheet["C110"] = cc_power_wiring_size
    design_basis_sheet["C111"] = cc_control_wiring_color
    design_basis_sheet["C112"] = cc_control_wiring_size
    design_basis_sheet["C113"] = cc_vdc_24_wiring_color
    design_basis_sheet["C114"] = cc_vdc_24_wiring_size
    design_basis_sheet["C115"] = cc_analog_signal_wiring_color
    design_basis_sheet["C116"] = cc_analog_signal_wiring_size
    design_basis_sheet["C117"] = cc_ct_wiring_color
    design_basis_sheet["C118"] = cc_ct_wiring_size
    design_basis_sheet["C119"] = cc_rtd_thermocouple_wiring_color
    design_basis_sheet["C120"] = cc_rtd_thermocouple_wiring_size
    design_basis_sheet["C121"] = cc_air_clearance_between_phase_to_phase_bus
    design_basis_sheet["C122"] = cc_air_clearance_between_phase_to_neutral_bus
    design_basis_sheet["C123"] = cc_cable_insulation_pvc
    design_basis_sheet["C124"] = cc_device_identification_of_components
    design_basis_sheet["C125"] = cc_general_note_internal_wiring

    design_basis_sheet["C127"] = cc_power_terminal_clipon
    design_basis_sheet["C128"] = cc_power_terminal_busbar_type
    design_basis_sheet["C129"] = cc_control_terminal
    design_basis_sheet["C130"] = cc_spare_terminal

    design_basis_sheet["C132"] = cc_push_button_start
    design_basis_sheet["C133"] = cc_push_button_stop
    design_basis_sheet["C134"] = cc_push_button_ess
    design_basis_sheet["C135"] = cc_forward_push_button_start
    design_basis_sheet["C136"] = cc_reverse_push_button_start
    design_basis_sheet["C137"] = cc_potentiometer
    design_basis_sheet["C138"] = cc_speed_increase_pb
    design_basis_sheet["C139"] = cc_speed_decrease_pb
    design_basis_sheet["C140"] = cc_alarm_acknowledge_and_lamp_test
    design_basis_sheet["C141"] = cc_test_dropdown
    design_basis_sheet["C142"] = cc_reset_dropdown
    design_basis_sheet["C143"] = cc_lamp_test_push_button

    design_basis_sheet["C145"] = cc_selector_switch_applicable

    design_basis_sheet["C147"] = cc_running_open
    design_basis_sheet["C148"] = cc_stopped_closed
    design_basis_sheet["C149"] = cc_trip

    design_basis_sheet["C152"] = cc_safe_field_motor_type
    design_basis_sheet["C153"] = cc_safe_field_motor_enclosure
    design_basis_sheet["C154"] = cc_safe_field_motor_material
    design_basis_sheet["C155"] = cc_safe_field_motor_qty
    design_basis_sheet["C156"] = cc_safe_field_motor_isolator_color_shade
    design_basis_sheet["C157"] = cc_safe_field_motor_cable_entry
    design_basis_sheet["C158"] = cc_safe_field_motor_canopy

    design_basis_sheet["D152"] = cc_hazardous_field_motor_type
    design_basis_sheet["D153"] = cc_hazardous_field_motor_enclosure
    design_basis_sheet["D154"] = cc_hazardous_field_motor_material
    design_basis_sheet["D155"] = cc_hazardous_field_motor_qty
    design_basis_sheet["D156"] = cc_hazardous_field_motor_isolator_color_shade
    design_basis_sheet["D157"] = cc_hazardous_field_motor_cable_entry
    design_basis_sheet["D158"] = cc_hazardous_field_motor_canopy

    design_basis_sheet["C160"] = cc_lpbs_push_button_start_color
    design_basis_sheet["C161"] = cc_forward_push_button_start
    design_basis_sheet["C162"] = cc_reverse_push_button_start
    design_basis_sheet["C163"] = cc_push_button_ess
    design_basis_sheet["C164"] = cc_lpbs_speed_increase
    design_basis_sheet["C165"] = cc_lpbs_speed_decrease
    design_basis_sheet["C166"] = cc_lpbs_indication_lamp_start_color
    design_basis_sheet["C167"] = cc_lpbs_indication_lamp_stop_color

    design_basis_sheet["C169"] = cc_safe_lpbs_type
    design_basis_sheet["C170"] = cc_safe_lpbs_enclosure
    design_basis_sheet["C171"] = cc_safe_lpbs_material
    design_basis_sheet["C172"] = cc_safe_lpbs_qty
    design_basis_sheet["C173"] = cc_safe_lpbs_color_shade
    design_basis_sheet["C174"] = cc_safe_lpbs_canopy
    design_basis_sheet["C175"] = cc_safe_lpbs_canopy_type

    design_basis_sheet["D169"] = cc_hazardous_lpbs_type
    design_basis_sheet["D170"] = cc_hazardous_lpbs_enclosure
    design_basis_sheet["D171"] = cc_hazardous_lpbs_material
    design_basis_sheet["D172"] = cc_hazardous_lpbs_qty
    design_basis_sheet["D173"] = cc_hazardous_lpbs_color_shade
    design_basis_sheet["D174"] = cc_hazardous_lpbs_canopy
    design_basis_sheet["D175"] = cc_hazardous_lpbs_canopy_type

    design_basis_sheet["C177"] = cc_ferrule
    design_basis_sheet["C178"] = cc_ferrule_note
    design_basis_sheet["C179"] = cc_device_identification_of_components

    design_basis_sheet["C181"] = cc_cooling_fans
    design_basis_sheet["C182"] = cc_louvers_and_filters

    cable_tray_data = frappe.db.get_list(
        "Cable Tray Layout", {"revision_id": revision_id}, "*"
    )
    cable_tray_data = cable_tray_data[0]

    ct_number_of_cores = cable_tray_data.get("number_of_cores")
    ct_specific_requirement = cable_tray_data.get("specific_requirement")
    ct_type_of_insulation = cable_tray_data.get("type_of_insulation")
    ct_color_scheme = cable_tray_data.get("color_scheme")
    ct_motor_voltage_drop_during_running = cable_tray_data.get(
        "motor_voltage_drop_during_running"
    )
    ct_copper_conductor = cable_tray_data.get("copper_conductor")
    ct_aluminium_conductor = cable_tray_data.get("aluminium_conductor")
    ct_cable_installation = cable_tray_data.get("cable_installation")
    ct_motor_voltage_drop_during_starting = cable_tray_data.get(
        "motor_voltage_drop_during_starting"
    )
    ct_voltage_grade = cable_tray_data.get("voltage_grade")
    ct_touching_factor_air = cable_tray_data.get("touching_factor_air")
    ct_touching_factor_burid = cable_tray_data.get("touching_factor_burid")
    ct_ambient_temp_factor_air = cable_tray_data.get("ambient_temp_factor_air")
    ct_ambient_temp_factor_burid = cable_tray_data.get("ambient_temp_factor_burid")
    ct_derating_factor_air = cable_tray_data.get("derating_factor_air")
    ct_derating_factor_burid = cable_tray_data.get("derating_factor_burid")
    ct_gland_make = cable_tray_data.get("gland_make")
    ct_moc = cable_tray_data.get("moc")
    ct_type_of_gland = cable_tray_data.get("type_of_gland")
    ct_future_space_on_trays = cable_tray_data.get("future_space_on_trays")
    ct_cable_placement = cable_tray_data.get("cable_placement")
    ct_orientation = cable_tray_data.get("orientation")
    ct_vertical_distance = cable_tray_data.get("vertical_distance")
    ct_horizontal_distance = cable_tray_data.get("horizontal_distance")
    ct_is_dry_area_selected = cable_tray_data.get("is_dry_area_selected")
    ct_cable_tray_moc = cable_tray_data.get("cable_tray_moc")
    ct_is_wet_area_selected = cable_tray_data.get("is_wet_area_selected")
    ct_cable_tray_moc_input = cable_tray_data.get("cable_tray_moc_input")
    is_pct_perforated_type_selected = cable_tray_data.get(
        "is_pct_perforated_type_selected"
    )
    pct_perforated_type_width = cable_tray_data.get("pct_perforated_type_width")
    pct_perforated_type_max_width = cable_tray_data.get("pct_perforated_type_max_width")
    pct_perforated_type_height = cable_tray_data.get("pct_perforated_type_height")
    pct_perforated_type_thickness = cable_tray_data.get("pct_perforated_type_thickness")
    is_pct_ladder_type_selected = cable_tray_data.get("is_pct_ladder_type_selected")
    pct_ladder_type_width = cable_tray_data.get("pct_ladder_type_width")
    pct_ladder_type_max_width = cable_tray_data.get("pct_ladder_type_max_width")
    pct_ladder_type_height = cable_tray_data.get("pct_ladder_type_height")
    pct_ladder_type_thickness = cable_tray_data.get("pct_ladder_type_thickness")
    is_pct_mesh_type_selected = cable_tray_data.get("is_pct_mesh_type_selected")
    pct_mesh_type_width = cable_tray_data.get("pct_mesh_type_width")
    pct_mesh_type_max_length = cable_tray_data.get("pct_mesh_type_max_length")
    pct_mesh_type_height = cable_tray_data.get("pct_mesh_type_height")
    pct_mesh_type_thickness = cable_tray_data.get("pct_mesh_type_thickness")
    is_pct_conduit_selected = cable_tray_data.get("is_pct_conduit_selected")
    pct_conduit_moc = cable_tray_data.get("pct_conduit_moc")
    pct_conduit_size = cable_tray_data.get("pct_conduit_size")
    is_cct_perforated_type_selected = cable_tray_data.get(
        "is_cct_perforated_type_selected"
    )
    cct_perforated_type_width = cable_tray_data.get("cct_perforated_type_width")
    cct_perforated_type_max_width = cable_tray_data.get("cct_perforated_type_max_width")
    cct_perforated_type_height = cable_tray_data.get("cct_perforated_type_height")
    cct_perforated_type_thickness = cable_tray_data.get("cct_perforated_type_thickness")
    is_cct_ladder_type_selected = cable_tray_data.get("is_cct_ladder_type_selected")
    cct_ladder_type_width = cable_tray_data.get("cct_ladder_type_width")
    cct_ladder_type_max_width = cable_tray_data.get("cct_ladder_type_max_width")
    cct_ladder_type_height = cable_tray_data.get("cct_ladder_type_height")
    cct_ladder_type_thickness = cable_tray_data.get("cct_ladder_type_thickness")
    is_cct_mesh_type_selected = cable_tray_data.get("is_cct_mesh_type_selected")
    cct_mesh_type_width = cable_tray_data.get("cct_mesh_type_width")
    cct_mesh_type_max_width = cable_tray_data.get("cct_mesh_type_max_width")
    cct_mesh_type_height = cable_tray_data.get("cct_mesh_type_height")
    cct_mesh_type_thickness = cable_tray_data.get("cct_mesh_type_thickness")
    is_cct_conduit_selected = cable_tray_data.get("is_cct_conduit_selected")
    cct_conduit_moc = cable_tray_data.get("cct_conduit_moc")
    cct_conduit_size = cable_tray_data.get("cct_conduit_size")
    is_sct_perforated_type_selected = cable_tray_data.get(
        "is_sct_perforated_type_selected"
    )
    sct_perforated_type_width = cable_tray_data.get("sct_perforated_type_width")
    sct_perforated_type_max_width = cable_tray_data.get("sct_perforated_type_max_width")
    sct_perforated_type_height = cable_tray_data.get("sct_perforated_type_height")
    sct_perforated_type_thickness = cable_tray_data.get("sct_perforated_type_thickness")
    is_sct_ladder_type_selected = cable_tray_data.get("is_sct_ladder_type_selected")
    sct_ladder_type_width = cable_tray_data.get("sct_ladder_type_width")
    sct_ladder_type_max_width = cable_tray_data.get("sct_ladder_type_max_width")
    sct_ladder_type_height = cable_tray_data.get("sct_ladder_type_height")
    sct_ladder_type_thickness = cable_tray_data.get("sct_ladder_type_thickness")
    is_sct_mesh_type_selected = cable_tray_data.get("is_sct_mesh_type_selected")
    sct_mesh_type_width = cable_tray_data.get("sct_mesh_type_width")
    sct_mesh_type_max_width = cable_tray_data.get("sct_mesh_type_max_width")
    sct_mesh_type_height = cable_tray_data.get("sct_mesh_type_height")
    sct_mesh_type_thickness = cable_tray_data.get("sct_mesh_type_thickness")
    is_sct_conduit_selected = cable_tray_data.get("is_sct_conduit_selected")
    sct_conduit_moc = cable_tray_data.get("sct_conduit_moc")
    sct_conduit_size = cable_tray_data.get("sct_conduit_size")
    cable_tray_cover = cable_tray_data.get("cable_tray_cover")

    design_basis_sheet["C184"] = ct_number_of_cores
    design_basis_sheet["C185"] = ct_specific_requirement
    design_basis_sheet["C186"] = ct_type_of_insulation
    design_basis_sheet["C187"] = ct_color_scheme
    design_basis_sheet["C188"] = f"{ct_motor_voltage_drop_during_starting} %"
    design_basis_sheet["C189"] = f"{ct_motor_voltage_drop_during_running} %"
    design_basis_sheet["C190"] = ct_voltage_grade
    design_basis_sheet["C191"] = f"{ct_copper_conductor} Sq. mm"
    design_basis_sheet["C192"] = f"{ct_aluminium_conductor} Sq. mm"
    design_basis_sheet["C193"] = ct_touching_factor_air
    design_basis_sheet["C194"] = ct_ambient_temp_factor_air
    design_basis_sheet["C195"] = ct_derating_factor_air
    design_basis_sheet["C196"] = ct_touching_factor_burid
    design_basis_sheet["C197"] = ct_ambient_temp_factor_burid
    design_basis_sheet["C198"] = ct_derating_factor_burid

    design_basis_sheet["C200"] = ct_gland_make
    design_basis_sheet["C201"] = ct_moc
    design_basis_sheet["C202"] = ct_type_of_gland

    design_basis_sheet["C206"] = cable_tray_cover
    design_basis_sheet["C207"] = f"{ct_future_space_on_trays} %"
    design_basis_sheet["C208"] = ct_cable_placement
    design_basis_sheet["C209"] = f"{ct_vertical_distance} mm"
    design_basis_sheet["C210"] = f"{ct_horizontal_distance} mm"
    design_basis_sheet["C211"] = ct_cable_tray_moc

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
        panel_id = project_panel.get("name")
        if project_panel.get("panel_main_type") == "MCC":

            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"revision_id": revision_id, "panel_id": panel_id}, "*"
            )
            mcc_panel_data = mcc_panel_data[0]

            panel_sheet = template_workbook.copy_worksheet(mcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")

            panel_sheet["B3"] = project_panel.get("panel_name")

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
            led_type_on_input = mcc_panel_data.get("led_type_on_input")
            is_indication_off_selected = mcc_panel_data.get(
                "is_indication_off_selected"
            )
            led_type_off_input = mcc_panel_data.get("led_type_off_input")
            is_indication_trip_selected = mcc_panel_data.get(
                "is_indication_trip_selected"
            )
            led_type_trip_input = mcc_panel_data.get("led_type_trip_input")
            is_blue_cb_spring_charge_selected = mcc_panel_data.get(
                "is_blue_cb_spring_charge_selected"
            )
            is_red_cb_in_service = mcc_panel_data.get("is_red_cb_in_service")
            is_white_healthy_trip_circuit_selected = mcc_panel_data.get(
                "is_white_healthy_trip_circuit_selected"
            )
            is_other_selected = mcc_panel_data.get("is_other_selected")
            led_type_other_input = mcc_panel_data.get("led_type_other_input")
            current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
            )
            control_transformer_coating = mcc_panel_data.get(
                "control_transformer_coating"
            )
            control_transformer_configuration = mcc_panel_data.get(
                "control_transformer_configuration"
            )
            current_transformer_number = mcc_panel_data.get(
                "current_transformer_number"
            )
            current_transformer_configuration = mcc_panel_data.get(
                "current_transformer_configuration"
            )
            alarm_annunciator = mcc_panel_data.get("alarm_annunciator")
            mi_analog = mcc_panel_data.get("mi_analog") or "NA"
            mi_digital = mcc_panel_data.get("mi_digital") or "NA"
            mi_communication_protocol = (
                mcc_panel_data.get("mi_communication_protocol") or "NA"
            )
            ga_moc_material = mcc_panel_data.get("ga_moc_material")
            door_thickness = mcc_panel_data.get("door_thickness")
            ga_moc_thickness_door = mcc_panel_data.get("ga_moc_thickness_door")
            ga_moc_thickness_covers = mcc_panel_data.get("ga_moc_thickness_covers")
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
            ga_panel_mounting_frame = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            is_marshalling_section_selected = mcc_panel_data.get(
                "is_marshalling_section_selected"
            )
            marshalling_section_text_area = mcc_panel_data.get(
                "marshalling_section_text_area"
            )
            is_cable_alley_section_selected = mcc_panel_data.get(
                "is_cable_alley_section_selected"
            )
            is_power_and_bus_separation_section_selected = mcc_panel_data.get(
                "is_power_and_bus_separation_section_selected"
            )
            is_both_side_extension_section_selected = mcc_panel_data.get(
                "is_both_side_extension_section_selected"
            )
            ga_gland_plate_3mm_drill_type = mcc_panel_data.get(
                "ga_gland_plate_3mm_drill_type"
            )
            ga_gland_plate_thickness = mcc_panel_data.get("ga_gland_plate_thickness")
            ga_busbar_chamber_position = mcc_panel_data.get(
                "ga_busbar_chamber_position"
            )
            ga_power_and_control_busbar_separation = mcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            ga_enclosure_protection_degree = mcc_panel_data.get(
                "ga_enclosure_protection_degree"
            )
            ga_cable_entry_position = mcc_panel_data.get("ga_cable_entry_position")
            general_requirments_for_construction = mcc_panel_data.get(
                "general_requirments_for_construction"
            )
            ppc_painting_standards = mcc_panel_data.get("ppc_painting_standards")
            ppc_interior_and_exterior_paint_shade = mcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = mcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            ppc_base_frame_paint_shade = mcc_panel_data.get(
                "ppc_base_frame_paint_shade"
            )
            ppc_minimum_coating_thickness = mcc_panel_data.get(
                "ppc_minimum_coating_thickness"
            )
            ppc_pretreatment_panel_standard = mcc_panel_data.get(
                "ppc_pretreatment_panel_standard"
            )
            vfd_auto_manual_selection = mcc_panel_data.get("vfd_auto_manual_selection")
            two_year_operational_spare = mcc_panel_data.get(
                "two_year_operational_spare"
            )
            commissioning_spare = mcc_panel_data.get("commissioning_spare")
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
            is_spg_applicable = mcc_panel_data.get("is_spg_applicable")
            spg_name_plate_unit_name = mcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = mcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = mcc_panel_data.get(
                "spg_name_plate_manufacturing_year"
            )
            spg_name_plate_weight = mcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = mcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = mcc_panel_data.get("spg_name_plate_part_code")
            special_note = mcc_panel_data.get("special_note")

            incomer_data = f"Upto {incomer_ampere}, {incomer_pole} Pole {incomer_type} \nAbove {incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type} "

            if is_indication_on_selected == "0" or is_indication_on_selected == 0:
                led_type_on_input = "Not Applicable"

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            panel_sheet["C5"] = na_to_string(incomer_data)
            panel_sheet["C6"] = led_type_on_input
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input
            panel_sheet["C9"] = na_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = na_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = na_to_string(is_white_healthy_trip_circuit_selected)
            panel_sheet["C12"] = na_to_string(alarm_annunciator)

            analog_data = (
                mi_analog.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
                if mi_analog is not None
                else "Not Applicable"
            )
            digital_data = (
                mi_digital.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
                if mi_digital is not None
                else "Not Applicable"
            )

            if "NA" in mi_analog:
                analog_data = "Not Applicable"

            if "NA" in mi_digital:
                digital_data = "Not Applicable"

            if "NA" in mi_communication_protocol:
                mi_communication_protocol = "Not Applicable"

            panel_sheet["C14"] = analog_data
            panel_sheet["C15"] = digital_data
            panel_sheet["C16"] = mi_communication_protocol

            panel_sheet["C18"] = current_transformer_coating
            panel_sheet["C19"] = current_transformer_number
            panel_sheet["C20"] = current_transformer_configuration

            panel_sheet["C22"] = ga_moc_material  # MOC
            panel_sheet["C23"] = (
                ga_moc_thickness_door  # Component Mounting Plate Thickness
            )
            panel_sheet["C24"] = door_thickness  # Door Thickness
            panel_sheet["C25"] = ga_moc_thickness_covers  # Top & Side Thickness
            panel_sheet["C26"] = ga_gland_plate_thickness  # Gland Plate Thickness
            panel_sheet["C27"] = ga_gland_plate_3mm_drill_type  # Gland Plate Type
            panel_sheet["C28"] = ga_mcc_compartmental  # Panel Front Type
            panel_sheet["C29"] = (
                ga_mcc_construction_front_type  # Type of Construction for Board
            )
            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_mcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            if (
                is_marshalling_section_selected == 0
                or is_marshalling_section_selected == "0"
            ):
                marshalling_section_text_area = "Not Applicable"

            panel_sheet["C35"] = marshalling_section_text_area  # Marshalling Section
            panel_sheet["C36"] = num_to_string(is_cable_alley_section_selected)
            panel_sheet["C37"] = num_to_string(
                is_power_and_bus_separation_section_selected
            )  # BUS
            panel_sheet["C38"] = num_to_string(
                is_both_side_extension_section_selected
            )  # Extension on Both sides
            panel_sheet["C39"] = ga_busbar_chamber_position  # Busbar Chamber position
            panel_sheet["C40"] = ga_power_and_control_busbar_separation  # BUSBAR
            panel_sheet["C41"] = ga_enclosure_protection_degree  # Degree of Enclosure
            panel_sheet["C42"] = ga_cable_entry_position  # BUSBAR

            panel_sheet["C44"] = "As per OEM Stanadard"
            panel_sheet["C45"] = ppc_interior_and_exterior_paint_shade
            panel_sheet["C46"] = ppc_component_mounting_plate_paint_shade

            panel_sheet["C47"] = ppc_minimum_coating_thickness
            panel_sheet["C48"] = "Black"
            panel_sheet["C49"] = ppc_pretreatment_panel_standard
            panel_sheet["C50"] = general_requirments_for_construction

            panel_sheet["C52"] = vfd_auto_manual_selection
            panel_sheet["C54"] = commissioning_spare
            panel_sheet["C55"] = two_year_operational_spare

            mcc_boiler_power_supply = f"{boiler_power_supply_vac}, {boiler_power_supply_phase}, {boiler_power_supply_frequency}"
            mcc_boiler_control_supply = f"{boiler_control_supply_vac}, {boiler_control_supply_phase}, {boiler_control_supply_frequency}"

            if boiler_evaporation == "NA":
                boiler_evaporation = "Not Applicable"
            else:
                boiler_evaporation = f"{boiler_evaporation} kg/Hr"

            if boiler_output == "NA":
                boiler_output = "Not Applicable"
            else:
                boiler_output = f"{boiler_output} MW"

            if boiler_connected_load == "NA":
                boiler_connected_load = "Not Applicable"
            else:
                boiler_connected_load = f"{boiler_connected_load} kW"

            if boiler_design_pressure == "NA":
                boiler_design_pressure = "Not Applicable"
            else:
                boiler_design_pressure = f"{boiler_design_pressure} kg/cm2(g)/Bar"

            if (
                is_punching_details_for_boiler_selected == "0"
                or is_punching_details_for_boiler_selected == 0
            ):
                boiler_connected_load = "Not Applicable"
                boiler_model = "Not Applicable"
                boiler_fuel = "Not Applicable"
                boiler_year = "Not Applicable"
                boiler_evaporation = "Not Applicable"
                boiler_output = "Not Applicable"
                boiler_connected_load = "Not Applicable"
                boiler_design_pressure = "Not Applicable"
                mcc_boiler_power_supply = "Not Applicable"
                mcc_boiler_control_supply = "Not Applicable"

            panel_sheet["C58"] = boiler_model
            panel_sheet["C59"] = boiler_fuel
            panel_sheet["C60"] = boiler_year
            panel_sheet["C61"] = mcc_boiler_power_supply
            panel_sheet["C62"] = mcc_boiler_control_supply
            panel_sheet["C63"] = boiler_evaporation
            panel_sheet["C64"] = boiler_output
            panel_sheet["C65"] = boiler_connected_load
            panel_sheet["C66"] = boiler_design_pressure

            mcc_heater_power_supply = f"{heater_power_supply_vac}, {heater_power_supply_phase}, {heater_power_supply_frequency}"
            mcc_heater_control_supply = f"{heater_control_supply_vac}, {heater_control_supply_phase}, {heater_control_supply_frequency}"

            if heater_evaporation == "NA":
                heater_evaporation = "Not Applicable"
            else:
                heater_evaporation = f"{heater_evaporation} kg/Hr"

            if heater_output == "NA":
                heater_output = "Not Applicable"
            else:
                heater_output = f"{heater_output} MW"

            if heater_connected_load == "NA":
                heater_connected_load = "Not Applicable"
            else:
                heater_connected_load = f"{heater_connected_load} kW"

            if heater_temperature == "NA":
                heater_temperature = "Not Applicable"
            else:
                heater_temperature = f"{heater_temperature} kg/cm2(g)/Bar"

            if (
                is_punching_details_for_heater_selected == "0"
                or is_punching_details_for_heater_selected == 0
            ):
                heater_model = "Not Applicable"
                heater_fuel = "Not Applicable"
                heater_year = "Not Applicable"
                mcc_heater_power_supply = "Not Applicable"
                mcc_heater_control_supply = "Not Applicable"
                heater_evaporation = "Not Applicable"
                heater_output = "Not Applicable"
                heater_connected_load = "Not Applicable"
                heater_temperature = "Not Applicable"

            panel_sheet["C68"] = heater_model
            panel_sheet["C69"] = heater_fuel
            panel_sheet["C70"] = heater_year
            panel_sheet["C71"] = mcc_heater_power_supply
            panel_sheet["C72"] = mcc_heater_control_supply
            panel_sheet["C73"] = heater_evaporation
            panel_sheet["C74"] = heater_output
            panel_sheet["C75"] = heater_connected_load
            panel_sheet["C76"] = heater_temperature

            panel_sheet["C78"] = na_to_string(spg_name_plate_unit_name)
            panel_sheet["C79"] = na_to_string(spg_name_plate_capacity)
            panel_sheet["C80"] = na_to_string(spg_name_plate_manufacturing_year)
            panel_sheet["C81"] = na_to_string(spg_name_plate_weight)
            panel_sheet["C82"] = na_to_string(spg_name_plate_oc_number)
            panel_sheet["C83"] = na_to_string(spg_name_plate_part_code)

        elif project_panel.get("panel_main_type") == "PCC":

            pcc_panel_data = frappe.db.get_list(
                "PCC Panel", {"revision_id": revision_id, "panel_id": panel_id}, "*"
            )
            pcc_panel_data = pcc_panel_data[0]

            panel_sheet = template_workbook.copy_worksheet(pcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")

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
            mi_analog = pcc_panel_data.get("mi_analog", "NA")
            mi_digital = pcc_panel_data.get("mi_digital", "NA")
            mi_communication_protocol = pcc_panel_data.get(
                "mi_communication_protocol", "NA"
            )
            ga_moc_material = pcc_panel_data.get("ga_moc_material")
            door_thickness = pcc_panel_data.get("door_thickness")
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
            ga_gland_plate_thickness = pcc_panel_data.get("ga_gland_plate_thickness")
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

            pcc_incomer_data = f"{incomer_ampere}, {incomer_pole} Pole {incomer_type} \n{incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type}"

            if is_indication_on_selected == "0" or is_indication_on_selected == 0:
                led_type_on_input = "Not Applicable"

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            panel_sheet["C5"] = na_to_string(pcc_incomer_data)
            panel_sheet["C6"] = led_type_on_input
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input
            panel_sheet["C9"] = na_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = na_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = na_to_string(is_white_healthy_trip_circuit_selected)
            panel_sheet["C12"] = na_to_string(alarm_annunciator)

            analog_data = (
                mi_analog.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
            )
            digital_data = (
                mi_digital.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
            )

            if "NA" in mi_analog:
                analog_data = "Not Applicable"

            if "NA" in mi_digital:
                digital_data = "Not Applicable"

            if "NA" in mi_communication_protocol:
                mi_communication_protocol = "Not Applicable"

            panel_sheet["C14"] = analog_data
            panel_sheet["C15"] = digital_data
            panel_sheet["C16"] = mi_communication_protocol

            panel_sheet["C18"] = current_transformer_coating
            panel_sheet["C19"] = current_transformer_number
            panel_sheet["C20"] = current_transformer_configuration

            # General

            panel_sheet["C22"] = ga_moc_material  # MOC
            panel_sheet["C23"] = (
                ga_moc_thickness_door  # Component Mounting Plate Thickness
            )
            panel_sheet["C24"] = door_thickness  # Door Thickness
            panel_sheet["C25"] = ga_moc_thickness_covers  # Top & Side Thickness
            panel_sheet["C26"] = ga_gland_plate_thickness  # Gland Plate Thickness
            panel_sheet["C27"] = ga_gland_plate_3mm_drill_type  # Gland Plate Type
            panel_sheet["C28"] = ga_pcc_compartmental  # Panel Front Type
            panel_sheet["C29"] = (
                ga_pcc_construction_front_type  # Type of Construction for Board
            )
            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_pcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            if (
                is_marshalling_section_selected == 0
                or is_marshalling_section_selected == "0"
            ):
                marshalling_section_text_area = "Not Applicable"

            panel_sheet["C35"] = marshalling_section_text_area  # Marshalling Section
            panel_sheet["C36"] = num_to_string(is_cable_alley_section_selected)
            panel_sheet["C37"] = num_to_string(
                is_power_and_bus_separation_section_selected
            )  # BUS
            panel_sheet["C38"] = num_to_string(
                is_both_side_extension_section_selected
            )  # Extension on Both sides
            panel_sheet["C39"] = ga_busbar_chamber_position  # Busbar Chamber position
            panel_sheet["C40"] = ga_power_and_control_busbar_separation  # BUSBAR
            panel_sheet["C41"] = ga_enclosure_protection_degree  # Degree of Enclosure
            panel_sheet["C42"] = ga_cable_entry_position  # BUSBAR

            # end

            panel_sheet["C44"] = "AS per OEM Standard"
            panel_sheet["C45"] = ppc_interior_and_exterior_paint_shade
            panel_sheet["C46"] = ppc_component_mounting_plate_paint_shade

            panel_sheet["C47"] = ppc_minimum_coating_thickness
            panel_sheet["C48"] = "Black"
            panel_sheet["C49"] = ppc_pretreatment_panel_standard
            panel_sheet["C50"] = general_requirments_for_construction

            panel_sheet["C52"] = commissioning_spare
            panel_sheet["C53"] = two_year_operational_spare

            pcc_boiler_power_supply = f"{boiler_power_supply_vac}, {boiler_power_supply_phase}, {boiler_power_supply_frequency}"
            pcc_boiler_control_supply = f"{boiler_control_supply_vac}, {boiler_control_supply_phase}, {boiler_control_supply_frequency}"

            if boiler_evaporation == "NA":
                boiler_evaporation = "Not Applicable"
            else:
                boiler_evaporation = f"{boiler_evaporation} kg/Hr"

            if boiler_output == "NA":
                boiler_output = "Not Applicable"
            else:
                boiler_output = f"{boiler_output} MW"

            if boiler_connected_load == "NA":
                boiler_connected_load = "Not Applicable"
            else:
                boiler_connected_load = f"{boiler_connected_load} kW"

            if boiler_design_pressure == "NA":
                boiler_design_pressure = "Not Applicable"
            else:
                boiler_design_pressure = f"{boiler_design_pressure} kg/cm2(g)/Bar"

            if (
                is_punching_details_for_boiler_selected == "0"
                or is_punching_details_for_boiler_selected == 0
            ):
                boiler_connected_load = "Not Applicable"
                boiler_model = "Not Applicable"
                boiler_fuel = "Not Applicable"
                boiler_year = "Not Applicable"
                boiler_evaporation = "Not Applicable"
                boiler_output = "Not Applicable"
                boiler_connected_load = "Not Applicable"
                boiler_design_pressure = "Not Applicable"
                pcc_boiler_power_supply = "Not Applicable"
                pcc_boiler_control_supply = "Not Applicable"

            panel_sheet["C56"] = boiler_model
            panel_sheet["C57"] = boiler_fuel
            panel_sheet["C58"] = boiler_year
            panel_sheet["C59"] = pcc_boiler_power_supply
            panel_sheet["C60"] = pcc_boiler_control_supply
            panel_sheet["C61"] = boiler_evaporation
            panel_sheet["C62"] = boiler_output
            panel_sheet["C63"] = boiler_connected_load
            panel_sheet["C64"] = boiler_design_pressure

            pcc_heater_power_supply = f"{heater_power_supply_vac}, {heater_power_supply_phase}, {heater_power_supply_frequency}"

            pcc_heater_control_supply = f"{heater_control_supply_vac}, {heater_control_supply_phase}, {heater_control_supply_frequency}"

            if heater_evaporation == "NA":
                heater_evaporation = "Not Applicable"
            else:
                heater_evaporation = f"{heater_evaporation} kg/Hr"

            if heater_output == "NA":
                heater_output = "Not Applicable"
            else:
                heater_output = f"{heater_output} MW"

            if heater_connected_load == "NA":
                heater_connected_load = "Not Applicable"
            else:
                heater_connected_load = f"{heater_connected_load} kW"

            if heater_temperature == "NA":
                heater_temperature = "Not Applicable"
            else:
                heater_temperature = f"{heater_temperature} kg/cm2(g)/Bar"

            if (
                is_punching_details_for_heater_selected == "0"
                or is_punching_details_for_heater_selected == 0
            ):
                heater_model = "Not Applicable"
                heater_fuel = "Not Applicable"
                heater_year = "Not Applicable"
                pcc_heater_power_supply = "Not Applicable"
                pcc_heater_control_supply = "Not Applicable"
                heater_evaporation = "Not Applicable"
                heater_output = "Not Applicable"
                heater_connected_load = "Not Applicable"
                heater_temperature = "Not Applicable"

            panel_sheet["C66"] = heater_model
            panel_sheet["C67"] = heater_fuel
            panel_sheet["C68"] = heater_year
            panel_sheet["C69"] = pcc_heater_power_supply
            panel_sheet["C70"] = pcc_heater_control_supply
            panel_sheet["C71"] = heater_evaporation
            panel_sheet["C72"] = heater_output
            panel_sheet["C73"] = heater_connected_load
            panel_sheet["C74"] = heater_temperature

            panel_sheet["C76"] = na_to_string(spg_name_plate_unit_name)
            panel_sheet["C77"] = na_to_string(spg_name_plate_capacity)
            panel_sheet["C78"] = na_to_string(spg_name_plate_manufacturing_year)
            panel_sheet["C79"] = na_to_string(spg_name_plate_weight)
            panel_sheet["C80"] = na_to_string(spg_name_plate_oc_number)
            panel_sheet["C81"] = na_to_string(spg_name_plate_part_code)

        else:
            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"revision_id": revision_id, "panel_id": panel_id}, "*"
            )
            mcc_panel_data = mcc_panel_data[0]

            panel_sheet = template_workbook.copy_worksheet(mcc_cum_plc_sheet)
            panel_sheet.title = project_panel.get("panel_name")

            panel_sheet["B3"] = project_panel.get("panel_name")

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
            led_type_on_input = mcc_panel_data.get("led_type_on_input")
            is_indication_off_selected = mcc_panel_data.get(
                "is_indication_off_selected"
            )
            led_type_off_input = mcc_panel_data.get("led_type_off_input")
            is_indication_trip_selected = mcc_panel_data.get(
                "is_indication_trip_selected"
            )
            led_type_trip_input = mcc_panel_data.get("led_type_trip_input")
            is_blue_cb_spring_charge_selected = mcc_panel_data.get(
                "is_blue_cb_spring_charge_selected"
            )
            is_red_cb_in_service = mcc_panel_data.get("is_red_cb_in_service")
            is_white_healthy_trip_circuit_selected = mcc_panel_data.get(
                "is_white_healthy_trip_circuit_selected"
            )
            is_other_selected = mcc_panel_data.get("is_other_selected")
            led_type_other_input = mcc_panel_data.get("led_type_other_input")
            current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
            )
            control_transformer_coating = mcc_panel_data.get(
                "control_transformer_coating"
            )
            control_transformer_configuration = mcc_panel_data.get(
                "control_transformer_configuration"
            )
            current_transformer_number = mcc_panel_data.get(
                "current_transformer_number"
            )
            current_transformer_configuration = mcc_panel_data.get(
                "current_transformer_configuration"
            )
            alarm_annunciator = mcc_panel_data.get("alarm_annunciator")
            mi_analog = mcc_panel_data.get("mi_analog", "NA")
            mi_digital = mcc_panel_data.get("mi_digital", "NA")
            mi_communication_protocol = mcc_panel_data.get(
                "mi_communication_protocol", "NA"
            )
            ga_moc_material = mcc_panel_data.get("ga_moc_material")
            door_thickness = mcc_panel_data.get("door_thickness")
            ga_moc_thickness_door = mcc_panel_data.get("ga_moc_thickness_door")
            ga_moc_thickness_covers = mcc_panel_data.get("ga_moc_thickness_covers")
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
            ga_panel_mounting_frame = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            is_marshalling_section_selected = mcc_panel_data.get(
                "is_marshalling_section_selected"
            )
            marshalling_section_text_area = mcc_panel_data.get(
                "marshalling_section_text_area"
            )
            is_cable_alley_section_selected = mcc_panel_data.get(
                "is_cable_alley_section_selected"
            )
            is_power_and_bus_separation_section_selected = mcc_panel_data.get(
                "is_power_and_bus_separation_section_selected"
            )
            is_both_side_extension_section_selected = mcc_panel_data.get(
                "is_both_side_extension_section_selected"
            )
            ga_gland_plate_3mm_drill_type = mcc_panel_data.get(
                "ga_gland_plate_3mm_drill_type"
            )
            ga_gland_plate_thickness = mcc_panel_data.get("ga_gland_plate_thickness")
            ga_busbar_chamber_position = mcc_panel_data.get(
                "ga_busbar_chamber_position"
            )
            ga_power_and_control_busbar_separation = mcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            ga_enclosure_protection_degree = mcc_panel_data.get(
                "ga_enclosure_protection_degree"
            )
            ga_cable_entry_position = mcc_panel_data.get("ga_cable_entry_position")
            general_requirments_for_construction = mcc_panel_data.get(
                "general_requirments_for_construction"
            )
            ppc_painting_standards = mcc_panel_data.get("ppc_painting_standards")
            ppc_interior_and_exterior_paint_shade = mcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = mcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            ppc_base_frame_paint_shade = mcc_panel_data.get(
                "ppc_base_frame_paint_shade"
            )
            ppc_minimum_coating_thickness = mcc_panel_data.get(
                "ppc_minimum_coating_thickness"
            )
            ppc_pretreatment_panel_standard = mcc_panel_data.get(
                "ppc_pretreatment_panel_standard"
            )
            vfd_auto_manual_selection = mcc_panel_data.get("vfd_auto_manual_selection")
            two_year_operational_spare = mcc_panel_data.get(
                "two_year_operational_spare"
            )
            commissioning_spare = mcc_panel_data.get("commissioning_spare")
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
            is_spg_applicable = mcc_panel_data.get("is_spg_applicable")
            spg_name_plate_unit_name = mcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = mcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = mcc_panel_data.get(
                "spg_name_plate_manufacturing_year"
            )
            spg_name_plate_weight = mcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = mcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = mcc_panel_data.get("spg_name_plate_part_code")
            special_note = mcc_panel_data.get("special_note")

            incomer_data = f"Upto {incomer_ampere}, {incomer_pole} Pole {incomer_type} \nAbove {incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type} "

            if is_indication_on_selected == "0" or is_indication_on_selected == 0:
                led_type_on_input = "Not Applicable"

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            panel_sheet["C5"] = na_to_string(incomer_data)
            panel_sheet["C6"] = led_type_on_input
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input
            panel_sheet["C9"] = na_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = na_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = na_to_string(is_white_healthy_trip_circuit_selected)
            panel_sheet["C12"] = na_to_string(alarm_annunciator)

            analog_data = (
                mi_analog.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
            )
            digital_data = (
                mi_digital.replace("[", "")
                .replace("]", "")
                .replace('"', "")
                .replace(",", ", ")
            )

            if "NA" in mi_analog:
                analog_data = "Not Applicable"

            if "NA" in mi_digital:
                digital_data = "Not Applicable"

            if "NA" in mi_communication_protocol:
                mi_communication_protocol = "Not Applicable"

            panel_sheet["C14"] = analog_data
            panel_sheet["C15"] = digital_data
            panel_sheet["C16"] = mi_communication_protocol

            panel_sheet["C18"] = current_transformer_coating
            panel_sheet["C19"] = current_transformer_number
            panel_sheet["C20"] = current_transformer_configuration

            panel_sheet["C22"] = ga_moc_material  # MOC
            panel_sheet["C23"] = (
                ga_moc_thickness_door  # Component Mounting Plate Thickness
            )
            panel_sheet["C24"] = door_thickness  # Door Thickness
            panel_sheet["C25"] = ga_moc_thickness_covers  # Top & Side Thickness
            panel_sheet["C26"] = ga_gland_plate_thickness  # Gland Plate Thickness
            panel_sheet["C27"] = ga_gland_plate_3mm_drill_type  # Gland Plate Type
            panel_sheet["C28"] = ga_mcc_compartmental  # Panel Front Type
            panel_sheet["C29"] = (
                ga_mcc_construction_front_type  # Type of Construction for Board
            )
            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_mcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            if (
                is_marshalling_section_selected == 0
                or is_marshalling_section_selected == "0"
            ):
                marshalling_section_text_area = "Not Applicable"

            panel_sheet["C35"] = marshalling_section_text_area  # Marshalling Section
            panel_sheet["C36"] = num_to_string(is_cable_alley_section_selected)
            panel_sheet["C37"] = num_to_string(
                is_power_and_bus_separation_section_selected
            )  # BUS
            panel_sheet["C38"] = num_to_string(
                is_both_side_extension_section_selected
            )  # Extension on Both sides
            panel_sheet["C39"] = ga_busbar_chamber_position  # Busbar Chamber position
            panel_sheet["C40"] = ga_power_and_control_busbar_separation  # BUSBAR
            panel_sheet["C41"] = ga_enclosure_protection_degree  # Degree of Enclosure
            panel_sheet["C42"] = ga_cable_entry_position  # BUSBAR

            panel_sheet["C44"] = "As per OEM Stanadard"
            panel_sheet["C45"] = ppc_interior_and_exterior_paint_shade
            panel_sheet["C46"] = ppc_component_mounting_plate_paint_shade

            panel_sheet["C47"] = ppc_minimum_coating_thickness
            panel_sheet["C48"] = "Black"
            panel_sheet["C49"] = ppc_pretreatment_panel_standard
            panel_sheet["C50"] = general_requirments_for_construction

            panel_sheet["C52"] = vfd_auto_manual_selection
            panel_sheet["C54"] = commissioning_spare
            panel_sheet["C55"] = two_year_operational_spare

            plc_panel_1 = frappe.db.get_list(
                "Panel PLC 1 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_1 = plc_panel_1[0]
            plc_panel_2 = frappe.db.get_list(
                "Panel PLC 2 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_2 = plc_panel_2[0]
            plc_panel_3 = frappe.db.get_list(
                "Panel PLC 3 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_3 = plc_panel_3[0]

            plc_panel = {**plc_panel_1, **plc_panel_2, **plc_panel_3}
            print("wait")
            # PLC fields
            # Supply Requirements
            panel_sheet["C58"] = na_to_string(
                plc_panel.get("ups_control_voltage", "NA")
            )
            panel_sheet["C59"] = na_to_string(
                plc_panel.get("non_ups_control_voltage", "NA")
            )
            panel_sheet["C60"] = num_to_string(
                plc_panel.get("is_bulk_power_supply_selected", "0")
            )

            # UPS
            ups_scope = plc_panel.get("ups_scope")
            panel_sheet["C62"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_make", "NA"))
            )
            panel_sheet["C63"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_input_voltage_3p", "NA"))
            )
            panel_sheet["C64"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_input_voltage_1p", "NA"))
            )
            panel_sheet["C65"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_output_voltage_1p", "NA"))
            )
            panel_sheet["C66"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_type", "NA"))
            )
            panel_sheet["C67"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else na_to_string(plc_panel.get("ups_battery_type", "NA"))
            )
            panel_sheet["C68"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else num_to_string(
                    plc_panel.get("is_ups_battery_mounting_rack_selected", "0")
                )
            )
            panel_sheet["C69"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else plc_panel.get("ups_battery_backup_time")
            )
            panel_sheet["C70"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else plc_panel.get("ups_redundancy")
            )

            # PLC Hardware
            panel_sheet["C72"] = ""
            panel_sheet["C73"] = plc_panel.get(
                "plc_cpu_system_series", "Not Applicable"
            )
            panel_sheet["C74"] = plc_panel.get(
                "plc_cpu_system_input_voltage", "Not Applicable"
            )
            panel_sheet["C75"] = plc_panel.get(
                "plc_cpu_system_memory_free_space_after_program", "Not Applicable"
            )

            # Redundancy
            panel_sheet["C77"] = num_to_string(
                plc_panel.get("is_power_supply_plc_cpu_system_selected", "0")
            )
            panel_sheet["C78"] = num_to_string(
                plc_panel.get("is_power_supply_input_output_module_selected", "0")
            )
            panel_sheet["C79"] = num_to_string(
                plc_panel.get("is_plc_input_output_modules_system_selected", "0")
            )
            panel_sheet["C80"] = num_to_string(
                plc_panel.get(
                    "is_plc_cpu_system_and_input_output_modules_system_selected", "0"
                )
            )
            panel_sheet["C81"] = num_to_string(
                plc_panel.get("is_plc_cpu_system_and_hmi_scada_selected", "0")
            )
            panel_sheet["C82"] = num_to_string(
                plc_panel.get("is_plc_cpu_system_and_third_party_devices_selected", "0")
            )
            panel_sheet["C83"] = num_to_string(
                plc_panel.get("is_plc_cpu_system_selected", "0")
            )

            # PLC Panel Mounted
            panel_sheet["C85"] = plc_panel.get("panel_mounted_ac", "Not Applicable")
            is_marshalling_cabinet_for_plc_and_ups_selected = plc_panel.get(
                "is_marshalling_cabinet_for_plc_and_ups_selected", 0
            )
            panel_sheet["C86"] = (
                plc_panel.get("marshalling_cabinet_for_plc_and_ups")
                if int(is_marshalling_cabinet_for_plc_and_ups_selected) == 1
                else "Not Applicable"
            )

            # Panel Mounted Push Buttons , Indication Lamps & Colors
            is_electronic_hooter_selected = plc_panel.get(
                "is_electronic_hooter_selected"
            )
            panel_sheet["C88"] = (
                plc_panel.get("electronic_hooter_acknowledge")
                if int(is_electronic_hooter_selected) == 1
                else "Not Applicable"
            )
            panel_sheet["C89"] = na_to_string(
                plc_panel.get("panel_power_supply_on_color", "NA")
            )
            panel_sheet["C90"] = na_to_string(
                plc_panel.get("panel_power_supply_off_color", "NA")
            )
            panel_sheet["C91"] = na_to_string(
                plc_panel.get("indicating_lamp_color_for_nonups_power_supply", "NA")
            )
            panel_sheet["C92"] = na_to_string(
                plc_panel.get("indicating_lamp_colour_for_ups_power_supply", "NA")
            )

            # DI Modules
            panel_sheet["C94"] = plc_panel.get("di_module_channel_density")
            panel_sheet["C95"] = plc_panel.get("loop_current")
            panel_sheet["C96"] = plc_panel.get("isolation")
            panel_sheet["C97"] = plc_panel.get("di_module_input_type")
            panel_sheet["C98"] = plc_panel.get("interrogation_voltage")
            panel_sheet["C99"] = plc_panel.get("scan_time")

            # DO Modules
            panel_sheet["C101"] = plc_panel.get("do_module_channel_density")
            panel_sheet["C102"] = plc_panel.get("do_module_loop_current")
            panel_sheet["C103"] = plc_panel.get("do_module_isolation")
            panel_sheet["C104"] = plc_panel.get("do_module_output_type")

            # Interposing Relay
            is_no_of_contacts_selected = plc_panel.get("is_no_of_contacts_selected")
            panel_sheet["C106"] = na_to_string(plc_panel.get("interposing_relay", "NA"))
            panel_sheet["C107"] = plc_panel.get("interposing_relay_contacts_rating")
            panel_sheet["C108"] = (
                plc_panel.get("no_of_contacts")
                if int(is_no_of_contacts_selected) == 1
                else "Not Applicable"
            )

            # AI Modules
            panel_sheet["C110"] = plc_panel.get("ai_module_channel_density")
            panel_sheet["C111"] = plc_panel.get("ai_module_loop_current")
            panel_sheet["C112"] = plc_panel.get("ai_module_isolation")
            panel_sheet["C113"] = plc_panel.get("ai_module_input_type")
            panel_sheet["C114"] = plc_panel.get("ai_module_scan_time")
            is_ai_module_hart_protocol_support_selected = plc_panel.get(
                "is_ai_module_hart_protocol_support_selected"
            )
            panel_sheet["C115"] = (
                "Applicable"
                if int(is_ai_module_hart_protocol_support_selected) == 1
                else "Not Applicable"
            )

            # AO Modules
            panel_sheet["C117"] = plc_panel.get("ao_module_channel_density")
            panel_sheet["C118"] = plc_panel.get("ao_module_loop_current")
            panel_sheet["C119"] = plc_panel.get("ao_module_isolation")
            panel_sheet["C120"] = plc_panel.get("ao_module_output_type")
            panel_sheet["C121"] = plc_panel.get("ao_module_scan_time")
            is_ao_module_hart_protocol_support_selected = plc_panel.get(
                "is_ao_module_hart_protocol_support_selected"
            )
            panel_sheet["C122"] = (
                "Applicable"
                if int(is_ao_module_hart_protocol_support_selected) == 1
                else "Not Applicable"
            )

            # RTD Modules
            panel_sheet["C124"] = plc_panel.get("rtd_module_channel_density")
            panel_sheet["C125"] = plc_panel.get("rtd_module_loop_current")
            panel_sheet["C126"] = plc_panel.get("rtd_module_isolation")
            panel_sheet["C127"] = plc_panel.get("rtd_module_input_type")
            panel_sheet["C128"] = plc_panel.get("rtd_module_scan_time")
            is_rtd_module_hart_protocol_support_selected = plc_panel.get(
                "is_rtd_module_hart_protocol_support_selected"
            )
            panel_sheet["C129"] = (
                "Applicable"
                if int(is_rtd_module_hart_protocol_support_selected) == 1
                else "Not Applicable"
            )

            # Thermocouple Modules
            panel_sheet["C131"] = plc_panel.get("thermocouple_module_channel_density")
            panel_sheet["C132"] = plc_panel.get("thermocouple_module_loop_current")
            panel_sheet["C133"] = plc_panel.get("thermocouple_module_isolation")
            panel_sheet["C134"] = plc_panel.get("thermocouple_module_input_type")
            panel_sheet["C135"] = plc_panel.get("thermocouple_module_scan_time")
            panel_sheet["C136"] = ""

            # Universal Modules
            panel_sheet["C138"] = ""
            panel_sheet["C139"] = ""
            panel_sheet["C140"] = ""
            panel_sheet["C141"] = ""
            panel_sheet["C142"] = ""
            panel_sheet["C143"] = ""

            # Terminal Block Connectors
            panel_sheet["C145"] = ""
            panel_sheet["C146"] = ""
            panel_sheet["C147"] = ""
            panel_sheet["C148"] = ""
            panel_sheet["C149"] = ""
            panel_sheet["C150"] = ""

            # HMI
            panel_sheet["C152"] = ""
            panel_sheet["C153"] = ""
            panel_sheet["C154"] = ""
            panel_sheet["C155"] = ""
            panel_sheet["C156"] = ""
            panel_sheet["C157"] = ""

            # Human Interface Device
            panel_sheet["C159"] = ""
            panel_sheet["C160"] = ""
            panel_sheet["C161"] = ""

            # Software License
            panel_sheet["C163"] = ""
            panel_sheet["C164"] = ""
            panel_sheet["C165"] = ""

            # Engineering/Operating SCADA Station
            panel_sheet["C167"] = ""
            panel_sheet["C168"] = ""
            panel_sheet["C169"] = ""
            panel_sheet["C170"] = ""
            panel_sheet["C171"] = ""
            panel_sheet["C172"] = ""
            panel_sheet["C173"] = ""
            panel_sheet["C174"] = ""
            panel_sheet["C175"] = ""
            panel_sheet["C176"] = ""
            panel_sheet["C177"] = ""
            panel_sheet["C178"] = ""
            panel_sheet["C179"] = ""

            # Communication
            panel_sheet["C181"] = ""
            panel_sheet["C182"] = ""
            panel_sheet["C183"] = ""
            panel_sheet["C184"] = ""
            panel_sheet["C185"] = ""
            panel_sheet["C186"] = ""
            panel_sheet["C187"] = ""
            panel_sheet["C188"] = ""
            panel_sheet["C189"] = ""
            panel_sheet["C190"] = ""

            # Burner Controller LMV
            panel_sheet["C192"] = ""
            panel_sheet["C193"] = ""
            panel_sheet["C194"] = ""

            # PLC Spares
            panel_sheet["C196"] = ""
            panel_sheet["C197"] = ""
            panel_sheet["C198"] = ""
            panel_sheet["C199"] = ""

    ###############################################################################################################

    # Load the workbook from the template path
    # template_workbook.save("design_basis.xlsx")

    template_workbook.remove(mcc_sheet)
    template_workbook.remove(pcc_sheet)
    template_workbook.remove(mcc_cum_plc_sheet)

    # Create a BytesIO stream to save the workbook
    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    # Prepare the response for file download
    frappe.local.response.filename = "generated_design_basis.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
