import frappe
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.utils import (
    check_value_kW,
    check_value_kW_below,
    handle_make_of_component,
    na_to_string,
    num_to_string,
)


def get_design_basis_sheet(
    design_basis_sheet, project_id, revision_id, division_name, make_of_components_data
):
    # DESIGN BASIS SHEET #
    project_info_data = frappe.get_doc("Project Information", project_id).as_dict()

    main_supply_lv = project_info_data.get("main_supply_lv", "Not Applicable")
    main_supply_lv_variation = project_info_data.get(
        "main_supply_lv_variation", "Not Applicable"
    )
    main_supply_lv_phase = project_info_data.get(
        "main_supply_lv_phase", "Not Applicable"
    )

    main_supply_mv = project_info_data.get("main_supply_mv", "Not Applicable")
    main_supply_mv_variation = project_info_data.get(
        "main_supply_mv_variation", "Not Applicable"
    )
    main_supply_mv_phase = project_info_data.get(
        "main_supply_mv_phase", "Not Applicable"
    )

    control_supply = project_info_data.get("control_supply", "Not Applicable")
    control_supply_variation = project_info_data.get(
        "control_supply_variation", "Not Applicable"
    )
    control_supply_phase = project_info_data.get(
        "control_supply_phase", "Not Applicable"
    )

    utility_supply = project_info_data.get("utility_supply", "Not Applicable")
    utility_supply_variation = project_info_data.get(
        "utility_supply_variation", "Not Applicable"
    )
    utility_supply_phase = project_info_data.get(
        "utility_supply_phase", "Not Applicable"
    )

    project_info_freq = project_info_data.get("frequency", "Not Applicable")
    preojct_info_freq_var = project_info_data.get(
        "frequency_variation", "Not Applicable"
    )

    project_info_fault = project_info_data.get("fault_level", "Not Applicable")
    project_info_sec = project_info_data.get("sec", "Not Applicable")

    general_info_data = frappe.db.get_list(
        "Design Basis General Info", {"revision_id": revision_id}, "*"
    )
    general_info_data = general_info_data[0]
    battery_limit = general_info_data.get("battery_limit", "Not Applicable")

    design_basis_sheet["C4"] = (
        "Not Applicable"
        if main_supply_mv == "NA"
        else f"{main_supply_mv}, {main_supply_mv_variation}%, {main_supply_mv_phase}"
    )
    design_basis_sheet["C5"] = (
        "Not Applicable"
        if main_supply_lv == "NA"
        else f"{main_supply_lv}, {main_supply_lv_variation}%, {main_supply_lv_phase}"
    )
    design_basis_sheet["C6"] = (
        control_supply
        if control_supply_variation == "NA"
        else f"{control_supply}, {control_supply_variation}%, {control_supply_phase}"
    )
    design_basis_sheet["C7"] = (
        utility_supply
        if utility_supply_variation == "NA"
        else f"{utility_supply}, {utility_supply_variation}%, {utility_supply_phase}"
    )
    design_basis_sheet["C8"] = f"{project_info_freq} Hz , {preojct_info_freq_var}%"
    design_basis_sheet["C9"] = f"{project_info_fault} kA, {project_info_sec} Sec"
    design_basis_sheet["C10"] = (
        f'{project_info_data.get("ambient_temperature_max")} Deg. C'
    )
    design_basis_sheet["C11"] = (
        f'{project_info_data.get("ambient_temperature_min")} Deg. C'
    )
    design_basis_sheet["C12"] = (
        f'{project_info_data.get("electrical_design_temperature")} Deg. C'
    )
    design_basis_sheet["C13"] = int(project_info_data.get("seismic_zone", 0))
    design_basis_sheet["C14"] = f'{project_info_data.get("max_humidity")}%'
    design_basis_sheet["C15"] = f'{project_info_data.get("min_humidity")}%'
    design_basis_sheet["C16"] = f'{project_info_data.get("avg_humidity")}%'
    design_basis_sheet["C17"] = f'{project_info_data.get("performance_humidity")}%'
    design_basis_sheet["C18"] = f'{project_info_data.get("altitude")} meters'

    main_packages_data_array = frappe.db.get_list(
        "Project Main Package",
        fields=["*"],
        filters={"revision_id": revision_id},
        order_by="creation asc",
    )

    temp_main_package_name = "Not Applicable"
    temp_safe_area = "Not Applicable"
    temp_hazardous_area = "Not Applicable"
    temp_area_of_classification = "Not Applicable"

    if len(main_packages_data_array) > 0:
        package_name_array = []
        safe_area_array = []
        hazard_area_array = []

        for package in main_packages_data_array:
            package_name = package.get("main_package_name")
            current_package_id = package.get("name")
            package_name_array.append(package_name)

            current_package_data = frappe.get_doc(
                "Project Main Package", current_package_id
            ).as_dict()
            sub_package_data = current_package_data["sub_packages"]

            if len(sub_package_data) > 0:
                for sub_package in sub_package_data:
                    if sub_package["area_of_classification"] == "Safe Area":
                        safe_area_array.append(sub_package["sub_package_name"])
                    else:
                        hazard_area_array.append(sub_package["sub_package_name"])
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
        standard = (
            area_classification_data[0] if len(area_classification_data) > 0 else ""
        )
        classification_1 = (
            area_classification_data[1] if len(area_classification_data) > 1 else ""
        )
        gas_group = (
            area_classification_data[2] if len(area_classification_data) > 2 else ""
        )
        temperature_class = (
            area_classification_data[3] if len(area_classification_data) > 3 else ""
        )

        area_classification_data = f"Standard-{standard}, {classification_1}, Gas Group-{gas_group}, Temperature Class-{temperature_class}"

        package_name_array = ", ".join(package_name_array)
        safe_area_sub_package = ", ".join(safe_area_array)
        hazard_area_sub_package = ", ".join(hazard_area_array)

        temp_main_package_name = package_name_array

        if len(safe_area_array) > 0:
            temp_safe_area = safe_area_sub_package
        if len(hazard_area_array) > 0:
            temp_hazardous_area = hazard_area_sub_package
            temp_area_of_classification = area_classification_data

    design_basis_sheet["C20"] = temp_main_package_name
    design_basis_sheet["C21"] = temp_safe_area
    design_basis_sheet["C22"] = temp_hazardous_area
    design_basis_sheet["C23"] = temp_area_of_classification
    design_basis_sheet["C24"] = battery_limit

    # MOTOR PARAMETERS

    motor_parameters_data = frappe.db.get_list(
        "Design Basis Motor Parameters", {"revision_id": revision_id}, "*"
    )
    motor_parameters_data = motor_parameters_data[0]

    hazardous_area_efficiency_level = na_to_string(
        motor_parameters_data.get("hazardous_area_efficiency_level", "Not Applicable")
    )
    hazardous_area_insulation_class = na_to_string(
        motor_parameters_data.get("hazardous_area_insulation_class", "Not Applicable")
    )
    hazardous_area_temperature_rise = na_to_string(
        motor_parameters_data.get("hazardous_area_temperature_rise", "Not Applicable")
    )
    hazardous_area_enclosure_ip_rating = na_to_string(
        motor_parameters_data.get(
            "hazardous_area_enclosure_ip_rating", "Not Applicable"
        )
    )
    hazardous_area_max_temperature = na_to_string(
        motor_parameters_data.get("hazardous_area_max_temperature", "Not Applicable")
    )
    hazardous_area_min_temperature = na_to_string(
        motor_parameters_data.get("hazardous_area_min_temperature", "Not Applicable")
    )
    hazardous_area_altitude = na_to_string(
        motor_parameters_data.get("hazardous_area_altitude", "Not Applicable")
    )
    hazardous_area_terminal_box_ip_rating = na_to_string(
        motor_parameters_data.get(
            "hazardous_area_terminal_box_ip_rating", "Not Applicable"
        )
    )
    hazardous_area_thermister = na_to_string(
        motor_parameters_data.get("hazardous_area_thermister", "Not Applicable")
    )
    hazardous_area_space_heater = na_to_string(
        motor_parameters_data.get("hazardous_area_space_heater", "Not Applicable")
    )
    hazardous_area_certification = na_to_string(
        motor_parameters_data.get("hazardous_area_certification", "Not Applicable")
    )
    hazardous_area_bearing_rtd = na_to_string(
        motor_parameters_data.get("hazardous_area_bearing_rtd", "Not Applicable")
    )
    hazardous_area_winding_rtd = na_to_string(
        motor_parameters_data.get("hazardous_area_winding_rtd", "Not Applicable")
    )
    hazardous_area_bearing_type = na_to_string(
        motor_parameters_data.get("hazardous_area_bearing_type", "Not Applicable")
    )
    hazardous_area_duty = na_to_string(
        motor_parameters_data.get("hazardous_area_duty", "Not Applicable")
    )
    hazardous_area_service_factor = na_to_string(
        motor_parameters_data.get("hazardous_area_service_factor", "Not Applicable")
    )
    hazardous_area_cooling_type = na_to_string(
        motor_parameters_data.get("hazardous_area_cooling_type", "Not Applicable")
    )
    hazardous_area_body_material = na_to_string(
        motor_parameters_data.get("hazardous_area_body_material", "Not Applicable")
    )
    hazardous_area_terminal_box_material = na_to_string(
        motor_parameters_data.get(
            "hazardous_area_terminal_box_material", "Not Applicable"
        )
    )
    hazardous_area_paint_type_and_shade = na_to_string(
        motor_parameters_data.get(
            "hazardous_area_paint_type_and_shade", "Not Applicable"
        )
    )
    hazardous_area_starts_hour_permissible = na_to_string(
        motor_parameters_data.get(
            "hazardous_area_starts_hour_permissible", "Not Applicable"
        )
    )

    hazardous_area_bearing_rtd = na_to_string(hazardous_area_bearing_rtd)
    hazardous_area_winding_rtd = na_to_string(hazardous_area_winding_rtd)

    hazardous_area_max_temperature = f"{hazardous_area_max_temperature} Deg. C"
    hazardous_area_min_temperature = f"{hazardous_area_min_temperature} Deg. C"
    hazardous_area_altitude = f"{hazardous_area_altitude} meters"

    if hazardous_area_thermister == "NA" or hazardous_area_thermister == "As per OEM":
        hazardous_area_thermister = "Not Applicable"
    elif hazardous_area_thermister == "All":
        hazardous_area_thermister = f"{hazardous_area_thermister} kW"

    if hazardous_area_bearing_rtd == "NA":
        hazardous_area_bearing_rtd = "Not Applicable"
    if hazardous_area_winding_rtd == "NA":
        hazardous_area_winding_rtd = "Not Applicable"

    if (
        general_info_data.get("is_package_selection_enabled") == 1
        or general_info_data.get("is_package_selection_enabled") == "1"
    ):
        if temp_hazardous_area == "Not Applicable":
            # hazard ke sare fields not applicable kar do
            hazardous_area_efficiency_level = "Not Applicable"
            hazardous_area_insulation_class = "Not Applicable"
            hazardous_area_temperature_rise = "Not Applicable"
            hazardous_area_enclosure_ip_rating = "Not Applicable"
            hazardous_area_max_temperature = "Not Applicable"
            hazardous_area_min_temperature = "Not Applicable"
            hazardous_area_altitude = "Not Applicable"
            hazardous_area_terminal_box_ip_rating = "Not Applicable"
            hazardous_area_thermister = "Not Applicable"
            hazardous_area_space_heater = "Not Applicable"
            hazardous_area_certification = "Not Applicable"
            hazardous_area_bearing_rtd = "Not Applicable"
            hazardous_area_winding_rtd = "Not Applicable"
            hazardous_area_bearing_type = "Not Applicable"
            hazardous_area_duty = "Not Applicable"
            hazardous_area_service_factor = "Not Applicable"
            hazardous_area_cooling_type = "Not Applicable"
            hazardous_area_body_material = "Not Applicable"
            hazardous_area_terminal_box_material = "Not Applicable"
            hazardous_area_paint_type_and_shade = "Not Applicable"
            hazardous_area_starts_hour_permissible = "Not Applicable"

    design_basis_sheet["C27"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_efficiency_level")
    )
    design_basis_sheet["C28"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_insulation_class")
    )
    design_basis_sheet["C29"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_temperature_rise", "Not Applicable")
    )
    design_basis_sheet["C30"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_enclosure_ip_rating")
    )
    design_basis_sheet["C31"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else f'{motor_parameters_data.get("safe_area_max_temperature", "Not Applicable")} Deg. C'
    )
    design_basis_sheet["C32"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else f'{motor_parameters_data.get("safe_area_min_temperature", "Not Applicable")} Deg. C'
    )
    design_basis_sheet["C33"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else f'{motor_parameters_data.get("safe_area_altitude", "Not Applicable")} meters'
    )
    design_basis_sheet["C34"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get(
            "safe_area_terminal_box_ip_rating", "Not Applicable"
        )
    )
    safe_area_thermister = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_thermister", "Not Applicable")
    )
    design_basis_sheet["C35"] = check_value_kW(safe_area_thermister)
    safe_area_space_heater = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_space_heater", "Not Applicable")
    )
    design_basis_sheet["C36"] = check_value_kW(safe_area_space_heater)
    design_basis_sheet["C37"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else motor_parameters_data.get("safe_area_certification", "Not Applicable")
    )
    safe_area_bearing_rtd = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get("safe_area_bearing_rtd", "Not Applicable")
        )
    )
    design_basis_sheet["C38"] = check_value_kW(safe_area_bearing_rtd)
    safe_area_winding_rtd = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get("safe_area_winding_rtd", "Not Applicable")
        )
    )
    design_basis_sheet["C39"] = check_value_kW(safe_area_winding_rtd)
    design_basis_sheet["C40"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get("safe_area_bearing_type", "Not Applicable")
        )
    )
    design_basis_sheet["C41"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(motor_parameters_data.get("safe_area_duty", "Not Applicable"))
    )
    safe_area_service_factor = na_to_string(
        motor_parameters_data.get("safe_area_service_factor", 0)
    )
    safe_area_service_factor = (
        int(safe_area_service_factor)
        if safe_area_service_factor != "Not Applicable"
        else "Not Applicable"
    )
    design_basis_sheet["C42"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else safe_area_service_factor
    )
    design_basis_sheet["C43"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get("safe_area_cooling_type", "Not Applicable")
        )
    )
    design_basis_sheet["C44"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get("safe_area_body_material", "Not Applicable")
        )
    )
    design_basis_sheet["C45"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get(
                "safe_area_terminal_box_material", "Not Applicable"
            )
        )
    )
    design_basis_sheet["C46"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get(
                "safe_area_paint_type_and_shade", "Not Applicable"
            )
        )
    )
    design_basis_sheet["C47"] = (
        "Not Applicable"
        if temp_safe_area == "Not Applicable"
        else na_to_string(
            motor_parameters_data.get(
                "safe_area_starts_hour_permissible", "Not Applicable"
            )
        )
    )

    design_basis_sheet["D27"] = hazardous_area_efficiency_level
    design_basis_sheet["D28"] = hazardous_area_insulation_class
    design_basis_sheet["D29"] = hazardous_area_temperature_rise
    design_basis_sheet["D30"] = hazardous_area_enclosure_ip_rating
    design_basis_sheet["D31"] = hazardous_area_max_temperature
    design_basis_sheet["D32"] = hazardous_area_min_temperature
    design_basis_sheet["D33"] = hazardous_area_altitude
    design_basis_sheet["D34"] = hazardous_area_terminal_box_ip_rating
    design_basis_sheet["D35"] = check_value_kW(hazardous_area_thermister)
    design_basis_sheet["D36"] = check_value_kW(hazardous_area_space_heater)
    design_basis_sheet["D37"] = hazardous_area_certification
    design_basis_sheet["D38"] = check_value_kW(hazardous_area_bearing_rtd)
    design_basis_sheet["D39"] = check_value_kW(hazardous_area_winding_rtd)
    design_basis_sheet["D40"] = hazardous_area_bearing_type
    design_basis_sheet["D41"] = hazardous_area_duty
    design_basis_sheet["D42"] = hazardous_area_service_factor
    design_basis_sheet["D43"] = hazardous_area_cooling_type
    design_basis_sheet["D44"] = hazardous_area_body_material
    design_basis_sheet["D45"] = hazardous_area_terminal_box_material
    design_basis_sheet["D46"] = hazardous_area_paint_type_and_shade
    design_basis_sheet["D47"] = hazardous_area_starts_hour_permissible

    # MAKE OF COMPONENTS

    motor = make_of_components_data.get("motor")
    cable = make_of_components_data.get("cable")
    lv_switchgear = make_of_components_data.get("lv_switchgear")
    panel_enclosure = make_of_components_data.get("panel_enclosure")
    variable_frequency_speed_drive_vfd_vsd = make_of_components_data.get(
        "variable_frequency_speed_drive_vfd_vsd"
    )
    soft_starter = make_of_components_data.get("soft_starter")
    plc = make_of_components_data.get("plc")
    gland_make_of_component = make_of_components_data.get("gland_make")

    design_basis_sheet["C50"] = na_to_string(handle_make_of_component(motor))
    design_basis_sheet["C51"] = na_to_string(handle_make_of_component(cable))
    design_basis_sheet["C52"] = handle_make_of_component(lv_switchgear)
    design_basis_sheet["C53"] = na_to_string(handle_make_of_component(panel_enclosure))
    design_basis_sheet["C54"] = na_to_string(
        handle_make_of_component(variable_frequency_speed_drive_vfd_vsd)
    )
    design_basis_sheet["C55"] = na_to_string(handle_make_of_component(soft_starter))
    design_basis_sheet["C56"] = na_to_string(handle_make_of_component(plc))

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
    cc_digital_meters = na_to_string(common_config_data.get("digital_meters"))
    cc_analog_meters = na_to_string(common_config_data.get("analog_meters"))
    cc_communication_protocol = na_to_string(
        common_config_data.get("communication_protocol")
    )
    cc_pole = na_to_string(common_config_data.get("pole"))
    cc_dm_standard = na_to_string(common_config_data.get("dm_standard"))
    cc_power_wiring_color = na_to_string(common_config_data.get("power_wiring_color"))
    cc_power_wiring_size = na_to_string(common_config_data.get("power_wiring_size"))
    cc_control_wiring_color = na_to_string(
        common_config_data.get("control_wiring_color")
    )
    cc_control_wiring_size = na_to_string(common_config_data.get("control_wiring_size"))
    cc_vdc_24_wiring_color = na_to_string(common_config_data.get("vdc_24_wiring_color"))
    cc_vdc_24_wiring_size = na_to_string(common_config_data.get("vdc_24_wiring_size"))
    cc_analog_signal_wiring_color = na_to_string(
        common_config_data.get("analog_signal_wiring_color")
    )
    cc_analog_signal_wiring_size = na_to_string(
        common_config_data.get("analog_signal_wiring_size")
    )
    cc_ct_wiring_color = na_to_string(common_config_data.get("ct_wiring_color"))
    cc_ct_wiring_size = na_to_string(common_config_data.get("ct_wiring_size"))
    cc_rtd_thermocouple_wiring_color = na_to_string(
        common_config_data.get("rtd_thermocouple_wiring_color", "Not Applicable")
    )
    cc_rtd_thermocouple_wiring_size = na_to_string(
        common_config_data.get("rtd_thermocouple_wiring_size", "Not Applicable")
    )
    cc_cable_insulation_pvc = na_to_string(
        common_config_data.get("cable_insulation_pvc", "Not Applicable")
    )
    cc_air_clearance_between_phase_to_phase_bus = na_to_string(
        common_config_data.get(
            "air_clearance_between_phase_to_phase_bus", "Not Applicable"
        )
    )
    cc_air_clearance_between_phase_to_neutral_bus = na_to_string(
        common_config_data.get(
            "air_clearance_between_phase_to_neutral_bus", "Not Applicable"
        )
    )
    cc_ferrule = na_to_string(common_config_data.get("ferrule"))
    cc_ferrule_note = na_to_string(common_config_data.get("ferrule_note"))
    cc_device_identification_of_components = na_to_string(
        common_config_data.get("device_identification_of_components")
    )
    cc_general_note_internal_wiring = na_to_string(
        common_config_data.get("general_note_internal_wiring")
    )
    cc_power_terminal_clipon = na_to_string(
        common_config_data.get("power_terminal_clipon")
    )
    cc_power_terminal_busbar_type = na_to_string(
        common_config_data.get("power_terminal_busbar_type")
    )
    cc_control_terminal = na_to_string(common_config_data.get("control_terminal"))
    cc_spare_terminal = na_to_string(common_config_data.get("spare_terminal"))
    cc_forward_push_button_start = na_to_string(
        common_config_data.get("forward_push_button_start")
    )
    cc_reverse_push_button_start = na_to_string(
        common_config_data.get("reverse_push_button_start")
    )
    cc_push_button_start = na_to_string(common_config_data.get("push_button_start"))
    cc_push_button_stop = na_to_string(common_config_data.get("push_button_stop"))
    cc_push_button_ess = na_to_string(common_config_data.get("push_button_ess"))
    cc_potentiometer = na_to_string(common_config_data.get("potentiometer"))
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

    cc_is_safe_lpbs_selected = common_config_data.get("is_safe_lpbs_selected")
    cc_is_hazardous_lpbs_selected = common_config_data.get("is_hazardous_lpbs_selected")
    cc_is_hazardous_area_isolator_selected = common_config_data.get(
        "is_hazardous_area_isolator_selected"
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

    cc_safe_lpbs_type = na_to_string(common_config_data.get("safe_lpbs_type"))
    cc_hazardous_lpbs_type = na_to_string(common_config_data.get("hazardous_lpbs_type"))
    cc_safe_lpbs_enclosure = na_to_string(common_config_data.get("safe_lpbs_enclosure"))
    cc_hazardous_lpbs_enclosure = na_to_string(
        common_config_data.get("hazardous_lpbs_enclosure")
    )
    cc_safe_lpbs_material = na_to_string(common_config_data.get("safe_lpbs_material"))
    cc_hazardous_lpbs_material = na_to_string(
        common_config_data.get("hazardous_lpbs_material")
    )
    cc_safe_lpbs_qty = na_to_string(common_config_data.get("safe_lpbs_qty"))
    cc_hazardous_lpbs_qty = na_to_string(common_config_data.get("hazardous_lpbs_qty"))
    cc_safe_lpbs_color_shade = na_to_string(
        common_config_data.get("safe_lpbs_color_shade")
    )
    cc_hazardous_lpbs_color_shade = na_to_string(
        common_config_data.get("hazardous_lpbs_color_shade")
    )
    cc_safe_lpbs_canopy = na_to_string(common_config_data.get("safe_lpbs_canopy"))
    cc_hazardous_lpbs_canopy = na_to_string(
        common_config_data.get("hazardous_lpbs_canopy")
    )
    cc_safe_lpbs_canopy_type = na_to_string(
        common_config_data.get("safe_lpbs_canopy_type")
    )
    cc_hazardous_lpbs_canopy_type = na_to_string(
        common_config_data.get("hazardous_lpbs_canopy_type")
    )
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

    cc_power_bus_main_busbar_selection = na_to_string(
        common_config_data.get("power_bus_main_busbar_selection")
    )
    cc_power_bus_heat_pvc_sleeve = na_to_string(
        common_config_data.get("power_bus_heat_pvc_sleeve")
    )
    cc_power_bus_material = na_to_string(common_config_data.get("power_bus_material"))
    cc_power_bus_current_density = na_to_string(
        common_config_data.get("power_bus_current_density")
    )
    cc_power_bus_rating_of_busbar = na_to_string(
        common_config_data.get("power_bus_rating_of_busbar")
    )
    cc_control_bus_main_busbar_selection = na_to_string(
        common_config_data.get("control_bus_main_busbar_selection")
    )
    cc_control_bus_heat_pvc_sleeve = na_to_string(
        common_config_data.get("control_bus_heat_pvc_sleeve")
    )
    cc_control_bus_material = na_to_string(
        common_config_data.get("control_bus_material")
    )
    cc_control_bus_current_density = na_to_string(
        common_config_data.get("control_bus_current_density")
    )
    cc_control_bus_rating_of_busbar = na_to_string(
        common_config_data.get("control_bus_rating_of_busbar")
    )
    cc_earth_bus_main_busbar_selection = na_to_string(
        common_config_data.get("earth_bus_main_busbar_selection")
    )
    cc_earth_bus_busbar_position = na_to_string(
        common_config_data.get("earth_bus_busbar_position")
    )
    cc_earth_bus_material = na_to_string(common_config_data.get("earth_bus_material"))
    cc_earth_bus_current_density = na_to_string(
        common_config_data.get("earth_bus_current_density")
    )
    cc_earth_bus_rating_of_busbar = na_to_string(
        common_config_data.get("earth_bus_rating_of_busbar")
    )
    cc_door_earthing = na_to_string(common_config_data.get("door_earthing"))
    cc_instrument_earth = na_to_string(common_config_data.get("instrument_earth"))
    cc_general_note_busbar_and_insulation_materials = common_config_data.get(
        "general_note_busbar_and_insulation_materials"
    )
    cc_cooling_fans = na_to_string(common_config_data.get("cooling_fans"))
    cc_louvers_and_filters = na_to_string(common_config_data.get("louvers_and_filters"))
    cc_current_transformer = na_to_string(common_config_data.get("current_transformer"))
    cc_current_transformer_coating = na_to_string(
        common_config_data.get("current_transformer_coating")
    )
    cc_current_transformer_quantity = na_to_string(
        common_config_data.get("current_transformer_quantity")
    )
    cc_current_transformer_configuration = na_to_string(
        common_config_data.get("current_transformer_configuration")
    )
    cc_control_transformer_type = na_to_string(
        common_config_data.get("control_transformer_type")
    )

    if (
        cc_is_control_transformer_applicable == "0"
        or cc_is_control_transformer_applicable == 0
    ):
        cc_control_transformer_primary_voltage = "NA"
        cc_control_transformer_secondary_voltage_copy = "NA"
        cc_control_transformer_coating = "NA"
        cc_control_transformer_quantity = "NA"
        cc_control_transformer_configuration = "NA"
        cc_control_transformer_type = "NA"

    design_basis_sheet["C59"] = na_to_string(cc_dm_standard)
    design_basis_sheet["C61"] = na_to_string(cc_control_transformer_primary_voltage)
    design_basis_sheet["C62"] = na_to_string(
        cc_control_transformer_secondary_voltage_copy
    )
    design_basis_sheet["C63"] = na_to_string(cc_control_transformer_coating)
    design_basis_sheet["C64"] = na_to_string(cc_control_transformer_quantity)
    design_basis_sheet["C65"] = na_to_string(cc_control_transformer_configuration)
    design_basis_sheet["C66"] = na_to_string(cc_control_transformer_type)

    apfc_relay = na_to_string(common_config_data.get("apfc_relay"))

    design_basis_sheet["C68"] = (
        f"{apfc_relay} Stage" if apfc_relay != "Not Applicable" else apfc_relay
    )

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

    design_basis_sheet["C91"] = check_value_kW_below(cc_dol_starter)
    design_basis_sheet["C92"] = check_value_kW(cc_star_delta_starter)

    design_basis_sheet["C93"] = cc_mcc_switchgear_type
    if division_name != "SPG":
        cc_switchgear_combination = "Not Applicable"
    else:
        if not "Fuseless" in cc_mcc_switchgear_type:
            cc_switchgear_combination = "Not Applicable"

    design_basis_sheet["C94"] = cc_switchgear_combination

    design_basis_sheet["C96"] = check_value_kW(cc_ammeter)
    design_basis_sheet["C97"] = na_to_string(cc_ammeter_configuration)
    design_basis_sheet["C98"] = handle_make_of_component(cc_analog_meters)
    design_basis_sheet["C99"] = handle_make_of_component(cc_digital_meters)
    design_basis_sheet["C100"] = na_to_string(cc_communication_protocol)

    design_basis_sheet["C102"] = check_value_kW(cc_current_transformer)
    design_basis_sheet["C103"] = na_to_string(cc_current_transformer_coating)
    design_basis_sheet["C104"] = na_to_string(cc_current_transformer_quantity)
    design_basis_sheet["C105"] = na_to_string(cc_current_transformer_configuration)

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

    design_basis_sheet["C127"] = na_to_string(cc_power_terminal_clipon)
    design_basis_sheet["C128"] = na_to_string(cc_power_terminal_busbar_type)
    design_basis_sheet["C129"] = na_to_string(cc_control_terminal)
    design_basis_sheet["C130"] = f"{cc_spare_terminal} %"

    design_basis_sheet["C132"] = na_to_string(cc_push_button_start)
    design_basis_sheet["C133"] = na_to_string(cc_push_button_stop)
    design_basis_sheet["C134"] = na_to_string(cc_push_button_ess)
    design_basis_sheet["C135"] = na_to_string(cc_forward_push_button_start)
    design_basis_sheet["C136"] = na_to_string(cc_reverse_push_button_start)
    design_basis_sheet["C137"] = num_to_string(cc_potentiometer)

    if cc_is_push_button_speed_selected == 0 or cc_is_push_button_speed_selected == "0":
        cc_speed_increase_pb = "Not Applicable"
        cc_speed_decrease_pb = "Not Applicable"

    design_basis_sheet["C138"] = cc_speed_increase_pb
    design_basis_sheet["C139"] = cc_speed_decrease_pb
    design_basis_sheet["C140"] = na_to_string(cc_alarm_acknowledge_and_lamp_test)
    design_basis_sheet["C141"] = na_to_string(cc_test_dropdown)
    design_basis_sheet["C142"] = na_to_string(cc_reset_dropdown)
    design_basis_sheet["C143"] = na_to_string(cc_lamp_test_push_button)

    if cc_selector_switch_applicable == "Applicable":
        cc_selector_switch_applicable = (
            f"{cc_selector_switch_applicable}, {cc_selector_switch_lockable}"
        )

    design_basis_sheet["C145"] = cc_selector_switch_applicable

    design_basis_sheet["C147"] = na_to_string(cc_running_open)
    design_basis_sheet["C148"] = na_to_string(cc_stopped_closed)
    design_basis_sheet["C149"] = na_to_string(cc_trip)

    if int(cc_is_field_motor_isolator_selected) == 0:
        cc_safe_field_motor_type = "Not Applicable"
        cc_safe_field_motor_enclosure = "Not Applicable"
        cc_safe_field_motor_material = "Not Applicable"
        cc_safe_field_motor_qty = "Not Applicable"
        cc_safe_field_motor_isolator_color_shade = "Not Applicable"
        cc_safe_field_motor_cable_entry = "Not Applicable"
        cc_safe_field_motor_canopy = "Not Applicable"

        cc_hazardous_field_motor_type = "Not Applicable"
        cc_hazardous_field_motor_enclosure = "Not Applicable"
        cc_hazardous_field_motor_material = "Not Applicable"
        cc_hazardous_field_motor_qty = "Not Applicable"
        cc_hazardous_field_motor_isolator_color_shade = "Not Applicable"
        cc_hazardous_field_motor_cable_entry = "Not Applicable"
        cc_hazardous_field_motor_canopy = "Not Applicable"
    else:
        if int(cc_is_safe_area_isolator_selected) == 0:
            cc_safe_field_motor_type = "Not Applicable"
            cc_safe_field_motor_enclosure = "Not Applicable"
            cc_safe_field_motor_material = "Not Applicable"
            cc_safe_field_motor_qty = "Not Applicable"
            cc_safe_field_motor_isolator_color_shade = "Not Applicable"
            cc_safe_field_motor_cable_entry = "Not Applicable"
            cc_safe_field_motor_canopy = "Not Applicable"

        if int(cc_is_hazardous_area_isolator_selected) == 0:
            cc_hazardous_field_motor_type = "Not Applicable"
            cc_hazardous_field_motor_enclosure = "Not Applicable"
            cc_hazardous_field_motor_material = "Not Applicable"
            cc_hazardous_field_motor_qty = "Not Applicable"
            cc_hazardous_field_motor_isolator_color_shade = "Not Applicable"
            cc_hazardous_field_motor_cable_entry = "Not Applicable"
            cc_hazardous_field_motor_canopy = "Not Applicable"

    design_basis_sheet["C152"] = cc_safe_field_motor_type
    design_basis_sheet["C153"] = na_to_string(cc_safe_field_motor_enclosure)

    if (
        cc_safe_field_motor_material == "CRCA"
        or cc_safe_field_motor_material == "SS 316"
        or cc_safe_field_motor_material == "SS 306"
    ):
        cc_safe_field_motor_material = (
            f"{cc_safe_field_motor_material}, {cc_safe_field_motor_thickness}"
        )
        cc_safe_field_motor_cable_entry = f"{cc_safe_field_motor_cable_entry}, 3 mm"
    elif cc_safe_field_motor_material == "NA":
        cc_safe_field_motor_material = "Not Applicable"

    design_basis_sheet["C154"] = cc_safe_field_motor_material
    design_basis_sheet["C155"] = na_to_string(cc_safe_field_motor_qty)
    design_basis_sheet["C156"] = na_to_string(cc_safe_field_motor_isolator_color_shade)
    design_basis_sheet["C157"] = cc_safe_field_motor_cable_entry
    design_basis_sheet["C158"] = na_to_string(cc_safe_field_motor_canopy)
    design_basis_sheet["D152"] = cc_hazardous_field_motor_type
    design_basis_sheet["D153"] = na_to_string(cc_hazardous_field_motor_enclosure)

    if (
        cc_hazardous_field_motor_material == "CRCA"
        or cc_hazardous_field_motor_material == "SS 316"
        or cc_hazardous_field_motor_material == "SS 306"
    ):
        cc_hazardous_field_motor_material = f"{cc_hazardous_field_motor_material}, {cc_hazardous_field_motor_thickness} mm"
        cc_hazardous_field_motor_cable_entry = (
            f"{cc_hazardous_field_motor_cable_entry}, 3 mm"
        )
    elif cc_hazardous_field_motor_material == "NA":
        cc_hazardous_field_motor_material = "Not Applicable"

    design_basis_sheet["D154"] = na_to_string(cc_hazardous_field_motor_material)
    design_basis_sheet["D155"] = na_to_string(cc_hazardous_field_motor_qty)
    design_basis_sheet["D156"] = na_to_string(
        cc_hazardous_field_motor_isolator_color_shade
    )
    design_basis_sheet["D157"] = cc_hazardous_field_motor_cable_entry
    design_basis_sheet["D158"] = na_to_string(cc_hazardous_field_motor_canopy)

    if int(cc_is_local_push_button_station_selected) == 0:
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

    else:
        if int(cc_is_safe_lpbs_selected) == 0:
            cc_safe_lpbs_type = "Not Applicable"
            cc_safe_lpbs_enclosure = "Not Applicable"
            cc_safe_lpbs_material = "Not Applicable"
            cc_safe_lpbs_qty = "Not Applicable"
            cc_safe_lpbs_color_shade = "Not Applicable"
            cc_safe_lpbs_canopy = "Not Applicable"
            cc_safe_lpbs_canopy_type = "Not Applicable"

        if int(cc_is_hazardous_lpbs_selected) == 0:
            cc_hazardous_lpbs_type = "Not Applicable"
            cc_hazardous_lpbs_enclosure = "Not Applicable"
            cc_hazardous_lpbs_material = "Not Applicable"
            cc_hazardous_lpbs_qty = "Not Applicable"
            cc_hazardous_lpbs_color_shade = "Not Applicable"
            cc_hazardous_lpbs_canopy = "Not Applicable"
            cc_hazardous_lpbs_canopy_type = "Not Applicable"

    design_basis_sheet["C160"] = na_to_string(cc_lpbs_push_button_start_color)
    design_basis_sheet["C161"] = na_to_string(cc_forward_push_button_start)
    design_basis_sheet["C162"] = na_to_string(cc_reverse_push_button_start)
    design_basis_sheet["C163"] = na_to_string(cc_push_button_ess)
    design_basis_sheet["C164"] = na_to_string(cc_lpbs_speed_increase)
    design_basis_sheet["C165"] = na_to_string(cc_lpbs_speed_decrease)
    design_basis_sheet["C166"] = na_to_string(cc_lpbs_indication_lamp_start_color)
    design_basis_sheet["C167"] = na_to_string(cc_lpbs_indication_lamp_stop_color)

    if (
        cc_safe_lpbs_material == "CRCA"
        or cc_safe_lpbs_material == "SS 316"
        or cc_safe_lpbs_material == "SS 306"
    ):
        cc_safe_lpbs_material = (
            f"{cc_safe_lpbs_material}, {cc_hazardous_field_motor_thickness}"
        )
        cc_hazardous_field_motor_cable_entry = (
            f"{cc_hazardous_field_motor_cable_entry}, 3 mm"
        )
    elif cc_safe_lpbs_material == "NA":
        cc_safe_lpbs_material = "Not Applicable"

    design_basis_sheet["C169"] = na_to_string(cc_safe_lpbs_type)
    design_basis_sheet["C170"] = na_to_string(cc_safe_lpbs_enclosure)
    design_basis_sheet["C171"] = na_to_string(cc_safe_lpbs_material)
    design_basis_sheet["C172"] = na_to_string(cc_safe_lpbs_qty)
    design_basis_sheet["C173"] = na_to_string(cc_safe_lpbs_color_shade)
    design_basis_sheet["C174"] = na_to_string(cc_safe_lpbs_canopy)
    design_basis_sheet["C175"] = na_to_string(cc_safe_lpbs_canopy_type)

    design_basis_sheet["D169"] = na_to_string(cc_hazardous_lpbs_type)
    design_basis_sheet["D170"] = na_to_string(cc_hazardous_lpbs_enclosure)
    design_basis_sheet["D171"] = na_to_string(cc_hazardous_lpbs_material)
    design_basis_sheet["D172"] = na_to_string(cc_hazardous_lpbs_qty)
    design_basis_sheet["D173"] = na_to_string(cc_hazardous_lpbs_color_shade)
    design_basis_sheet["D174"] = na_to_string(cc_hazardous_lpbs_canopy)
    design_basis_sheet["D175"] = na_to_string(cc_hazardous_lpbs_canopy_type)

    design_basis_sheet["C177"] = na_to_string(cc_ferrule)
    design_basis_sheet["C178"] = cc_ferrule_note
    design_basis_sheet["C179"] = cc_device_identification_of_components

    design_basis_sheet["C181"] = cc_cooling_fans
    design_basis_sheet["C182"] = cc_louvers_and_filters

    cable_tray_data = frappe.db.get_list(
        "Cable Tray Layout", {"revision_id": revision_id}, "*"
    )
    cable_tray_data = cable_tray_data[0]

    ct_copper_conductor = na_to_string(
        cable_tray_data.get("copper_conductor", "Not Applicable")
    )
    ct_aluminium_conductor = na_to_string(
        cable_tray_data.get("aluminium_conductor", "Not Applicable")
    )

    ct_touching_factor_air = na_to_string(
        cable_tray_data.get("touching_factor_air", "Not Applicable")
    )
    ct_touching_factor_burid = na_to_string(
        cable_tray_data.get("touching_factor_burid", "Not Applicable")
    )
    ct_ambient_temp_factor_air = na_to_string(
        cable_tray_data.get("ambient_temp_factor_air")
    )
    ct_ambient_temp_factor_burid = na_to_string(
        cable_tray_data.get("ambient_temp_factor_burid")
    )
    ct_derating_factor_air = na_to_string(cable_tray_data.get("derating_factor_air"))
    ct_derating_factor_burid = na_to_string(
        cable_tray_data.get("derating_factor_burid")
    )
    ct_moc = na_to_string(cable_tray_data.get("moc"))
    ct_type_of_gland = na_to_string(cable_tray_data.get("type_of_gland"))
    ct_future_space_on_trays = na_to_string(
        cable_tray_data.get("future_space_on_trays")
    )
    ct_cable_placement = na_to_string(cable_tray_data.get("cable_placement"))
    ct_vertical_distance = na_to_string(cable_tray_data.get("vertical_distance"))
    ct_horizontal_distance = na_to_string(cable_tray_data.get("horizontal_distance"))
    ct_cable_tray_moc = na_to_string(cable_tray_data.get("cable_tray_moc"))
    ct_cable_tray_moc_input = na_to_string(cable_tray_data.get("cable_tray_moc_input"))

    cable_tray_cover = na_to_string(cable_tray_data.get("cable_tray_cover"))

    design_basis_sheet["C184"] = na_to_string(cable_tray_data.get("number_of_cores"))
    design_basis_sheet["C185"] = na_to_string(
        cable_tray_data.get("specific_requirement")
    )
    design_basis_sheet["C186"] = na_to_string(cable_tray_data.get("type_of_insulation"))
    design_basis_sheet["C187"] = na_to_string(cable_tray_data.get("color_scheme"))
    ct_motor_voltage_drop_during_starting = na_to_string(
        cable_tray_data.get("motor_voltage_drop_during_starting", "Not Applicable")
    )
    design_basis_sheet["C188"] = (
        f"{ct_motor_voltage_drop_during_starting} %"
        if ct_motor_voltage_drop_during_starting != "Not Applicable"
        else "Not Applicable"
    )

    ct_motor_voltage_drop_during_running = na_to_string(
        cable_tray_data.get("motor_voltage_drop_during_running", "Not Applicable")
    )
    design_basis_sheet["C189"] = (
        f"{ct_motor_voltage_drop_during_running} %"
        if ct_motor_voltage_drop_during_running != "Not Applicable"
        else "Not Applicable"
    )

    design_basis_sheet["C190"] = na_to_string(
        cable_tray_data.get("voltage_grade", "Not Applicable")
    )

    if ct_copper_conductor == "All":
        ct_copper_conductor = "All"
    else:
        ct_copper_conductor = f"{ct_copper_conductor} Sq. mm & Below"

    if "NA" in ct_aluminium_conductor:
        ct_aluminium_conductor = "Not Applicable"
    else:
        if ct_aluminium_conductor == "All":
            ct_aluminium_conductor = "All"
        else:
            ct_aluminium_conductor = f"{ct_aluminium_conductor} Sq. mm & Above"

    design_basis_sheet["C191"] = ct_copper_conductor
    design_basis_sheet["C192"] = ct_aluminium_conductor
    design_basis_sheet["C193"] = ct_touching_factor_air
    design_basis_sheet["C194"] = ct_ambient_temp_factor_air
    design_basis_sheet["C195"] = ct_derating_factor_air
    design_basis_sheet["C196"] = ct_touching_factor_burid
    design_basis_sheet["C197"] = ct_ambient_temp_factor_burid
    design_basis_sheet["C198"] = ct_derating_factor_burid

    design_basis_sheet["C200"] = handle_make_of_component(gland_make_of_component)
    design_basis_sheet["C201"] = ct_moc
    design_basis_sheet["C202"] = ct_type_of_gland

    gland_type_safe_area = "Not Applicable"
    gland_type_hazardous_area = "Not Applicable"

    if temp_hazardous_area != "Not Applicable":
        gland_type_hazardous_area = (
            f"{temp_area_of_classification}, with Dual Certification"
        )

    if temp_safe_area != "Not Applicable":
        gland_type_safe_area = "Weatherproof"

    design_basis_sheet["C203"] = gland_type_safe_area
    design_basis_sheet["C204"] = gland_type_hazardous_area

    design_basis_sheet["C206"] = num_to_string(cable_tray_cover)
    design_basis_sheet["C207"] = f"{ct_future_space_on_trays} %"
    design_basis_sheet["C208"] = ct_cable_placement
    design_basis_sheet["C209"] = f"{ct_vertical_distance} mm"
    design_basis_sheet["C210"] = f"{ct_horizontal_distance} mm"

    if ct_cable_tray_moc == "MS - Hot dipped Galvanised":
        ct_cable_tray_moc = f"{ct_cable_tray_moc}, {ct_cable_tray_moc_input}"

    design_basis_sheet["C211"] = ct_cable_tray_moc

    earthing_layout_data = frappe.db.get_list(
        "Layout Earthing", {"revision_id": revision_id}, "*"
    )
    earthing_layout_data = earthing_layout_data[0]

    soil_resistivity = na_to_string(earthing_layout_data.get("soil_resistivity"))

    design_basis_sheet["C213"] = na_to_string(
        earthing_layout_data.get("earthing_system")
    )
    design_basis_sheet["C214"] = na_to_string(earthing_layout_data.get("earth_strip"))
    design_basis_sheet["C215"] = na_to_string(earthing_layout_data.get("earth_pit"))
    design_basis_sheet["C216"] = (
        f"{soil_resistivity} ohm"
        if soil_resistivity != "Not Applicable"
        else "Not Applicable"
    )
    return design_basis_sheet
