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
    try:
        # Retrieve the payload from the request
        payload = frappe.local.form_dict
        revision_id = payload["revision_id"]

        project_data = frappe.db.get_list("Project", {"name": project_id}, "*")
        design_basis_revision_data = frappe.db.get_list(
            "Design Basis Revision History", {"name": revision_id}, ["*"]
        )

        # print(project_data[0]["project_id"], "project_data")

        # Define the path to the template
        template_path = frappe.get_app_path(
            "thermax_backend", "templates", "design_basis_template.xlsx"
        )

        template_workbook = load_workbook(template_path)

        project_id = design_basis_revision_data[0].get("project_id")
        project_description = design_basis_revision_data[0].get("description")
        project_status = design_basis_revision_data[0].get("status")
        owner = design_basis_revision_data[0].get("owner")

        ###############################################################################################################

        # Loading the Sheets of templates

        cover_sheet = {}  # template_workbook["COVER"]
        design_basis_sheet = {}  # template_workbook["Design Basis"]
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

        revision_date = datetime.strptime(modified, "%Y-%m-%d %H:%M:%S.%f").strftime(
            "%d-%m-%Y"
        )

        # COVER SHEET ################################################################################################################################

        division_name = project_data[0].get("division")
        project_name = project_data[0].get("project_name")
        project_oc_number = project_data[0].get("project_oc_number")
        approver = project_data[0].get("approver")
        client_name = project_data[0].get("client_name")
        consultant_name = project_data[0].get("consultant_name")
        modified = project_data[0].get("modified")

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

        # cover_sheet["A3"] = division_name.upper()
        # cover_sheet["A3"] = division_name.upper()

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

        # DESIGN BASIS SHEET ################################################################################################################################
        project_info_data = frappe.db.get_list(
            "Project Information", {"proejct_id": project_id}, ["*"]
        )
        project_info_data = project_info_data[0]

        main_supply_lv = project_info_data.get("main_supply_lv")
        main_supply_lv_variation = project_info_data.get("main_supply_lv_variation")
        main_supply_lv_phase = project_info_data.get("main_supply_lv_phase")
        lv_data = (
            f"{main_supply_lv}, {main_supply_lv_variation}, {main_supply_lv_phase}"
        )
        if main_supply_lv == "NA":
            lv_data = "Not Applicable"

        main_supply_mv = project_info_data.get("main_supply_mv")
        main_supply_mv_variation = project_info_data.get("main_supply_mv_variation")
        main_supply_mv_phase = project_info_data.get("main_supply_mv_phase")
        mv_data = (
            f"{main_supply_mv}, {main_supply_mv_variation}, {main_supply_mv_phase}"
        )

        if main_supply_mv == "NA":
            mv_data = "Not Applicable"

        control_supply = project_info_data.get("control_supply")
        control_supply_variation = project_info_data.get("control_supply_variation")
        control_supply_phase = project_info_data.get("control_supply_phase")
        control_supply_data = f"{control_supply}, Variation: {control_supply_variation}, {control_supply_phase}"
        if control_supply_variation == "NA":
            control_supply_data = control_supply

        utility_supply = project_info_data.get("utility_supply")
        utility_supply_variation = project_info_data.get("utility_supply_variation")
        utility_supply_phase = project_info_data.get("utility_supply_phase")
        utility_supply_data = f"{utility_supply}, Variation: {utility_supply_variation}, {utility_supply_phase}"
        if utility_supply_variation == "NA":
            utility_supply_data = utility_supply

        project_info_freq = project_info_data.get("frequency")
        preojct_info_freq_var = project_info_data.get("frequency_variation")
        project_info_frequency_data = (
            f"{project_info_freq} Hz , Variation: {preojct_info_freq_var}"
        )

        project_info_fault = project_info_data.get("fault_level")
        project_info_sec = project_info_data.get("sec")
        fault_data = f"{project_info_fault} KA, {project_info_sec} Sec"

        ambient_temperature_max = project_info_data.get("ambient_temperature_max")
        ambient_temperature_min = project_info_data.get("ambient_temperature_min")
        electrical_design_temperature = project_info_data.get(
            "electrical_design_temperature"
        )
        seismic_zone = project_info_data("seismic_zone")
        min_humidity = project_info_data("min_humidity")
        max_humidity = project_info_data("max_humidity")
        avg_humidity = project_info_data("avg_humidity")
        performance_humidity = project_info_data("performance_humidity")
        altitude = project_info_data("altitude")

        general_info_data = frappe.db.get_list(
            "Design Basis General Info", {"revision_id": revision_id}, "*"
        )
        general_info_data = general_info_data[0]
        battery_limit = general_info_data.get("battery_limit")

        main_packages_data = frappe.db.get_list(
            "Project Main Package",
            fields=["*"],
            filters={"revision_id": revision_id},
            order_by="creation asc",
        )

        for main_package in main_packages_data:
            # Get all Sub Package records
            sub_packages = frappe.get_doc(
                "Project Main Package", main_package["name"]
            ).sub_packages
            main_package["sub_packages"] = sub_packages

        design_basis_sheet["C4"] = mv_data
        design_basis_sheet["C5"] = lv_data
        design_basis_sheet["C6"] = control_supply_data
        design_basis_sheet["C7"] = utility_supply_data
        design_basis_sheet["C8"] = project_info_frequency_data
        design_basis_sheet["C9"] = fault_data
        design_basis_sheet["C10"] = f"{ambient_temperature_max} Deg. C"
        design_basis_sheet["C11"] = f"{ambient_temperature_min} Deg. C"
        design_basis_sheet["C12"] = f"{electrical_design_temperature} Deg. C"
        design_basis_sheet["C13"] = seismic_zone
        design_basis_sheet["C14"] = f"{max_humidity} %"
        design_basis_sheet["C15"] = f"{min_humidity} %"
        design_basis_sheet["C16"] = f"{avg_humidity} %"
        design_basis_sheet["C17"] = f"{performance_humidity} %"
        design_basis_sheet["C18"] = f"{altitude} meters"

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

        safe_area_efficiency_level = motor_parameters_data.get(
            "safe_area_efficiency_level"
        )
        safe_area_insulation_class = motor_parameters_data.get(
            "safe_area_insulation_class"
        )
        safe_area_temperature_rise = motor_parameters_data.get(
            "safe_area_temperature_rise"
        )
        safe_area_enclosure_ip_rating = motor_parameters_data.get(
            "safe_area_enclosure_ip_rating"
        )
        safe_area_max_temperature = motor_parameters_data.get(
            "safe_area_max_temperature"
        )
        safe_area_min_temperature = motor_parameters_data.get(
            "safe_area_min_temperature"
        )
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

        is_hazardous_area_present = motor_parameters_data.get(
            "is_hazardous_area_present"
        )
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
        hazardous_area_thermister = motor_parameters_data.get(
            "hazardous_area_thermister"
        )
        hazardous_area_space_heater = motor_parameters_data.get(
            "hazardous_area_space_heater"
        )
        hazardous_area_certification = motor_parameters_data.get(
            "hazardous_area_certification"
        )
        hazardous_area_bearing_rtd = motor_parameters_data.get(
            "hazardous_area_bearing_rtd"
        )
        hazardous_area_winding_rtd = motor_parameters_data.get(
            "hazardous_area_winding_rtd"
        )
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
        design_basis_sheet["C42"] = safe_area_service_factor
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
        design_basis_sheet["D34"] = (
            f"{hazardous_area_terminal_box_ip_rating} kW & Above"
        )
        design_basis_sheet["D35"] = f"{hazardous_area_thermister} kW & Above"
        design_basis_sheet["D36"] = f"{hazardous_area_space_heater} kW & Above"
        design_basis_sheet["D37"] = f"{hazardous_area_certification} kW & Above"
        design_basis_sheet["D38"] = f"{hazardous_area_bearing_rtd} kW & Above"
        design_basis_sheet["D39"] = f"{hazardous_area_winding_rtd} kW & Above"
        design_basis_sheet["D40"] = hazardous_area_bearing_type
        design_basis_sheet["D41"] = hazardous_area_duty
        design_basis_sheet["D42"] = hazardous_area_service_factor
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
        preferred_panel_enclosure = make_of_components_data.get(
            "preferred_panel_enclosure"
        )
        preferred_vfdvsd = make_of_components_data.get("preferred_vfdvsd")
        preferred_soft_starter = make_of_components_data.get("preferred_soft_starter")
        preferred_plc = make_of_components_data.get("preferred_plc")
        # preferred_gland_make = make_of_components_data.get("preferred_gland_make") ,

        design_basis_sheet["C50"] = motor
        design_basis_sheet["C51"] = cable
        design_basis_sheet["C52"] = lv_switchgear
        design_basis_sheet["C53"] = panel_enclosure
        design_basis_sheet["C54"] = variable_frequency_speed_drive_vfd_vsd
        design_basis_sheet["C55"] = soft_starter
        design_basis_sheet["C56"] = plc

        design_basis_sheet["D50"] = preferred_motor
        design_basis_sheet["D51"] = preferred_cable
        design_basis_sheet["D52"] = preferred_lv_switchgear
        design_basis_sheet["D53"] = preferred_panel_enclosure
        design_basis_sheet["D54"] = preferred_vfdvsd
        design_basis_sheet["D55"] = preferred_soft_starter
        design_basis_sheet["D56"] = preferred_plc

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
        control_transformer_coating = common_config_data.get(
            "control_transformer_coating"
        )
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

        design_basis_sheet["71"] = f"{ammeter} kW & Above"
        design_basis_sheet["72"] = ammeter_configuration
        design_basis_sheet["73"] = analog_meters
        design_basis_sheet["74"] = digital_meters
        design_basis_sheet["75"] = communication_protocol

        current_transformer = common_config_data.get("current_transformer")
        current_transformer_coating = common_config_data.get(
            "current_transformer_coating"
        )
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

        design_basis_sheet["84"] = dm_standard

        power_wiring_color = common_config_data.get("power_wiring_color")
        power_wiring_size = common_config_data.get("power_wiring_size")
        control_wiring_color = common_config_data.get("control_wiring_color")
        control_wiring_size = common_config_data.get("control_wiring_size")
        vdc_24_wiring_color = common_config_data.get("vdc_24_wiring_color")
        vdc_24_wiring_size = common_config_data.get("vdc_24_wiring_size")
        analog_signal_wiring_color = common_config_data.get(
            "analog_signal_wiring_color"
        )
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
        power_terminal_busbar_type = common_config_data.get(
            "power_terminal_busbar_type"
        )
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

        selector_switch_applicable = common_config_data.get(
            "selector_switch_applicable"
        )
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
        safe_field_motor_enclosure = common_config_data.get(
            "safe_field_motor_enclosure"
        )
        safe_field_motor_material = common_config_data.get("safe_field_motor_material")
        # safe_field_motor_thickness = common_config_data.get("safe_field_motor_thickness")
        safe_field_motor_qty = common_config_data.get("safe_field_motor_qty")
        safe_field_motor_isolator_color_shade = common_config_data.get(
            "safe_field_motor_isolator_color_shade"
        )
        safe_field_motor_cable_entry = common_config_data.get(
            "safe_field_motor_cable_entry"
        )
        safe_field_motor_canopy = common_config_data.get("safe_field_motor_canopy")
        # safe_field_motor_canopy_type = common_config_data.get("safe_field_motor_canopy_type")

        hazardous_field_motor_type = common_config_data.get(
            "hazardous_field_motor_type"
        )
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
        hazardous_lpbs_color_shade = common_config_data.get(
            "hazardous_lpbs_color_shade"
        )
        hazardous_lpbs_canopy = common_config_data.get("hazardous_lpbs_canopy")
        hazardous_lpbs_canopy_type = common_config_data.get(
            "hazardous_lpbs_canopy_type"
        )
        hazardous_lpbs_thickness = common_config_data.get("hazardous_lpbs_thickness")

        design_basis_sheet["D145"] = hazardous_lpbs_type
        design_basis_sheet["D146"] = hazardous_lpbs_enclosure
        design_basis_sheet["D147"] = hazardous_lpbs_material
        design_basis_sheet["D148"] = hazardous_lpbs_qty
        design_basis_sheet["D149"] = hazardous_lpbs_color_shade
        design_basis_sheet["D150"] = hazardous_lpbs_canopy
        design_basis_sheet["D151"] = hazardous_lpbs_canopy_type

        apfc_relay = common_config_data.get("apfc_relay")
        design_basis_sheet["D153"] = f"{apfc_relay} Stage"

        power_bus_main_busbar_selection = common_config_data.get(
            "power_bus_main_busbar_selection"
        )
        power_bus_heat_pvc_sleeve = common_config_data.get("power_bus_heat_pvc_sleeve")
        power_bus_material = common_config_data.get("power_bus_material")
        power_bus_current_density = common_config_data.get("power_bus_current_density")
        power_bus_rating_of_busbar = common_config_data.get(
            "power_bus_rating_of_busbar"
        )

        design_basis_sheet["C155"] = power_bus_main_busbar_selection
        design_basis_sheet["C156"] = power_bus_heat_pvc_sleeve
        design_basis_sheet["C157"] = power_bus_material
        design_basis_sheet["C158"] = power_bus_current_density
        design_basis_sheet["C159"] = power_bus_rating_of_busbar

        control_bus_main_busbar_selection = common_config_data.get(
            "control_bus_main_busbar_selection"
        )
        control_bus_heat_pvc_sleeve = common_config_data.get(
            "control_bus_heat_pvc_sleeve"
        )
        control_bus_material = common_config_data.get("control_bus_material")
        control_bus_current_density = common_config_data.get(
            "control_bus_current_density"
        )
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
        earth_bus_rating_of_busbar = common_config_data.get(
            "earth_bus_rating_of_busbar"
        )
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
        device_identification_of_components = common_config_data.get(
            "device_identification_of_components"
        )

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
        two_year_operational_spare = common_config_data.get(
            "two_year_operational_spare"
        )

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

        design_basis_sheet["183"] = number_of_cores
        design_basis_sheet["184"] = specific_requirement
        design_basis_sheet["185"] = type_of_insulation
        design_basis_sheet["186"] = color_scheme
        design_basis_sheet["187"] = f"{motor_voltage_drop_during_starting} %"
        design_basis_sheet["188"] = f"{motor_voltage_drop_during_running} %"
        design_basis_sheet["189"] = voltage_grade
        design_basis_sheet["190"] = f"{copper_conductor} Sq. mm"
        design_basis_sheet["191"] = f"{aluminium_conductor} Sq. mm"
        design_basis_sheet["192"] = touching_factor_air
        design_basis_sheet["193"] = ambient_temp_factor_air
        design_basis_sheet["194"] = derating_factor_air
        design_basis_sheet["195"] = touching_factor_burid
        design_basis_sheet["196"] = ambient_temp_factor_burid
        design_basis_sheet["197"] = derating_factor_burid

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
        ###############################################################################################################

        # Load the workbook from the template path
        # template_workbook = load_workbook(template_path)
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

    except Exception as e:
        # Log the error for debugging purposes
        frappe.log_error(message=str(e), title="Error in get_design_basis_excel")

        # Return a user-friendly error message
        return _("An error occurred while generating the file: {0}").format(str(e))
