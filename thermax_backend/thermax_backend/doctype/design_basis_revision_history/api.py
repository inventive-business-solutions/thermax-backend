import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy


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
    payload = frappe.local.form_dict
    metadata = payload.get("metadata")
    project = payload.get("project")
    project_info = payload.get("projectInfo")
    general_info = payload.get("generalInfo")
    motor_parameters = payload.get("motorParameters")
    make_of_components = payload.get("makeOfComponents")
    common_configuration = payload.get("commonConfigurations")
    cable_tray_layout = payload.get("cableTrayLayoutData")
    earthing_layout_data = payload.get("earthingLayoutData")

    project_panels = payload.get("projectPanelData")

    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "design_basis_template.xlsx"
    )
    template_workbook = load_workbook(template_path)

    cover_sheet = template_workbook["COVER"]
    design_basis_sheet = template_workbook["Design Basis"]
    revision_sheet = template_workbook["REVISION"]
    mcc_sheet = template_workbook["MCC"]
    pcc_sheet = template_workbook["PCC"]
    mcc_cum_plc_sheet = template_workbook["MCC CUM PLC"]

    # Cover Sheet

    cover_sheet["A3"] = metadata.get("division_name")
    cover_sheet["D7"] = project.get("client_name")
    cover_sheet["D8"] = project.get("consultant_name")
    cover_sheet["D9"] = project.get("project_name")
    cover_sheet["D10"] = project.get("project_oc_number")

    cover_sheet["E36"] = metadata.get("prepared_by_initials")
    cover_sheet["F36"] = metadata.get("checked_by_initials")
    cover_sheet["G36"] = metadata.get("approved_by_initials")

    # Revision Sheet

    """
        Design Basis Sheet
    """
    # General Information

    design_basis_sheet["C8"] = general_info.get("battery_limit")
    design_basis_sheet["C9"] = project_info.get("main_supply_mv")
    design_basis_sheet["C10"] = project_info.get("main_supply_lv")
    design_basis_sheet["C11"] = project_info.get("control_supply")
    design_basis_sheet["C12"] = project_info.get("utility_supply")
    design_basis_sheet["C13"] = project_info.get("frequency")
    design_basis_sheet["C14"] = project_info.get("fault_level")
    design_basis_sheet["C15"] = project_info.get("ambient_temperature_max")
    design_basis_sheet["C16"] = project_info.get("ambient_temperature_min")
    design_basis_sheet["C17"] = project_info.get("electrical_design_temperature")
    design_basis_sheet["C18"] = project_info.get("seismic_zone")

    """
        Motor Details
    """
    # Safe Area
    design_basis_sheet["E21"] = motor_parameters.get("safe_area_efficiency_level")
    design_basis_sheet["E22"] = motor_parameters.get("safe_area_insulation_class")
    design_basis_sheet["E23"] = motor_parameters.get("safe_area_temperature_rise")
    design_basis_sheet["E24"] = motor_parameters.get("safe_area_enclosure_ip_rating")
    design_basis_sheet["E25"] = motor_parameters.get("safe_area_max_temperature")
    design_basis_sheet["E26"] = motor_parameters.get("safe_area_min_temperature")
    design_basis_sheet["E27"] = motor_parameters.get("safe_area_altitude")
    design_basis_sheet["E28"] = motor_parameters.get("safe_area_terminal_box_ip_rating")
    design_basis_sheet["E29"] = motor_parameters.get("safe_area_thermister")
    design_basis_sheet["E30"] = motor_parameters.get("safe_area_space_heater")
    design_basis_sheet["E31"] = motor_parameters.get("hazardous_area_certification")
    design_basis_sheet["E32"] = motor_parameters.get("safe_area_bearing_rtd")
    design_basis_sheet["E33"] = motor_parameters.get("safe_area_winding_rtd")
    design_basis_sheet["E34"] = motor_parameters.get("safe_area_bearing_type")
    design_basis_sheet["E35"] = motor_parameters.get("safe_area_duty")
    design_basis_sheet["E36"] = motor_parameters.get("safe_area_service_factor")
    design_basis_sheet["E37"] = motor_parameters.get("safe_area_cooling_type")
    design_basis_sheet["E38"] = motor_parameters.get("safe_area_body_material")
    design_basis_sheet["E39"] = motor_parameters.get("safe_area_terminal_box_material")
    design_basis_sheet["E40"] = motor_parameters.get("safe_area_paint_type_and_shade")
    design_basis_sheet["E41"] = motor_parameters.get(
        "safe_area_starts_hour_permissible"
    )

    # Hazardous Area
    design_basis_sheet["E43"] = motor_parameters.get("hazardous_area_efficiency_level")
    design_basis_sheet["E44"] = motor_parameters.get("hazardous_area_insulation_class")
    design_basis_sheet["E45"] = motor_parameters.get("hazardous_area_temperature_rise")
    design_basis_sheet["E46"] = motor_parameters.get(
        "hazardous_area_enclosure_ip_rating"
    )
    design_basis_sheet["E47"] = motor_parameters.get("hazardous_area_max_temperature")
    design_basis_sheet["E48"] = motor_parameters.get("hazardous_area_min_temperature")
    design_basis_sheet["E49"] = motor_parameters.get("hazardous_area_altitude")
    design_basis_sheet["E50"] = motor_parameters.get(
        "hazardous_area_terminal_box_ip_rating"
    )
    design_basis_sheet["E51"] = motor_parameters.get("hazardous_area_thermister")
    design_basis_sheet["E52"] = motor_parameters.get("hazardous_area_space_heater")
    design_basis_sheet["E53"] = motor_parameters.get("hazardous_area_certification")
    design_basis_sheet["E54"] = motor_parameters.get("hazardous_area_bearing_rtd")
    design_basis_sheet["E55"] = motor_parameters.get("hazardous_area_winding_rtd")
    design_basis_sheet["E56"] = motor_parameters.get("hazardous_area_bearing_type")
    design_basis_sheet["E57"] = motor_parameters.get("hazardous_area_duty")
    design_basis_sheet["E58"] = motor_parameters.get("hazardous_area_service_factor")
    design_basis_sheet["E59"] = motor_parameters.get("hazardous_area_cooling_type")
    design_basis_sheet["E60"] = motor_parameters.get("hazardous_area_body_material")
    design_basis_sheet["E61"] = motor_parameters.get(
        "hazardous_area_terminal_box_material"
    )
    design_basis_sheet["E62"] = motor_parameters.get(
        "hazardous_area_paint_type_and_shade"
    )
    design_basis_sheet["E63"] = motor_parameters.get(
        "hazardous_area_starts_hour_permissible"
    )

    """
        Make of Components
    """
    design_basis_sheet["E66"] = make_of_components.get("motor")
    design_basis_sheet["E67"] = make_of_components.get("cable")
    design_basis_sheet["E68"] = make_of_components.get("lv_switchgear")
    design_basis_sheet["E69"] = make_of_components.get("panel_enclosure")
    design_basis_sheet["E70"] = make_of_components.get(
        "variable_frequency_speed_drive_vfd_vsd"
    )
    design_basis_sheet["E71"] = make_of_components.get("soft_starter")
    design_basis_sheet["E72"] = make_of_components.get("plc")

    """
        Common Configuration
    """
    design_basis_sheet["E74"] = common_configuration.get("dol_starter")
    design_basis_sheet["E75"] = common_configuration.get("star_delta_starter")
    design_basis_sheet["E76"] = common_configuration.get("ammeter")
    design_basis_sheet["E77"] = common_configuration.get("ammeter_configuration")
    design_basis_sheet["E78"] = common_configuration.get("mcc_switchgear_type")
    design_basis_sheet["E79"] = common_configuration.get("switchgear_combination")
    design_basis_sheet["E80"] = common_configuration.get("pole")
    design_basis_sheet["E81"] = common_configuration.get("supply_feeder_standard")
    design_basis_sheet["E82"] = common_configuration.get("dm_standard")
    design_basis_sheet["E83"] = common_configuration.get("testing_standard")

    """
        Wiring
    """
    design_basis_sheet["E85"] = common_configuration.get("power_wiring_color")
    design_basis_sheet["E86"] = common_configuration.get("power_wiring_size")
    design_basis_sheet["E87"] = common_configuration.get("control_wiring_color")
    design_basis_sheet["E88"] = common_configuration.get("control_wiring_size")
    design_basis_sheet["E89"] = common_configuration.get("vdc_24_wiring_color")
    design_basis_sheet["E90"] = common_configuration.get("vdc_24_wiring_size")
    design_basis_sheet["E91"] = common_configuration.get("analog_signal_wiring_color")
    design_basis_sheet["E92"] = common_configuration.get("analog_signal_wiring_size")
    design_basis_sheet["E93"] = common_configuration.get("ct_wiring_color")
    design_basis_sheet["E94"] = common_configuration.get("ct_wiring_size")
    design_basis_sheet["E95"] = common_configuration.get("cable_insulation_pvc")
    design_basis_sheet["E96"] = common_configuration.get("ferrule")
    design_basis_sheet["E97"] = common_configuration.get("common_requirement")

    """
        Terminal
    """
    design_basis_sheet["E99"] = common_configuration.get("spare_terminal")

    """
        Push Button Color
    """
    design_basis_sheet["E101"] = common_configuration.get("push_button_start")
    design_basis_sheet["E102"] = common_configuration.get("push_button_stop")
    design_basis_sheet["E103"] = common_configuration.get("push_button_ess")
    design_basis_sheet["E104"] = common_configuration.get("speed_increase_pb")
    design_basis_sheet["E105"] = common_configuration.get("speed_decrease_pb")
    design_basis_sheet["E106"] = common_configuration.get(
        "alarm_acknowledge_and_lamp_test"
    )
    design_basis_sheet["E107"] = common_configuration.get("test_reset")

    """
        Selector Switch
    """
    design_basis_sheet["E109"] = common_configuration.get("selector_switch_applicable")

    """
        Indicating Lamp
    """
    design_basis_sheet["E111"] = common_configuration.get("running_open")
    design_basis_sheet["E112"] = common_configuration.get("stopped_closed")
    design_basis_sheet["E113"] = common_configuration.get("trip")

    """
        Field Motor Isolator(General Specifications)
    """
    design_basis_sheet["E115"] = common_configuration.get("field_motor_type")
    design_basis_sheet["E116"] = common_configuration.get("field_motor_enclosure")
    design_basis_sheet["E117"] = common_configuration.get("field_motor_material")
    design_basis_sheet["E118"] = common_configuration.get("field_motor_qty")
    design_basis_sheet["E119"] = common_configuration.get(
        "field_motor_isolator_color_shade"
    )
    design_basis_sheet["E120"] = common_configuration.get("field_motor_cable_entry")
    design_basis_sheet["E121"] = common_configuration.get("field_motor_canopy_on_top")

    """
        Local Push Button Station (General Specifications)				
    """
    design_basis_sheet["E123"] = common_configuration.get("lpbs_type")
    design_basis_sheet["E124"] = common_configuration.get("lpbs_enclosure")
    design_basis_sheet["E125"] = common_configuration.get("lpbs_material")
    design_basis_sheet["E126"] = common_configuration.get("lpbs_qty")
    design_basis_sheet["E127"] = common_configuration.get("lpbs_color_shade")
    design_basis_sheet["E128"] = common_configuration.get("lpbs_canopy_on_top")
    design_basis_sheet["E129"] = common_configuration.get(
        "lpbs_push_button_start_color"
    )
    design_basis_sheet["E130"] = common_configuration.get(
        "lpbs_indication_lamp_start_color"
    )
    design_basis_sheet["E131"] = common_configuration.get(
        "lpbs_indication_lamp_stop_color"
    )
    design_basis_sheet["E132"] = common_configuration.get("lpbs_speed_increase")
    design_basis_sheet["E133"] = common_configuration.get("lpbs_speed_decrease")

    """
        Power Bus
    """
    design_basis_sheet["E135"] = common_configuration.get(
        "power_bus_main_busbar_selection"
    )
    design_basis_sheet["E136"] = common_configuration.get("power_bus_heat_pvc_sleeve")
    design_basis_sheet["E137"] = common_configuration.get("power_bus_material")
    design_basis_sheet["E138"] = common_configuration.get("power_bus_current_density")
    design_basis_sheet["E139"] = common_configuration.get("power_bus_rating_of_busbar")

    """
        Control Bus
    """
    design_basis_sheet["E141"] = common_configuration.get(
        "control_bus_main_busbar_selection"
    )
    design_basis_sheet["E142"] = common_configuration.get("control_bus_heat_pvc_sleeve")
    design_basis_sheet["E143"] = common_configuration.get("control_bus_material")
    design_basis_sheet["E144"] = common_configuration.get("control_bus_current_density")
    design_basis_sheet["E145"] = common_configuration.get(
        "control_bus_rating_of_busbar"
    )

    """
        Earth Bus
    """
    design_basis_sheet["E147"] = common_configuration.get(
        "earth_bus_main_busbar_selection"
    )
    design_basis_sheet["E148"] = common_configuration.get("earth_bus_heat_pvc_sleeve")
    design_basis_sheet["E149"] = common_configuration.get("earth_bus_material")
    design_basis_sheet["E150"] = common_configuration.get("earth_bus_current_density")
    design_basis_sheet["E151"] = common_configuration.get("earth_bus_rating_of_busbar")

    """
        Metering for Feeder
    """
    design_basis_sheet["E153"] = common_configuration.get("metering_for_feeders")

    """
        Others
    """
    design_basis_sheet["E155"] = common_configuration.get("cooling_fans")
    design_basis_sheet["E156"] = common_configuration.get("louvers_and_filters")
    design_basis_sheet["E157"] = common_configuration.get("alarm_annunciator")

    """
        Spares
    """
    design_basis_sheet["E159"] = common_configuration.get("commissioning_spare")
    design_basis_sheet["E160"] = common_configuration.get("two_year_operational_spare")

    """
        APFC
    """
    design_basis_sheet["E162"] = common_configuration.get("apfc_relay")

    """
        Power Cable
    """
    design_basis_sheet["E164"] = cable_tray_layout.get("number_of_cores")
    design_basis_sheet["E165"] = cable_tray_layout.get("specific_requirement")
    design_basis_sheet["E166"] = cable_tray_layout.get("type_of_insulation")
    design_basis_sheet["E167"] = cable_tray_layout.get("color_scheme")
    design_basis_sheet["E168"] = cable_tray_layout.get(
        "motor_voltage_drop_during_starting"
    )
    design_basis_sheet["E169"] = cable_tray_layout.get(
        "motor_voltage_drop_during_running"
    )
    design_basis_sheet["E170"] = cable_tray_layout.get("copper_conductor")
    design_basis_sheet["E171"] = cable_tray_layout.get("aluminiun_conductor")
    design_basis_sheet["E172"] = cable_tray_layout.get("voltage_grade")
    design_basis_sheet["E173"] = cable_tray_layout.get("touching_factor_for_air")
    design_basis_sheet["E174"] = cable_tray_layout.get(
        "ambient_temperature_factor_for_air"
    )
    design_basis_sheet["E175"] = cable_tray_layout.get("derating_factor_for_air")
    design_basis_sheet["E176"] = cable_tray_layout.get("touching_factor_for_buried")
    design_basis_sheet["E177"] = cable_tray_layout.get(
        "ambient_temperature_factor_for_buried"
    )
    design_basis_sheet["E178"] = cable_tray_layout.get("derating_factor_for_buried")

    """
        Gland
    """
    design_basis_sheet["E180"] = cable_tray_layout.get("gland_make")
    design_basis_sheet["E181"] = cable_tray_layout.get("moc")
    design_basis_sheet["E182"] = cable_tray_layout.get("type_of_gland")
    design_basis_sheet["E183"] = cable_tray_layout.get("safe_area_gland_type")
    design_basis_sheet["E184"] = cable_tray_layout.get("hazardous_area_gland_type")

    """
        Cable Tray
    """
    design_basis_sheet["E186"] = cable_tray_layout.get("future_space_on_trays")
    design_basis_sheet["E187"] = cable_tray_layout.get("cable_placement")
    design_basis_sheet["E188"] = cable_tray_layout.get("orientation")
    design_basis_sheet["E189"] = cable_tray_layout.get("vertical_distance")
    design_basis_sheet["E190"] = cable_tray_layout.get("horizontal_distance")
    design_basis_sheet["E191"] = cable_tray_layout.get("dry_area")
    design_basis_sheet["E192"] = cable_tray_layout.get("wet_area")

    """
        Earthing
    """
    design_basis_sheet["E194"] = earthing_layout_data.get("earthing_system")
    design_basis_sheet["E195"] = earthing_layout_data.get("earth_strip")
    design_basis_sheet["E196"] = earthing_layout_data.get("earth_pit")
    design_basis_sheet["E197"] = earthing_layout_data.get("soil_resistivity")

    for project_panel in project_panels:
        if project_panel.get("panel_main_type") == "MCC":
            panel_sheet = template_workbook.copy_worksheet(mcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            panel_data = project_panel.get("panelData")
            """
                Selection Details
            """
            panel_sheet["E5"] = (
                f"Upto - {panel_data.get('incomer_ampere')} - {panel_data.get('incomer_pole')} Pole {panel_data.get('incomer_type')} > {panel_data.get('incomer_above_ampere')} - {panel_data.get('incomer_above_pole')} Pole {panel_data.get('incomer_above_type')}"
            )
            panel_sheet["E6"] = panel_data.get("led_type_indication_lamp")
            panel_sheet["E7"] = panel_data.get("current_transformer_coating")
            panel_sheet["E8"] = panel_data.get("current_transformer_number")
            panel_sheet["E9"] = panel_data.get("control_transformer_coating")
            panel_sheet["E10"] = panel_data.get("control_transformer_configuration")
            panel_sheet["E11"] = panel_data.get("alarm_annunciator")

            """
                Metering Instruments for Incomer				
            """
            panel_sheet["E13"] = f"Analog - {panel_data.get("mi_analog")} ; Digital - { panel_data.get("mi_digital")} ; Communication Protocol - { panel_data.get("mi_communication_protocol") }"

            """
                General Arrangement				
            """
            panel_sheet["E15"] = panel_data.get("ga_moc_material")
            panel_sheet["E16"] = panel_data.get("ga_moc_thickness_door")
            panel_sheet["E17"] = panel_data.get("ga_moc_thickness_covers")
            panel_sheet["E18"] = f"{panel_data.get("ga_mcc_compartmental"), {panel_data.get("ga_mcc_construction_front_type")}, {panel_data.get("ga_mcc_construction_drawout_type")}, {panel_data.get("ga_mcc_construction_type")}}"
            panel_sheet["E19"] = panel_data.get("busbar_material_of_construction")
            panel_sheet["E20"] = panel_data.get("ga_current_density")
            panel_sheet["E21"] = panel_data.get("ga_panel_mounting_frame")
            panel_sheet["E22"] = panel_data.get("ga_panel_mounting_height")
            panel_sheet["E23"] = panel_data.get("is_marshalling_section_selected")
            panel_sheet["E24"] = panel_data.get("is_cable_alley_section_selected")
            panel_sheet["E25"] = panel_data.get("is_power_and_bus_separation_section_selected")
            panel_sheet["E26"] = panel_data.get("is_both_side_extension_section_selected")
            panel_sheet["E27"] = panel_data.get("ga_gland_plate_3mm_drill_type")
            panel_sheet["E28"] = panel_data.get("ga_gland_plate_3mm_attachment_type")
            panel_sheet["E29"] = panel_data.get("ga_busbar_chamber_position")
            panel_sheet["E30"] = panel_data.get("ga_power_and_control_busbar_separation")
            panel_sheet["E31"] = panel_data.get("ga_enclosure_protection_degree")
            panel_sheet["E32"] = panel_data.get("ga_cable_entry_position")

            """
                Painting / Powder Coating			
            """
            panel_sheet["E34"] = panel_data.get("ppc_painting_standards")
            panel_sheet["E35"] = panel_data.get("ppc_interior_and_exterior_paint_shade")
            panel_sheet["E36"] = panel_data.get("ppc_component_mounting_plate_paint_shade")
            panel_sheet["E37"] = panel_data.get("ppc_minimum_coating_thickness")
            panel_sheet["E38"] = panel_data.get("ppc_base_frame_paint_shade")
            panel_sheet["E39"] = panel_data.get("ppc_pretreatment_panel_standard")

            """
                VFD
            """
            panel_sheet["E41"] = panel_data.get("vfd_auto_manual_selection")

            """
                Punching Details
            """

            # Punching Details for Boiler
            panel_sheet["E44"] = panel_data.get("boiler_model")
            panel_sheet["E45"] = panel_data.get("boiler_fuel")
            panel_sheet["E46"] = panel_data.get("boiler_year")
            panel_sheet["E47"] = f"{panel_data.get("boiler_power_supply_vac")} VAC {panel_data.get("boiler_power_supply_phase")} Phase {panel_data.get("boiler_power_supply_frequency")} Hz"
            panel_sheet["E48"] = f"{panel_data.get("boiler_control_supply_vac")} {panel_data.get("boiler_control_supply_phase")} {panel_data.get("boiler_control_supply_frequency")}"
            panel_sheet["E49"] = panel_data.get("boiler_evaporation")
            panel_sheet["E50"] = panel_data.get("boiler_output")
            panel_sheet["E51"] = panel_data.get("boiler_connected_load")
            panel_sheet["E52"] = panel_data.get("boiler_design_pressure")

            # Punching Details for Heater
            panel_sheet["E54"] = panel_data.get("heater_model")
            panel_sheet["E55"] = panel_data.get("heater_fuel")
            panel_sheet["E56"] = panel_data.get("heater_year")
            panel_sheet["E57"] = f"{panel_data.get("heater_power_supply_vac")} VAC {panel_data.get("heater_power_supply_phase")} Phase {panel_data.get("heater_power_supply_frequency")} Hz"
            panel_sheet["E58"] = f"{panel_data.get("heater_control_supply_vac")} {panel_data.get("heater_control_supply_phase")} {panel_data.get("heater_control_supply_frequency")}"
            panel_sheet["E59"] = panel_data.get("heater_evaporation")
            panel_sheet["E60"] = panel_data.get("heater_output")
            panel_sheet["E61"] = panel_data.get("heater_connected_load")
            panel_sheet["E62"] = panel_data.get("heater_temperature")

            # Name Plate Details for SPG
            panel_sheet["E64"] = panel_data.get("spg_name_plate_unit_name")
            panel_sheet["E65"] = panel_data.get("spg_name_plate_capacity")
            panel_sheet["E66"] = panel_data.get("spg_name_plate_manufacturing_year")
            panel_sheet["E67"] = panel_data.get("spg_name_plate_weight")
            panel_sheet["E68"] = panel_data.get("spg_name_plate_oc_number")
            panel_sheet["E69"] = panel_data.get("spg_name_plate_part_code")


        if project_panel.get("panel_main_type") == "PCC":
            panel_sheet = template_workbook.copy_worksheet(pcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            panel_data = project_panel.get("panelData")

            """
                Selection Details
            """
            panel_sheet["E5"] = f"Upto - {panel_data.get('incomer_ampere')} - {panel_data.get('incomer_pole')} Pole {panel_data.get('incomer_type')} > {panel_data.get('incomer_above_ampere')} - {panel_data.get('incomer_above_pole')} Pole {panel_data.get('incomer_above_type')}"
            panel_sheet["E6"] = panel_data.get("led_type_indication_lamp")
            panel_sheet["E7"] = panel_data.get("control_transformer_coating")
            panel_sheet["E8"] = panel_data.get("control_transformer_configuration")
            panel_sheet["E9"] = panel_data.get("alarm_annunciator")

            """
                Metering Instruments for Incomer				
            """
            panel_sheet["E11"] = f"Analog - {panel_data.get('mi_analog')} ; Digital - { panel_data.get('mi_digital')} ; Communication Protocol - { panel_data.get('mi_communication_protocol') }"

            """
                General Arrangement
            """
            panel_sheet["E13"] = panel_data.get("ga_moc_material")
            panel_sheet["E14"] = panel_data.get("ga_moc_thickness_door")
            panel_sheet["E15"] = panel_data.get("ga_moc_thickness_covers")
            panel_sheet["E16"] = f"{panel_data.get('ga_mcc_compartmental'), {panel_data.get('ga_mcc_construction_front_type')}, {panel_data.get('ga_mcc_construction_drawout_type')}, {panel_data.get('ga_mcc_construction_type')}}"
            panel_sheet["E17"] = panel_data.get("busbar_material_of_construction")
            panel_sheet["E18"] = panel_data.get("ga_current_density")
            panel_sheet["E19"] = panel_data.get("ga_panel_mounting_frame")
            panel_sheet["E20"] = panel_data.get("ga_panel_mounting_height")
            panel_sheet["E21"] = panel_data.get("is_marshalling_section_selected")
            panel_sheet["E22"] = panel_data.get("is_cable_alley_section_selected")
            panel_sheet["E23"] = panel_data.get("is_power_and_bus_separation_section_selected")
            panel_sheet["E24"] = panel_data.get("is_both_side_extension_section_selected")
            panel_sheet["E25"] = panel_data.get("ga_gland_plate_3mm_drill_type")
            panel_sheet["E26"] = panel_data.get("ga_gland_plate_3mm_attachment_type")
            panel_sheet["E27"] = panel_data.get("ga_busbar_chamber_position")
            panel_sheet["E28"] = panel_data.get("ga_power_and_control_busbar_separation")
            panel_sheet["E29"] = panel_data.get("ga_enclosure_protection_degree")
            panel_sheet["E30"] = panel_data.get("ga_cable_entry_position")

            """
                Painting / Powder Coating
            """
            panel_sheet["E32"] = panel_data.get("ppc_painting_standards")
            panel_sheet["E33"] = panel_data.get("ppc_interior_and_exterior_paint_shade")
            panel_sheet["E34"] = panel_data.get("ppc_component_mounting_plate_paint_shade")
            panel_sheet["E35"] = panel_data.get("ppc_minimum_coating_thickness")
            panel_sheet["E36"] = panel_data.get("ppc_base_frame_paint_shade")
            panel_sheet["E37"] = panel_data.get("ppc_pretreatment_panel_standard")

            """
                Punching Details
            """
            # Punching Details for Boiler
            panel_sheet["E40"] = panel_data.get("boiler_model")
            panel_sheet["E41"] = panel_data.get("boiler_fuel")
            panel_sheet["E42"] = panel_data.get("boiler_year")
            panel_sheet["E43"] = f"{panel_data.get('boiler_power_supply_vac')} VAC {panel_data.get('boiler_power_supply_phase')} Phase {panel_data.get('boiler_power_supply_frequency')} Hz"
            panel_sheet["E44"] = f"{panel_data.get('boiler_control_supply_vac')} {panel_data.get('boiler_control_supply_phase')} {panel_data.get('boiler_control_supply_frequency')}"
            panel_sheet["E45"] = panel_data.get("boiler_evaporation")
            panel_sheet["E46"] = panel_data.get("boiler_output")
            panel_sheet["E47"] = panel_data.get("boiler_connected_load")
            panel_sheet["E48"] = panel_data.get("boiler_design_pressure")

            # Punching Details for Heater
            panel_sheet["E50"] = panel_data.get("heater_model")
            panel_sheet["E51"] = panel_data.get("heater_fuel")
            panel_sheet["E52"] = panel_data.get("heater_year")
            panel_sheet["E53"] = f"{panel_data.get('heater_power_supply_vac')} VAC {panel_data.get('heater_power_supply_phase')} Phase {panel_data.get('heater_power_supply_frequency')} Hz"
            panel_sheet["E54"] = f"{panel_data.get('heater_control_supply_vac')} {panel_data.get('heater_control_supply_phase')} {panel_data.get('heater_control_supply_frequency')}"
            panel_sheet["E55"] = panel_data.get("heater_evaporation")
            panel_sheet["E56"] = panel_data.get("heater_output")
            panel_sheet["E57"] = panel_data.get("heater_connected_load")
            panel_sheet["E58"] = panel_data.get("heater_temperature")

            """
                Name Plate Details for SPG
            """
            panel_sheet["E60"] = panel_data.get("spg_name_plate_unit_name")
            panel_sheet["E61"] = panel_data.get("spg_name_plate_capacity")
            panel_sheet["E62"] = panel_data.get("spg_name_plate_manufacturing_year")
            panel_sheet["E63"] = panel_data.get("spg_name_plate_weight")
            panel_sheet["E64"] = panel_data.get("spg_name_plate_oc_number")
            panel_sheet["E65"] = panel_data.get("spg_name_plate_part_code")
            


        if project_panel.get("panel_main_type") == "MCC cum PCC":
            panel_sheet = template_workbook.copy_worksheet(mcc_cum_plc_sheet)
            panel_sheet.title = project_panel.get("panel_name")

    template_workbook.remove(mcc_sheet)
    template_workbook.remove(pcc_sheet)
    template_workbook.remove(mcc_cum_plc_sheet)

    excel_save_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "generated_design_basis.xlsx"
    )

    template_workbook.save(excel_save_path)
    return metadata
