import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
from collections import defaultdict
import io
from datetime import datetime

# revision_id = "st486uu99i"

def num_to_string(value):

    if int(value) == 0:
        return "Not Applicable"

    return "Applicable"

def na_to_string(value):

    if value == "NA":
        return "Not Applicable"

    return "Applicable"


@frappe.whitelist()
def get_panel_specification_excel():
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    panel_spec_revision_data = frappe.get_doc(
        "Panel Specifications Revisions", revision_id, "*"
    ).as_dict()

    project_id = panel_spec_revision_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", {"project_id": project_id}
    ).as_dict()

    project_revision_id = design_basis_revision_data.get("name")

    # Loading the workbook
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "power_cum_plc_panel_specification_template.xlsx"
    )
    template_workbook = load_workbook(template_path)


    cover_sheet = template_workbook["COVER"]
    mcc_cum_plc_panel_sheet = template_workbook["MCC CUM PLC PANEL"]
    plc_specification_sheet = template_workbook["PLC SPECIFICATION"]

    dynamic_document_list_data = frappe.get_doc(
        "Dynamic Document List",
        project_id,
        "*"
    ).as_dict()

    static_document_list_data = frappe.get_doc(
        "Static Document List",
        project_id,
        "*"
    ).as_dict()

    project_panel_data = frappe.db.get_list(
        "Project Panel Data", {"revision_id": project_revision_id}, "*", order_by="creation asc"
    )

    project_info_data = frappe.get_doc(
        "Project Information",
        project_id,
        "*"
    ).as_dict()

    
    for project_panel in project_panel_data:
        panel_id = project_panel.get("name")

        if project_panel.get("type") != "PCC" and project_panel.get("type") != "MCC":
            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"panel_id": panel_id}, "*"
            )

            if len(mcc_panel_data) == 0:
                continue
            mcc_panel_data = mcc_panel_data[0]

            plc_panel_1 = frappe.db.get_list(
                "Panel PLC 1 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_1 = plc_panel_1[0] if len(plc_panel_1) > 0 else {}
            plc_panel_2 = frappe.db.get_list(
                "Panel PLC 2 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_2 = plc_panel_2[0] if len(plc_panel_2) > 0 else {}
            plc_panel_3 = frappe.db.get_list(
                "Panel PLC 3 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_3 = plc_panel_3[0] if len(plc_panel_3) > 0 else {}

            plc_panel = {**plc_panel_1, **plc_panel_2, **plc_panel_3}
            # PLC fields

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

            
            # PLC SPECIFICATION SHEET 

            plc_specification_sheet["C9"] = "TBD"
            plc_specification_sheet["C10"] = "TBD"
            plc_specification_sheet["C11"] = "TBD"


            plc_specification_sheet["C13"] = plc_panel.get("ups_control_voltage", "NA")
            plc_specification_sheet["C14"] = plc_panel.get("non_ups_control_voltage", "NA")
            freq_data = f"{project_info_data.get("frequency")}, {project_info_data.get("frequency_variation")}"
            plc_specification_sheet["C15"] = freq_data


            plc_specification_sheet["C17"] = project_info_data.get("project_location")
            plc_specification_sheet["C18"] = project_info_data.get("ambient_temperature_max")
            plc_specification_sheet["C19"] = project_info_data.get("ambient_temperature_min")
            plc_specification_sheet["C20"] = project_info_data.get("electrical_design_temperature")
            plc_specification_sheet["C21"] = project_info_data.get("seismic_zone")
            plc_specification_sheet["C22"] = project_info_data.get("electrical_design_temperature")

            plc_specification_sheet["C24"] = common_config_data.get("supply_feeder_standard")
            plc_specification_sheet["C25"] = common_config_data.get("supply_feeder_standard")
            plc_specification_sheet["C26"] = mcc_panel_data.get("ga_mcc_construction_front_type")
            plc_specification_sheet["C27"] = mcc_panel_data.get("ga_mcc_compartmental")
            plc_specification_sheet["C28"] = mcc_panel_data.get("ga_panel_mounting_frame")
            plc_specification_sheet["C29"] = mcc_panel_data.get("ga_panel_mounting_height")
            # plc_specification_sheet["C30"] = mcc_panel_data.get("ga_panel_mounting_height")
            
            plc_specification_sheet["C32"] = mcc_panel_data.get("ga_moc_thickness_door")
            plc_specification_sheet["C33"] = mcc_panel_data.get("door_thickness")
            plc_specification_sheet["C34"] = mcc_panel_data.get("ga_moc_thickness_covers")
            plc_specification_sheet["C35"] = mcc_panel_data.get("ga_gland_plate_thickness")
            plc_specification_sheet["C36"] = mcc_panel_data.get("ga_cable_entry_position")
            plc_specification_sheet["C37"] = mcc_panel_data.get("ga_busbar_chamber_position")

            plc_specification_sheet["C40"] = mcc_panel_data.get("ppc_interior_and_exterior_paint_shade")
            plc_specification_sheet["C41"] = mcc_panel_data.get("ppc_component_mounting_plate_paint_shade")
            # plc_specification_sheet["C42"] = mcc_panel_data.get("ppc_interior_and_exterior_paint_shade")
            plc_specification_sheet["C43"] = mcc_panel_data.get("ppc_minimum_coating_thickness")
            plc_specification_sheet["C44"] = mcc_panel_data.get("ppc_pretreatment_panel_standard")
            plc_specification_sheet["C45"] = mcc_panel_data.get("general_requirments_for_construction")

            plc_specification_sheet["C47"] = common_config_data.get("power_wiring_color")
            plc_specification_sheet["C48"] = common_config_data.get("power_wiring_size")
            plc_specification_sheet["C49"] = common_config_data.get("control_wiring_color")
            plc_specification_sheet["C50"] = common_config_data.get("control_wiring_size")
            plc_specification_sheet["C51"] = common_config_data.get("vdc_24_wiring_color")
            plc_specification_sheet["C52"] = common_config_data.get("vdc_24_wiring_size")
            plc_specification_sheet["C53"] = common_config_data.get("analog_signal_wiring_color")
            plc_specification_sheet["C54"] = common_config_data.get("analog_signal_wiring_size")
            plc_specification_sheet["C55"] = common_config_data.get("rtd_thermocouple_wiring_color")
            plc_specification_sheet["C56"] = common_config_data.get("rtd_thermocouple_wiring_size")
            plc_specification_sheet["C57"] = common_config_data.get("cable_insulation_pvc")
            plc_specification_sheet["C58"] = common_config_data.get("general_note_internal_wiring")


            # plc_specification_sheet["C61"] = plc_panel.get("general_note_internal_wiring")
            hooter_data = plc_panel.get("is_electronic_hooter_selected")
            plc_specification_sheet["C62"] = num_to_string(hooter_data)
            hooter_acknowledge_data = plc_panel.get("electronic_hooter_acknowledge")
            plc_specification_sheet["C63"] = na_to_string(hooter_acknowledge_data)

            plc_specification_sheet["C65"] = plc_panel.get("panel_power_supply_on_color")
            plc_specification_sheet["C66"] = plc_panel.get("panel_power_supply_off_color")
    
            plc_specification_sheet["C68"] = common_config_data.get("power_terminal_clipon")
            plc_specification_sheet["C69"] = common_config_data.get("power_terminal_busbar_type")
            plc_specification_sheet["C70"] = plc_panel.get("di_module_terminal")
            plc_specification_sheet["C71"] = plc_panel.get("do_module_terminal")
            plc_specification_sheet["C72"] = plc_panel.get("ai_module_terminal")
            plc_specification_sheet["C73"] = plc_panel.get("ao_module_terminal")
            plc_specification_sheet["C74"] = plc_panel.get("rtd_module_terminal")
            plc_specification_sheet["C75"] = common_config_data.get("thermocouple_module_terminal")

            plc_specification_sheet["C78"] = common_config_data.get("ferrule")
            plc_specification_sheet["C79"] = common_config_data.get("ferrule_note")
            plc_specification_sheet["C80"] = common_config_data.get("device_identification_of_components")

            cooling_fans_data = common_config_data.get("cooling_fans")
            louvers_and_filters_data = common_config_data.get("louvers_and_filters")
            plc_specification_sheet["C82"] = num_to_string(cooling_fans_data)
            plc_specification_sheet["C82"] = num_to_string(louvers_and_filters_data)

            plc_specification_sheet["C85"] = common_config_data.get("commissioning_spare")
            plc_specification_sheet["C86"] = common_config_data.get("two_year_operational_spare")

            plc_specification_sheet["C88"] = plc_panel.get("ups_scope")
            plc_specification_sheet["C89"] = plc_panel.get("ups_input_voltage_3p")
            plc_specification_sheet["C90"] = plc_panel.get("ups_input_voltage_1p")
            plc_specification_sheet["C91"] = plc_panel.get("ups_output_voltage_1p")
            plc_specification_sheet["C92"] = plc_panel.get("ups_type")
            plc_specification_sheet["C93"] = plc_panel.get("ups_battery_type")
            plc_specification_sheet["C94"] = plc_panel.get("ups_battery_backup_time")
            is_ups_battery_mounting_rack_selected_data = plc_panel.get("is_ups_battery_mounting_rack_selected")
            plc_specification_sheet["C95"] = num_to_string(is_ups_battery_mounting_rack_selected_data)
            plc_specification_sheet["C96"] = plc_panel.get("ups_redundancy")

            plc_specification_sheet["C98"] = plc_panel.get("hmi_hardware_make")
            plc_specification_sheet["C99"] = plc_panel.get("plc_cpu_system_series")
            plc_specification_sheet["C100"] = plc_panel.get("plc_cpu_system_input_voltage")
            plc_specification_sheet["C101"] = plc_panel.get("plc_cpu_system_memory_free_space_after_program")

            plc_specification_sheet["C104"] = plc_panel.get("di_module_channel_density")
            plc_specification_sheet["C105"] = plc_panel.get("di_module_loop_current")
            plc_specification_sheet["C106"] = plc_panel.get("di_module_isolation")
            plc_specification_sheet["C107"] = plc_panel.get("di_module_input_type")
            plc_specification_sheet["C108"] = plc_panel.get("di_module_interrogation_voltage")
            plc_specification_sheet["C109"] = plc_panel.get("di_module_scan_time")

            plc_specification_sheet["C111"] = plc_panel.get("do_module_channel_density")
            plc_specification_sheet["C112"] = plc_panel.get("do_module_loop_current")
            plc_specification_sheet["C113"] = plc_panel.get("do_module_isolation")
            plc_specification_sheet["C114"] = plc_panel.get("do_module_output_type")

            plc_specification_sheet["C116"] = plc_panel.get("interposing_relay")
            plc_specification_sheet["C117"] = plc_panel.get("interposing_relay_contacts_rating")

            is_no_of_contacts_selected_data = plc_panel.get("is_no_of_contacts_selected")
            no_of_contact_data = plc_panel.get("no_of_contacts")
            if int(is_no_of_contacts_selected_data) == 0:
                no_of_contact_data = "Not Applicable"
            plc_specification_sheet["C118"] = no_of_contact_data


            plc_specification_sheet["C120"] = plc_panel.get("ai_module_channel_density")
            plc_specification_sheet["C121"] = plc_panel.get("ai_module_loop_current")
            plc_specification_sheet["C122"] = plc_panel.get("ai_module_isolation")
            plc_specification_sheet["C123"] = plc_panel.get("ai_module_input_type")
            plc_specification_sheet["C124"] = plc_panel.get("ai_module_scan_time")
            is_ai_module_hart_protocol_support_selected_data = plc_panel.get("is_ai_module_hart_protocol_support_selected")
            plc_specification_sheet["C125"] = num_to_string(is_ai_module_hart_protocol_support_selected_data)

            plc_specification_sheet["C127"] = plc_panel.get("ao_module_channel_density")
            plc_specification_sheet["C128"] = plc_panel.get("ao_module_loop_current")
            plc_specification_sheet["C129"] = plc_panel.get("ao_module_isolation")
            plc_specification_sheet["C130"] = plc_panel.get("ao_module_output_type")
            plc_specification_sheet["C131"] = plc_panel.get("ao_module_scan_time")
            is_ao_module_hart_protocol_support_selected_data = plc_panel.get("is_ao_module_hart_protocol_support_selected")
            plc_specification_sheet["C132"] = num_to_string(is_ao_module_hart_protocol_support_selected_data)
            

            plc_specification_sheet["C134"] = plc_panel.get("rtd_module_channel_density")
            plc_specification_sheet["C135"] = plc_panel.get("rtd_module_loop_current")
            plc_specification_sheet["C136"] = plc_panel.get("rtd_module_isolation")
            plc_specification_sheet["C137"] = plc_panel.get("rtd_module_input_type")
            plc_specification_sheet["C138"] = plc_panel.get("rtd_module_scan_time")
            is_rtd_module_hart_protocol_support_selected_data = plc_panel.get("is_rtd_module_hart_protocol_support_selected")
            plc_specification_sheet["C139"] = num_to_string(is_rtd_module_hart_protocol_support_selected_data)

            plc_specification_sheet["C141"] = plc_panel.get("thermocouple_module_channel_density")
            plc_specification_sheet["C142"] = plc_panel.get("thermocouple_module_loop_current")
            plc_specification_sheet["C143"] = plc_panel.get("thermocouple_module_isolation")
            plc_specification_sheet["C144"] = plc_panel.get("thermocouple_module_input_type")
            plc_specification_sheet["C145"] = plc_panel.get("thermocouple_module_scan_time")
            is_thermocouple_module_hart_protocol_support_selected_data = plc_panel.get("is_thermocouple_module_hart_protocol_support_selected")
            plc_specification_sheet["C146"] = num_to_string(is_thermocouple_module_hart_protocol_support_selected_data)

            plc_specification_sheet["C148"] = plc_panel.get("universal_module_channel_density")
            plc_specification_sheet["C149"] = plc_panel.get("universal_module_loop_current")
            plc_specification_sheet["C150"] = plc_panel.get("universal_module_isolation")
            plc_specification_sheet["C151"] = plc_panel.get("universal_module_input_type")
            plc_specification_sheet["C152"] = plc_panel.get("universal_module_scan_time")
            is_universal_module_hart_protocol_support_selected_data = plc_panel.get("is_universal_module_hart_protocol_support_selected")
            plc_specification_sheet["C153"] = num_to_string(is_universal_module_hart_protocol_support_selected_data)

            plc_specification_sheet["C156"] = plc_panel.get("hmi_size")
            plc_specification_sheet["C157"] = plc_panel.get("hmi_quantity")
            plc_specification_sheet["C158"] = plc_panel.get("hmi_hardware_make")
            plc_specification_sheet["C159"] = plc_panel.get("hmi_series")
            plc_specification_sheet["C160"] = plc_panel.get("hmi_input_voltage")
            plc_specification_sheet["C161"] = plc_panel.get("hmi_battery_backup")

            is_engineering_station_quantity_selected_data = plc_panel.get("is_engineering_station_quantity_selected")
            engineering_station_quantity_data = plc_panel.get("engineering_station_quantity")

            if int(is_engineering_station_quantity_selected_data) == 0:
                engineering_station_quantity_data = "Not Applicable"

            plc_specification_sheet["C163"] = engineering_station_quantity_data


            is_engineering_cum_operating_station_quantity_selected_data = plc_panel.get("is_engineering_cum_operating_station_quantity_selected")
            engineering_cum_operating_station_quantity_data = plc_panel.get("engineering_cum_operating_station_quantity")

            if int(is_engineering_cum_operating_station_quantity_selected_data) == 0:
                engineering_cum_operating_station_quantity_data = "Not Applicable"

            plc_specification_sheet["C164"] = engineering_cum_operating_station_quantity_data


            is_operating_station_quantity_selected_data = plc_panel.get("is_operating_station_quantity_selected")
            operating_station_quantity_data = plc_panel.get("operating_station_quantity")

            if int(is_operating_station_quantity_selected_data) == 0:
                operating_station_quantity_data = "Not Applicable"

            plc_specification_sheet["C165"] = operating_station_quantity_data


            plc_specification_sheet["C167"] = plc_panel.get("scada_runtime_license_quantity")
            plc_specification_sheet["C168"] = plc_panel.get("scada_program_development_license_quantity")
            plc_specification_sheet["C169"] = plc_panel.get("plc_programming_software_license_quantity")

            is_power_supply_plc_cpu_system_selected_data = plc_panel.get("is_power_supply_plc_cpu_system_selected")
            is_power_supply_input_output_module_selected_data = plc_panel.get("is_power_supply_input_output_module_selected")
            is_plc_input_output_modules_system_selected_data = plc_panel.get("is_plc_input_output_modules_system_selected")
            is_plc_cpu_system_and_input_output_modules_system_selected_data = plc_panel.get("is_plc_cpu_system_and_input_output_modules_system_selected")
            is_plc_cpu_system_and_hmi_scada_selected_data = plc_panel.get("is_plc_cpu_system_and_hmi_scada_selected")
            is_plc_cpu_system_and_third_party_devices_selected_data = plc_panel.get("is_plc_cpu_system_and_third_party_devices_selected")
            is_plc_cpu_system_selected_data = plc_panel.get("is_plc_cpu_system_selected")

            plc_specification_sheet["C171"] =  num_to_string(is_power_supply_plc_cpu_system_selected_data)
            plc_specification_sheet["C172"] =  num_to_string(is_power_supply_input_output_module_selected_data)
            plc_specification_sheet["C173"] =  num_to_string(is_plc_input_output_modules_system_selected_data)
            plc_specification_sheet["C174"] =  num_to_string(is_plc_cpu_system_and_input_output_modules_system_selected_data)
            plc_specification_sheet["C175"] =  num_to_string(is_plc_cpu_system_and_hmi_scada_selected_data)
            plc_specification_sheet["C176"] =  num_to_string(is_plc_cpu_system_and_third_party_devices_selected_data)
            plc_specification_sheet["C177"] =  num_to_string(is_plc_cpu_system_selected_data)


            plc_specification_sheet["C179"] =  plc_panel.get("system_hardware")
            plc_specification_sheet["C180"] =  plc_panel.get("pc_hardware_specifications")
            plc_specification_sheet["C181"] =  plc_panel.get("monitor_size")
            plc_specification_sheet["C182"] =  plc_panel.get("windows_operating_system")
            plc_specification_sheet["C183"] =  plc_panel.get("hardware_between_plc_and_scada_pc")
            plc_specification_sheet["C184"] =  plc_panel.get("is_printer_with_suitable_communication_cable_selected")
            plc_specification_sheet["C185"] =  plc_panel.get("printer_type")
            plc_specification_sheet["C186"] =  plc_panel.get("printer_size")
            plc_specification_sheet["C187"] =  plc_panel.get("printer_quantity")
            plc_specification_sheet["C188"] =  plc_panel.get("is_furniture_selected")
            plc_specification_sheet["C189"] =  plc_panel.get("is_console_with_chair_selected")
            plc_specification_sheet["C190"] =  plc_panel.get("is_plc_logic_diagram_selected")
            plc_specification_sheet["C191"] =  plc_panel.get("is_loop_drawing_for_complete_project_selected")

            plc_specification_sheet["C193"] =  plc_panel.get("interface_signal_and_control_logic_implementation")
            plc_specification_sheet["C194"] =  plc_panel.get("differential_pressure_flow_linearization")
            plc_specification_sheet["C195"] =  plc_panel.get("third_party_comm_protocol_for_plc_cpu_system")
            plc_specification_sheet["C196"] =  plc_panel.get("third_party_communication_protocol")
            plc_specification_sheet["C197"] =  plc_panel.get("hardware_between_plc_and_third_party")
            plc_specification_sheet["C198"] =  plc_panel.get("hardware_between_plc_and_client_system")

            plc_specification_sheet["C199"] =  plc_panel.get("client_system_communication")
            plc_specification_sheet["C200"] =  plc_panel.get("hardware_between_plc_and_client_system")

            plc_specification_sheet["C201"] =  plc_panel.get("is_iiot_selected")
            plc_specification_sheet["C202"] =  plc_panel.get("iiot_gateway_mounting")
            plc_specification_sheet["C203"] =  plc_panel.get("iiot_gateway_note")
            plc_specification_sheet["C204"] =  plc_panel.get("is_burner_controller_lmv_mounting_selected")
            plc_specification_sheet["C205"] =  plc_panel.get("burner_controller_lmv_mounting")
            plc_specification_sheet["C206"] =  plc_panel.get("burner_controller_lmv_note")

        elif project_panel.get("type") == "MCC":
            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"panel_id": panel_id}, "*"
            )

            if len(mcc_panel_data) == 0:
                continue
            mcc_panel_data = mcc_panel_data[0]

            plc_panel_1 = frappe.db.get_list(
                "Panel PLC 1 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_1 = plc_panel_1[0] if len(plc_panel_1) > 0 else {}
            plc_panel_2 = frappe.db.get_list(
                "Panel PLC 2 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_2 = plc_panel_2[0] if len(plc_panel_2) > 0 else {}
            plc_panel_3 = frappe.db.get_list(
                "Panel PLC 3 - 3",
                {"panel_id": panel_id},
                "*",
            )
            plc_panel_3 = plc_panel_3[0] if len(plc_panel_3) > 0 else {}

            plc_panel = {**plc_panel_1, **plc_panel_2, **plc_panel_3}
            # PLC fields

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

            mcc_cum_plc_panel_sheet["C9"] = common_config_data.get("mcc_switchgear_type")
            system_voltage_data = (
                f"{project_info_data.get("main_supply_lv")}, {project_info_data.get("main_supply_lv_variation")}, {project_info_data.get("main_supply_lv_phase")}"
            )
            mcc_cum_plc_panel_sheet["C11"] = system_voltage_data

            freq_data = (
                f"{project_info_data.get("frequency")}, {project_info_data.get("frequency_variation")}"
            )
            mcc_cum_plc_panel_sheet["C12"] = freq_data

            fault_data = (
                f"{project_info_data.get("fault_level")} kA for {project_info_data.get("sec")}"
            )
            mcc_cum_plc_panel_sheet["C13"] = fault_data

            utility_data = (
                f"{project_info_data("utility_supply")}, {project_info_data.get("utility_supply_variation")}, {project_info_data.get("utility_supply_phase")}"
            )
            mcc_cum_plc_panel_sheet["C14"] = utility_data

            control_data = (
                f"{project_info_data.get("control_supply")}, {project_info_data.get("control_supply_variation")}, {project_info_data.get("control_supply_variation")}"
            )
            mcc_cum_plc_panel_sheet["C15"] = control_data
            mcc_cum_plc_panel_sheet["C16"] = project_info_data.get("frequency")

            mcc_cum_plc_panel_sheet["C18"] = project_info_data.get("project_location")
            mcc_cum_plc_panel_sheet["C19"] = project_info_data.get("ambient_temperature_max")
            mcc_cum_plc_panel_sheet["C20"] = project_info_data.get("ambient_temperature_min")
            mcc_cum_plc_panel_sheet["C21"] = project_info_data.get("electrical_design_temperature")
            mcc_cum_plc_panel_sheet["C22"] = project_info_data.get("seismic_zone")
            # mcc_cum_plc_panel_sheet["C23"] = project_info_data.get("area_classification")
            mcc_cum_plc_panel_sheet["C24"] = project_info_data.get("altitude")
            mcc_cum_plc_panel_sheet["C25"] = project_info_data.get("min_humidity")
            mcc_cum_plc_panel_sheet["C26"] = project_info_data.get("max_humidity")
            mcc_cum_plc_panel_sheet["C27"] = project_info_data.get("avg_humidity")
            mcc_cum_plc_panel_sheet["C28"] = project_info_data.get("performance_humidity")

            mcc_cum_plc_panel_sheet["C30"] = project_info_data.get("supply_feeder_standard")
            # mcc_cum_plc_panel_sheet["C31"] = project_info_data.get("supply_feeder_standard")
            mcc_cum_plc_panel_sheet["C32"] = project_info_data.get("ga_mcc_construction_front_type")
            mcc_cum_plc_panel_sheet["C33"] = project_info_data.get("ga_mcc_compartmental")
            mcc_cum_plc_panel_sheet["C34"] = project_info_data.get("incoming_drawout_type")
            mcc_cum_plc_panel_sheet["C35"] = project_info_data.get("outgoing_drawout_type")
            mcc_cum_plc_panel_sheet["C36"] = project_info_data.get("ga_mcc_construction_type")
            mcc_cum_plc_panel_sheet["C37"] = project_info_data.get("ga_panel_mounting_frame")
            mcc_cum_plc_panel_sheet["C38"] = project_info_data.get("ga_panel_mounting_height")
            # mcc_cum_plc_panel_sheet["C39"] = project_info_data.get("supply_feeder_standard")
            mcc_cum_plc_panel_sheet["C40"] = project_info_data.get("marshalling_section_text_area")
            mcc_cum_plc_panel_sheet["C41"] = project_info_data.get("is_cable_alley_section_selected")

            mcc_cum_plc_panel_sheet["C43"] = project_info_data.get("ga_moc_thickness_door")
            mcc_cum_plc_panel_sheet["C44"] = project_info_data.get("door_thickness")
            mcc_cum_plc_panel_sheet["C45"] = project_info_data.get("ga_moc_thickness_covers")
            mcc_cum_plc_panel_sheet["C46"] = project_info_data.get("ga_gland_plate_thickness")
            mcc_cum_plc_panel_sheet["C47"] = project_info_data.get("ga_gland_plate_3mm_drill_type")
            mcc_cum_plc_panel_sheet["C48"] = project_info_data.get("ga_cable_entry_position")
            mcc_cum_plc_panel_sheet["C49"] = project_info_data.get("ga_busbar_chamber_position")
            # mcc_cum_plc_panel_sheet["C50"] = project_info_data.get("is_cable_alley_section_selected")
            mcc_cum_plc_panel_sheet["C51"] = project_info_data.get("ga_power_and_control_busbar_separation")

            mcc_cum_plc_panel_sheet["C53"] = project_info_data.get("ppc_interior_and_exterior_paint_shade")
            mcc_cum_plc_panel_sheet["C54"] = project_info_data.get("ppc_component_mounting_plate_paint_shade")
            # mcc_cum_plc_panel_sheet["C55"] = project_info_data.get("ppc_interior_and_exterior_paint_shade")
            mcc_cum_plc_panel_sheet["C56"] = project_info_data.get("ppc_minimum_coating_thickness")
            mcc_cum_plc_panel_sheet["C56"] = project_info_data.get("ppc_pretreatment_panel_standard")
            mcc_cum_plc_panel_sheet["C56"] = project_info_data.get("general_requirments_for_construction")

            mcc_cum_plc_panel_sheet["C61"] = common_config_data.get("power_bus_main_busbar_selection")
            mcc_cum_plc_panel_sheet["C62"] = common_config_data.get("power_bus_material")
            mcc_cum_plc_panel_sheet["C63"] = common_config_data.get("power_bus_current_density")
            mcc_cum_plc_panel_sheet["C64"] = common_config_data.get("power_bus_rating_of_busbar")
            mcc_cum_plc_panel_sheet["C65"] = common_config_data.get("power_bus_heat_pvc_sleeve")

            mcc_cum_plc_panel_sheet["C67"] = common_config_data.get("control_bus_main_busbar_selection")
            mcc_cum_plc_panel_sheet["C68"] = common_config_data.get("control_bus_material")
            mcc_cum_plc_panel_sheet["C69"] = common_config_data.get("control_bus_current_density")
            mcc_cum_plc_panel_sheet["C70"] = common_config_data.get("control_bus_rating_of_busbar")
            mcc_cum_plc_panel_sheet["C71"] = common_config_data.get("control_bus_heat_pvc_sleeve")
            
            mcc_cum_plc_panel_sheet["C73"] = common_config_data.get("earth_bus_main_busbar_selection")
            mcc_cum_plc_panel_sheet["C74"] = common_config_data.get("earth_bus_material")
            mcc_cum_plc_panel_sheet["C75"] = common_config_data.get("earth_bus_current_density")
            mcc_cum_plc_panel_sheet["C76"] = common_config_data.get("earth_bus_rating_of_busbar")
            mcc_cum_plc_panel_sheet["C77"] = common_config_data.get("earth_bus_busbar_position")
            
            mcc_cum_plc_panel_sheet["C78"] = common_config_data.get("door_earthing")
            mcc_cum_plc_panel_sheet["C79"] = common_config_data.get("instrument_earth")
            mcc_cum_plc_panel_sheet["C80"] = common_config_data.get("general_note_busbar_and_insulation_materials")
            
            
            mcc_cum_plc_panel_sheet["C82"] = common_config_data.get("power_wiring_color")
            mcc_cum_plc_panel_sheet["C83"] = common_config_data.get("power_wiring_size")
            mcc_cum_plc_panel_sheet["C84"] = common_config_data.get("control_wiring_color")
            mcc_cum_plc_panel_sheet["C85"] = common_config_data.get("control_wiring_size")
            mcc_cum_plc_panel_sheet["C86"] = common_config_data.get("vdc_24_wiring_color")
            mcc_cum_plc_panel_sheet["C87"] = common_config_data.get("vdc_24_wiring_size")
            mcc_cum_plc_panel_sheet["C88"] = common_config_data.get("analog_signal_wiring_color")
            mcc_cum_plc_panel_sheet["C89"] = common_config_data.get("analog_signal_wiring_size")
            mcc_cum_plc_panel_sheet["C90"] = common_config_data.get("rtd_thermocouple_wiring_color")
            mcc_cum_plc_panel_sheet["C91"] = common_config_data.get("rtd_thermocouple_wiring_size")
            mcc_cum_plc_panel_sheet["C92"] = common_config_data.get("cable_insulation_pvc")
            
            mcc_cum_plc_panel_sheet["C93"] = common_config_data.get("common_requirement")
            mcc_cum_plc_panel_sheet["C94"] = common_config_data.get("air_clearance_between_phase_to_phase_bus")
            # mcc_cum_plc_panel_sheet["C95"] = common_config_data.get("air_clearance_between_phase_to_phase_bus")
            mcc_cum_plc_panel_sheet["C96"] = common_config_data.get("general_note_internal_wiring")