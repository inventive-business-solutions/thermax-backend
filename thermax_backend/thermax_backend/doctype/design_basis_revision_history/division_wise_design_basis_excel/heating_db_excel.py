import frappe
from thermax_backend.thermax_backend.doctype.design_basis_revision_history.division_wise_design_basis_excel.utils import (
    handle_make_of_component,
    handle_none_to_number,
    handle_none_to_string,
    num_to_string,
)


def get_heating_db_excel(
    template_workbook,
    mcc_sheet,
    pcc_sheet,
    mcc_cum_plc_sheet,
    project_data,
    make_of_components_data,
    revision_id,
):
    project_panel_data = frappe.db.get_list(
        "Project Panel Data", {"revision_id": revision_id}, "*", order_by="creation asc"
    )

    for project_panel in project_panel_data:
        panel_id = project_panel.get("name")
        if project_panel.get("panel_main_type") == "MCC":
            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"panel_id": panel_id}, "*"
            )
            panel_sheet = template_workbook.copy_worksheet(mcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            if len(mcc_panel_data) == 0:
                continue
            mcc_panel_data = mcc_panel_data[0]

            panel_sheet["B3"] = project_panel.get("panel_name")

            incomer_ampere = mcc_panel_data.get("incomer_ampere")
            incomer_pole = mcc_panel_data.get("incomer_pole")
            incomer_type = mcc_panel_data.get("incomer_type")
            incomer_above_ampere = mcc_panel_data.get("incomer_above_ampere")
            incomer_above_pole = mcc_panel_data.get("incomer_above_pole")
            incomer_above_type = mcc_panel_data.get("incomer_above_type")

            is_indication_on_selected = handle_none_to_number(
                mcc_panel_data.get("is_indication_on_selected")
            )
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

            current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
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

            ga_panel_mounting_frame = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            is_marshalling_section_selected = handle_none_to_number(
                mcc_panel_data.get("is_marshalling_section_selected")
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

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            incomer_data = f"Upto {incomer_ampere}, {incomer_pole} Pole {incomer_type} \nAbove {incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type}"
            panel_sheet["C5"] = incomer_data

            panel_sheet["C6"] = (
                "Not Applicable"
                if is_indication_on_selected == 0
                else handle_none_to_string(mcc_panel_data.get("led_type_on_input"))
            )
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input

            if incomer_type is None or "ACB" not in incomer_type:
                is_blue_cb_spring_charge_selected = "NA"
                is_red_cb_in_service = "NA"
                is_white_healthy_trip_circuit_selected = "NA"

            panel_sheet["C9"] = handle_none_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = handle_none_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = handle_none_to_string(
                is_white_healthy_trip_circuit_selected
            )
            panel_sheet["C12"] = handle_none_to_string(alarm_annunciator)

            panel_sheet["C14"] = handle_make_of_component(mi_analog)
            panel_sheet["C15"] = handle_make_of_component(mi_digital)
            panel_sheet["C16"] = handle_none_to_string(mi_communication_protocol)

            panel_sheet["C18"] = handle_none_to_string(current_transformer_coating)
            panel_sheet["C19"] = handle_none_to_string(current_transformer_number)
            panel_sheet["C20"] = handle_none_to_string(
                current_transformer_configuration
            )

            panel_sheet["C22"] = ga_moc_material  # MOC
            panel_sheet["C23"] = handle_none_to_string(
                ga_moc_thickness_door
            )  # Component Mounting Plate Thickness
            panel_sheet["C24"] = handle_none_to_string(door_thickness)  # Door Thickness
            panel_sheet["C25"] = handle_none_to_string(
                ga_moc_thickness_covers
            )  # Top & Side Thickness
            panel_sheet["C26"] = handle_none_to_string(
                ga_gland_plate_thickness
            )  # Gland Plate Thickness
            panel_sheet["C27"] = handle_none_to_string(
                ga_gland_plate_3mm_drill_type
            )  # Gland Plate Type
            panel_sheet["C28"] = ga_mcc_compartmental  # Panel Front Type
            panel_sheet["C29"] = (
                ga_mcc_construction_front_type  # Type of Construction for Board
            )
            if (ga_mcc_compartmental is None) or ("Non" in ga_mcc_compartmental):
                incoming_drawout_type = "Not Applicable"
                outgoing_drawout_type = "Not Applicable"

            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_mcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            panel_sheet["C35"] = (
                "Not Applicable"
                if is_marshalling_section_selected == 0
                else marshalling_section_text_area
            )
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

        elif project_panel.get("panel_main_type") == "PCC":

            pcc_panel_data = frappe.db.get_list(
                "PCC Panel", {"panel_id": panel_id}, "*"
            )
            panel_sheet = template_workbook.copy_worksheet(pcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            if len(pcc_panel_data) == 0:
                continue
            pcc_panel_data = pcc_panel_data[0]

            panel_sheet["B3"] = project_data.get("panel_name")

            incomer_ampere = handle_none_to_string(pcc_panel_data.get("incomer_ampere"))
            incomer_pole = handle_none_to_string(pcc_panel_data.get("incomer_pole"))
            incomer_type = handle_none_to_string(pcc_panel_data.get("incomer_type"))
            incomer_above_ampere = handle_none_to_string(
                pcc_panel_data.get("incomer_above_ampere")
            )
            incomer_above_pole = handle_none_to_string(
                pcc_panel_data.get("incomer_above_pole")
            )
            incomer_above_type = handle_none_to_string(
                pcc_panel_data.get("incomer_above_type")
            )

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

            current_transformer_coating = pcc_panel_data.get(
                "current_transformer_coating"
            )
            current_transformer_number = pcc_panel_data.get(
                "current_transformer_number"
            )
            current_transformer_configuration = pcc_panel_data.get(
                "current_transformer_configuration"
            )
            alarm_annunciator = handle_none_to_string(
                pcc_panel_data.get("alarm_annunciator")
            )
            mi_analog = handle_none_to_string(
                pcc_panel_data.get("mi_analog", "Not Applicable")
            )
            mi_digital = handle_none_to_string(
                pcc_panel_data.get("mi_digital", "Not Applicable")
            )
            mi_communication_protocol = handle_none_to_string(
                pcc_panel_data.get("mi_communication_protocol", "Not Applicable")
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

            ga_panel_mounting_frame = pcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = pcc_panel_data.get("ga_panel_mounting_height")
            is_marshalling_section_selected = handle_none_to_number(
                pcc_panel_data.get("is_marshalling_section_selected")
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
            ppc_interior_and_exterior_paint_shade = pcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            ppc_component_mounting_plate_paint_shade = pcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
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

            pcc_incomer_data = f"Upto {incomer_ampere}, {incomer_pole} Pole {incomer_type} \nAbove {incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type}"

            if is_indication_on_selected == "0" or is_indication_on_selected == 0:
                led_type_on_input = "Not Applicable"

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            panel_sheet["C5"] = handle_none_to_string(pcc_incomer_data)
            panel_sheet["C6"] = led_type_on_input
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input
            if not "ACB" in incomer_type:
                is_blue_cb_spring_charge_selected = "NA"
                is_red_cb_in_service = "NA"
                is_white_healthy_trip_circuit_selected = "NA"

            panel_sheet["C9"] = handle_none_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = handle_none_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = handle_none_to_string(
                is_white_healthy_trip_circuit_selected
            )
            panel_sheet["C12"] = handle_none_to_string(alarm_annunciator)

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

            if (ga_pcc_compartmental is None) or ("Non" in ga_pcc_compartmental):
                incoming_drawout_type = "Not Applicable"
                outgoing_drawout_type = "Not Applicable"

            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_pcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            panel_sheet["C35"] = (
                "Not Applicable"
                if is_marshalling_section_selected == 0
                else marshalling_section_text_area
            )  # Marshalling Section
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

        else:
            mcc_panel_data = frappe.db.get_list(
                "MCC Panel", {"panel_id": panel_id}, "*"
            )
            panel_sheet = template_workbook.copy_worksheet(mcc_cum_plc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            if len(mcc_panel_data) == 0:
                continue
            mcc_panel_data = mcc_panel_data[0]

            panel_sheet["B3"] = project_panel.get("panel_name")

            incomer_ampere = mcc_panel_data.get("incomer_ampere")
            incomer_pole = mcc_panel_data.get("incomer_pole")
            incomer_type = mcc_panel_data.get("incomer_type")
            incomer_above_ampere = mcc_panel_data.get("incomer_above_ampere")
            incomer_above_pole = mcc_panel_data.get("incomer_above_pole")
            incomer_above_type = mcc_panel_data.get("incomer_above_type")

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

            current_transformer_coating = mcc_panel_data.get(
                "current_transformer_coating"
            )

            current_transformer_number = mcc_panel_data.get(
                "current_transformer_number"
            )
            current_transformer_configuration = mcc_panel_data.get(
                "current_transformer_configuration"
            )
            alarm_annunciator = handle_none_to_string(
                mcc_panel_data.get("alarm_annunciator")
            )
            mi_analog = handle_none_to_string(mcc_panel_data.get("mi_analog"))
            mi_digital = handle_none_to_string(mcc_panel_data.get("mi_digital"))
            mi_communication_protocol = handle_none_to_string(
                mcc_panel_data.get("mi_communication_protocol")
            )
            ga_moc_material = handle_none_to_string(
                mcc_panel_data.get("ga_moc_material")
            )
            door_thickness = handle_none_to_string(mcc_panel_data.get("door_thickness"))
            ga_moc_thickness_door = handle_none_to_string(
                mcc_panel_data.get("ga_moc_thickness_door")
            )
            ga_moc_thickness_covers = handle_none_to_string(
                mcc_panel_data.get("ga_moc_thickness_covers")
            )
            ga_mcc_compartmental = handle_none_to_string(
                mcc_panel_data.get("ga_mcc_compartmental")
            )
            ga_mcc_construction_front_type = mcc_panel_data.get(
                "ga_mcc_construction_front_type"
            )
            incoming_drawout_type = handle_none_to_string(
                mcc_panel_data.get("incoming_drawout_type")
            )
            outgoing_drawout_type = handle_none_to_string(
                mcc_panel_data.get("outgoing_drawout_type")
            )
            ga_mcc_construction_type = handle_none_to_string(
                mcc_panel_data.get("ga_mcc_construction_type")
            )

            ga_panel_mounting_frame = handle_none_to_string(
                mcc_panel_data.get("ga_panel_mounting_frame")
            )
            ga_panel_mounting_height = handle_none_to_string(
                mcc_panel_data.get("ga_panel_mounting_height")
            )
            is_marshalling_section_selected = handle_none_to_number(
                mcc_panel_data.get("is_marshalling_section_selected")
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

            incomer_data = f"Upto {incomer_ampere}, {incomer_pole} Pole {incomer_type} \nAbove {incomer_above_ampere}, {incomer_above_pole} Pole {incomer_above_type} "

            if is_indication_on_selected == "0" or is_indication_on_selected == 0:
                led_type_on_input = "Not Applicable"

            if is_indication_off_selected == "0" or is_indication_off_selected == 0:
                led_type_off_input = "Not Applicable"

            if is_indication_trip_selected == "0" or is_indication_trip_selected == 0:
                led_type_trip_input = "Not Applicable"

            panel_sheet["C5"] = handle_none_to_string(incomer_data)
            panel_sheet["C6"] = led_type_on_input
            panel_sheet["C7"] = led_type_off_input
            panel_sheet["C8"] = led_type_trip_input
            panel_sheet["C9"] = handle_none_to_string(is_blue_cb_spring_charge_selected)
            panel_sheet["C10"] = handle_none_to_string(is_red_cb_in_service)
            panel_sheet["C11"] = handle_none_to_string(
                is_white_healthy_trip_circuit_selected
            )
            panel_sheet["C12"] = handle_none_to_string(alarm_annunciator)

            if "NA" in mi_communication_protocol:
                mi_communication_protocol = "Not Applicable"

            panel_sheet["C14"] = handle_none_to_string(current_transformer_coating)
            panel_sheet["C15"] = handle_none_to_string(current_transformer_number)
            panel_sheet["C16"] = handle_none_to_string(
                current_transformer_configuration
            )

            panel_sheet["C18"] = handle_make_of_component(mi_analog)
            panel_sheet["C19"] = handle_make_of_component(mi_digital)
            panel_sheet["C20"] = handle_make_of_component(mi_communication_protocol)

            panel_sheet["C22"] = ga_moc_material  # MOC
            panel_sheet["C23"] = handle_none_to_string(
                ga_moc_thickness_door
            )  # Component Mounting Plate Thickness
            panel_sheet["C24"] = handle_none_to_string(door_thickness)  # Door Thickness
            panel_sheet["C25"] = handle_none_to_string(
                ga_moc_thickness_covers
            )  # Top & Side Thickness
            panel_sheet["C26"] = handle_none_to_string(
                ga_gland_plate_thickness
            )  # Gland Plate Thickness
            panel_sheet["C27"] = handle_none_to_string(
                ga_gland_plate_3mm_drill_type
            )  # Gland Plate Type
            panel_sheet["C28"] = ga_mcc_compartmental  # Panel Front Type
            panel_sheet["C29"] = (
                ga_mcc_construction_front_type  # Type of Construction for Board
            )
            if "Non" in ga_mcc_compartmental:
                incoming_drawout_type = "Not Applicable"
                outgoing_drawout_type = "Not Applicable"

            panel_sheet["C30"] = incoming_drawout_type
            panel_sheet["C31"] = outgoing_drawout_type
            panel_sheet["C32"] = ga_mcc_construction_type  # Panel Construction Type
            panel_sheet["C33"] = ga_panel_mounting_frame  # Panel Mounting
            panel_sheet["C34"] = (
                f"{ga_panel_mounting_height} mm"  # Height of Base Frame
            )

            panel_sheet["C35"] = (
                "Not Applicable"
                if is_marshalling_section_selected == 0
                else marshalling_section_text_area
            )  # Marshalling Section
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

            panel_sheet["C202"] = boiler_model
            panel_sheet["C203"] = boiler_fuel
            panel_sheet["C204"] = boiler_year
            panel_sheet["C205"] = mcc_boiler_power_supply
            panel_sheet["C206"] = mcc_boiler_control_supply
            panel_sheet["C207"] = boiler_evaporation
            panel_sheet["C208"] = boiler_output
            panel_sheet["C209"] = boiler_connected_load
            panel_sheet["C210"] = boiler_design_pressure

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

            panel_sheet["C212"] = heater_model
            panel_sheet["C213"] = heater_fuel
            panel_sheet["C214"] = heater_year
            panel_sheet["C215"] = mcc_heater_power_supply
            panel_sheet["C216"] = mcc_heater_control_supply
            panel_sheet["C217"] = heater_evaporation
            panel_sheet["C218"] = heater_output
            panel_sheet["C219"] = heater_connected_load
            panel_sheet["C220"] = heater_temperature

            plc_panel_1 = frappe.db.get_list(
                "Panel PLC 1 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_1 = plc_panel_1[0] if len(plc_panel_1) > 0 else {}
            plc_panel_2 = frappe.db.get_list(
                "Panel PLC 2 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_2 = plc_panel_2[0] if len(plc_panel_2) > 0 else {}
            plc_panel_3 = frappe.db.get_list(
                "Panel PLC 3 - 3",
                {"revision_id": revision_id, "panel_id": panel_id},
                "*",
            )
            plc_panel_3 = plc_panel_3[0] if len(plc_panel_3) > 0 else {}

            plc_panel = {**plc_panel_1, **plc_panel_2, **plc_panel_3}
            # PLC fields
            # Supply Requirements
            panel_sheet["C58"] = handle_none_to_string(
                plc_panel.get("ups_control_voltage", "NA")
            )
            panel_sheet["C59"] = handle_none_to_string(
                plc_panel.get("non_ups_control_voltage", "NA")
            )
            panel_sheet["C60"] = num_to_string(
                plc_panel.get("is_bulk_power_supply_selected", "0")
            )

            # UPS
            ups_scope = plc_panel.get("ups_scope")
            panel_sheet["C62"] = ups_scope
            panel_sheet["C63"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_input_voltage_3p", "NA"))
            )
            panel_sheet["C64"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_input_voltage_1p", "NA"))
            )
            panel_sheet["C65"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_output_voltage_1p", "NA"))
            )
            panel_sheet["C66"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_type", "NA"))
            )
            panel_sheet["C67"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_battery_type", "NA"))
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
                else handle_none_to_string(
                    plc_panel.get("ups_battery_backup_time", "NA")
                )
            )
            panel_sheet["C70"] = (
                "Not Applicable"
                if ups_scope == "Client Scope"
                else handle_none_to_string(plc_panel.get("ups_redundancy", "NA"))
            )

            plc = make_of_components_data.get("plc")

            # PLC Hardware
            panel_sheet["C72"] = handle_make_of_component(plc)
            panel_sheet["C73"] = plc_panel.get(
                "plc_cpu_system_series", "Not Applicable"
            )
            panel_sheet["C74"] = plc_panel.get(
                "plc_cpu_system_input_voltage", "Not Applicable"
            )
            plc_cpu = plc_panel.get(
                "plc_cpu_system_memory_free_space_after_program", "Not Applicable"
            )
            panel_sheet["C75"] = f"{plc_cpu} %"
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
            is_marshalling_cabinet_for_plc_and_ups_selected = handle_none_to_number(
                plc_panel.get("is_marshalling_cabinet_for_plc_and_ups_selected", 0)
            )
            panel_sheet["C86"] = (
                plc_panel.get("marshalling_cabinet_for_plc_and_ups")
                if is_marshalling_cabinet_for_plc_and_ups_selected == 1
                else "Not Applicable"
            )

            # Panel Mounted Push Buttons , Indication Lamps & Colors
            is_electronic_hooter_selected = handle_none_to_number(
                plc_panel.get("is_electronic_hooter_selected")
            )
            panel_sheet["C88"] = (
                plc_panel.get("electronic_hooter_acknowledge")
                if is_electronic_hooter_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C89"] = handle_none_to_string(
                plc_panel.get("panel_power_supply_on_color", "NA")
            )
            panel_sheet["C90"] = handle_none_to_string(
                plc_panel.get("panel_power_supply_off_color", "NA")
            )
            panel_sheet["C91"] = handle_none_to_string(
                plc_panel.get("indicating_lamp_color_for_nonups_power_supply", "NA")
            )
            panel_sheet["C92"] = handle_none_to_string(
                plc_panel.get("indicating_lamp_colour_for_ups_power_supply", "NA")
            )

            # # DI Modules
            panel_sheet["C94"] = plc_panel.get("di_module_channel_density")
            panel_sheet["C95"] = plc_panel.get("di_module_loop_current")
            panel_sheet["C96"] = handle_none_to_string(
                plc_panel.get("di_module_isolation")
            )  # UI Error
            panel_sheet["C97"] = plc_panel.get("di_module_input_type")
            panel_sheet["C98"] = handle_none_to_string(
                plc_panel.get("di_module_interrogation_voltage")
            )  # UI Error
            panel_sheet["C99"] = plc_panel.get("di_module_scan_time")

            # DO Modules
            panel_sheet["C101"] = plc_panel.get("do_module_channel_density")
            panel_sheet["C102"] = plc_panel.get("do_module_loop_current")
            panel_sheet["C103"] = handle_none_to_string(
                plc_panel.get("do_module_isolation")
            )
            panel_sheet["C104"] = plc_panel.get("do_module_output_type")

            # # Interposing Relay
            is_no_of_contacts_selected = handle_none_to_number(
                plc_panel.get("is_no_of_contacts_selected")
            )
            panel_sheet["C106"] = handle_none_to_string(
                plc_panel.get("interposing_relay", "NA")
            )
            panel_sheet["C107"] = handle_none_to_string(
                plc_panel.get("interposing_relay_contacts_rating")
            )
            panel_sheet["C108"] = (
                plc_panel.get("no_of_contacts")
                if is_no_of_contacts_selected == 1
                else "Not Applicable"
            )

            # AI Modules
            panel_sheet["C110"] = plc_panel.get("ai_module_channel_density")
            panel_sheet["C111"] = plc_panel.get("ai_module_loop_current")
            panel_sheet["C112"] = handle_none_to_string(
                plc_panel.get("ai_module_isolation")
            )
            panel_sheet["C113"] = plc_panel.get("ai_module_input_type")
            panel_sheet["C114"] = plc_panel.get("ai_module_scan_time")
            is_ai_module_hart_protocol_support_selected = handle_none_to_number(
                plc_panel.get("is_ai_module_hart_protocol_support_selected")
            )
            panel_sheet["C115"] = (
                "Applicable"
                if is_ai_module_hart_protocol_support_selected == 1
                else "Not Applicable"
            )

            # AO Modules
            panel_sheet["C117"] = plc_panel.get("ao_module_channel_density")
            panel_sheet["C118"] = plc_panel.get("ao_module_loop_current")
            panel_sheet["C119"] = handle_none_to_string(
                plc_panel.get("ao_module_isolation")
            )
            panel_sheet["C120"] = plc_panel.get("ao_module_output_type")
            panel_sheet["C121"] = plc_panel.get("ao_module_scan_time")
            is_ao_module_hart_protocol_support_selected = handle_none_to_number(
                plc_panel.get("is_ao_module_hart_protocol_support_selected")
            )
            panel_sheet["C122"] = (
                "Applicable"
                if is_ao_module_hart_protocol_support_selected == 1
                else "Not Applicable"
            )

            # # RTD Modules
            panel_sheet["C124"] = plc_panel.get("rtd_module_channel_density")
            panel_sheet["C125"] = plc_panel.get("rtd_module_loop_current")
            panel_sheet["C126"] = handle_none_to_string(
                plc_panel.get("rtd_module_isolation")
            )
            panel_sheet["C127"] = plc_panel.get("rtd_module_input_type")
            panel_sheet["C128"] = plc_panel.get("rtd_module_scan_time")
            is_rtd_module_hart_protocol_support_selected = handle_none_to_number(
                plc_panel.get("is_rtd_module_hart_protocol_support_selected")
            )
            panel_sheet["C129"] = (
                "Applicable"
                if is_rtd_module_hart_protocol_support_selected == 1
                else "Not Applicable"
            )

            # Thermocouple Modules
            panel_sheet["C131"] = plc_panel.get("thermocouple_module_channel_density")
            panel_sheet["C132"] = plc_panel.get("thermocouple_module_loop_current")
            panel_sheet["C133"] = handle_none_to_string(
                plc_panel.get("thermocouple_module_isolation")
            )
            panel_sheet["C134"] = plc_panel.get("thermocouple_module_input_type")
            panel_sheet["C135"] = plc_panel.get("thermocouple_module_scan_time")
            is_thermocouple_module_hart_protocol_support_selected = (
                handle_none_to_number(
                    plc_panel.get(
                        "is_thermocouple_module_hart_protocol_support_selected"
                    )
                )
            )
            panel_sheet["C136"] = (
                "Applicable"
                if is_thermocouple_module_hart_protocol_support_selected == 1
                else "Not Applicable"
            )

            # Universal Modules
            panel_sheet["C138"] = plc_panel.get("universal_module_channel_density")
            panel_sheet["C139"] = plc_panel.get("universal_module_loop_current")
            panel_sheet["C140"] = handle_none_to_string(
                plc_panel.get("universal_module_isolation")
            )
            panel_sheet["C141"] = plc_panel.get("universal_module_input_type")
            panel_sheet["C142"] = plc_panel.get("universal_module_scan_time")
            is_universal_module_hart_protocol_support_selected = handle_none_to_number(
                plc_panel.get("is_universal_module_hart_protocol_support_selected")
            )
            panel_sheet["C143"] = (
                "Applicable"
                if is_universal_module_hart_protocol_support_selected == 1
                else "Not Applicable"
            )

            # Terminal Block Connectors
            panel_sheet["C145"] = handle_none_to_string(
                plc_panel.get("di_module_terminal", "NA")
            )
            panel_sheet["C146"] = handle_none_to_string(
                plc_panel.get("do_module_terminal", "NA")
            )
            panel_sheet["C147"] = handle_none_to_string(
                plc_panel.get("ai_module_terminal", "NA")
            )
            panel_sheet["C148"] = handle_none_to_string(
                plc_panel.get("ao_module_terminal", "NA")
            )
            panel_sheet["C149"] = handle_none_to_string(
                plc_panel.get("rtd_module_terminal", "NA")
            )
            panel_sheet["C150"] = handle_none_to_string(
                plc_panel.get("thermocouple_module_terminal", "NA")
            )

            # HMI
            is_hmi_selected = handle_none_to_number(plc_panel.get("is_hmi_selected", 0))

            hmi_size = handle_none_to_string(plc_panel.get("hmi_size", "NA"))
            panel_sheet["C152"] = (
                f"{hmi_size} inch" if is_hmi_selected == 1 else "Not Applicable"
            )
            panel_sheet["C153"] = (
                plc_panel.get("hmi_quantity", 0)
                if is_hmi_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C154"] = (
                handle_none_to_string(plc_panel.get("hmi_hardware_make", "NA"))
                if is_hmi_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C155"] = (
                handle_none_to_string(plc_panel.get("hmi_series", "NA"))
                if is_hmi_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C156"] = (
                handle_none_to_string(plc_panel.get("hmi_input_voltage", "NA"))
                if is_hmi_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C157"] = (
                handle_none_to_string(plc_panel.get("hmi_battery_backup", "NA"))
                if is_hmi_selected == 1
                else "Not Applicable"
            )

            # Human Interface Device
            is_engineering_station_quantity_selected = handle_none_to_number(
                plc_panel.get("is_engineering_station_quantity_selected", 0)
            )

            panel_sheet["C159"] = (
                plc_panel.get("engineering_station_quantity", 0)
                if is_engineering_station_quantity_selected == 1
                else "Not Applicable"
            )

            is_engineering_cum_operating_station_quantity_selected = (
                handle_none_to_number(
                    plc_panel.get(
                        "is_engineering_cum_operating_station_quantity_selected", 0
                    )
                )
            )

            panel_sheet["C160"] = (
                plc_panel.get("engineering_cum_operating_station_quantity", 0)
                if is_engineering_cum_operating_station_quantity_selected == 1
                else "Not Applicable"
            )

            is_operating_station_quantity_selected = handle_none_to_number(
                plc_panel.get("is_operating_station_quantity_selected", 0)
            )

            panel_sheet["C161"] = (
                plc_panel.get("operating_station_quantity", 0)
                if is_operating_station_quantity_selected == 1
                else "Not Applicable"
            )

            # Software License
            is_scada_program_development_license_quantity_selected = (
                handle_none_to_number(
                    plc_panel.get(
                        "is_scada_program_development_license_quantity_selected", 0
                    )
                )
            )

            panel_sheet["C163"] = (
                plc_panel.get("scada_program_development_license_quantity", 0)
                if is_scada_program_development_license_quantity_selected == 1
                else "Not Applicable"
            )

            is_scada_runtime_license_quantity_selected = handle_none_to_number(
                plc_panel.get("is_scada_runtime_license_quantity_selected", 0)
            )

            panel_sheet["C164"] = (
                plc_panel.get("scada_runtime_license_quantity", 0)
                if is_scada_runtime_license_quantity_selected == 1
                else "Not Applicable"
            )

            is_plc_progamming_software_license_quantity = handle_none_to_number(
                plc_panel.get("is_plc_progamming_software_license_quantity", 0)
            )

            panel_sheet["C165"] = (
                plc_panel.get("plc_programming_software_license_quantity", 0)
                if is_plc_progamming_software_license_quantity == 1
                else "Not Applicable"
            )

            # Engineering/Operating SCADA Station
            panel_sheet["C167"] = plc_panel.get("system_hardware", "Not Applicable")
            panel_sheet["C168"] = plc_panel.get(
                "pc_hardware_specifications", "Not Applicable"
            )
            monitor_size_data = handle_none_to_string(plc_panel.get("monitor_size"))

            panel_sheet["C169"] = f"{monitor_size_data} inch"
            panel_sheet["C170"] = plc_panel.get(
                "windows_operating_system", "Not Applicable"
            )
            panel_sheet["C171"] = plc_panel.get(
                "hardware_between_plc_and_scada_pc", "Not Applicable"
            )

            is_printer_with_suitable_communication_cable_selected = (
                handle_none_to_number(
                    plc_panel.get(
                        "is_printer_with_suitable_communication_cable_selected", 0
                    )
                )
            )

            panel_sheet["C172"] = (
                "Applicable"
                if is_printer_with_suitable_communication_cable_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C173"] = (
                plc_panel.get("printer_type", 0)
                if is_printer_with_suitable_communication_cable_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C174"] = (
                plc_panel.get("printer_size", 0)
                if is_printer_with_suitable_communication_cable_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C175"] = (
                plc_panel.get("printer_quantity", 0)
                if is_printer_with_suitable_communication_cable_selected == 1
                else "Not Applicable"
            )

            panel_sheet["C176"] = (
                "Applicable"
                if plc_panel.get("is_furniture_selected", 0) == 1
                else "Not Applicable"
            )
            panel_sheet["C177"] = (
                "Applicable"
                if plc_panel.get("is_console_with_chair_selected", 0) == 1
                else "Not Applicable"
            )
            panel_sheet["C178"] = (
                "Applicable"
                if plc_panel.get("is_plc_logic_diagram_selected", 0) == 1
                else "Not Applicable"
            )
            panel_sheet["C179"] = (
                "Applicable"
                if plc_panel.get("is_loop_drawing_for_complete_project_selected", 0)
                == 1
                else "Not Applicable"
            )

            # Communication
            panel_sheet["C181"] = handle_none_to_string(
                plc_panel.get(
                    "interface_signal_and_control_logic_implementation",
                    "Not Applicable",
                )
            )
            panel_sheet["C182"] = handle_none_to_string(
                plc_panel.get(
                    "differential_pressure_flow_linearization", "Not Applicable"
                )
            )
            panel_sheet["C183"] = handle_none_to_string(
                plc_panel.get(
                    "third_party_comm_protocol_for_plc_cpu_system", "Not Applicable"
                )
            )
            panel_sheet["C184"] = plc_panel.get(
                "third_party_communication_protocol", "Not Applicable"
            )
            panel_sheet["C185"] = plc_panel.get(
                "hardware_between_plc_and_third_party", "Not Applicable"
            )

            is_client_system_comm_with_plc_cpu_selected = handle_none_to_number(
                plc_panel.get("is_client_system_comm_with_plc_cpu_selected", 0)
            )

            panel_sheet["C186"] = (
                "Applicable"
                if is_client_system_comm_with_plc_cpu_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C187"] = (
                plc_panel.get("client_system_communication", 0)
                if is_client_system_comm_with_plc_cpu_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C188"] = (
                plc_panel.get("hardware_between_plc_and_client_system", 0)
                if is_client_system_comm_with_plc_cpu_selected == 1
                else "Not Applicable"
            )

            is_iiot_selected = plc_panel.get("is_iiot_selected", 0)
            panel_sheet["C189"] = (
                plc_panel.get("iiot_gateway_mounting", 0)
                if is_iiot_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C190"] = (
                plc_panel.get("iiot_gateway_note", 0)
                if is_iiot_selected == 1
                else "Not Applicable"
            )

            # Burner Controller LMV
            is_burner_controller_lmv_mounting_selected = plc_panel.get(
                "is_burner_controller_lmv_mounting_selected", 0
            )

            panel_sheet["C192"] = (
                plc_panel.get("burner_controller_lmv_mounting", 0)
                if is_burner_controller_lmv_mounting_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C193"] = (
                plc_panel.get("hardware_between_plc_and_burner_controller_lmv", 0)
                if is_burner_controller_lmv_mounting_selected == 1
                else "Not Applicable"
            )
            panel_sheet["C194"] = (
                plc_panel.get("burner_controller_lmv_note", 0)
                if is_burner_controller_lmv_mounting_selected == 1
                else "Not Applicable"
            )

            # PLC Spares
            panel_sheet["C196"] = plc_panel.get(
                "spare_input_and_output_notes", "Not Applicable"
            )
            panel_sheet["C197"] = plc_panel.get("commissioning_spare", "Not Applicable")
            panel_sheet["C198"] = plc_panel.get(
                "two_year_operational_spare", "Not Applicable"
            )
            panel_sheet["C199"] = plc_panel.get(
                "project_specific_notes", "Not Applicable"
            )

    return template_workbook
