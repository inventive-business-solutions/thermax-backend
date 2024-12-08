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
    payload = frappe.local.form_dict
    metadata = payload.get("metadata")
    project = payload.get("project")
    document_revisions = payload.get("documentRevisions")
    project_package_data = payload.get("projectMainPkgData")
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


    def na_To_string(value):
        if value == "NA" :
            return "Not Applicable"
        return value
        
    def number_To_string(value):
        if value == 0 :
            return "Not Applicable"
        elif value == 1 :
            return "Applicable"
        else:
            return value

    cover_sheet = template_workbook["COVER"]
    design_basis_sheet = template_workbook["Design Basis"]
    revision_sheet = template_workbook["REVISION"]
    mcc_sheet = template_workbook["MCC"]
    pcc_sheet = template_workbook["PCC"]
    mcc_cum_plc_sheet = template_workbook["MCC CUM PLC"]

    # revision id


    current_revision_id = general_info.get("revision_id")

    # Cover Sheet
    division_name = metadata.get("division_name").upper()  # Get the division name and convert to uppercase
    if division_name == "WWS SPG":
        cover_sheet["A3"] = "Water & Waste Solution".upper()  # Replace with desired text
        cover_sheet["A4"] = "411 026"
    elif division_name == "Enviro".upper():
        cover_sheet["A4"] = "411 026"
    else:
        cover_sheet["A3"] = division_name.upper()  # Otherwise, use the original division name

    # cover_sheet["A3"] = metadata.get("division_name").upper()

    revision_date = document_revisions[0].get("modified")
    revision_date = datetime.strptime(revision_date, "%Y-%m-%d %H:%M:%S.%f").strftime("%d-%m-%Y")

    cover_sheet["C36"] = revision_date
    cover_sheet["D7"] = project.get("client_name").upper()
    cover_sheet["D8"] = project.get("consultant_name").upper()
    cover_sheet["D9"] = project.get("project_name").upper()
    cover_sheet["D10"] = project.get("project_oc_number").upper()

    cover_sheet["D36"] = document_revisions[0].get("status") # from payload

    project_owner = project.get("owner")
    project_approver = project.get("approver")

    division_current = metadata.get("division_name")

    prepped_by_initial = frappe.db.get_value("Thermax Extended User", project_owner, "name_initial")
    checked_by_initial = frappe.db.get_value("Thermax Extended User", project_approver, "name_initial")
    super_user_initial = frappe.db.get_value("Thermax Extended User",{"is_superuser":1, "division":division_current}, "name_initial")
    
    cover_sheet["E36"] = prepped_by_initial
    cover_sheet["F36"] = checked_by_initial
    cover_sheet["G36"] = super_user_initial

    # Revision Sheet

    """
        Design Basis Sheet
    """

    # from datetime import datetime

# Assuming 'revision_sheet' is already defined and is a valid object
# and 'document_revisions' is a list of dictionaries containing the revision data.

# Start from row 6 (assuming you want to fill from row 6 onwards)
    
    start_row = 6
    modified_revision_date = document_revisions[0].get("modified")
    modified_revision_date = datetime.strptime(modified_revision_date, "%Y-%m-%d %H:%M:%S.%f").strftime("%d-%m-%Y")

    if(len(document_revisions) > 1):

        for idx, revision in enumerate(document_revisions):
            # Extracting the modified date and formatting it
            modified_revision_date = revision.get("modified")
            
            if modified_revision_date:
                modified_revision_date = datetime.strptime(modified_revision_date, "%Y-%m-%d %H:%M:%S.%f").strftime("%d-%m-%Y")
            else:
                modified_revision_date = ""  # Handle cases where 'modified' might be None

                # Update the revision_sheet with the current revision data
                revision_sheet[f"B{start_row + idx}"] = revision.get("idx")
                revision_sheet[f"D{start_row + idx}"] = modified_revision_date
                revision_sheet[f"E{start_row + idx}"] = revision.get("status")
    else:
        some = document_revisions[0].get("idx")
        revision_sheet[f"B6"] = f"R{some}"
        revision_sheet[f"D6"] = modified_revision_date
        revision_sheet[f"E6"] = document_revisions[0].get("status")

   
    
    # General Information

    main_supply_lv = project_info.get("main_supply_lv")
    main_supply_lv_variation = project_info.get("main_supply_lv_variation")
    main_supply_lv_phase = project_info.get("main_supply_lv_phase")

    main_supply_mv = project_info.get("main_supply_mv")
    main_supply_mv_variation = project_info.get("main_supply_mv_variation")
    main_supply_mv_phase = project_info.get("main_supply_mv_phase")

    control_supply = project_info.get("control_supply")
    control_supply_variation = project_info.get("control_supply_variation")
    control_supply_phase = project_info.get("control_supply_phase")
    
    utility_supply = project_info.get("utility_supply")
    utility_supply_variation = project_info.get("utility_supply_variation")
    utility_supply_phase = project_info.get("utility_supply_phase")

    utility_supply_data = f"{utility_supply}, Variation: {utility_supply_variation}, {utility_supply_phase}"
    if utility_supply_variation == "NA":
        utility_supply_data = utility_supply

    control_supply_data = f"{control_supply}, Variation: {control_supply_variation}, {control_supply_phase}"
    if control_supply_variation == "NA" :
        control_supply_data = control_supply

    mv_data = f"{main_supply_mv}, Variation: {main_supply_mv_variation}, {main_supply_mv_phase}"
    if main_supply_mv == "NA" :
        mv_data = "Not Applicable"

    project_info_freq = project_info.get("frequency")
    preojct_info_freq_var = project_info.get("frequency_variation")
    project_info_frequency_data = f"{project_info_freq} Hz , Variation: {preojct_info_freq_var}"

    project_info_fault = project_info.get("fault_level")
    project_info_sec = project_info.get("sec")
    fault_data = f"{project_info_fault} KA, {project_info_sec} Sec"

    ambient_temperature_max_data = project_info.get("ambient_temperature_max")
    ambient_temperature_min_data = project_info.get("ambient_temperature_min")



    # Initialize variables
    variable1 = ""  # For Safe Area subpackage names
    variable2 = ""  # For other subpackage names

    # Loop through the subpackage array
    if len(project_package_data) >= 1:
        for main_package in project_package_data:
            for sub_package in main_package['sub_packages']:
                if sub_package['area_of_classification'] == 'Safe Area':
                    # Append to variable1 with a comma if it's not empty
                    if variable1:
                        variable1 += ", "
                    variable1 += sub_package['sub_package_name']
                else:
                    # Append to variable2 with a comma if it's not empty
                    if variable2:
                        variable2 += ", "
                    variable2 += sub_package['sub_package_name']

        
        design_basis_sheet["C4"] = project_package_data[0].get("main_package_name")
        design_basis_sheet["C5"] = variable1
        design_basis_sheet["C6"] = variable2
    
    else:
        design_basis_sheet["C4"] = "Not Applicable"
        design_basis_sheet["C5"] = "Not Applicable"
        design_basis_sheet["C6"] = "Not Applicable"

    area_classification_data = frappe.db.get_value("Project Main Package", {"revision_id":current_revision_id}, ["standard","zone","gas_group","temperature_class"])

    default_values = {
        "standard": "default_standard",  # Replace with your actual default value
        "zone": "default_zone",          # Replace with your actual default value
        "gas_group": "default_gas_group",# Replace with your actual default value
        "temperature_class": "default_temperature_class" # Replace with your actual default value
    }

    area_classification_data = [
        value if value is not None else default_values[field]
        for value, field in zip(area_classification_data, default_values.keys())
    ]

    design_basis_sheet["C7"] = f"Standard-{area_classification_data[0]}, {area_classification_data[1]}, Gas Group-{area_classification_data[2]}, Temperature Class-{area_classification_data[3]}"
    design_basis_sheet["C8"] = general_info.get("battery_limit")
    design_basis_sheet["C9"] = mv_data
    design_basis_sheet["C10"] = f"{main_supply_lv}, Variation: {main_supply_lv_variation}, {main_supply_lv_phase}"
    design_basis_sheet["C11"] = control_supply_data
    design_basis_sheet["C12"] = utility_supply_data
    design_basis_sheet["C13"] = project_info_frequency_data
    design_basis_sheet["C14"] = fault_data
    design_basis_sheet["C15"] = f"{ambient_temperature_max_data} Deg C"
    design_basis_sheet["C16"] = f"{ambient_temperature_min_data} Deg C"

    electrical_design_temp_data = project_info.get("electrical_design_temperature")
    design_basis_sheet["C17"] = f"{electrical_design_temp_data} Deg C"
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
    design_basis_sheet["E31"] = "Not Applicable"
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

    def clean_and_replace(value):
        # Clean the string by removing brackets and quotes
        cleaned_value = value.replace('[', '').replace(']', '').replace('"', '').strip()
        # Replace "NA" with "Not Applicable"
        if cleaned_value == "NA":
            return "Not Applicable"
        return cleaned_value




    motor_string = clean_and_replace(make_of_components.get("motor", "NA"))
    cable_string = clean_and_replace(make_of_components.get("cable", "NA"))
    lv_switchgear_string = clean_and_replace(make_of_components.get("lv_switchgear", "NA"))
    panel_enclosure_string = clean_and_replace(make_of_components.get("panel_enclosure", "NA"))
    variable_frequency_string = clean_and_replace(make_of_components.get("variable_frequency_speed_drive_vfd_vsd", "NA"))
    soft_starter_string = clean_and_replace(make_of_components.get("soft_starter", "NA"))
    plc_string = clean_and_replace(make_of_components.get("plc", "NA"))
    
    design_basis_sheet["E66"] = motor_string
    design_basis_sheet["E67"] = cable_string
    design_basis_sheet["E68"] = lv_switchgear_string
    design_basis_sheet["E69"] = panel_enclosure_string
    design_basis_sheet["E70"] = variable_frequency_string
    design_basis_sheet["E71"] = soft_starter_string
    design_basis_sheet["E72"] = plc_string

    """
        Common Configuration
    """
    switchgear_combination_data = common_configuration.get("switchgear_combination")
    mcc_switchgear_type = common_configuration.get("mcc_switchgear_type")
    if not division_name == "WWS SPG" and "Fuseless" not in mcc_switchgear_type:
        switchgear_combination_data =  "Not Applicable"

    dol_starter = common_configuration.get("dol_starter")
    design_basis_sheet["E74"] = na_To_string(dol_starter)
    design_basis_sheet["E75"] = common_configuration.get("star_delta_starter")
    design_basis_sheet["E76"] = common_configuration.get("ammeter")
    design_basis_sheet["E77"] = common_configuration.get("ammeter_configuration")
    design_basis_sheet["E78"] = common_configuration.get("mcc_switchgear_type")
    design_basis_sheet["E79"] = switchgear_combination_data
    design_basis_sheet["E80"] = common_configuration.get("pole")
    design_basis_sheet["E81"] = common_configuration.get("supply_feeder_standard")
    
    dm_standard = common_configuration.get("dm_standard")
    design_basis_sheet["E82"] = na_To_string(dm_standard)
    testing_standard = common_configuration.get("testing_standard")
    design_basis_sheet["E83"] = na_To_string(testing_standard)

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
    spare_terminal_data = common_configuration.get("spare_terminal")
    design_basis_sheet["E99"] = f"{spare_terminal_data} %"

    """
        Push Button Color
    """
    speed_increase_pb = common_configuration.get("speed_increase_pb")
    is_push_button_speed_selected = common_configuration.get("is_push_button_speed_selected")
    speed_decrease_pb = common_configuration.get("speed_decrease_pb")
    
    if(is_push_button_speed_selected == "0"):
        speed_increase_pb = "Not Applicable"
        speed_decrease_pb = "Not Applicable"

    push_button_start = common_configuration.get("push_button_start")
    push_button_stop = common_configuration.get("push_button_stop")

    design_basis_sheet["E101"] = na_To_string(push_button_start)
    design_basis_sheet["E102"] = na_To_string(push_button_stop)
    design_basis_sheet["E103"] = common_configuration.get("push_button_ess")
    design_basis_sheet["E104"] = speed_increase_pb
    design_basis_sheet["E105"] = speed_decrease_pb
    alarm_acknowledge_and_lamp_test = common_configuration.get(
        "alarm_acknowledge_and_lamp_test"
    )
    design_basis_sheet["E106"] = na_To_string(alarm_acknowledge_and_lamp_test)
    test_reset = common_configuration.get("test_reset")
    design_basis_sheet["E107"] = na_To_string(test_reset)

    """
        Selector Switch
    """
    selector_switch_applicable_data = common_configuration.get("selector_switch_applicable")
    selector_switch_locable_data = common_configuration.get("selector_switch_lockable").replace("'", "").strip()
    selector_switch__data = f"{selector_switch_applicable_data}, {selector_switch_locable_data}"

    if selector_switch_applicable_data == "Not Applicable":
        selector_switch__data = "Not Applicable"
        
    design_basis_sheet["E109"] = selector_switch__data

    """
        Indicating Lamp
    """
    design_basis_sheet["E111"] = common_configuration.get("running_open")
    design_basis_sheet["E112"] = common_configuration.get("stopped_closed")
    design_basis_sheet["E113"] = common_configuration.get("trip")

    """
        Field Motor Isolator(General Specifications)
    """
    field_motor_type = common_configuration.get("field_motor_type")
    field_motor_enclosure = common_configuration.get("field_motor_enclosure")
    field_motor_material = common_configuration.get("field_motor_material")
    field_motor_qty = common_configuration.get("field_motor_qty")
    field_motor_isolator_color_shade = common_configuration.get("field_motor_isolator_color_shade")
    field_motor_cable_entry = common_configuration.get("field_motor_cable_entry")
    field_motor_canopy_on_top = common_configuration.get("field_motor_canopy_on_top")

    design_basis_sheet["E115"] = na_To_string(field_motor_type)
    design_basis_sheet["E116"] = na_To_string(field_motor_enclosure)
    design_basis_sheet["E117"] = na_To_string(field_motor_material)
    design_basis_sheet["E118"] = na_To_string(field_motor_qty)
    design_basis_sheet["E119"] = na_To_string(field_motor_isolator_color_shade)
    design_basis_sheet["E121"] = na_To_string(field_motor_cable_entry)
    design_basis_sheet["E120"] = na_To_string(field_motor_canopy_on_top)

    """
        Local Push Button Station (General Specifications)				
    """

    lpbs_type = common_configuration.get("lpbs_type")
    lpbs_enclosure = common_configuration.get("lpbs_enclosure")
    lpbs_material = common_configuration.get("lpbs_material")
    lpbs_qty = common_configuration.get("lpbs_qty")
    lpbs_color_shade = common_configuration.get("lpbs_color_shade")
    lpbs_canopy_on_top = common_configuration.get("lpbs_canopy_on_top")
    lpbs_push_button_start_color = common_configuration.get("lpbs_push_button_start_color")
    lpbs_indication_lamp_start_color = common_configuration.get("lpbs_indication_lamp_start_color")
    lpbs_indication_lamp_stop_color = common_configuration.get("lpbs_indication_lamp_stop_color")
    lpbs_speed_increase = common_configuration.get("lpbs_speed_increase")
    lpbs_speed_decrease = common_configuration.get("lpbs_speed_decrease")

    design_basis_sheet["E123"] = na_To_string(lpbs_type)
    design_basis_sheet["E124"] = na_To_string(lpbs_enclosure)
    design_basis_sheet["E125"] = na_To_string(lpbs_material)
    design_basis_sheet["E126"] = na_To_string(lpbs_qty)
    design_basis_sheet["E127"] = na_To_string(lpbs_color_shade)
    design_basis_sheet["E128"] = na_To_string(lpbs_canopy_on_top)
    design_basis_sheet["E129"] = na_To_string(lpbs_push_button_start_color)
    design_basis_sheet["E130"] = na_To_string(lpbs_indication_lamp_start_color)
    design_basis_sheet["E131"] = na_To_string(lpbs_indication_lamp_stop_color)
    design_basis_sheet["E132"] = na_To_string(lpbs_speed_increase)
    design_basis_sheet["E133"] = na_To_string(lpbs_speed_decrease)

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
    design_basis_sheet["E148"] = common_configuration.get("earth_bus_busbar_position")
    design_basis_sheet["E149"] = common_configuration.get("earth_bus_material")
    design_basis_sheet["E150"] = common_configuration.get("earth_bus_current_density")
    design_basis_sheet["E151"] = common_configuration.get("earth_bus_rating_of_busbar")

    """
        Metering for Feeder
    """
    metering_for_feeders = common_configuration.get("metering_for_feeders")
    design_basis_sheet["E153"] = na_To_string(metering_for_feeders)

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
    apfc_relay = common_configuration.get("apfc_relay")
    if apfc_relay == "NA":
        apfc_relay = "Not Applicable"
    else :
        apfc_relay = f"{apfc_relay} Stage"

    design_basis_sheet["E162"] = apfc_relay

    """
        Power Cable
    """
    design_basis_sheet["E164"] = cable_tray_layout.get("number_of_cores")
    design_basis_sheet["E165"] = cable_tray_layout.get("specific_requirement")
    design_basis_sheet["E166"] = cable_tray_layout.get("type_of_insulation")
    design_basis_sheet["E167"] = cable_tray_layout.get("color_scheme")
    motor_voltage_drop_during_starting = cable_tray_layout.get(
        "motor_voltage_drop_during_starting"
    )
    design_basis_sheet["E168"] = f"{motor_voltage_drop_during_starting} %"
    motor_voltage_drop_during_running = cable_tray_layout.get(
        "motor_voltage_drop_during_running"
    )
    design_basis_sheet["E169"] = f"{motor_voltage_drop_during_running} %"
    copper_conductor = cable_tray_layout.get("copper_conductor")
    aluminium_conductor = cable_tray_layout.get("aluminium_conductor")
    design_basis_sheet["E170"] = f"{copper_conductor} Sq. mm"
    design_basis_sheet["E171"] = f"{aluminium_conductor} Sq. mm"
    design_basis_sheet["E172"] = cable_tray_layout.get("voltage_grade")
    design_basis_sheet["E173"] = cable_tray_layout.get("touching_factor_air")
    design_basis_sheet["E174"] = cable_tray_layout.get(
        "ambient_temp_factor_air"
    )
    design_basis_sheet["E175"] = cable_tray_layout.get("derating_factor_air")
    design_basis_sheet["E176"] = cable_tray_layout.get("touching_factor_burid")
    design_basis_sheet["E177"] = cable_tray_layout.get(
        "ambient_temp_factor_burid"
    )
    design_basis_sheet["E178"] = cable_tray_layout.get("derating_factor_burid")

    """
        Gland
    """
    design_basis_sheet["E180"] = cable_tray_layout.get("gland_make")
    design_basis_sheet["E181"] = cable_tray_layout.get("moc")
    design_basis_sheet["E182"] = cable_tray_layout.get("type_of_gland")

    #logic for Gland Type

    safe_area_gland_data = "Not Applicable"
    hazardous_area_gland_data = "Not Applicable"

    if len(project_package_data) > 0:
        for main_package in project_package_data:
            for sub_package in main_package['sub_packages']:
                if sub_package['area_of_classification'] == 'Safe Area':
                    # Append to variable1 with a comma if it's not empty
                    safe_area_gland_data = "Weatherproof"
                else:
                    # Append to variable2 with a comma if it's not empty
                    hazardous_area_gland_data = "Flameproof"

    design_basis_sheet["E183"] = safe_area_gland_data
    design_basis_sheet["E184"] = hazardous_area_gland_data

    """
        Cable Tray
    """

    design_basis_sheet["E186"] = cable_tray_layout.get("future_space_on_trays")
    design_basis_sheet["E187"] = cable_tray_layout.get("cable_placement")
    design_basis_sheet["E188"] = cable_tray_layout.get("orientation")
    vertical_distance_data = cable_tray_layout.get("vertical_distance")
    design_basis_sheet["E189"] = f"{vertical_distance_data} mm"
    horizontal_distance_data = cable_tray_layout.get("horizontal_distance")
    design_basis_sheet["E190"] = f"{horizontal_distance_data} mm"
    dry_area = cable_tray_layout.get("dry_area")
    wet_area = cable_tray_layout.get("wet_area")
    design_basis_sheet["E191"] = na_To_string(dry_area)
    design_basis_sheet["E192"] = na_To_string(wet_area)

    """
        Earthing
    """
    design_basis_sheet["E194"] = earthing_layout_data.get("earthing_system")

    earth_strip = earthing_layout_data.get("earth_strip")
    earth_pit = earthing_layout_data.get("earth_pit")
    design_basis_sheet["E195"] = na_To_string(earth_strip)
    design_basis_sheet["E196"] = na_To_string(earth_pit)

    soil_resistivity_data = earthing_layout_data.get("soil_resistivity")
    design_basis_sheet["E197"] = f"{soil_resistivity_data} ohm"


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

            
            indication_lamp_led_data = panel_data.get("is_led_type_lamp_selected")
            if indication_lamp_led_data == "NA":
                indication_lamp_led_data = "OFF"
            else:
                indication_lamp_led_data = "ON"
            
            others_data = panel_data.get("led_type_other_input")
            indication_data = indication_lamp_led_data
            if others_data and not others_data == "NA":
                indication_data = f"{indication_lamp_led_data}, {others_data}"

            indication_lamp_led_data = panel_data.get("is_led_type_lamp_selected")
            if indication_lamp_led_data == "NA":
                indication_data = "OFF"
            else:
                indication_data = "ON"
            
            panel_sheet["E6"] = indication_data
            current_transformer_coating = panel_data.get("current_transformer_coating")
            current_transformer_number = panel_data.get("current_transformer_number")
            panel_sheet["E7"] = na_To_string(current_transformer_coating)
            panel_sheet["E8"] = na_To_string(current_transformer_number)
            control_transformer_coating = panel_data.get("control_transformer_coating")
            panel_sheet["E9"] = na_To_string(control_transformer_coating)
            control_transformer_configuration = panel_data.get("control_transformer_configuration")
            panel_sheet["E10"] = na_To_string(control_transformer_configuration)
            panel_sheet["E11"] = panel_data.get("alarm_annunciator")

            """
                Metering Instruments for Incomer				
            """

            panel_sheet["E13"] = (
                f"Analog - {panel_data.get('mi_analog')} ; Digital - {panel_data.get('mi_digital')} ; Communication Protocol - {panel_data.get('mi_communication_protocol')}"
            )

            """
                General Arrangement				
            """


            panel_sheet["E15"] = panel_data.get("ga_moc_material")
            ga_moc_thickness_door = panel_data.get("ga_moc_thickness_door")
            panel_sheet["E16"] = (f"{ga_moc_thickness_door} mm")
            ga_moc_thickness_covers = panel_data.get("ga_moc_thickness_covers")
            panel_sheet["E17"] = (f"{ga_moc_thickness_covers} mm")
            ga_data = f"{panel_data.get('ga_mcc_compartmental'), {panel_data.get('ga_mcc_construction_front_type')}, {panel_data.get('ga_mcc_construction_drawout_type')}, {panel_data.get('ga_mcc_construction_type')}}"
            ga_data = ga_data.replace("{", "").replace("}", "").replace("'","").replace("(", "").replace(")", "").strip()
            panel_sheet["E18"] = ga_data
            panel_sheet["E19"] = panel_data.get("busbar_material_of_construction")
            panel_sheet["E20"] = panel_data.get("ga_current_density")
            panel_sheet["E21"] = panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height =  panel_data.get("ga_panel_mounting_height")
            panel_sheet["E22"] = (f"{ga_panel_mounting_height} mm")

            is_marshalling_section_selected = panel_data.get("is_marshalling_section_selected")
            is_cable_alley_section_selected = panel_data.get("is_cable_alley_section_selected")
            is_power_and_bus_separation_section_selected = panel_data.get("is_power_and_bus_separation_section_selected")
            is_both_side_extension_section_selected = panel_data.get("is_both_side_extension_section_selected")
            panel_sheet["E23"] = number_To_string(is_marshalling_section_selected)
            panel_sheet["E24"] = number_To_string(is_cable_alley_section_selected)
            panel_sheet["E25"] = number_To_string(is_power_and_bus_separation_section_selected)
            panel_sheet["E26"] = number_To_string(is_both_side_extension_section_selected)
            
            panel_sheet["E27"] = panel_data.get("ga_gland_plate_3mm_drill_type")
            panel_sheet["E28"] = panel_data.get("ga_gland_plate_3mm_attachment_type")
            panel_sheet["E29"] = panel_data.get("ga_busbar_chamber_position")
            panel_sheet["E30"] = panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            panel_sheet["E31"] = panel_data.get("ga_enclosure_protection_degree")
            panel_sheet["E32"] = panel_data.get("ga_cable_entry_position")

            """
                Painting / Powder Coating			
            """
            panel_sheet["E34"] = panel_data.get("ppc_painting_standards")
            panel_sheet["E35"] = panel_data.get("ppc_interior_and_exterior_paint_shade")
            panel_sheet["E36"] = panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
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
            boiler_mcc_power_vac_data = panel_data.get('boiler_power_supply_vac')
            boiler_mcc_power_supply_data = (f"{panel_data.get('boiler_power_supply_vac')}, {panel_data.get('boiler_power_supply_phase')}, {panel_data.get('boiler_power_supply_frequency')} Hz")
            if boiler_mcc_power_vac_data == "NA":
                boiler_mcc_power_supply_data = "Not Applicable"

            boiler_mcc_control_vac_data = panel_data.get('boiler_control_supply_vac')
            boiler_mcc_control_supply_data = (f"{panel_data.get('boiler_control_supply_vac')}, {panel_data.get('boiler_control_supply_phase')}, {panel_data.get('boiler_control_supply_frequency')} Hz")
            if boiler_mcc_control_vac_data == "NA":
                boiler_mcc_control_supply_data = "Not Applicable"

            # Punching Details for Boiler
            boiler_model = panel_data.get("boiler_model")
            boiler_fuel = panel_data.get("boiler_fuel")
            boiler_year = panel_data.get("boiler_year")
            boiler_evaporation = panel_data.get("boiler_evaporation")
            boiler_output = panel_data.get("boiler_output")
            boiler_connected_load = panel_data.get("boiler_connected_load")
            boiler_design_pressure = panel_data.get("boiler_design_pressure")

            is_boiler_selected =  frappe.db.get_value("MCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_boiler_selected"])
            if is_boiler_selected == 0:
                boiler_mcc_power_supply_data = "NA"
                boiler_mcc_control_supply_data = "NA"
                boiler_model = "NA"
                boiler_fuel = "NA"
                boiler_year = "NA"
                boiler_evaporation = "NA"
                boiler_output = "NA"
                boiler_connected_load = "NA"
                boiler_design_pressure = "NA"

            panel_sheet["E44"] = na_To_string(boiler_model)
            panel_sheet["E45"] = na_To_string(boiler_fuel)
            panel_sheet["E46"] = na_To_string(boiler_year)
            panel_sheet["E47"] = na_To_string(boiler_mcc_power_supply_data)
            panel_sheet["E48"] = na_To_string(boiler_mcc_control_supply_data)
            panel_sheet["E49"] = na_To_string(boiler_evaporation)
            panel_sheet["E50"] = na_To_string(boiler_output)
            panel_sheet["E51"] = na_To_string(boiler_connected_load)
            panel_sheet["E52"] = na_To_string(boiler_design_pressure)

            # Punching Details for Heater

            heater_mcc_power_vac_data = panel_data.get('boiler_power_supply_vac')
            heater_mcc_power_supply_data = (f"{panel_data.get('heater_power_supply_vac')}, {panel_data.get('heater_power_supply_phase')}, {panel_data.get('heater_power_supply_frequency')} Hz")
            if heater_mcc_power_vac_data == "NA":
                heater_mcc_power_supply_data = "Not Applicable"

            heater_mcc_control_vac_data = panel_data.get('boiler_control_supply_vac')
            heater_mcc_control_supply_data = (f"{panel_data.get('heater_control_supply_vac')}, {panel_data.get('heater_control_supply_phase')}, {panel_data.get('heater_control_supply_frequency')}")
            if heater_mcc_control_vac_data == "NA":
                heater_mcc_control_supply_data = "Not Applicable"

            heater_model = panel_data.get("heater_model")
            heater_fuel = panel_data.get("heater_fuel")
            heater_year = panel_data.get("heater_year")
            heater_evaporation = panel_data.get("heater_evaporation")
            heater_output = panel_data.get("heater_output")
            heater_connected_load = panel_data.get("heater_connected_load")
            heater_temperature = panel_data.get("heater_temperature")

            is_heater_selected =  frappe.db.get_value("MCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_heater_selected"])
            if is_heater_selected == 0:
                heater_mcc_power_supply_data = "NA"
                heater_mcc_control_supply_data = "NA"
                heater_model = "NA"
                heater_fuel = "NA"
                heater_year = "NA"
                heater_evaporation = "NA"
                heater_output = "NA"
                heater_connected_load = "NA"
                heater_temperature = "NA"

            panel_sheet["E54"] = na_To_string(heater_model)
            panel_sheet["E55"] = na_To_string(heater_fuel)
            panel_sheet["E56"] = na_To_string(heater_year)
            panel_sheet["E57"] = na_To_string(heater_mcc_power_supply_data)
            panel_sheet["E58"] = na_To_string(heater_mcc_control_supply_data)
            panel_sheet["E59"] = na_To_string(heater_evaporation)
            panel_sheet["E60"] = na_To_string(heater_output)
            panel_sheet["E61"] = na_To_string(heater_connected_load)
            panel_sheet["E62"] = na_To_string(heater_temperature)

            # Name Plate Details for SPG
            spg_name_plate_unit_name = panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = panel_data.get("spg_name_plate_manufacturing_year")
            spg_name_plate_weight = panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = panel_data.get("spg_name_plate_part_code")

            panel_sheet["E64"] = na_To_string(spg_name_plate_unit_name)
            panel_sheet["E65"] = na_To_string(spg_name_plate_capacity)
            panel_sheet["E66"] = na_To_string(spg_name_plate_manufacturing_year)
            panel_sheet["E67"] = na_To_string(spg_name_plate_weight)
            panel_sheet["E68"] = na_To_string(spg_name_plate_oc_number)
            panel_sheet["E69"] = na_To_string(spg_name_plate_part_code)

        if project_panel.get("panel_main_type") == "PCC":
            panel_sheet = template_workbook.copy_worksheet(pcc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            panel_data = project_panel.get("panelData")

            """
                Selection Details
            """
            panel_sheet["E5"] = (
                f"Upto - {panel_data.get('incomer_ampere')} - {panel_data.get('incomer_pole')} Pole {panel_data.get('incomer_type')} > {panel_data.get('incomer_above_ampere')} - {panel_data.get('incomer_above_pole')} Pole {panel_data.get('incomer_above_type')}"
            )

            indication_lamp_led_data = panel_data.get("is_led_type_lamp_selected")
            indication_lamp_led_data = number_To_string(indication_lamp_led_data)
            others_data = panel_data.get("led_type_other_input")
            indication_data = indication_lamp_led_data
            if others_data and not others_data == "NA":
                indication_data = f"{indication_lamp_led_data}, {others_data}"

            indication_lamp_led_data = panel_data.get("is_led_type_lamp_selected")
            if indication_lamp_led_data == "NA":
                indication_data = "OFF"
            else:
                indication_data = "ON"
                
            panel_sheet["E6"] = indication_data
            control_transformer_coating = panel_data.get("control_transformer_coating")
            panel_sheet["E7"] = na_To_string(control_transformer_coating)
            control_transformer_configuration = panel_data.get("control_transformer_configuration")
            panel_sheet["E8"] = na_To_string(control_transformer_configuration)
            panel_sheet["E9"] = panel_data.get("alarm_annunciator")

            """
                Metering Instruments for Incomer				
            """
            panel_sheet["E11"] = (
                f"Analog - {panel_data.get('mi_analog')} ; Digital - { panel_data.get('mi_digital')} ; Communication Protocol - { panel_data.get('mi_communication_protocol') }"
            )

            """
                General Arrangement
            """
            panel_sheet["E13"] = panel_data.get("ga_moc_material")
            ga_moc_thickness_door = panel_data.get("ga_moc_thickness_door")            
            panel_sheet["E14"] = f"{ga_moc_thickness_door} mm"
            ga_moc_thickness_covers = panel_data.get("ga_moc_thickness_covers")
            panel_sheet["E15"] = f"{ga_moc_thickness_covers} mm"

            ga_data = f"{panel_data.get('ga_pcc_compartmental'), {panel_data.get('ga_pcc_construction_front_type')}, {panel_data.get('ga_pcc_construction_drawout_type')}, {panel_data.get('ga_pcc_construction_type')}}"
            ga_data = ga_data.replace("{", "").replace("}", "").replace("'","").replace("(", "").replace(")", "").strip()            
            panel_sheet["E16"] = ga_data

            panel_sheet["E17"] = panel_data.get("busbar_material_of_construction")
            panel_sheet["E18"] = panel_data.get("ga_current_density")
            panel_sheet["E19"] = panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height =  panel_data.get("ga_panel_mounting_height")
            panel_sheet["E20"] = f"{ga_panel_mounting_height} mm"

            is_marshalling_section_selected = panel_data.get("is_marshalling_section_selected")
            is_cable_alley_section_selected = panel_data.get("is_cable_alley_section_selected")
            is_power_and_bus_separation_section_selected = panel_data.get("is_power_and_bus_separation_section_selected")
            is_both_side_extension_section_selected = panel_data.get("is_both_side_extension_section_selected")
            panel_sheet["E21"] = number_To_string(is_marshalling_section_selected)
            panel_sheet["E22"] = number_To_string(is_cable_alley_section_selected)
            panel_sheet["E23"] = number_To_string(is_power_and_bus_separation_section_selected)
            panel_sheet["E24"] = number_To_string(is_both_side_extension_section_selected)

            panel_sheet["E25"] = panel_data.get("ga_gland_plate_3mm_drill_type")
            panel_sheet["E26"] = panel_data.get("ga_gland_plate_3mm_attachment_type")
            panel_sheet["E27"] = panel_data.get("ga_busbar_chamber_position")
            panel_sheet["E28"] = panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            panel_sheet["E29"] = panel_data.get("ga_enclosure_protection_degree")
            panel_sheet["E30"] = panel_data.get("ga_cable_entry_position")

            """
                Painting / Powder Coating
            """
            panel_sheet["E32"] = panel_data.get("ppc_painting_standards")
            panel_sheet["E33"] = panel_data.get("ppc_interior_and_exterior_paint_shade")
            panel_sheet["E34"] = panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            panel_sheet["E35"] = panel_data.get("ppc_minimum_coating_thickness")
            panel_sheet["E36"] = panel_data.get("ppc_base_frame_paint_shade")
            panel_sheet["E37"] = panel_data.get("ppc_pretreatment_panel_standard")

            """
                Punching Details
            """

            boiler_pcc_power_vac_data = panel_data.get('boiler_power_supply_vac')
            boiler_pcc_power_supply_data = (f"{panel_data.get('boiler_power_supply_vac')}, {panel_data.get('boiler_power_supply_phase')}, {panel_data.get('boiler_power_supply_frequency')} Hz")
            if boiler_pcc_power_vac_data == "NA":
                boiler_pcc_power_supply_data = "Not Applicable"

            boiler_pcc_control_vac_data = panel_data.get('boiler_control_supply_vac')
            boiler_pcc_control_supply_data = (f"{panel_data.get('boiler_control_supply_vac')}, {panel_data.get('boiler_control_supply_phase')}, {panel_data.get('boiler_control_supply_frequency')} Hz")
            if boiler_pcc_control_vac_data == "NA":
                boiler_pcc_control_supply_data = "Not Applicable"
            
            # Punching Details for Boiler

            boiler_model = panel_data.get("boiler_model")
            boiler_fuel = panel_data.get("boiler_fuel")
            boiler_year = panel_data.get("boiler_year")
            boiler_evaporation = panel_data.get("boiler_evaporation")
            boiler_output = panel_data.get("boiler_output")
            boiler_connected_load = panel_data.get("boiler_connected_load")
            boiler_design_pressure = panel_data.get("boiler_design_pressure")

            is_boiler_selected =  frappe.db.get_value("PCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_boiler_selected"])
            if is_boiler_selected == 0:
                boiler_pcc_power_supply_data = "NA"
                boiler_pcc_control_supply_data = "NA"
                boiler_model = "NA"
                boiler_fuel = "NA"
                boiler_year = "NA"
                boiler_evaporation = "NA"
                boiler_output = "NA"
                boiler_connected_load = "NA"
                boiler_design_pressure = "NA"


            panel_sheet["E40"] = na_To_string(boiler_model)
            panel_sheet["E41"] = na_To_string(boiler_fuel)
            panel_sheet["E42"] = na_To_string(boiler_year)
            panel_sheet["E43"] = na_To_string(boiler_pcc_power_supply_data)
            panel_sheet["E44"] = na_To_string(boiler_pcc_control_supply_data)
            panel_sheet["E45"] = na_To_string(boiler_evaporation)
            panel_sheet["E46"] = na_To_string(boiler_output)
            panel_sheet["E47"] = na_To_string(boiler_connected_load)
            panel_sheet["E48"] = na_To_string(boiler_design_pressure)

            # Punching Details for Heater

            heater_pcc_power_vac_data = panel_data.get('boiler_power_supply_vac')
            heater_pcc_power_supply_data = (f"{panel_data.get('heater_power_supply_vac')}, {panel_data.get('heater_power_supply_phase')}, {panel_data.get('heater_power_supply_frequency')} Hz")
            if heater_pcc_power_vac_data == "NA":
                heater_pcc_power_supply_data = "Not Applicable"

            heater_pcc_control_vac_data = panel_data.get('boiler_control_supply_vac')
            heater_pcc_control_supply_data = (f"{panel_data.get('heater_control_supply_vac')}, {panel_data.get('heater_control_supply_phase')}, {panel_data.get('heater_control_supply_frequency')}")
            if heater_pcc_control_vac_data == "NA":
                heater_pcc_control_supply_data = "Not Applicable"

            heater_model = panel_data.get("heater_model")
            heater_fuel = panel_data.get("heater_fuel")
            heater_year = panel_data.get("heater_year")
            heater_evaporation = panel_data.get("heater_evaporation")
            heater_output = panel_data.get("heater_output")
            heater_connected_load = panel_data.get("heater_connected_load")
            heater_temperature = panel_data.get("heater_temperature")

            is_heater_selected =  frappe.db.get_value("PCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_heater_selected"])
            if is_heater_selected == 0:
                heater_pcc_power_supply_data = "NA"
                heater_pcc_control_supply_data = "NA"
                heater_model = "NA"
                heater_fuel = "NA"
                heater_year = "NA"
                heater_evaporation = "NA"
                heater_output = "NA"
                heater_connected_load = "NA"
                heater_temperature = "NA"


            panel_sheet["E50"] = na_To_string(heater_model)
            panel_sheet["E51"] = na_To_string(heater_fuel)
            panel_sheet["E52"] = na_To_string(heater_year)
            panel_sheet["E53"] = na_To_string(heater_pcc_power_supply_data)
            panel_sheet["E54"] = na_To_string(heater_pcc_control_supply_data)
            panel_sheet["E55"] = na_To_string(heater_evaporation)
            panel_sheet["E56"] = na_To_string(heater_output)
            panel_sheet["E57"] = na_To_string(heater_connected_load)
            panel_sheet["E58"] = na_To_string(heater_temperature)

            """
                Name Plate Details for SPG
            """
            spg_name_plate_unit_name = panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = panel_data.get("spg_name_plate_manufacturing_year")
            spg_name_plate_weight = panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = panel_data.get("spg_name_plate_part_code")

            panel_sheet["E60"] = na_To_string(spg_name_plate_unit_name)
            panel_sheet["E61"] = na_To_string(spg_name_plate_capacity)
            panel_sheet["E62"] = na_To_string(spg_name_plate_manufacturing_year)
            panel_sheet["E63"] = na_To_string(spg_name_plate_weight)
            panel_sheet["E64"] = na_To_string(spg_name_plate_oc_number)
            panel_sheet["E65"] = na_To_string(spg_name_plate_part_code)

        if project_panel.get("panel_main_type") == "MCC cum PCC":
            panel_sheet = template_workbook.copy_worksheet(mcc_cum_plc_sheet)
            panel_sheet.title = project_panel.get("panel_name")
            mcc_panel_data = project_panel.get("mccPanelData")
            plc_panel_data = project_panel.get("plcPanelData")

            # """
            #     Selection Details
            # """
            panel_sheet["E5"] = (
                f"Upto - {mcc_panel_data.get('incomer_ampere')} - {mcc_panel_data.get('incomer_pole')} Pole {mcc_panel_data.get('incomer_type')} > {mcc_panel_data.get('incomer_above_ampere')} - {mcc_panel_data.get('incomer_above_pole')} Pole {mcc_panel_data.get('incomer_above_type')}"
            )
            
            indication_lamp_led_data = mcc_panel_data.get("is_led_type_lamp_selected")
            indication_lamp_led_data = number_To_string(indication_lamp_led_data)
            others_data = mcc_panel_data.get("led_type_other_input")
            indication_data = indication_lamp_led_data
            if others_data and not others_data == "NA":
                indication_data = f"{indication_lamp_led_data}, {others_data}"

            indication_lamp_led_data = mcc_panel_data.get("is_led_type_lamp_selected")
            if indication_lamp_led_data == "NA":
                indication_data = "OFF"
            else:
                indication_data = "ON"

            panel_sheet["E6"] = indication_data
            current_transformer_coating = mcc_panel_data.get("current_transformer_coating")
            current_transformer_number = mcc_panel_data.get("current_transformer_number")
            panel_sheet["E7"] = na_To_string(current_transformer_coating)
            panel_sheet["E8"] = na_To_string(current_transformer_number)
            control_transformer_coating = mcc_panel_data.get("control_transformer_coating")
            panel_sheet["E9"] = na_To_string(control_transformer_coating)
            control_transformer_configuration = mcc_panel_data.get("control_transformer_configuration")
            panel_sheet["E10"] = na_To_string(control_transformer_configuration)
            panel_sheet["E11"] = mcc_panel_data.get("alarm_annunciator")

            """
                Metering Instruments for Incomer				
            """
            panel_sheet["E13"] = (
                f"Analog - {mcc_panel_data.get('mi_analog')} ; Digital - { mcc_panel_data.get('mi_digital')} ; Communication Protocol - { mcc_panel_data.get('mi_communication_protocol') }"
            )

            """
                General Arrangement				
            """
            panel_sheet["E15"] = mcc_panel_data.get("ga_moc_material")
            ga_moc_thickness_door = mcc_panel_data.get("ga_moc_thickness_door")
            panel_sheet["E16"] = (f"{ga_moc_thickness_door} mm")
            ga_moc_thickness_covers = mcc_panel_data.get("ga_moc_thickness_covers")
            panel_sheet["E17"] = (f"{ga_moc_thickness_covers} mm")
            
            ga_data = f"{mcc_panel_data.get('ga_mcc_compartmental'), {mcc_panel_data.get('ga_mcc_construction_front_type')}, {mcc_panel_data.get('ga_mcc_construction_drawout_type')}, {mcc_panel_data.get('ga_mcc_construction_type')}}"
            ga_data = ga_data.replace("{", "").replace("}", "").replace("'","").replace("(", "").replace(")", "").strip()
            panel_sheet["E18"] = ga_data
            
            panel_sheet["E19"] = mcc_panel_data.get("busbar_material_of_construction")
            panel_sheet["E20"] = mcc_panel_data.get("ga_current_density")
            panel_sheet["E21"] = mcc_panel_data.get("ga_panel_mounting_frame")
            ga_panel_mounting_height = mcc_panel_data.get("ga_panel_mounting_height")
            panel_sheet["E22"] = (f"{ga_panel_mounting_height} mm")
            
            
            is_marshalling_section_selected = mcc_panel_data.get("is_marshalling_section_selected")
            is_cable_alley_section_selected = mcc_panel_data.get("is_cable_alley_section_selected")
            is_power_and_bus_separation_section_selected = mcc_panel_data.get("is_power_and_bus_separation_section_selected")
            is_both_side_extension_section_selected = mcc_panel_data.get("is_both_side_extension_section_selected")
            panel_sheet["E23"] = number_To_string(is_marshalling_section_selected)
            panel_sheet["E24"] = number_To_string(is_cable_alley_section_selected)
            panel_sheet["E25"] = number_To_string(is_power_and_bus_separation_section_selected)
            panel_sheet["E26"] = number_To_string(is_both_side_extension_section_selected)
            
            
            panel_sheet["E27"] = mcc_panel_data.get("ga_gland_plate_3mm_drill_type")
            panel_sheet["E28"] = mcc_panel_data.get(
                "ga_gland_plate_3mm_attachment_type"
            )
            panel_sheet["E29"] = mcc_panel_data.get("ga_busbar_chamber_position")
            panel_sheet["E30"] = mcc_panel_data.get(
                "ga_power_and_control_busbar_separation"
            )
            panel_sheet["E31"] = mcc_panel_data.get("ga_enclosure_protection_degree")
            panel_sheet["E32"] = mcc_panel_data.get("ga_cable_entry_position")

            """
                Painting / Powder Coating			
            """
            panel_sheet["E34"] = mcc_panel_data.get("ppc_painting_standards")
            panel_sheet["E35"] = mcc_panel_data.get(
                "ppc_interior_and_exterior_paint_shade"
            )
            panel_sheet["E36"] = mcc_panel_data.get(
                "ppc_component_mounting_plate_paint_shade"
            )
            panel_sheet["E37"] = mcc_panel_data.get("ppc_minimum_coating_thickness")
            panel_sheet["E38"] = mcc_panel_data.get("ppc_base_frame_paint_shade")
            panel_sheet["E39"] = mcc_panel_data.get("ppc_pretreatment_panel_standard")

            """
                VFD
            """
            panel_sheet["E41"] = mcc_panel_data.get("vfd_auto_manual_selection")

            """
                Punching Details
            """

            # Punching Details for Boiler

            boiler_mcc_power_vac_data = mcc_panel_data.get('boiler_power_supply_vac')
            boiler_mcc_power_supply_data = (f"{mcc_panel_data.get('boiler_power_supply_vac')}, {mcc_panel_data.get('boiler_power_supply_phase')}, {mcc_panel_data.get('boiler_power_supply_frequency')} Hz")
            if boiler_mcc_power_vac_data == "NA":
                boiler_mcc_power_supply_data = "Not Applicable"

            boiler_mcc_control_vac_data = mcc_panel_data.get('boiler_control_supply_vac')
            boiler_mcc_control_supply_data = (f"{mcc_panel_data.get('boiler_control_supply_vac')}, {mcc_panel_data.get('boiler_control_supply_phase')}, {mcc_panel_data.get('boiler_control_supply_frequency')} Hz")
            if boiler_mcc_control_vac_data == "NA":
                boiler_mcc_control_supply_data = "Not Applicable"

            # Punching Details for Boiler
            boiler_model = mcc_panel_data.get("boiler_model")
            boiler_fuel = mcc_panel_data.get("boiler_fuel")
            boiler_year = mcc_panel_data.get("boiler_year")
            boiler_evaporation = mcc_panel_data.get("boiler_evaporation")
            boiler_output = mcc_panel_data.get("boiler_output")
            boiler_connected_load = mcc_panel_data.get("boiler_connected_load")
            boiler_design_pressure = mcc_panel_data.get("boiler_design_pressure")

            is_boiler_selected =  frappe.db.get_value("MCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_boiler_selected"])
            if is_boiler_selected == 0:
                boiler_pcc_power_supply_data = "NA"
                boiler_pcc_control_supply_data = "NA"
                boiler_model = "NA"
                boiler_fuel = "NA"
                boiler_year = "NA"
                boiler_evaporation = "NA"
                boiler_output = "NA"
                boiler_connected_load = "NA"
                boiler_design_pressure = "NA"


            panel_sheet["E44"] = na_To_string(boiler_model)
            panel_sheet["E45"] = na_To_string(boiler_fuel)
            panel_sheet["E46"] = na_To_string(boiler_year)
            panel_sheet["E47"] = na_To_string(boiler_mcc_power_supply_data)
            panel_sheet["E48"] = na_To_string(boiler_mcc_control_supply_data)
            panel_sheet["E49"] = na_To_string(boiler_evaporation)
            panel_sheet["E50"] = na_To_string(boiler_output)
            panel_sheet["E51"] = na_To_string(boiler_connected_load)
            panel_sheet["E52"] = na_To_string(boiler_design_pressure)

            # # Punching Details for Heater
            heater_mcc_power_vac_data = mcc_panel_data.get('boiler_power_supply_vac')
            heater_mcc_power_supply_data = (f"{mcc_panel_data.get('heater_power_supply_vac')}, {mcc_panel_data.get('heater_power_supply_phase')}, {mcc_panel_data.get('heater_power_supply_frequency')} Hz")
            if heater_mcc_power_vac_data == "NA":
                heater_mcc_power_supply_data = "Not Applicable"

            heater_mcc_control_vac_data = mcc_panel_data.get('boiler_control_supply_vac')
            heater_mcc_control_supply_data = (f"{mcc_panel_data.get('heater_control_supply_vac')}, {mcc_panel_data.get('heater_control_supply_phase')}, {mcc_panel_data.get('heater_control_supply_frequency')}")
            if heater_mcc_control_vac_data == "NA":
                heater_mcc_control_supply_data = "Not Applicable"

            heater_model = mcc_panel_data.get("heater_model")
            heater_fuel = mcc_panel_data.get("heater_fuel")
            heater_year = mcc_panel_data.get("heater_year")
            heater_evaporation = mcc_panel_data.get("heater_evaporation")
            heater_output = mcc_panel_data.get("heater_output")
            heater_connected_load = mcc_panel_data.get("heater_connected_load")
            heater_temperature = mcc_panel_data.get("heater_temperature")

            heater_selected =  frappe.db.get_value("MCC Panel",{"revision_id": current_revision_id},["is_punching_details_for_heater_selected"])
            if is_boiler_selected == 0:
                heater_pcc_power_supply_data = "NA"
                heater_pcc_control_supply_data = "NA"
                heater_model = "NA"
                heater_fuel = "NA"
                heater_year = "NA"
                heater_evaporation = "NA"
                heater_output = "NA"
                heater_connected_load = "NA"
                heater_design_pressure = "NA"


            panel_sheet["E54"] = na_To_string(heater_model)
            panel_sheet["E55"] = na_To_string(heater_fuel)
            panel_sheet["E56"] = na_To_string(heater_year)
            panel_sheet["E57"] = na_To_string(heater_mcc_power_supply_data)
            panel_sheet["E58"] = na_To_string(heater_mcc_control_supply_data)
            panel_sheet["E59"] = na_To_string(heater_evaporation)
            panel_sheet["E60"] = na_To_string(heater_output)
            panel_sheet["E61"] = na_To_string(heater_connected_load)
            panel_sheet["E62"] = na_To_string(heater_temperature)


            # # Name Plate Details for SPG
            spg_name_plate_unit_name = mcc_panel_data.get("spg_name_plate_unit_name")
            spg_name_plate_capacity = mcc_panel_data.get("spg_name_plate_capacity")
            spg_name_plate_manufacturing_year = mcc_panel_data.get("spg_name_plate_manufacturing_year")
            spg_name_plate_weight = mcc_panel_data.get("spg_name_plate_weight")
            spg_name_plate_oc_number = mcc_panel_data.get("spg_name_plate_oc_number")
            spg_name_plate_part_code = mcc_panel_data.get("spg_name_plate_part_code")

            panel_sheet["E64"] = na_To_string(spg_name_plate_unit_name)
            panel_sheet["E65"] = na_To_string(spg_name_plate_capacity)
            panel_sheet["E66"] = na_To_string(spg_name_plate_manufacturing_year)
            panel_sheet["E67"] = na_To_string(spg_name_plate_weight)
            panel_sheet["E68"] = na_To_string(spg_name_plate_oc_number)
            panel_sheet["E69"] = na_To_string(spg_name_plate_part_code)

            """
                PLC
            """
            # UPS
            ups_scope = plc_panel_data.get("ups_scope")
            ups_type = plc_panel_data.get("ups_type")
            ups_battery_type = plc_panel_data.get("ups_battery_type")
            is_ups_battery_mounting_rack_selected = plc_panel_data.get("is_ups_battery_mounting_rack_selected")
            ups_battery_backup_time = plc_panel_data.get("ups_battery_backup_time")

            if ups_scope == "Client Scope" or ups_scope == "Thermax Scope":
                ups_type = "Not Applicable"
                ups_battery_type = "Not Applicable"
                is_ups_battery_mounting_rack_selected = 0
                ups_battery_backup_time = "Not Applicable"

            panel_sheet["E72"] = ups_scope
            panel_sheet["E73"] = ups_type
            panel_sheet["E74"] =  ups_battery_type
            
            if is_ups_battery_mounting_rack_selected == 1:
                is_ups_battery_mounting_rack_selected = "Applicable"
            else:
                is_ups_battery_mounting_rack_selected = "Not Applicable"
            panel_sheet["E75"] = is_ups_battery_mounting_rack_selected
            panel_sheet["E76"] = ups_battery_backup_time

            # PLC Hardware
            # panel_sheet["E78"] = plc_panel_data.get("approved_plc_hardware_make")
            panel_sheet["E78"] = plc_string
            is_bulk_power_supply_selected = plc_panel_data.get("is_bulk_power_supply_selected")
            panel_sheet["E79"] = number_To_string(is_bulk_power_supply_selected)
            panel_sheet["E80"] = plc_panel_data.get(
                "plc_cpu_or_processor_module_or_series"
            )
            panel_sheet["E81"] = plc_panel_data.get(
                "plc_communication_between_cpu_and_io_card"
            )
            third_party_protocol = plc_panel_data.get(
                "third_party_communication_protocol"
            )
            is_third_party_communication_protocol_selected = plc_panel_data.get("is_third_party_communication_protocol_selected")
            if is_third_party_communication_protocol_selected == "0":
                third_party_protocol = "Not Applicable"
                
            panel_sheet["E82"] = na_To_string(third_party_protocol)
            is_client_system_communication_selected = plc_panel_data.get("is_client_system_communication_selected")
            if is_client_system_communication_selected == "0":
                client_system_communication = "Not Applicable"
            client_system_communication = plc_panel_data.get("client_system_communication")
            panel_sheet["E83"] = client_system_communication

            # Redundancy

            is_power_supply_redundancy_selected = plc_panel_data.get("is_power_supply_redundancy_selected")
            is_io_redundancy_selected = plc_panel_data.get("is_io_redundancy_selected")
            is_cpu_and_io_card_redundancy_selected = plc_panel_data.get("is_cpu_and_io_card_redundancy_selected")
            is_cpu_and_hmi_scada_card_redundancy_selected = plc_panel_data.get("is_cpu_and_hmi_scada_card_redundancy_selected")
            is_cpu_and_third_party_services_redundancy_selected = plc_panel_data.get("is_cpu_and_third_party_services_redundancy_selected")

            is_cpu_redundancy_selected = plc_panel_data.get("is_cpu_redundancy_selected")
            if is_cpu_redundancy_selected == "0":
                cpu_redundancy = "Not Applicable"
            cpu_redundancy = plc_panel_data.get("cpu_redundancy")

            panel_sheet["E85"] = number_To_string(is_power_supply_redundancy_selected)
            panel_sheet["E86"] = number_To_string(is_io_redundancy_selected)
            panel_sheet["E87"] = number_To_string(is_cpu_and_io_card_redundancy_selected)
            panel_sheet["E88"] = number_To_string(is_cpu_and_hmi_scada_card_redundancy_selected)
            panel_sheet["E89"] = number_To_string(is_cpu_and_third_party_services_redundancy_selected)
            panel_sheet["E90"] = na_To_string(cpu_redundancy)

            # PLC Panel
            panel_sheet["E92"] = plc_panel_data.get("plc_panel_memory")
            panel_sheet["E93"] = plc_panel_data.get("panel_mounted_ac")
            panel_sheet["E94"] = plc_panel_data.get("control_voltage")
            is_plc_and_ups_marshalling_cabinet_selected = plc_panel_data.get("is_plc_and_ups_marshalling_cabinet_selected")
                

            marshalling_cabinet_for_plc_and_ups = plc_panel_data.get(
                "marshalling_cabinet_for_plc_and_ups"
            )
            if is_plc_and_ups_marshalling_cabinet_selected == "0":
                marshalling_cabinet_for_plc_and_ups = "Not Applicable"
            panel_sheet["E95"] = na_To_string(marshalling_cabinet_for_plc_and_ups)

            # Indicating Lamp, Push Button & Isolation Switch
            panel_sheet["E97"] = plc_panel_data.get("push_button_colour_acknowledge")
            panel_sheet["E98"] = plc_panel_data.get("push_button_color_reset")
            panel_sheet["E99"] = plc_panel_data.get(
                "indicating_lamp_color_for_nonups_power_supply"
            )
            panel_sheet["E100"] = plc_panel_data.get(
                "indicating_lamp_colour_for_ups_power_supply"
            )

            # DI Modules
            panel_sheet["E102"] = plc_panel_data.get("di_module_density")
            panel_sheet["E103"] = plc_panel_data.get("di_module_input_type")
            panel_sheet["E104"] = plc_panel_data.get("interrogation_voltage")
            panel_sheet["E105"] = plc_panel_data.get("di_module_scan_time")

            # DO Modules
            panel_sheet["E107"] = plc_panel_data.get("do_module_density")
            panel_sheet["E108"] = plc_panel_data.get("do_module_output_type")
            panel_sheet["E109"] = plc_panel_data.get(
                "output_contact_rating_of_interposing_relay"
            )
            panel_sheet["E110"] = plc_panel_data.get(
                "output_status_on_processor_or_module_failure"
            )
            do_module_no_of_contact = plc_panel_data.get("do_module_no_of_contact")

            is_no_of_contact_selected = plc_panel_data.get("is_no_of_contact_selected")

            if is_no_of_contact_selected == "0":
                do_module_no_of_contact = "Not Applicable"
            panel_sheet["E111"] = do_module_no_of_contact

            # AI Modules
            panel_sheet["E113"] = plc_panel_data.get("ai_module_density")
            panel_sheet["E114"] = plc_panel_data.get("ai_module_output_type")
            panel_sheet["E115"] = plc_panel_data.get("ai_module_scan_time")
            
            is_ai_module_hart_protocol_support_selected =  plc_panel_data.get("is_ai_module_hart_protocol_support_selected")
            panel_sheet["E116"] = number_To_string(is_ai_module_hart_protocol_support_selected)

            # RTD / TC Modules

            rtd_tc_module_density = plc_panel_data.get("rtd_tc_module_density")
            rtd_tc_module_input_type = plc_panel_data.get("rtd_tc_module_input_type")
            rtd_tc_module_scan_time = plc_panel_data.get("rtd_tc_module_scan_time")

            is_rtd_tc_moduule_selected_controlled = plc_panel_data.get("is_rtd_tc_moduule_selected")
            if is_rtd_tc_moduule_selected_controlled == "0":
                rtd_tc_module_density = "Not Applicable"
                rtd_tc_module_input_type = "Not Applicable"
                rtd_tc_module_scan_time = "Not Applicable"

            panel_sheet["E118"] = rtd_tc_module_density
            panel_sheet["E119"] = rtd_tc_module_input_type
            panel_sheet["E120"] = rtd_tc_module_scan_time

            is_rtd_tc_module_hart_protocol_support_selected =  plc_panel_data.get("is_rtd_tc_module_hart_protocol_support_selected")
            panel_sheet["E121"] = number_To_string(is_rtd_tc_module_hart_protocol_support_selected)

            # AO Modules
            panel_sheet["E123"] = plc_panel_data.get("ao_module_density")
            panel_sheet["E124"] = plc_panel_data.get("ao_module_output_type")
            panel_sheet["E125"] = plc_panel_data.get("ao_module_scan_time")
            
            is_ao_module_hart_protocol_support_selected = plc_panel_data.get("is_ao_module_hart_protocol_support_selected")
            panel_sheet["E126"] = number_To_string(is_ao_module_hart_protocol_support_selected)

            # PLC Spare
            plc_spare_io_count = plc_panel_data.get("plc_spare_io_count")
            is_plc_spare_io_count_selected = plc_panel_data.get("is_plc_spare_io_count_selected")

            if is_plc_spare_io_count_selected == "0":
                plc_spare_io_count = "Not Applicable"

            panel_sheet["E128"] = plc_spare_io_count
            plc_spare_memory = plc_panel_data.get("plc_spare_memory")
            is_plc_spare_memory_selected = plc_panel_data.get("is_plc_spare_memory_selected")
            if is_plc_spare_memory_selected == "0":
                plc_spare_memory = "Not Applicable"

            panel_sheet["E129"] = na_To_string(plc_spare_memory)

            # Human Interface Device
            is_no_of_hid_es_selected = plc_panel_data.get("is_no_of_hid_es_selected")
            is_no_of_hid_os_selected = plc_panel_data.get("is_no_of_hid_os_selected")
            is_no_of_hid_hmi_selected = plc_panel_data.get("is_no_of_hid_hmi_selected")
            is_hid_hmi_size_selected = plc_panel_data.get("is_hid_hmi_size_selected")

            no_of_hid_es = plc_panel_data.get("no_of_hid_es")
            no_of_hid_os = plc_panel_data.get("no_of_hid_os")
            no_of_hid_hmi = plc_panel_data.get("no_of_hid_hmi")
            hid_hmi_size = plc_panel_data.get("hid_hmi_size")

            if is_no_of_hid_es_selected == 0 :
                no_of_hid_es = "Not Applicable"

            if is_no_of_hid_os_selected == 0 :
                no_of_hid_os = "Not Applicable"

            if is_no_of_hid_hmi_selected == 0 :
                no_of_hid_hmi = "Not Applicable"

            if is_hid_hmi_size_selected == 0 :
                hid_hmi_size = "Not Applicable"
            else:
                hid_hmi_size = f"{hid_hmi_size} inch"



            panel_sheet["E131"] = no_of_hid_es
            panel_sheet["E132"] = no_of_hid_os
            panel_sheet["E133"] = no_of_hid_hmi
            panel_sheet["E134"] = hid_hmi_size

            # Software

            is_scada_development_license_selected = plc_panel_data.get("is_scada_development_license_selected")
            is_scada_runtime_license_selected = plc_panel_data.get("is_scada_runtime_license_selected")
            is_hmi_development_license_selected = plc_panel_data.get("is_hmi_development_license_selected")
            is_plc_programming_license_software_selected = plc_panel_data.get("is_plc_programming_license_software_selected")


            no_of_scada_development_license = plc_panel_data.get("no_of_scada_development_license")
            no_of_scada_runtime_license = plc_panel_data.get("no_of_scada_runtime_license")
            no_of_hmi_development_license = plc_panel_data.get("no_of_hmi_development_license")
            no_of_plc_programming_license_software = plc_panel_data.get("no_of_plc_programming_license_software")

            if is_scada_development_license_selected == 0:
                no_of_scada_development_license = "Not Applicable"

            if is_scada_runtime_license_selected == 0:
                no_of_scada_runtime_license = "Not Applicable"

            if is_hmi_development_license_selected == 0:
                no_of_hmi_development_license = "Not Applicable"

            if is_plc_programming_license_software_selected == 0:
                no_of_plc_programming_license_software = "Not Applicable"

            
            panel_sheet["E136"] = no_of_scada_development_license
            panel_sheet["E137"] = no_of_scada_runtime_license
            panel_sheet["E138"] = no_of_hmi_development_license
            panel_sheet["E139"] = no_of_plc_programming_license_software

            # Engineering / Operating SCADA Station

            system_hardware = plc_panel_data.get("system_hardware")
            commercial_grade_pc = plc_panel_data.get("commercial_grade_pc")
            monitor_size = plc_panel_data.get("monitor_size")
            windows_operating_system = plc_panel_data.get("windows_operating_system")
            is_printer_with_cable_selected = plc_panel_data.get("is_printer_with_cable_selected")
            printer_with_communication_cable = plc_panel_data.get("printer_with_communication_cable")
            no_of_printer = plc_panel_data.get("no_of_printer")
            printer_cable = plc_panel_data.get("printer_cable")
            is_furniture_for_scada_station_selected = plc_panel_data.get("is_furniture_for_scada_station_selected")
            furniture_for_scada_station = plc_panel_data.get("furniture_for_scada_station")
            hardware_between_plc_and_scada_pc = plc_panel_data.get("hardware_between_plc_and_scada_pc")

            if is_printer_with_cable_selected == 0:
                printer_with_communication_cable = "Not Applicable"
                no_of_printer = "Not Applicable"
                printer_cable = "Not Applicable"

            if is_furniture_for_scada_station_selected == 0:
                furniture_for_scada_station = "Not Applicable"


            panel_sheet["E141"] = na_To_string(system_hardware)
            panel_sheet["E142"] = na_To_string(commercial_grade_pc)
            panel_sheet["E143"] = f"{na_To_string(monitor_size)} inch"
            panel_sheet["E144"] = na_To_string(windows_operating_system)
            panel_sheet["E145"] = na_To_string(printer_with_communication_cable)
            panel_sheet["E146"] = na_To_string(no_of_printer)
            panel_sheet["E147"] = na_To_string(printer_cable)
            panel_sheet["E148"] = na_To_string(furniture_for_scada_station)
            panel_sheet["E149"] = na_To_string(hardware_between_plc_and_scada_pc)

            hardware_between_plc_and_third_party = plc_panel_data.get(
                "hardware_between_plc_and_third_party"
            )
            panel_sheet["E150"] = na_To_string(hardware_between_plc_and_third_party)
            hardware_between_plc_and_client_system = plc_panel_data.get(
                "hardware_between_plc_and_client_system"
            )
            hardware_between_plc_and_client_system = na_To_string(hardware_between_plc_and_client_system)
            panel_sheet["E151"] = na_To_string(hardware_between_plc_and_client_system)
            iiot_requirement = plc_panel_data.get("mandatory_spares")
            panel_sheet["E152"] = na_To_string(iiot_requirement)
            mandatory_spares = plc_panel_data.get("mandatory_spares")
            panel_sheet["E153"] = na_To_string(mandatory_spares)

    template_workbook.remove(mcc_sheet)
    template_workbook.remove(pcc_sheet)
    template_workbook.remove(mcc_cum_plc_sheet)

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "generated_design_basis.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
