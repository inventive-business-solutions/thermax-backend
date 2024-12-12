import frappe


def create_notes_sheet(notes_sheet, project):
    """
    Creates the notes sheet for the load list Excel file.
    """
    project_id = project.get("name")

    project_info_data = frappe.get_doc("Project Information", project_id).as_dict()

    main_supply_lv_data = project_info_data.get("main_supply_lv")
    frequency_data = project_info_data.get("frequency")
    lv_phase_data = project_info_data.get("main_supply_lv_phase")

    notes_sheet["B23"] = (
        f"Customer to provide: {main_supply_lv_data}, {frequency_data}, {lv_phase_data}"
    )
    return notes_sheet
