import io
import frappe
from frappe import _
from thermax_backend.thermax_backend.doctype.load_list_revisions.create_load_list_common_sheets import (
    create_load_list_common_sheets,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.create_load_list_sheet import (
    create_load_list_excel,
)


@frappe.whitelist()
def get_load_list_excel():
    """
    Generates an Excel sheet for the electrical load list based on the specified division.
    """
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")
    revision_data = frappe.get_doc("Load List Revisions", revision_id).as_dict()
    project_id = revision_data.get("project_id")

    project = frappe.get_doc("Project", project_id).as_dict()

    template_workbook = create_load_list_common_sheets(project, revision_data)
    template_workbook = create_load_list_excel(
        revision_data=revision_data,
        project=project,
        template_workbook=template_workbook,
    )
<<<<<<< HEAD
    checked_by_initial = frappe.db.get_value(
        "Thermax Extended User", project_approver, "name_initial"
    )
    super_user_initial = frappe.db.get_value(
        "Thermax Extended User",
        {"is_superuser": 1, "division": division_name},
        "name_initial",
    )

    electrical_load_list_data = revision_data.get("electrical_load_list_data")
    unique_panels = {item["panel"] for item in electrical_load_list_data}
    panels_data = {panel: [] for panel in unique_panels}

    for item in electrical_load_list_data:
        panel_name = item["panel"]
        panels_data[panel_name].append(item)

    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "heating_load_list_template.xlsx"
    )

    template_workbook = load_workbook(template_path)

    cover_sheet = template_workbook["COVER"]
    revision_sheet = template_workbook["REVISION"]
    notes_sheet = template_workbook["NOTES"]

    # Cover Sheet

    cover_sheet["A3"] = division_name.upper()
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

    revision_date = revision_data.get("modified")

    cover_sheet["C36"] = revision_date.strftime("%d-%m-%Y")
    cover_sheet["D7"] = project.get("client_name").upper()
    cover_sheet["D8"] = project.get("consultant_name").upper()
    cover_sheet["D9"] = project.get("project_name").upper()
    cover_sheet["D10"] = project.get("project_oc_number").upper()

    electrical_load_list_name = frappe.db.get_value("Static Document List", {"project_id": project_id}, "electrical_load_list")
    cover_sheet["D11"] = electrical_load_list_name
    
    # cover_sheet["D36"] = revision_data.get("status")  # from payload
    cover_sheet["D36"] = "ISSUED FOR APPROVAL" # from payload

    cover_sheet["E36"] = prepped_by_initial
    cover_sheet["F36"] = checked_by_initial
    cover_sheet["G36"] = super_user_initial


    # REVISION SHEET

    start_row = 6

    if revision_lists_len > 1 :
        for index, revision in enumerate(revision_lists) :
            modified_revision_date = revision.get("modified")

            if modified_revision_date:
                modified_revision_date = datetime.strptime(
                    modified_revision_date, "%Y-%m-%d %H:%M:%S.%f"
                ).strftime("%d-%m-%Y")
            else:
                modified_revision_date = (
                    ""  # Handle cases where 'modified' might be None
                )

                # Update the revision_sheet with the current revision data
                revision_sheet[f"B{start_row + index}"] = f"R{index}"
                revision_sheet[f"D{start_row + index}"] = modified_revision_date
                revision_sheet[f"E{start_row + index}"] = "ISSUED FOR APPROVAL"
    else:
        revision_sheet[f"B6"] = f"R{revision_lists_len-1}"
        revision_sheet[f"D6"] = "modified_revision_date"
        revision_sheet[f"E6"] = "ISSUED FOR APPROVAL"


    # NOTES PAGE 

    project_info_data = frappe.db.get_list("Project Information", {"project_id": project_id}, ["main_supply_lv", "frequency", "main_supply_lv_phase"])
    main_supply_lv_data  = project_info_data[0].main_supply_lv
    frequency_data  = project_info_data[0].frequency
    lv_phase_data  = project_info_data[0].main_supply_lv_phase
    
    notes_sheet["B23"] = f"Customer to provide: {main_supply_lv_data}, {frequency_data}, {lv_phase_data}"

    # LOAD LIST OUTPUT

    load_list_output_sheet = template_workbook["LOAD LIST OUTPUT"]
    all_panels_sheet = template_workbook.copy_worksheet(load_list_output_sheet)

    all_panels_sheet = create_load_list_excel(
        electrical_load_list_data=electrical_load_list_data,
        load_list_output_sheet=all_panels_sheet,
        division_name=division_name,
    )

    for panel_name, panel_data in panels_data.items():
        panel_sheet = template_workbook.copy_worksheet(load_list_output_sheet)
        panel_sheet.title = panel_name

        panel_sheet = create_load_list_excel(
            electrical_load_list_data=panel_data,
            load_list_output_sheet=panel_sheet,
            division_name=division_name,
        )

    template_workbook.remove(load_list_output_sheet)
    all_panels_sheet.title = "LOAD LIST OUTPUT"
=======
>>>>>>> 11a05c2 (division wise load list templates added)

    # template_workbook.save("electrical_load_list.xlsx")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "generated_design_basis.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
