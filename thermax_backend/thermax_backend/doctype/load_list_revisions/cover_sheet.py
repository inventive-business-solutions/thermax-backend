import frappe


def create_cover_sheet(cover_sheet, division_name, project, revision_data):
    """
    Creates the cover sheet for the load list Excel file.
    """
    project_id = project.get("name")
    project_owner = project.get("owner")
    project_approver = project.get("approver")
    prepped_by_initial = frappe.db.get_value(
        "Thermax Extended User", project_owner, "name_initial"
    )
    checked_by_initial = frappe.db.get_value(
        "Thermax Extended User", project_approver, "name_initial"
    )
    super_user_initial = frappe.db.get_value(
        "Thermax Extended User",
        {"is_superuser": 1, "division": division_name},
        "name_initial",
    )

    electrical_load_list_name = frappe.db.get_value(
        "Static Document List",
        {"project_id": project_id},
        "electrical_load_list",
    )

    revision_date = revision_data.get("modified")
    latest_revision_data = revision_date.strftime("%d-%m-%Y")
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

    cover_sheet["C36"] = latest_revision_data
    cover_sheet["D7"] = project.get("client_name").upper()
    cover_sheet["D8"] = project.get("consultant_name").upper()
    cover_sheet["D9"] = project.get("project_name").upper()
    cover_sheet["D10"] = project.get("project_oc_number").upper()
    cover_sheet["D11"] = electrical_load_list_name
    # cover_sheet["D36"] = revision_data.get("status")  # from payload
    cover_sheet["D36"] = "ISSUED FOR APPROVAL"  # from payload

    cover_sheet["E36"] = prepped_by_initial
    cover_sheet["F36"] = checked_by_initial
    cover_sheet["G36"] = super_user_initial

    return cover_sheet
