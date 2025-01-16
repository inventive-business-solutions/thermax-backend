import frappe


def create_cover_sheet(cover_sheet, project_data, revision_data, division_name):
    project_id = revision_data.get("project_id")

    revision_data_with_pid = frappe.db.get_list(
        "Design Basis Revision History", {"project_id": project_id}, "*"
    )

    static_documents = frappe.get_doc("Static Document List", project_id).as_dict()

    project_name = project_data.get("project_name")
    project_oc_number = project_data.get("project_oc_number")
    approver = project_data.get("approver")
    consultant_name = project_data.get("consultant_name")
    owner = revision_data.get("owner")

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

    cover_sheet["A3"] = division_name.upper()
    cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = project_data.get("client_name").upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = static_documents.get("electrical_load_list", "TBD").upper()

    index = 33

    for i in range(len(revision_data_with_pid) - 1, -1, -1):
        current_revision = revision_data_with_pid[i]
        revision_date = current_revision.get("modified")
        project_description = current_revision.get("description")

        cover_sheet[f"B{index}"] = f"R{len(revision_data_with_pid) - i - 1}"
        cover_sheet[f"C{index}"] = revision_date
        cover_sheet[f"D{index}"] = "Issued for Approval"
        cover_sheet[f"E{index}"] = prepped_by_initial
        cover_sheet[f"F{index}"] = checked_by_initial
        cover_sheet[f"G{index}"] = super_user_initial
        index = index - 1

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
    return cover_sheet
