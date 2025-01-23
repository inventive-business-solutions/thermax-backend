import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
import io
from datetime import datetime 

@frappe.whitelist()
def get_motor_canopy_excel(): 
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    motor_canopy_revision_data = frappe.get_doc(
        "Motor Canopy Revisions",
        revision_id,
        "*"
    ).as_dict()

    project_id = motor_canopy_revision_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", 
        {
            "project_id": project_id
        },
        "*"
    ).as_dict()

    project_data = frappe.get_doc("Project", project_id, "*").as_dict()



     # Loading the workbook 
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "motor_canopy_template.xlsx"
    )
    template_workbook = load_workbook(template_path)


    project_description = design_basis_revision_data.get("description")
    project_status = design_basis_revision_data.get("status")
    owner = design_basis_revision_data.get("owner")

    division_name = project_data.get("division")
    project_name = project_data.get("project_name")
    project_oc_number = project_data.get("project_oc_number")
    approver = project_data.get("approver")
    client_name = project_data.get("client_name")
    consultant_name = project_data.get("consultant_name")
    modified = project_data.get("modified")

    cover_sheet = template_workbook["COVER"]
    combine_list_sheet = template_workbook["COMBINE LIST "]
    motor_bom_sheet = template_workbook["MOTOR BOM"]


     # # cover page sheet populating

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

    revision_date = modified.strftime("%d-%m-%Y")
    revision_data_with_pid = frappe.db.get_list("Design Basis Revision History", {"project_id": project_id}, "*")
    static_document_list_data = frappe.get_doc("Static Document List", {"project_id":project_id}, "*").as_dict()

    motor_canopy_list_and_specification = static_document_list_data.get("motor_canopy_list_and_specification")

    cover_sheet["A3"] = division_name.upper()
    # cover_sheet["D6"] = project_name.upper()
    cover_sheet["D7"] = client_name.upper()
    cover_sheet["D8"] = consultant_name.upper()
    cover_sheet["D9"] = project_name.upper()
    cover_sheet["D10"] = project_oc_number.upper()
    cover_sheet["D11"] = motor_canopy_list_and_specification

    index = 33

    for i in range(len(revision_data_with_pid) - 1, -1, -1):
        cover_sheet[f"B{index}"] = f"R{len(revision_data_with_pid) - i - 1}"
        cover_sheet[f"C{index}"] = revision_date
        if (len(revision_data_with_pid) - i - 1) == 0:
            project_description = "Issued for Approval"
        cover_sheet[f"D{index}"] = project_description
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


    motor_canopy_data_table = motor_canopy_revision_data.get("motor_canopy_data")
    index = 3
    # for i in range(len(motor_canopy_data_table)):
    for data in motor_canopy_data_table:
        combine_list_sheet[f"A{index}"] = index - 2
        combine_list_sheet[f"B{index}"] = data.get("tag_number")
        combine_list_sheet[f"C{index}"] = data.get("service_description")
        combine_list_sheet[f"D{index}"] = data.get("kw_rating")
        combine_list_sheet[f"E{index}"] = data.get("quantity")
        combine_list_sheet[f"F{index}"] = data.get("rpm")
        combine_list_sheet[f"G{index}"] = data.get("motor_mounting_type")
        combine_list_sheet[f"H{index}"] = data.get("motor_frame_size")
        combine_list_sheet[f"I{index}"] = data.get("motor_location")
        combine_list_sheet[f"J{index}"] = data.get("moc")
        combine_list_sheet[f"K{index}"] = data.get("canopy_model_number")
        combine_list_sheet[f"L{index}"] = data.get("canopy_leg_length")
        combine_list_sheet[f"M{index}"] = data.get("canopy_cut_out")
        combine_list_sheet[f"N{index}"] = data.get("part_code")
        combine_list_sheet[f"O{index}"] = data.get("motor_scope")
        combine_list_sheet[f"P{index}"] = data.get("remark")
        index = index + 1


    index = 3
    for data in motor_canopy_data_table:
        motor_bom_sheet[f"A{index}"] = index - 2
        motor_bom_sheet[f"B{index}"] = data.get("description")
        motor_bom_sheet[f"C{index}"] = data.get("part_code")
        motor_bom_sheet[f"D{index}"] = data.get("quantity")
        index = index + 1


    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "local_isolator_specification_template.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
