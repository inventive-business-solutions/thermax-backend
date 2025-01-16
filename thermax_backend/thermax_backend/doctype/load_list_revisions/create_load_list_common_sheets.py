import frappe
from openpyxl import load_workbook
from thermax_backend.thermax_backend.doctype.load_list_revisions.cover_sheet import (
    create_cover_sheet,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.notes_sheet import (
    create_notes_sheet,
)


def create_load_list_common_sheets(project, revision_data, incomer_power_supply):
    """
    Generates the common sheets for the load list Excel file.
    """
    division_name = project.get("division")

    template_path = ""

    if division_name == "Heating":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "heating_load_list_template.xlsx"
        )
    elif division_name == "WWS SPG":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "spg_load_list_template.xlsx"
        )
    elif division_name == "Enviro":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "enviro_load_list_template.xlsx"
        )
    elif division_name == "WWS IPG":
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "ipg_load_list_template.xlsx"
        )
    else:
        template_path = frappe.frappe.get_app_path(
            "thermax_backend", "templates", "heating_load_list_template.xlsx"
        )

    template_workbook = load_workbook(template_path)

    cover_sheet = template_workbook["COVER"]
    notes_sheet = template_workbook["NOTES"]

    cover_sheet = create_cover_sheet(
        cover_sheet=cover_sheet,
        project_data=project,
        revision_data=revision_data,
        division_name=division_name,
    )
    notes_sheet = create_notes_sheet(
        notes_sheet=notes_sheet, incomer_power_supply=incomer_power_supply
    )

    return template_workbook
