import frappe
from openpyxl import load_workbook
from thermax_backend.thermax_backend.doctype.load_list_revisions.cover_sheet import (
    create_cover_sheet,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.notes_sheet import (
    create_notes_sheet,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.revision_sheet import (
    create_revision_sheet,
)


def create_load_list_common_sheets(project, revision_data):
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
    revision_sheet = template_workbook["REVISION"]
    notes_sheet = template_workbook["NOTES"]

    cover_sheet = create_cover_sheet(cover_sheet, division_name, project, revision_data)
    revision_sheet = create_revision_sheet(revision_sheet, project)
    notes_sheet = create_notes_sheet(notes_sheet, project)

    return template_workbook
