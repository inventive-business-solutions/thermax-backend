import io
import frappe
from frappe import _
from openpyxl import load_workbook


@frappe.whitelist()
def get_voltage_drop_excel():
    payload = frappe.local.form_dict
    cable_schedule_revision_id = payload.get("cable_schedule_revision_id")
    design_basis_revision_id = payload.get("design_basis_revision_id")

    cable_schedule_revision = frappe.get_doc(
        "Cable Schedule Revisions", cable_schedule_revision_id
    ).as_dict()
    cable_schedule_data = cable_schedule_revision.get("cable_schedule_data")

    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "voltage_drop_calculation_template.xlsx"
    )

    template_workbook = load_workbook(template_path)
    voltage_drop_calculation_sheet = template_workbook["VOLTAGE DROP CALCULATION"]

    total_rows = len(cable_schedule_data)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 28  # Column AB (AB is the 28th column)

    for row in range(dynamic_start_row_number, template_row_number + total_rows):
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = voltage_drop_calculation_sheet.cell(
                row=template_row_number, column=col
            )
            # Get the target cell
            target_cell = voltage_drop_calculation_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = voltage_drop_calculation_sheet.column_dimensions[
                    column_letter
                ].width
                voltage_drop_calculation_sheet.column_dimensions[
                    column_letter
                ].width = template_width

    for index, data in enumerate(cable_schedule_data):
        row = template_row_number + index
        voltage_drop_calculation_sheet.cell(row=row, column=1, value=data.get("idx"))

    template_workbook.save("voltage_dropdown_calculation.xlsx")

    return _("Voltage Drop Excel Created")
