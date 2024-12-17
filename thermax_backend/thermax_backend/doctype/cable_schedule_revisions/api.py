import io
import frappe
from frappe import _
from thermax_backend.thermax_backend.doctype.cable_schedule_revisions.cable_schedule_excel import (
    create_cable_schedule_excel,
)
from thermax_backend.thermax_backend.doctype.cable_schedule_revisions.voltage_drop_excel import (
    create_voltage_drop_excel,
)


@frappe.whitelist()
def get_voltage_drop_excel():
    """
    POST request to generate an Excel sheet for the voltage drop calculation based on the specified revision ID.
    """
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")
    template_workbook = create_voltage_drop_excel(revision_id)

    # template_workbook.save("voltage_dropdown_calculation.xlsx")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "voltage_drop_calculation.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("Voltage Drop Excel Created")


@frappe.whitelist()
def get_cable_schedule_excel():
    """
    POST request to generate an Excel sheet for the cable schedule based on the specified revision ID.
    """
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")
    template_workbook = create_cable_schedule_excel(revision_id)

    # template_workbook.save("cable_schedule.xlsx")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "cable_schedule.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("Cable Schedule Excel Created")
