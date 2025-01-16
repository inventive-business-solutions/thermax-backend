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

    # template_workbook.save("electrical_load_list.xlsx")

    output = io.BytesIO()
    template_workbook.save(output)
    output.seek(0)

    frappe.local.response.filename = "generated_design_basis.xlsx"
    frappe.local.response.filecontent = output.getvalue()
    frappe.local.response.type = "binary"

    return _("File generated successfully.")
