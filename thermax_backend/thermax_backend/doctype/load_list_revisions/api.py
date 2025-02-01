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
    incomer_db_data = frappe.get_all(
        "Incomer Details", fields=["*"], order_by="incomer_rating ASC"
    )

    project_id = project.get("name")

    project_info_data = frappe.get_doc("Project Information", project_id).as_dict()

    main_supply_lv_data = project_info_data.get("main_supply_lv")
    frequency_data = project_info_data.get("frequency")
    lv_phase_data = project_info_data.get("main_supply_lv_phase")

    incomer_power_supply = (
        f"{main_supply_lv_data}, {frequency_data} Hz, {lv_phase_data}"
    )

    template_workbook = create_load_list_common_sheets(
        project=project,
        revision_data=revision_data,
        incomer_power_supply=incomer_power_supply,
    )
    template_workbook = create_load_list_excel(
        template_workbook=template_workbook,
        revision_data=revision_data,
        project=project,
        incomer_power_supply=incomer_power_supply,
        incomer_db_data=incomer_db_data,
    )

    template_workbook.save("electrical_load_list.xlsx")

    # output = io.BytesIO()
    # template_workbook.save(output)
    # output.seek(0)

    # frappe.local.response.filename = "electrical_load_list.xlsx"
    # frappe.local.response.filecontent = output.getvalue()
    # frappe.local.response.type = "binary"

    return _("File generated successfully.")
