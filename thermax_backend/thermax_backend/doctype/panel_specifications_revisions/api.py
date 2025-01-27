import frappe
from frappe import _
from openpyxl import Workbook, load_workbook
from copy import copy
from collections import defaultdict
import io
from datetime import datetime

# revision_id = "st486uu99i"


@frappe.whitelist()
def get_panel_specification_excel():
    payload = frappe.local.form_dict
    revision_id = payload.get("revision_id")

    panel_spec_revision_data = frappe.get_doc(
        "Panel Specifications Revisions", revision_id, "*"
    ).as_dict()

    project_id = panel_spec_revision_data.get("project_id")

    design_basis_revision_data = frappe.get_doc(
        "Design Basis Revision History", {"project_id": project_id}
    ).as_dict()

    project_revision_id = design_basis_revision_data.get("name")

    # Loading the workbook
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "power_cum_plc_panel_specification_template.xlsx"
    )
    template_workbook = load_workbook(template_path)


    cover_sheet = template_workbook["COVER"]
    mcc_cum_plc_panel_sheet = template_workbook["MCC CUM PLC PANEL"]
    plc_specification_sheet = template_workbook["PLC SPECIFICATION"]

    dynamic_document_list_data = frappe.get_doc(
        "Dynamic Document List",
        project_id,
        "*"
    ).as_dict()

    static_document_list_data = frappe.get_doc(
        "Static Document List",
        project_id,
        "*"
    ).as_dict()

    project_panel_data = frappe.db.get_list(
        "Project Panel Data", {"revision_id": project_revision_id}, "*", order_by="creation asc"
    )

    project_info_data = frappe.get_doc(
        "Project Information",
        project_id,
        "*"
    ).as_dict()

    plc_data = frappe.db.get_list(
        "Panel PLC 1 - 3",
        {"revision_id": project_revision_id}
    )

    



    # PLC SPECIFICATION SHEET 

    

    plc_specification_sheet["C9"] = "TBD"
    plc_specification_sheet["C10"] = "TBD"
    plc_specification_sheet["C11"] = "TBD"




    