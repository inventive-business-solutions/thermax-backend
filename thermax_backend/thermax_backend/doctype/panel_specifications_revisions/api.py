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
    instruction_sheet = template_workbook["INSTRUCTION PAGE"]
    specification_sheet = template_workbook["SPECIFICATION"]
    safe_area_motor_list_sheet = template_workbook["SAFE AREA MOTOR LIST  "]
    safe_area_motor_bom_sheet = template_workbook[" SAFE AREA MOTOR BOM"]
    hazardous_area_motor_list_sheet = template_workbook[" HAZARDOUS AREA MOTOR LIST "]
    hazardous_area_motor_bom_sheet = template_workbook[" HAZARDOUS AREA MOTOR BOM  "]