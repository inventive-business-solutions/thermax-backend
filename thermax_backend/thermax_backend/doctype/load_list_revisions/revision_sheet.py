import frappe
from datetime import datetime


def create_revision_sheet(revision_sheet, project):
    """
    Creates and populates the revision sheet for the load list Excel file.

    Args:
        revision_sheet (object): The Excel worksheet object for the revision sheet.
        project (dict): Dictionary containing project details, including the project ID.

    Returns:
        object: The updated revision sheet with revision details.
    """
    project_id = project.get("name")
    revision_lists = frappe.db.get_list(
        "Load List Revisions", {"project_id": project_id}, "*"
    )
    start_row = 6

    if revision_lists:
        for index, revision in enumerate(revision_lists):
            revision_date = revision.get("modified")
            formatted_date = revision_date.strftime("%d-%m-%Y")

            # Populate the revision sheet
            row = start_row + index
            revision_sheet[f"B{row}"] = f"R{index}"
            revision_sheet[f"D{row}"] = formatted_date
            revision_sheet[f"E{row}"] = "ISSUED FOR APPROVAL"
    else:
        # Handle the case where no revisions exist
        revision_sheet["B6"] = "R0"
        revision_sheet["D6"] = ""
        revision_sheet["E6"] = "ISSUED FOR APPROVAL"

    return revision_sheet