import frappe


def create_notes_sheet(notes_sheet, incomer_power_supply):
    """
    Creates the notes sheet for the load list Excel file.
    """

    notes_sheet["B23"] = f"Customer to provide: {incomer_power_supply}"
    return notes_sheet
