import re
from collections import defaultdict

import frappe
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation

from thermax_backend.thermax_backend.doctype.cable_schedule_revisions.excel_styles import (
    cell_styles,
    get_center_border_style,
    get_left_center_style,
)


def get_yes_no_dropdown():
    """
    Get excel cell dropdown values for yes/no selection.
    """
    dropdown_values = ["Yes", "No"]
    formula = f'"{",".join(dropdown_values)}"'
    dropdown = DataValidation(type="list", formula1=formula, allow_blank=True)
    dropdown.error = "Invalid value"
    dropdown.errorTitle = "Invalid Input"
    dropdown.prompt = "Please select a value from the dropdown"
    dropdown.promptTitle = "Dropdown Input"
    return dropdown


def get_size_selection_dropdown():
    """
    Get excel cell dropdown values for cable size selection.
    """
    formula = "Dropdowns!A:A"
    dropdown = DataValidation(type="list", formula1=formula, allow_blank=True)
    dropdown.error = "Invalid value"
    dropdown.errorTitle = "Invalid Input"
    dropdown.prompt = "Please select a value from the dropdown"
    dropdown.promptTitle = "Dropdown Input"
    return dropdown


def extract_cable_type(type_of_cable):
    """
    Extracts the cable type from the type_of_cable string.
    """
    if not type_of_cable:
        return ""

    try:
        # Ensure the input is a string
        type_of_cable = str(type_of_cable)

        # Split the string by hyphen and check if there are at least two parts
        parts = type_of_cable.split("-")
        if len(parts) > 1:
            return parts[1].strip()
        else:
            return ""
    except Exception as e:
        return ""


def extract_cable_name(type_of_cable):
    """
    Extracts the cable name from the type_of_cable string.
    """
    type_of_cable = str(type_of_cable)
    if not type_of_cable:
        return ""  # Return empty string if type_of_cable is None or empty

    if "Power" in type_of_cable:
        return "PC"
    elif "Control" in type_of_cable:
        return "CC"
    elif "Signal" in type_of_cable:
        return "SC"
    else:
        return "Unknown"


def extract_number(pair_core):
    """
    Extracts the numeric part from the pair_core string.
    """
    pair_core = str(pair_core)
    if not pair_core:
        return 0  # Return 0 if pair_core is None or empty

    match = re.search(r"\d+", pair_core)  # Find the first occurrence of a number
    return int(match.group()) if match else 0  # Return the number if found, otherwise 0


def distribute_data_by_panel(rearranged_data):
    """
    Distributes the rearranged data into a dictionary organized by unique panel names.
    """
    panel_data = {}

    for item in rearranged_data:
        panel_name = item.get("panel_name")  # Extract the panel name
        if not panel_name:
            continue  # Skip items without a panel name

        if panel_name not in panel_data:
            panel_data[panel_name] = []  # Initialize a list for a new panel name

        panel_data[panel_name].append(
            item
        )  # Append the current item to the appropriate panel list

    return panel_data


def rearrange_cable_schedule_data(cable_schedule_data):
    """
    Rearranges the cable schedule data to be used in the Excel sheet.
    """
    rearranged_data = []
    for cable_schedule_index in cable_schedule_data:
        cable_schedule = cable_schedule_data[cable_schedule_index]
        cables = cable_schedule_data[cable_schedule_index].get("cables", [])
        for cable in cables:
            cable["motor_name"] = cable_schedule.get("motor_name")
            rearranged_data.append(cable)

    rearranged_data = distribute_data_by_panel(rearranged_data)

    return rearranged_data


def create_other_division_excel(cable_schedule_data):
    """
    Creates an Excel sheet for the cable schedule based on the specified revision ID.
    """
    template_path = frappe.frappe.get_app_path(
        "thermax_backend", "templates", "other_division_cable_schedule_template.xlsx"
    )

    template_workbook = load_workbook(template_path)
    center_border_style = get_center_border_style()
    left_center_style = get_left_center_style()
    template_workbook.add_named_style(center_border_style)
    template_workbook.add_named_style(left_center_style)

    cable_schedule_sheet = template_workbook["MCC CABLE SCHDULE"]
    size_selection_dropdowns = get_size_selection_dropdown()
    cable_schedule_sheet.add_data_validation(size_selection_dropdowns)

    yes_no_dropdown = get_yes_no_dropdown()
    cable_schedule_sheet.add_data_validation(yes_no_dropdown)

    panel_wise_data = rearrange_cable_schedule_data(cable_schedule_data)

    template_row_number = 8
    current_row = template_row_number
    template_range_start_col = 1  # Column A
    template_range_end_col = 32  # Column AF (AF is the 32nd column)

    for panel_name, panel_cables in panel_wise_data.items():
        # Merge and style the panel name row
        cable_schedule_sheet.cell(current_row, 1).style = "left_center_style"
        cable_schedule_sheet.merge_cells(
            start_row=current_row, start_column=1, end_row=current_row, end_column=32
        )
        cable_schedule_sheet.row_dimensions[current_row].height = 40

        cable_schedule_sheet.cell(current_row, 1).value = panel_name

        current_row += 1  # Move to the next row after panel name

        motor_name_groups = defaultdict(list)

        for pc in panel_cables:
            motor_name = pc.get(
                "motor_name", "Unknown"
            )  # Default to "Unknown" if motor_name is missing
            motor_name_groups[motor_name].append(pc)

        motor_name_groups = dict(motor_name_groups)

        for motor_number_index, (motor_name, motor_cables) in enumerate(
            motor_name_groups.items()
        ):
            total_panel_cores = 0
            for cable in motor_cables:
                pair_core = cable.get("pair_core")
                number_of_cores = extract_number(pair_core)
                total_panel_cores += number_of_cores

            cable_schedule_sheet.cell(current_row, 1).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=1,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=1,
            )
            cable_schedule_sheet.cell(current_row, 1).value = motor_number_index + 1

            # 7th column G: FROM PANEL / JB NO
            cable_schedule_sheet.cell(current_row, 7).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=7,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=7,
            )
            cable_schedule_sheet.cell(current_row, 7).value = panel_name

            # 8th column H: FROM PANEL / JB DESCRIPTION
            cable_schedule_sheet.cell(current_row, 8).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=8,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=8,
            )
            cable_schedule_sheet.cell(current_row, 8).value = panel_name

            # 9th column I: SYSTEM VOLTAGE
            cable_schedule_sheet.cell(current_row, 9).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=9,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=9,
            )
            supply_voltage = (
                motor_cables[0].get("voltage") if len(motor_cables) > 1 else 0
            )
            cable_schedule_sheet.cell(current_row, 9).value = supply_voltage

            # 10th column J: KW RATING
            cable_schedule_sheet.cell(current_row, 10).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=10,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=10,
            )
            cable_schedule_sheet.cell(current_row, 10).value = (
                motor_cables[0].get("kw") if len(motor_cables) > 1 else 0
            )

            for cable_index, cable in enumerate(motor_cables, start=1):
                # Extract the number of cores from the pair_core string
                pair_core = cable.get("pair_core")
                number_of_cores = extract_number(pair_core)

                # Calculate the ending row for the merge
                end_row = current_row + max(
                    number_of_cores - 1, 0
                )  # Ensure non-negative range

                # 2nd column B: CABLE NO.
                cable_schedule_sheet.cell(current_row, 2).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=2,
                    end_row=end_row,
                    end_column=2,
                )
                cable_schedule_sheet.cell(current_row, 2).value = extract_cable_name(
                    cable.get("type_of_cable")
                )

                # 3rd column C: CABLE TAG
                cable_schedule_sheet.cell(current_row, 3).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=3,
                    end_row=end_row,
                    end_column=3,
                )
                cable_schedule_sheet.cell(current_row, 3).value = (
                    f"""=CONCATENATE(B{current_row},"/",G{current_row},"-",P{current_row})"""
                )

                # 11th column K: TYPE
                cable_schedule_sheet.cell(current_row, 11).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=11,
                    end_row=end_row,
                    end_column=11,
                )
                cable_schedule_sheet.cell(current_row, 11).value = cable.get(
                    "starter_type"
                )

                # 12th column L: CABLE SIZE
                cable_schedule_sheet.cell(current_row, 12).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=12,
                    end_row=end_row,
                    end_column=12,
                )
                cable_schedule_sheet.cell(current_row, 12).value = (
                    f"{cable.get('pair_core')} X {cable.get('sizemm2')} SQ.MM. {cable.get('cable_material')} {cable.get('type_of_insulation')} ARMOURED CABLE"
                )

                # 13th column M: CABLE TYPE
                cable_schedule_sheet.cell(current_row, 13).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=13,
                    end_row=end_row,
                    end_column=13,
                )
                cable_schedule_sheet.cell(current_row, 13).value = extract_cable_type(
                    cable.get("type_of_cable")
                )

                # 16th column P: TAG NOS.
                cable_schedule_sheet.cell(current_row, 16).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=16,
                    end_row=end_row,
                    end_column=16,
                )
                cable_schedule_sheet.cell(current_row, 16).value = cable.get(
                    "tag_number"
                )

                # 18th column R: TO EQUIPMENT DESCRIPTION
                cable_schedule_sheet.cell(current_row, 18).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=18,
                    end_row=end_row,
                    end_column=18,
                )
                cable_schedule_sheet.cell(current_row, 18).value = cable.get(
                    "service_description"
                )

                # 19th column S: CABLE LENGTH (IN MTR.)
                cable_schedule_sheet.cell(current_row, 19).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=19,
                    end_row=end_row,
                    end_column=19,
                )
                cable_schedule_sheet.cell(current_row, 19).value = cable.get(
                    "appx_length"
                )

                # 20th column T: CABLE OD
                cable_schedule_sheet.cell(current_row, 20).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=20,
                    end_row=end_row,
                    end_column=20,
                )
                cable_schedule_sheet.cell(current_row, 20).value = cable.get("cable_od")

                # 21st column U: ENTRY AVAILABLE AT PANEL / JB
                cable_schedule_sheet.cell(current_row, 21).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=21,
                    end_row=end_row,
                    end_column=21,
                )
                cable_schedule_sheet.cell(current_row, 21).value = "Plate"

                # 22nd column V: SIZE SELECTED AT PANEL / JB
                cable_schedule_sheet.cell(current_row, 22).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=22,
                    end_row=end_row,
                    end_column=22,
                )
                size_selection_dropdowns.add(cable_schedule_sheet[f"V{current_row}"])

                # 23rd column W: GLAND CAT NO AT PANEL / JB
                cable_schedule_sheet.cell(current_row, 23).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=23,
                    end_row=end_row,
                    end_column=23,
                )
                cable_schedule_sheet.cell(current_row, 23).value = (
                    f"""=IF('GLAND SELEC. INPUT & NOTES SHT'!$H$17="Ni PLATED BRASS",IF($AQ{current_row}="NARMOURED CABLE",$AX{current_row},IF($AQ{current_row}=" ARMOURED CABLE",IF($AT{current_row}="M",$AV{current_row},IF($AT{current_row}=" ",$AU{current_row},IF($AT{current_row}="N",$AW{current_row},"NA"))))),IF($AQ{current_row}="NARMOURED CABLE",$BK{current_row},IF($AQ{current_row}=" ARMOURED CABLE",IF($BG{current_row}="M",$BI{current_row},IF($BG{current_row}=" ",$BH{current_row},IF($BG{current_row}="N",$BJ{current_row},"NA"))))))"""
                )

                # 24th column X: SHROUD REQUIREMENT AT PANEL / JB
                cable_schedule_sheet.cell(current_row, 24).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=24,
                    end_row=end_row,
                    end_column=24,
                )
                yes_no_dropdown.add(cable_schedule_sheet[f"V{current_row}"])

                # 25th column Y: SHROUD CAT NO  AT PANEL / JB
                cable_schedule_sheet.cell(current_row, 25).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=25,
                    end_row=end_row,
                    end_column=25,
                )
                cable_schedule_sheet.cell(current_row, 25).value = (
                    f"""=IF($X{current_row}="YES","PVC SHROUD FOR "&$W{current_row},"NA")"""
                )

                # 26th column Z: ENTRY AVAILABLE AT EQUIPMENT
                cable_schedule_sheet.cell(current_row, 26).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=26,
                    end_row=end_row,
                    end_column=26,
                )

                # 27th column AA: SIZE SELECTED AT EQUIPMENT
                cable_schedule_sheet.cell(current_row, 27).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=27,
                    end_row=end_row,
                    end_column=27,
                )
                size_selection_dropdowns.add(cable_schedule_sheet[f"AA{current_row}"])

                # 28th column AB: GLAND CAT NO AT EQUIPMENT
                cable_schedule_sheet.cell(current_row, 28).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=28,
                    end_row=end_row,
                    end_column=28,
                )
                cable_schedule_sheet.cell(current_row, 28).value = (
                    f"""=IF('GLAND SELEC. INPUT & NOTES SHT'!$H$17="Ni PLATED BRASS",IF($AQ{current_row}="NARMOURED CABLE",$BD{current_row},IF($AQ{current_row}=" ARMOURED CABLE",IF($AZ{current_row}="M",$BB{current_row},IF($AZ{current_row}=" ",$BA{current_row},IF($AZ{current_row}="N",$BC{current_row},"NA"))))),IF($AQ{current_row}="NARMOURED CABLE",$BQ{current_row},IF($AQ{current_row}=" ARMOURED CABLE",IF($BM{current_row}="M",$BO{current_row},IF($BM{current_row}=" ",$BN{current_row},IF($BM{current_row}="N",$BP{current_row},"NA"))))))"""
                )

                # 29th column AC: SHROUD REQUIREMENT AT EQUIPMENT
                cable_schedule_sheet.cell(current_row, 29).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=29,
                    end_row=end_row,
                    end_column=29,
                )
                yes_no_dropdown.add(cable_schedule_sheet[f"AC{current_row}"])

                # 30th column AD: SHROUD CAT NO AT EQUIPMENT
                cable_schedule_sheet.cell(current_row, 30).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=30,
                    end_row=end_row,
                    end_column=30,
                )
                cable_schedule_sheet.cell(current_row, 30).value = (
                    f"""=IF($AC{current_row}="YES","PVC SHROUD FOR "&$AB{current_row},"NA")"""
                )

                # 31st column AE: CABLE TRAY ROUTING
                cable_schedule_sheet.cell(current_row, 31).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=31,
                    end_row=end_row,
                    end_column=31,
                )

                # 32nd column AF: REMARKS
                cable_schedule_sheet.cell(current_row, 32).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=32,
                    end_row=end_row,
                    end_column=32,
                )
                cable_schedule_sheet.cell(current_row, 32).value = cable.get("comment")

                for row in range(number_of_cores):
                    # 4th column D: FEEDER NO
                    cable_schedule_sheet.cell(current_row + row, 4).style = (
                        "center_border_style"
                    )
                    cable_schedule_sheet.row_dimensions[current_row + row].height = 40

                    # 5th column E: FROM TB NO
                    cable_schedule_sheet.cell(current_row + row, 5).style = (
                        "center_border_style"
                    )

                    # 6th column F: FROM FERRULE NO
                    cable_schedule_sheet.cell(current_row + row, 6).style = (
                        "center_border_style"
                    )
                    cable_schedule_sheet.cell(current_row + row, 6).value = (
                        f"""=CONCATENATE(O{current_row + row},"/",E{current_row + row})"""
                    )

                    # 14th column N: SCOPE
                    cable_schedule_sheet.cell(current_row + row, 14).style = (
                        "center_border_style"
                    )
                    cable_schedule_sheet.cell(current_row + row, 14).value = (
                        f"""=CONCATENATE(E{current_row + row},"/",O{current_row + row})"""
                    )

                    # 15th column O: TO EQUIPMENT NO
                    cable_schedule_sheet.cell(current_row + row, 15).style = (
                        "center_border_style"
                    )
                    cable_schedule_sheet.cell(current_row + row, 15).value = (
                        f"""=P{current_row + row}&" - "&Q{current_row + row}"""
                    )

                    # 17th column Q: TO TB NO.
                    cable_schedule_sheet.cell(current_row + row, 17).style = (
                        "center_border_style"
                    )

                # Update the current_row to the row after the merged cells
                current_row = end_row + 1

    return template_workbook
