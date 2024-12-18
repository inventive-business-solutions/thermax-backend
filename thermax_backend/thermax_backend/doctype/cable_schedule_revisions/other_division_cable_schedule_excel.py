from collections import defaultdict
import frappe
import re
from openpyxl import load_workbook
from thermax_backend.thermax_backend.doctype.cable_schedule_revisions.excel_styles import (
    cell_styles,
    get_center_border_style,
    get_left_center_style,
)


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
            cable_schedule_sheet.cell(current_row, 7).value = (
                f"FROM PANEL / JB NO {current_row}"
            )

            # 8th column H: FROM PANEL / JB DESCRIPTION
            cable_schedule_sheet.cell(current_row, 8).style = "center_border_style"
            cable_schedule_sheet.merge_cells(
                start_row=current_row,
                start_column=8,
                end_row=current_row + max(total_panel_cores - 1, 0),
                end_column=8,
            )
            cable_schedule_sheet.cell(current_row, 8).value = (
                f"FROM PANEL / JB DESCRIPTION {current_row}"
            )

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
                cable_schedule_sheet.cell(current_row, 2).value = cable.get("name")

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
                cable_schedule_sheet.cell(current_row, 12).value = cable.get("sizemm2")

                # 13th column M: CABLE TYPE
                cable_schedule_sheet.cell(current_row, 13).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=13,
                    end_row=end_row,
                    end_column=13,
                )
                cable_schedule_sheet.cell(current_row, 13).value = cable.get(
                    "type_of_cable"
                )

                # 16th column P: TAG NOS.
                cable_schedule_sheet.cell(current_row, 16).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=16,
                    end_row=end_row,
                    end_column=16,
                )
                cable_schedule_sheet.cell(current_row, 16).value = cable.get("tag_no")

                # 18th column R: TO EQUIPMENT DESCRIPTION
                cable_schedule_sheet.cell(current_row, 18).style = "center_border_style"
                cable_schedule_sheet.merge_cells(
                    start_row=current_row,
                    start_column=18,
                    end_row=end_row,
                    end_column=18,
                )
                cable_schedule_sheet.cell(current_row, 18).value = (
                    f"TO PANEL / JB NO {current_row}"
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
