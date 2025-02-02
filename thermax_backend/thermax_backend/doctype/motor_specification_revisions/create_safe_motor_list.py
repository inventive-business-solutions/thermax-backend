def create_safe_area_motor_list_sheet(safe_area_motor_list_sheet, safe_data):
    total_rows = len(safe_data)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 30  # Column AD (30 the column number)

    template_row_height = safe_area_motor_list_sheet.row_dimensions[
        template_row_number
    ].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        safe_area_motor_list_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = safe_area_motor_list_sheet.cell(
                row=template_row_number, column=col
            )
            # Get the target cell
            target_cell = safe_area_motor_list_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = safe_area_motor_list_sheet.column_dimensions[
                    column_letter
                ].width
                safe_area_motor_list_sheet.column_dimensions[column_letter].width = (
                    template_width
                )

    index = 3

    for data in safe_data:

        motor_rating_data = data.get("working_kw")
        kw_data = "W"

        if float(motor_rating_data) == 0:
            kw_data = "S"
            motor_rating_data = data.get("standby_kw")

        safe_area_motor_list_sheet[f"A{index}"] = index - 2
        safe_area_motor_list_sheet[f"B{index}"] = data.get("tag_number")
        safe_area_motor_list_sheet[f"C{index}"] = data.get("service_description")
        safe_area_motor_list_sheet[f"D{index}"] = kw_data
        safe_area_motor_list_sheet[f"E{index}"] = motor_rating_data
        safe_area_motor_list_sheet[f"F{index}"] = data.get("rpm")
        safe_area_motor_list_sheet[f"G{index}"] = data.get("type_of_mounting")
        safe_area_motor_list_sheet[f"H{index}"] = data.get("motor_frame_size")
        safe_area_motor_list_sheet[f"I{index}"] = data.get("motor_gd2")
        safe_area_motor_list_sheet[f"J{index}"] = data.get("gd2_of_driven_equipment")
        safe_area_motor_list_sheet[f"K{index}"] = data.get("bkw")
        safe_area_motor_list_sheet[f"L{index}"] = data.get("type_of_couplings")
        safe_area_motor_list_sheet[f"M{index}"] = data.get("motor_location")
        safe_area_motor_list_sheet[f"N{index}"] = data.get("supply_voltage")
        safe_area_motor_list_sheet[f"O{index}"] = 50
        safe_area_motor_list_sheet[f"P{index}"] = data.get("starter_type")
        safe_area_motor_list_sheet[f"Q{index}"] = data.get("cable_size")
        safe_area_motor_list_sheet[f"R{index}"] = data.get("space_heater")
        roller_bearing = "No"
        if data.get("type_of_bearing") == "Roller":
            roller_bearing = "Yes"

        safe_area_motor_list_sheet[f"S{index}"] = roller_bearing

        insulated_bearing = "No"
        if "nsulat" in data.get("type_of_bearing"):
            insulated_bearing = "Yes"

        safe_area_motor_list_sheet[f"T{index}"] = insulated_bearing
        safe_area_motor_list_sheet[f"U{index}"] = data.get("thermistor")
        safe_area_motor_list_sheet[f"V{index}"] = data.get("bearing_rtd")
        safe_area_motor_list_sheet[f"W{index}"] = data.get("winding_rtd")
        safe_area_motor_list_sheet[f"X{index}"] = data.get("efficiency")
        safe_area_motor_list_sheet[f"Y{index}"] = data.get("motor_rated_current")
        safe_area_motor_list_sheet[f"Z{index}"] = data.get("power_factor")
        # safe_area_motor_list_sheet[f"AA{index}"] =
        safe_area_motor_list_sheet[f"AB{index}"] = data.get("make")
        safe_area_motor_list_sheet[f"AC{index}"] = data.get("part_code")
        safe_area_motor_list_sheet[f"AD{index}"] = data.get("remark")
        index = index + 1

    return safe_area_motor_list_sheet
