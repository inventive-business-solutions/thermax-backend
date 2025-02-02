def create_hazardous_area_motor_list_sheet(
    hazardous_area_motor_list_sheet, hazard_data
):
    total_rows = len(hazard_data)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 30  # Column AD (30 the column number)

    template_row_height = hazardous_area_motor_list_sheet.row_dimensions[
        template_row_number
    ].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        hazardous_area_motor_list_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = hazardous_area_motor_list_sheet.cell(
                row=template_row_number, column=col
            )
            # Get the target cell
            target_cell = hazardous_area_motor_list_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = hazardous_area_motor_list_sheet.column_dimensions[
                    column_letter
                ].width
                hazardous_area_motor_list_sheet.column_dimensions[
                    column_letter
                ].width = template_width

    index = 3
    for data in hazard_data:
        motor_rating_data = data.get("working_kw")
        kw_data = "W"

        if float(motor_rating_data) == 0:
            kw_data = "S"
            motor_rating_data = data.get("standby_kw")

        hazardous_area_motor_list_sheet[f"A{index}"] = index - 2
        hazardous_area_motor_list_sheet[f"B{index}"] = data.get("tag_number")
        hazardous_area_motor_list_sheet[f"C{index}"] = data.get("service_description")
        hazardous_area_motor_list_sheet[f"D{index}"] = kw_data
        hazardous_area_motor_list_sheet[f"E{index}"] = motor_rating_data
        hazardous_area_motor_list_sheet[f"F{index}"] = data.get("rpm")
        hazardous_area_motor_list_sheet[f"G{index}"] = data.get("type_of_mounting")
        hazardous_area_motor_list_sheet[f"H{index}"] = data.get("motor_frame_size")
        hazardous_area_motor_list_sheet[f"I{index}"] = data.get("motor_gd2")
        hazardous_area_motor_list_sheet[f"J{index}"] = data.get(
            "gd2_of_driven_equipment"
        )
        hazardous_area_motor_list_sheet[f"K{index}"] = data.get("bkw")
        hazardous_area_motor_list_sheet[f"L{index}"] = data.get("type_of_couplings")
        hazardous_area_motor_list_sheet[f"M{index}"] = data.get("motor_location")
        hazardous_area_motor_list_sheet[f"N{index}"] = data.get("supply_voltage")
        hazardous_area_motor_list_sheet[f"O{index}"] = 50
        hazardous_area_motor_list_sheet[f"P{index}"] = data.get("starter_type")
        hazardous_area_motor_list_sheet[f"Q{index}"] = data.get("cable_size")
        hazardous_area_motor_list_sheet[f"R{index}"] = data.get("space_heater")
        roller_bearing = "No"
        if data.get("type_of_bearing") == "Roller":
            roller_bearing = "Yes"

        hazardous_area_motor_list_sheet[f"S{index}"] = roller_bearing

        insulated_bearing = "No"
        if "nsulat" in data.get("type_of_bearing"):
            insulated_bearing = "Yes"

        hazardous_area_motor_list_sheet[f"T{index}"] = insulated_bearing
        hazardous_area_motor_list_sheet[f"U{index}"] = data.get("thermistor")
        hazardous_area_motor_list_sheet[f"V{index}"] = data.get("bearing_rtd")
        hazardous_area_motor_list_sheet[f"W{index}"] = data.get("winding_rtd")
        hazardous_area_motor_list_sheet[f"X{index}"] = data.get("efficiency")
        hazardous_area_motor_list_sheet[f"Y{index}"] = data.get("motor_rated_current")
        hazardous_area_motor_list_sheet[f"Z{index}"] = data.get("power_factor")
        # safe_area_motor_list_sheet[f"AA{index}"] =
        hazardous_area_motor_list_sheet[f"AB{index}"] = data.get("make")
        hazardous_area_motor_list_sheet[f"AC{index}"] = data.get("part_code")
        hazardous_area_motor_list_sheet[f"AD{index}"] = data.get("remark")
        index = index + 1

    return hazardous_area_motor_list_sheet
