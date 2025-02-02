def create_hazardous_area_lpbs_excel(
    lpbs_hazard_sheet, hazard_motor_details, hazardous_lpbs_canopy
):
    total_rows = len(hazard_motor_details)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 30  # Column AD (30 the column number)

    template_row_height = lpbs_hazard_sheet.row_dimensions[template_row_number].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        lpbs_hazard_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = lpbs_hazard_sheet.cell(row=template_row_number, column=col)
            # Get the target cell
            target_cell = lpbs_hazard_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = lpbs_hazard_sheet.column_dimensions[
                    column_letter
                ].width
                lpbs_hazard_sheet.column_dimensions[column_letter].width = (
                    template_width
                )

    index = 3  # Starting index
    for i, motor_detail in enumerate(hazard_motor_details, start=index):
        lpbs_hazard_sheet[f"A{i}"] = i - index + 1  # Sequence number
        lpbs_hazard_sheet[f"B{i}"] = motor_detail.get("tag_number")
        lpbs_hazard_sheet[f"C{i}"] = motor_detail.get("service_description")
        lpbs_hazard_sheet[f"D{i}"] = round(float(motor_detail.get("working_kw", 0)), 2)
        lpbs_hazard_sheet[f"E{i}"] = motor_detail.get("lpbs_type")

        hazard_lpbs_motor_location = motor_detail.get("motor_location")
        lpbs_hazard_sheet[f"K{i}"] = hazard_lpbs_motor_location

        canopy_required = "YES"

        if (
            hazardous_lpbs_canopy == "Outdoor"
            and hazard_lpbs_motor_location == "OUTDOOR"
        ):
            canopy_required = "YES"
        else:
            canopy_required = "NO"

        lpbs_hazard_sheet[f"F{i}"] = canopy_required
        lpbs_hazard_sheet[f"G{i}"] = motor_detail.get("standard")
        lpbs_hazard_sheet[f"H{i}"] = motor_detail.get("zone")
        lpbs_hazard_sheet[f"I{i}"] = motor_detail.get("gas_group")
        lpbs_hazard_sheet[f"J{i}"] = motor_detail.get("temperature_class")
        lpbs_hazard_sheet[f"L{i}"] = motor_detail.get("gland_size")

    type_count = {
        motor_type["lpbs_type"]: sum(
            1 for t in hazard_motor_details if t["lpbs_type"] == motor_type["lpbs_type"]
        )
        for motor_type in hazard_motor_details
    }
    a = index + 1 + total_rows
    val = 1
    for key, value in type_count.items():
        lpbs_hazard_sheet[f"B{a}"] = val
        lpbs_hazard_sheet[f"C{a}"] = "LPBS Type"
        lpbs_hazard_sheet[f"D{a}"] = key
        lpbs_hazard_sheet[f"E{a}"] = value
        # lpbs_safe_sheet[f"O{a}"] = "No.s"
        a = a + 1
        val = val + 1

    lpbs_hazard_sheet[f"C{a}"] = "Total"
    lpbs_hazard_sheet[f"E{a}"] = int(len(hazard_motor_details))

    return lpbs_hazard_sheet
