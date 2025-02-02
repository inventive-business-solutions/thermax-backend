def create_safe_area_lpbs_excel(lpbs_safe_sheet, safe_motor_details, safe_lpbs_canopy):
    total_rows = len(safe_motor_details)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 30  # Column AD (30 the column number)

    template_row_height = lpbs_safe_sheet.row_dimensions[template_row_number].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        lpbs_safe_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = lpbs_safe_sheet.cell(row=template_row_number, column=col)
            # Get the target cell
            target_cell = lpbs_safe_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = lpbs_safe_sheet.column_dimensions[column_letter].width
                lpbs_safe_sheet.column_dimensions[column_letter].width = template_width

    index = 3

    for i, motor in enumerate(safe_motor_details, start=index):
        lpbs_safe_sheet[f"A{i}"] = i - index + 1  # Ensure sequential numbering
        lpbs_safe_sheet[f"B{i}"] = motor.get("tag_number")
        lpbs_safe_sheet[f"C{i}"] = motor.get("service_description")
        lpbs_safe_sheet[f"D{i}"] = round(float(motor.get("working_kw", 0)), 2)
        lpbs_safe_sheet[f"E{i}"] = motor.get("lpbs_type")
        lpbs_safe_motor_location = motor.get("motor_location")
        lpbs_safe_sheet[f"G{i}"] = lpbs_safe_motor_location

        canopy_required = "YES"

        if safe_lpbs_canopy == "Outdoor" and lpbs_safe_motor_location == "OUTDOOR":
            canopy_required = "YES"
        else:
            canopy_required = "NO"

        lpbs_safe_sheet[f"F{i}"] = canopy_required
        lpbs_safe_sheet[f"H{i}"] = motor.get("gland_size")

    type_count = {
        motor_type["lpbs_type"]: sum(
            1 for t in safe_motor_details if t["lpbs_type"] == motor_type["lpbs_type"]
        )
        for motor_type in safe_motor_details
    }
    a = index + 1 + total_rows
    val = 1
    for key, value in type_count.items():
        lpbs_safe_sheet[f"B{a}"] = val
        lpbs_safe_sheet[f"C{a}"] = "LPBS Type"
        lpbs_safe_sheet[f"D{a}"] = key
        lpbs_safe_sheet[f"E{a}"] = value
        # lpbs_safe_sheet[f"O{a}"] = "No.s"
        a = a + 1
        val = val + 1

    lpbs_safe_sheet[f"C{a}"] = "Total"
    lpbs_safe_sheet[f"E{a}"] = total_rows

    return lpbs_safe_sheet
