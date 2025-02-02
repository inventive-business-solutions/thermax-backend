def create_hazardous_area_isolator_excel(
    isolator_hazard_area_sheet,
    hazard_motor_details,
    safe_isolator_data,
    hazard_isolator_data,
):
    total_rows = len(hazard_motor_details)
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 30  # Column AD (30 the column number)

    template_row_height = isolator_hazard_area_sheet.row_dimensions[
        template_row_number
    ].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        isolator_hazard_area_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = isolator_hazard_area_sheet.cell(
                row=template_row_number, column=col
            )
            # Get the target cell
            target_cell = isolator_hazard_area_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = isolator_hazard_area_sheet.column_dimensions[
                    column_letter
                ].width
                isolator_hazard_area_sheet.column_dimensions[column_letter].width = (
                    template_width
                )

    index = 3

    for index, motor in enumerate(hazard_motor_details, start=index):
        isolator_hazard_area_sheet[f"A{index}"] = index - 2
        isolator_hazard_area_sheet[f"B{index}"] = motor.get("tag_number", "")
        isolator_hazard_area_sheet[f"C{index}"] = motor.get("service_description", "")
        isolator_hazard_area_sheet[f"D{index}"] = round(
            float(motor.get("working_kw", 0)), 2
        )
        isolator_hazard_area_sheet[f"E{index}"] = ""

        motor_location = motor.get("motor_location", "")
        area = motor.get("area", "")

        isolator_hazard_area_sheet[f"K{index}"] = motor_location

        canopy_data = safe_isolator_data if area == "Safe" else hazard_isolator_data
        canopy = canopy_data.get("canopy", "")

        canopy_required = "YES"

        if canopy == "Outdoor" and motor_location == "OUTDOOR":
            canopy_required = "YES"
        else:
            canopy_required = "NO"

        isolator_hazard_area_sheet[f"F{index}"] = canopy_required

        isolator_hazard_area_sheet[f"G{index}"] = motor.get("standard", "")
        isolator_hazard_area_sheet[f"H{index}"] = motor.get("zone", "")
        isolator_hazard_area_sheet[f"I{index}"] = motor.get("gas_group", "")
        isolator_hazard_area_sheet[f"J{index}"] = motor.get("temprature_class", "")
        isolator_hazard_area_sheet[f"L{index}"] = motor.get("gland_size", "")

    isolator_hazard_area_sheet[f"C{index + 5}"] = "Total Quantity"
    isolator_hazard_area_sheet[f"D{index + 5}"] = int(len(hazard_motor_details))
    isolator_hazard_area_sheet[f"E{index + 5}"] = "Nos"

    return isolator_hazard_area_sheet
