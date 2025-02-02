def create_motor_bom_sheet(
    motor_bom_sheet,
    area_data,
):
    for data in area_data:
        standby_kw = round(float(data.get("standby_kw")), 2)
        working_kw = round(float(data.get("working_kw")), 2)
        non_zero_kw = standby_kw if standby_kw != 0 else working_kw
        data["non_zero_kw"] = non_zero_kw

    area_data.sort(key=lambda x: x["non_zero_kw"])
    count_dict = {}
    rpm_to_pole_mapping = {1000: 6, 1500: 4, 3000: 2}
    # Iterate through each motor detail in the correct structure
    for motor in area_data:
        motor_make = motor.get("make")
        non_zero_kw = motor.get("non_zero_kw")
        rpm = motor.get("rpm")
        pole = rpm_to_pole_mapping.get(rpm, 1)
        type_of_mounting = motor.get("type_of_mounting")
        efficiency = motor.get("efficiency")
        description = f"{motor_make} Make LT Motor: {non_zero_kw} kW, {pole} POLE, {type_of_mounting}, {efficiency}"
        if description in count_dict:
            count_dict[description] += 1
        else:
            count_dict[description] = 1

    index = 3

    total_hazard_data_rows = len(count_dict.keys())
    template_row_number = 3
    dynamic_start_row_number = template_row_number + 1
    template_range_start_col = 1  # Column A
    template_range_end_col = 6  # Column F (6 the column number)

    template_row_height = motor_bom_sheet.row_dimensions[template_row_number].height

    for row in range(
        dynamic_start_row_number, template_row_number + total_hazard_data_rows
    ):  # Rows 4 to (4 + total_rows)
        # Apply the row height for the current row
        motor_bom_sheet.row_dimensions[row].height = template_row_height

        # Iterate through columns A to P (1 to 16)
        for col in range(template_range_start_col, template_range_end_col + 1):
            # Get the template cell
            template_cell = motor_bom_sheet.cell(row=template_row_number, column=col)
            # Get the target cell
            target_cell = motor_bom_sheet.cell(row=row, column=col)
            # Copy the style from the template cell
            target_cell._style = template_cell._style

            # Apply column width (only once per column)
            column_letter = template_cell.column_letter
            if (
                row == dynamic_start_row_number
            ):  # Apply width only on the first iteration for each column
                template_width = motor_bom_sheet.column_dimensions[column_letter].width
                motor_bom_sheet.column_dimensions[column_letter].width = template_width

    for key, count in count_dict.items():
        motor_bom_sheet[f"A{index}"] = index - 2
        motor_bom_sheet[f"C{index}"] = key
        motor_bom_sheet[f"D{index}"] = "NOS"
        motor_bom_sheet[f"E{index}"] = count
        index += 1
    return create_motor_bom_sheet
