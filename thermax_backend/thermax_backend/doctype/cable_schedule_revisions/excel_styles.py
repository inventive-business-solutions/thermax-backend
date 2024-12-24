from openpyxl.styles import Font, Border, Side, Alignment, NamedStyle

cell_styles = {
    "all_b_thin": Border(
        top=Side(style="thin"),
        bottom=Side(style="thin"),
        left=Side(style="thin"),
        right=Side(style="thin"),
    ),
    "left_center": Alignment(horizontal="left", vertical="center"),
    "center": Alignment(horizontal="center", vertical="center"),
    "bold": Font(bold=True),
}


def get_center_border_style():
    return NamedStyle(
        name="center_border_style",
        alignment=cell_styles.get("center"),
        border=cell_styles.get("all_b_thin"),
    )


def get_center_border_bold_style():
    return NamedStyle(
        name="center_border_bold_style",
        alignment=cell_styles.get("center"),
        border=cell_styles.get("all_b_thin"),
        font=cell_styles.get("bold"),
    )


def get_left_center_style():
    return NamedStyle(
        name="left_center_style",
        font=cell_styles.get("bold"),
        alignment=cell_styles.get("left_center"),
        border=cell_styles.get("all_b_thin"),
    )
