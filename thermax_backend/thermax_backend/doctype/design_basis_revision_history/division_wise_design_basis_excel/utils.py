def num_to_string(value):
    if value == 1 or value == "1":
        return "Applicable"
    return "Not Applicable"


def na_to_string(value):
    if (value is None) or (value == "NA") or (value == "None"):
        return "Not Applicable"
    return value


def handle_make_of_component(component):
    component = (
        component.replace('"', "").replace("[", "").replace("]", "").replace(",", ", ")
        if component
        else "NA"
    )
    component = na_to_string(component)
    return component


def check_value_kW_below(value):
    value = na_to_string(value)
    if value == "All":
        return f"{value} kW"
    elif value == "Not Applicable":
        return value
    else:
        return f"{value} kW and Below"


def check_value_kW(value):
    value = na_to_string(value)
    if value == "As per OEM Standard" or value == "Not Applicable":
        return value
    elif value == "All":
        return f"{value} kW"
    else:
        return f"{value} kW and Above"
