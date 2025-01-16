def num_to_string(value):
    if value == 1 or value == "1":
        return "Applicable"
    return "Not Applicable"


def na_to_string(value):
    if (value is None) or ("NA" in value):
        return "Not Applicable"
    return value


def handle_make_of_component(component):
    component = (
        component.replace('"', "").replace("[", "").replace("]", "").replace(",", ", ")
        if component
        else "NA"
    )
    if "NA" in component:
        return "Not Applicable"
    else:
        return component
