from thermax_backend.thermax_backend.doctype.load_list_revisions.division_wise_load_list_excel.enviro_load_list_sheet import (
    get_enviro_load_list_excel,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.division_wise_load_list_excel.heating_load_list_sheet import (
    get_heating_load_list_excel,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.division_wise_load_list_excel.ipg_load_list_sheet import (
    get_ipg_load_list_excel,
)
from thermax_backend.thermax_backend.doctype.load_list_revisions.division_wise_load_list_excel.spg_load_list_sheet import (
    get_spg_load_list_excel,
)


def create_load_list_excel(
    template_workbook, revision_data, project, incomer_power_supply
):
    """
    Generates an Excel sheet for the electrical load list based on the specified division.

    Depending on the division name, the function delegates the creation of the load list
    to a specialized function (`create_heating_spg_load_list_excel`) for divisions like
    "Heating" or "WWS SPG". This ensures that division-specific logic is applied.

    Args:
        electrical_load_list_data (dict or list): The data representing the electrical load list.
        load_list_output_sheet (object): The Excel worksheet object where the load list will be written.
        division_name (str): The name of the division for which the load list is being created.
                             Must be "Heating", "WWS SPG", or other valid divisions.

    Returns:
        object: The updated Excel worksheet object containing the load list.
    """
    division_name = project.get("division")
    electrical_load_list_data = revision_data.get("electrical_load_list_data")
    unique_panels = {item["panel"] for item in electrical_load_list_data}
    panels_data = {panel: [] for panel in unique_panels}

    for item in electrical_load_list_data:
        panel_name = item["panel"]
        panels_data[panel_name].append(item)

    if division_name == "Heating":
        template_workbook = get_heating_load_list_excel(
            template_workbook=template_workbook,
            electrical_load_list_data=electrical_load_list_data,
            panels_data=panels_data,
            incomer_power_supply=incomer_power_supply,
        )
    elif division_name == "WWS SPG":
        template_workbook = get_spg_load_list_excel(
            electrical_load_list_data=electrical_load_list_data,
            panels_data=panels_data,
            template_workbook=template_workbook,
            incomer_power_supply=incomer_power_supply,
        )
    elif division_name == "Enviro":
        template_workbook = get_enviro_load_list_excel(
            electrical_load_list_data=electrical_load_list_data,
            panels_data=panels_data,
            template_workbook=template_workbook,
            incomer_power_supply=incomer_power_supply,
        )
    elif division_name == "WWS IPG":
        template_workbook = get_ipg_load_list_excel(
            electrical_load_list_data=electrical_load_list_data,
            panels_data=panels_data,
            template_workbook=template_workbook,
            incomer_power_supply=incomer_power_supply,
        )
    else:
        raise ValueError(f"Load list template is not present for : {division_name}")

    return template_workbook
