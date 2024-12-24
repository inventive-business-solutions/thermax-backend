import frappe
from frappe import _

import io


def create_dropdown_options(data, dropdown_key):
    if data is None:
        return []
    
    return [
        {**item, 'value': item[dropdown_key], 'label': item[dropdown_key]}
        for item in data
    ]

@frappe.whitelist()
def get_make_of_component_dropdowns():

    fields = frappe.local.form_dict
    result= {}

    for doctype, key in fields.items():
        try:
            record = frappe.get_all(doctype, fields=[key], order_by="name asc")
            formatted_options = create_dropdown_options(record, key)

            result[doctype] = formatted_options

        except Exception as e:
            result[doctype] =  {"Error": str(e)}
    
    return result