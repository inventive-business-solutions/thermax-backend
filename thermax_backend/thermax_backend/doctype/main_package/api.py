import frappe

@frappe.whitelist()
def get_main_package_list():
    # Get all Main Package records and link Sub Package records
    main_packages = frappe.get_all("Main Package", fields=["name", "package_name"], order_by="creation asc")

    result = []

    for main_package in main_packages:
        # Fetch the full document to access the get_children method
        main_package_doc = frappe.get_doc('Main Package', main_package['name'])
        
        sub_packages = sorted(main_package_doc.sub_packages, key=lambda x: x.creation)
        
        # Combine the main package data with its sub packages
        main_package_with_subs = {
            **main_package,
            'sub_packages': sub_packages
        }
        
        result.append(main_package_with_subs)
    return result