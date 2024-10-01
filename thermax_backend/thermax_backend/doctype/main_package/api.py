from frappe import frappe

@frappe.whitelist()
def get_main_package_list():
    # Get all Main Package records
    main_packages = frappe.db.get_list("Main Package", fields=["*"], order_by="creation asc")  # Specify the fields you want to retrieve

    for main_package in main_packages:
        sub_packages = frappe.db.get_list("Sub Package", fields=["*"], filters={"main_package_name": main_package.get('name')}, order_by="creation asc")

        main_package["sub_packages"] = sub_packages

    
    
    return main_packages