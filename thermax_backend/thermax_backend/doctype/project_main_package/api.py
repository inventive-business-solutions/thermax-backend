from frappe import frappe

@frappe.whitelist()
def get_project_main_package_list():
    revision_id = frappe.request.args.get("revision_id")
    # Get all Main Package records
    main_packages = frappe.db.get_list("Project Main Package", fields=["*"], filters={"revision_id": revision_id}, order_by="creation asc")

    for main_package in main_packages:
        # Get all Sub Package records
        sub_packages = frappe.get_doc("Project Main Package", main_package["name"]).sub_packages
        main_package["sub_packages"] = sub_packages    
    
    return main_packages