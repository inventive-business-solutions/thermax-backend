import frappe


@frappe.whitelist()
def get_data_for_sld_generation(project_id):
    """
    Get data for SLD generation
    """
    sld_data = {}
    # Project Details
    project = frappe.get_doc("Project", project_id).as_dict()

    # User Details
    thermax_extended_user = frappe.get_doc(
        "Thermax Extended User", project.get("owner")
    ).as_dict()

    # SLD Revisions
    sld_revision = frappe.db.get_list(
        "SLD Revisions",
        fields=["*"],
        filters={"project_id": project_id},
        order_by="creation asc",
    )
    for index, revision in enumerate(sld_revision):
        revision["revision_number"] = f"R{index}"

    # Static Document List
    static_document_list = frappe.get_doc("Static Document List", project_id).as_dict()

    # Desgin Basis Revision List
    design_basis_revision = frappe.db.get_list(
        "Design Basis Revision History",
        fields=["*"],
        filters={"project_id": project_id},
        order_by="creation asc",
    )

    latest_design_basis_revision = design_basis_revision[-1]
    latest_revision_id = latest_design_basis_revision.get("name")

    # Project Panel Data
    project_panels = frappe.db.get_list(
        "Project Panel Data",
        fields=["*"],
        filters={"revision_id": latest_revision_id},
        order_by="creation asc",
    )
    for panel in project_panels:
        panel_id = panel.get("name")
        dynamic_doc = frappe.get_doc("Dynamic Document List", panel_id).as_dict()
        panel["dynamic_doc_name"] = dynamic_doc

    sld_data["user"] = thermax_extended_user
    sld_data["project"] = project
    sld_data["sld_revision"] = sld_revision
    sld_data["static_document_list"] = static_document_list
    sld_data["panels"] = project_panels
    return sld_data
