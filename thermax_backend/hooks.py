app_name = "thermax_backend"
app_title = "Thermax Backend"
app_publisher = "Abhishek Bankar"
app_description = "Thermax buisiness logic app"
app_email = "abhishek.bankar@inventivebizsol.com"
app_license = "mit"

# Apps
# ------------------

# required_apps = []

# Each item in the list will be shown as an app in the apps page
# add_to_apps_screen = [
# 	{
# 		"name": "thermax_backend",
# 		"logo": "/assets/thermax_backend/logo.png",
# 		"title": "Thermax Backend",
# 		"route": "/thermax_backend",
# 		"has_permission": "thermax_backend.api.permission.has_app_permission"
# 	}
# ]

# Includes in <head>
# ------------------

# include js, css files in header of desk.html
# app_include_css = "/assets/thermax_backend/css/thermax_backend.css"
# app_include_js = "/assets/thermax_backend/js/thermax_backend.js"

# include js, css files in header of web template
# web_include_css = "/assets/thermax_backend/css/thermax_backend.css"
# web_include_js = "/assets/thermax_backend/js/thermax_backend.js"

# include custom scss in every website theme (without file extension ".scss")
# website_theme_scss = "thermax_backend/public/scss/website"

# include js, css files in header of web form
# webform_include_js = {"doctype": "public/js/doctype.js"}
# webform_include_css = {"doctype": "public/css/doctype.css"}

# include js in page
# page_js = {"page" : "public/js/file.js"}

# include js in doctype views
# doctype_js = {"doctype" : "public/js/doctype.js"}
# doctype_list_js = {"doctype" : "public/js/doctype_list.js"}
# doctype_tree_js = {"doctype" : "public/js/doctype_tree.js"}
# doctype_calendar_js = {"doctype" : "public/js/doctype_calendar.js"}

# Svg Icons
# ------------------
# include app icons in desk
# app_include_icons = "thermax_backend/public/icons.svg"

# Home Pages
# ----------

# application home page (will override Website Settings)
# home_page = "login"

# website user home page (by Role)
# role_home_page = {
# 	"Role": "home_page"
# }

# Generators
# ----------

# automatically create page for each record of this doctype
# website_generators = ["Web Page"]

# Jinja
# ----------

# add methods and filters to jinja environment
# jinja = {
# 	"methods": "thermax_backend.utils.jinja_methods",
# 	"filters": "thermax_backend.utils.jinja_filters"
# }

# Installation
# ------------

# before_install = "thermax_backend.install.before_install"
# after_install = "thermax_backend.install.after_install"

# Uninstallation
# ------------

# before_uninstall = "thermax_backend.uninstall.before_uninstall"
# after_uninstall = "thermax_backend.uninstall.after_uninstall"

# Integration Setup
# ------------------
# To set up dependencies/integrations with other apps
# Name of the app being installed is passed as an argument

# before_app_install = "thermax_backend.utils.before_app_install"
# after_app_install = "thermax_backend.utils.after_app_install"

# Integration Cleanup
# -------------------
# To clean up dependencies/integrations with other apps
# Name of the app being uninstalled is passed as an argument

# before_app_uninstall = "thermax_backend.utils.before_app_uninstall"
# after_app_uninstall = "thermax_backend.utils.after_app_uninstall"

# Desk Notifications
# ------------------
# See frappe.core.notifications.get_notification_config

# notification_config = "thermax_backend.notifications.get_notification_config"

# Permissions
# -----------
# Permissions evaluated in scripted ways

# permission_query_conditions = {
# 	"Event": "frappe.desk.doctype.event.event.get_permission_query_conditions",
# }
#
# has_permission = {
# 	"Event": "frappe.desk.doctype.event.event.has_permission",
# }

# DocType Class
# ---------------
# Override standard doctype classes

# override_doctype_class = {
# 	"ToDo": "custom_app.overrides.CustomToDo"
# }

# Document Events
# ---------------
# Hook on document methods and events

# doc_events = {
# 	"*": {
# 		"on_update": "method",
# 		"on_cancel": "method",
# 		"on_trash": "method"
# 	}
# }

# Scheduled Tasks
# ---------------

# scheduler_events = {
# 	"all": [
# 		"thermax_backend.tasks.all"
# 	],
# 	"daily": [
# 		"thermax_backend.tasks.daily"
# 	],
# 	"hourly": [
# 		"thermax_backend.tasks.hourly"
# 	],
# 	"weekly": [
# 		"thermax_backend.tasks.weekly"
# 	],
# 	"monthly": [
# 		"thermax_backend.tasks.monthly"
# 	],
# }

# Testing
# -------

# before_tests = "thermax_backend.install.before_tests"

# Overriding Methods
# ------------------------------

override_whitelisted_methods = {
    "main_package.get_main_package_list": "thermax_backend.thermax_backend.doctype.main_package.api.get_main_package_list",
    "thermax_extended_user.trigger_email_verification_mail": "thermax_backend.thermax_backend.doctype.thermax_extended_user.api.trigger_email_verification_mail",
    "thermax_extended_user.trigger_send_credentials": "thermax_backend.thermax_backend.doctype.thermax_extended_user.api.trigger_send_credentials",
    "thermax_extended_user.get_user_by_role": "thermax_backend.thermax_backend.doctype.thermax_extended_user.api.get_user_by_role",
    "project_main_package.get_project_main_package_list": "thermax_backend.thermax_backend.doctype.project_main_package.api.get_project_main_package_list",
    "thermax_extended_user.trigger_delete_user": "thermax_backend.thermax_backend.doctype.thermax_extended_user.api.trigger_delete_user",
    "project.trigger_approver_notification_mail": "thermax_backend.thermax_backend.doctype.project.api.trigger_approver_notification_mail",
    "db_revision.trigger_review_submission_mail": "thermax_backend.thermax_backend.doctype.design_basis_revision_history.api.trigger_review_submission_mail",
    "db_revision.trigger_review_resubmission_mail": "thermax_backend.thermax_backend.doctype.design_basis_revision_history.api.trigger_review_resubmission_mail",
    "db_revision.trigger_review_approval_mail": "thermax_backend.thermax_backend.doctype.design_basis_revision_history.api.trigger_review_approval_mail",
    "db_revision.get_design_basis_excel": "thermax_backend.thermax_backend.doctype.design_basis_revision_history.api.get_design_basis_excel",
    "load_list_revisions.get_load_list_excel": "thermax_backend.thermax_backend.doctype.load_list_revisions.api.get_load_list_excel",
    "cable_schedule_revisions.get_voltage_drop_excel": "thermax_backend.thermax_backend.doctype.cable_schedule_revisions.api.get_voltage_drop_excel",
    "cable_schedule_revisions.get_cable_schedule_excel": "thermax_backend.thermax_backend.doctype.cable_schedule_revisions.api.get_cable_schedule_excel",
    "project_information.get_project_info_dropdown_data": "thermax_backend.thermax_backend.doctype.project_information.api.get_project_info_dropdown_data",
    "local_isolator_revisions.get_local_isolator_excel":"thermax_backend.thermax_backend.doctype.local_isolator_revisions.api.get_local_isolator_excel",
    "lpbs_specification_revisions.get_lpbs_specification_excel":"thermax_backend.thermax_backend.doctype.lpbs_specification_revisions.api.get_lpbs_specification_excel",
    "motor_specification_revisions.get_motor_specification_excel":"thermax_backend.thermax_backend.doctype.motor_specification_revisions.api.get_motor_specification_excel",
    "motor_canopy_revisions.get_motor_canopy_excel":"thermax_backend.thermax_backend.doctype.motor_canopy_revisions.api.get_motor_canopy_excel",
    "panel_specifications_revisions.get_panel_specification_excel":"thermax_backend.thermax_backend.doctype.panel_specifications_revisions.api.get_panel_specification_excel",

    
    "design_basis_make_of_component.get_make_of_component_dropdowns": "thermax_backend.thermax_backend.doctype.design_basis_make_of_component.api.get_make_of_component_dropdowns",
    "design_basis_motor_parameters.get_motor_parameters_dropdowns": "thermax_backend.thermax_backend.doctype.design_basis_motor_parameters.api.get_motor_parameters_dropdowns",
    "common_configuration.get_common_config_dropdown": "thermax_backend.thermax_backend.doctype.common_configuration.api.get_common_config_dropdown",
    "pcc_panel.get_pcc_panel_dropdown": "thermax_backend.thermax_backend.doctype.pcc_panel.api.get_pcc_panel_dropdown",
    "mcc_panel.get_mcc_panel_dropdown": "thermax_backend.thermax_backend.doctype.mcc_panel.api.get_mcc_panel_dropdown",
    "plc_panel.get_plc_panel_dropdown": "thermax_backend.thermax_backend.doctype.plc_panel.api.get_plc_panel_dropdown",
    "layout_earthing.get_layout_earthing_dropdown": "thermax_backend.thermax_backend.doctype.layout_earthing.api.get_layout_earthing_dropdown",
    "cable_tray_layout.get_cable_tray_layout_dropdown": "thermax_backend.thermax_backend.doctype.cable_tray_layout.api.get_cable_tray_layout_dropdown",
    "sld_revisions.get_data_for_sld_generation": "thermax_backend.thermax_backend.doctype.sld_revisions.api.get_data_for_sld_generation",
    "project.send_custom_mail": "thermax_backend.thermax_backend.doctype.project.api.send_custom_mail",
}
#
# each overriding function accepts a `data` argument;
# generated from the base implementation of the doctype dashboard,
# along with any modifications made in other Frappe apps
# override_doctype_dashboards = {
# 	"Task": "thermax_backend.task.get_dashboard_data"
# }

# exempt linked doctypes from being automatically cancelled
#
# auto_cancel_exempted_doctypes = ["Auto Repeat"]

# Ignore links to specified DocTypes when deleting documents
# -----------------------------------------------------------

# ignore_links_on_delete = ["Communication", "ToDo"]

# Request Events
# ----------------
# before_request = ["thermax_backend.utils.before_request"]
# after_request = ["thermax_backend.utils.after_request"]

# Job Events
# ----------
# before_job = ["thermax_backend.utils.before_job"]
# after_job = ["thermax_backend.utils.after_job"]

# User Data Protection
# --------------------

# user_data_fields = [
# 	{
# 		"doctype": "{doctype_1}",
# 		"filter_by": "{filter_by}",
# 		"redact_fields": ["{field_1}", "{field_2}"],
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_2}",
# 		"filter_by": "{filter_by}",
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_3}",
# 		"strict": False,
# 	},
# 	{
# 		"doctype": "{doctype_4}"
# 	}
# ]

# Authentication and authorization
# --------------------------------

# auth_hooks = [
# 	"thermax_backend.auth.validate"
# ]

# Automatically update python controller files with type annotations for this app.
# export_python_type_annotations = True

# default_log_clearing_doctypes = {
# 	"Logging DocType Name": 30  # days to retain logs
# }
