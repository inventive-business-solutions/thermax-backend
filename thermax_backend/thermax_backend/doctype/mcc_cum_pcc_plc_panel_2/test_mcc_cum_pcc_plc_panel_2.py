# Copyright (c) 2024, Abhishek Bankar and Contributors
# See license.txt

# import frappe
from frappe.tests import IntegrationTestCase, UnitTestCase


# On IntegrationTestCase, the doctype test records and all
# link-field test record depdendencies are recursively loaded
# Use these module variables to add/remove to/from that list
EXTRA_TEST_RECORD_DEPENDENCIES = []  # eg. ["User"]
IGNORE_TEST_RECORD_DEPENDENCIES = []  # eg. ["User"]


class TestMCCCumPCCPLCPanel2(UnitTestCase):
	"""
	Unit tests for MCCCumPCCPLCPanel2.
	Use this class for testing individual functions and methods.
	"""

	pass


class TestMCCCumPCCPLCPanel2(IntegrationTestCase):
	"""
	Integration tests for MCCCumPCCPLCPanel2.
	Use this class for testing interactions between multiple components.
	"""

	pass
