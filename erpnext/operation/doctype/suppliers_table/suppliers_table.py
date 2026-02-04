# Copyright (c) 2024, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class SuppliersTable(Document):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from frappe.types import DF

		attached: DF.Attach | None
		best_option: DF.Check
		parent: DF.Data
		parentfield: DF.Data
		parenttype: DF.Data
		supplier_name: DF.Data | None
		total: DF.Data | None
		unitp: DF.Data | None
	# end: auto-generated types
	pass
