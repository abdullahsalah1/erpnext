# Copyright (c) 2024, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class POtable(Document):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from frappe.types import DF

		details: DF.Data | None
		notes_ملاحظات: DF.Data | None
		parent: DF.Data
		parentfield: DF.Data
		parenttype: DF.Data
		price: DF.Currency
		quantity: DF.Data | None
		total: DF.Float
		نوع_الوحدة__unit_type: DF.Link | None
	# end: auto-generated types
	pass
