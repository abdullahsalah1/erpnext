# Copyright (c) 2024, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class PurchasesOrder(Document):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from erpnext.operation.doctype.po_table.po_table import POtable
		from frappe.types import DF

		amended_from: DF.Link | None
		branch: DF.Literal["Sana'a", "Aden", "Marib"]
		date: DF.Date
		department: DF.Data | None
		item: DF.Table[POtable]
		naming_series: DF.Literal["P-O-"]
		posting_date: DF.Datetime | None
		purpose_of_the_order: DF.Data
		requested_by: DF.Link
	# end: auto-generated types
	pass
