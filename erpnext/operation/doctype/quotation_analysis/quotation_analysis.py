# Copyright (c) 2024, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt

# import frappe
from frappe.model.document import Document


class QuotationAnalysis(Document):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from erpnext.operation.doctype.suppliers_table.suppliers_table import SuppliersTable
		from frappe.types import DF

		amount: DF.Data | None
		based_on_the_above_analysis_the_supplier_is_selected: DF.Data | None
		estimated_price: DF.Currency
		item: DF.Data | None
		purchase_order: DF.Link | None
		reasons_to_choose_a_supplier: DF.Data | None
		requested_by: DF.Link | None
		suppliers: DF.Table[SuppliersTable]
		the_discount_percentage_if_any: DF.Currency
		the_discount_value_if_any: DF.Currency
		the_general_total_after_discount: DF.Currency
		total: DF.Currency
	# end: auto-generated types
	pass
