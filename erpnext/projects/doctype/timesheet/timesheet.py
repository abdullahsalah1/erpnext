# Copyright (c) 2015, Frappe Technologies Pvt. Ltd. and contributors
# For license information, please see license.txt


import json

import frappe
from frappe import _
from frappe.model.document import Document
from frappe.utils import add_to_date, flt, get_datetime, getdate, time_diff_in_hours
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from frappe.utils.file_manager import save_file

from erpnext.controllers.queries import get_match_cond
from erpnext.setup.utils import get_exchange_rate


class OverlapError(frappe.ValidationError):
	pass


class OverWorkLoggedError(frappe.ValidationError):
	pass


class Timesheet(Document):
	# begin: auto-generated types
	# This code is auto-generated. Do not modify anything in this block.

	from typing import TYPE_CHECKING

	if TYPE_CHECKING:
		from frappe.types import DF

		from erpnext.projects.doctype.timesheet_detail.timesheet_detail import TimesheetDetail

		amended_from: DF.Link | None
		annual_leave_rnr: DF.Float
		approved_by: DF.Data | None
		approval_date: DF.Date | None
		base_total_billable_amount: DF.Currency
		base_total_billed_amount: DF.Currency
		base_total_costing_amount: DF.Currency
		code: DF.Data | None
		company: DF.Link | None
		corresponding_unit: DF.Data | None
		currency: DF.Link | None
		customer: DF.Link | None
		date: DF.Date | None
		department: DF.Link | None
		employee: DF.Link | None
		employee_name: DF.Data | None
		end_date: DF.Date | None
		exchange_rate: DF.Float
		leave_without_pay: DF.Float
		name: DF.Data | None
		name_signature: DF.Data | None
		national_holidays: DF.Float
		naming_series: DF.Literal["TS-.YYYY.-"]
		note: DF.TextEditor | None
		parent_project: DF.Link | None
		payment_period: DF.Literal["Jan-25", "Feb-25", "Mar-25", "Apr-25", "May-25", "Jun-25", "Jul-25", "Aug-25", "Sep-25", "Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26", "Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27", "Jul-27", "Aug-27", "Sep-27", "Oct-27", "Nov-27", "Dec-27", "Jan-28", "Feb-28", "Mar-28", "Apr-28", "May-28", "Jun-28", "Jul-28", "Aug-28", "Sep-28", "Oct-28", "Nov-28", "Dec-28", "Jan-29", "Feb-29", "Mar-29", "Apr-29", "May-29", "Jun-29", "Jul-29", "Aug-29", "Sep-29", "Oct-29", "Nov-29", "Dec-29", "Jan-30", "Feb-30", "Mar-30", "Apr-30", "May-30", "Jun-30", "Jul-30", "Aug-30", "Sep-30"] | None
		per_billed: DF.Percent
		position: DF.Data | None
		revision_date: DF.Date | None
		roles: DF.Data | None
		sales_invoice: DF.Link | None
		sick_leave: DF.Float
		start_date: DF.Date | None
		status: DF.Literal["Draft", "Submitted", "Billed", "Payslip", "Completed", "Cancelled"]
		time_logs: DF.Table[TimesheetDetail]
		title: DF.Data | None
		total_billable_amount: DF.Currency
		total_billable_hours: DF.Float
		total_billed_amount: DF.Currency
		total_billed_hours: DF.Float
		total_costing_amount: DF.Currency
		total_hours: DF.Float
		user: DF.Link | None
		version: DF.Data | None
		work_week: DF.Literal["Monday-Friday", "Sunday-Thursday"] | None
	# end: auto-generated types

	def validate(self):
		self.set_status()
		# Guard time_logs-dependent logic when child table is not present
		self.validate_dates()
		self.calculate_hours()
		self.validate_time_logs()
		self.update_cost()
		self.calculate_total_amounts()
		self.calculate_percentage_billed()
		self.set_dates()

	def calculate_hours(self):
		for row in (self.get("time_logs") or []):
			if row.to_time and row.from_time:
				row.hours = time_diff_in_hours(row.to_time, row.from_time)

	def calculate_total_amounts(self):
		self.total_hours = 0.0
		self.total_billable_hours = 0.0
		self.total_billed_hours = 0.0
		self.total_billable_amount = self.base_total_billable_amount = 0.0
		self.total_costing_amount = self.base_total_costing_amount = 0.0
		self.total_billed_amount = self.base_total_billed_amount = 0.0

		for d in (self.get("time_logs") or []):
			self.update_billing_hours(d)
			self.update_time_rates(d)

			self.total_hours += flt(d.hours)
			self.total_costing_amount += flt(d.costing_amount)
			self.base_total_costing_amount += flt(d.base_costing_amount)
			if d.is_billable:
				self.total_billable_hours += flt(d.billing_hours)
				self.total_billable_amount += flt(d.billing_amount)
				self.base_total_billable_amount += flt(d.base_billing_amount)
				self.total_billed_amount += flt(d.billing_amount) if d.sales_invoice else 0.0
				self.base_total_billed_amount += flt(d.base_billing_amount) if d.sales_invoice else 0.0
				self.total_billed_hours += flt(d.billing_hours) if d.sales_invoice else 0.0

	def calculate_percentage_billed(self):
		self.per_billed = 0
		if self.total_billed_amount > 0 and self.total_billable_amount > 0:
			self.per_billed = (self.total_billed_amount * 100) / self.total_billable_amount
		elif self.total_billed_hours > 0 and self.total_billable_hours > 0:
			self.per_billed = (self.total_billed_hours * 100) / self.total_billable_hours

	def update_billing_hours(self, args):
		if args.is_billable:
			if flt(args.billing_hours) == 0.0:
				args.billing_hours = args.hours
			elif flt(args.billing_hours) > flt(args.hours):
				frappe.msgprint(
					_("Warning - Row {0}: Billing Hours are more than Actual Hours").format(args.idx),
					indicator="orange",
					alert=True,
				)
		else:
			args.billing_hours = 0

	def set_status(self):
		self.status = {"0": "Draft", "1": "Submitted", "2": "Cancelled"}[str(self.docstatus or 0)]

		if flt(self.per_billed, self.precision("per_billed")) >= 100.0:
			self.status = "Billed"

		if self.sales_invoice:
			self.status = "Completed"

	def set_dates(self):
		if self.docstatus < 2 and (self.get("time_logs") or []):
			start_date = min(getdate(d.from_time) for d in self.get("time_logs"))
			end_date = max(getdate(d.to_time) for d in self.get("time_logs"))

			if start_date and end_date:
				self.start_date = getdate(start_date)
				self.end_date = getdate(end_date)

	def before_cancel(self):
		self.set_status()

	def on_cancel(self):
		self.update_task_and_project()

	def on_submit(self):
		self.validate_mandatory_fields()
		self.update_task_and_project()

	def validate_mandatory_fields(self):
		for data in (self.get("time_logs") or []):
			if not data.from_time and not data.to_time:
				frappe.throw(_("Row {0}: From Time and To Time is mandatory.").format(data.idx))

			if not data.activity_type and self.employee:
				frappe.throw(_("Row {0}: Activity Type is mandatory.").format(data.idx))

			if flt(data.hours) == 0.0:
				frappe.throw(_("Row {0}: Hours value must be greater than zero.").format(data.idx))

	def update_task_and_project(self):
		tasks, projects = [], []

		for data in (self.get("time_logs") or []):
			if data.task and data.task not in tasks:
				task = frappe.get_doc("Task", data.task)
				task.update_time_and_costing()
				task.save(ignore_permissions=True)
				tasks.append(data.task)

			if data.project and data.project not in projects:
				projects.append(data.project)

		for project in projects:
			project_doc = frappe.get_doc("Project", project)
			project_doc.update_project()
			project_doc.save(ignore_permissions=True)

	def validate_dates(self):
		for data in (self.get("time_logs") or []):
			if data.from_time and data.to_time and time_diff_in_hours(data.to_time, data.from_time) < 0:
				frappe.throw(_("To date cannot be before from date"))

	def validate_time_logs(self):
		for data in (self.get("time_logs") or []):
			self.set_to_time(data)
			self.validate_overlap(data)
			self.set_project(data)
			self.validate_project(data)

	def set_to_time(self, data):
		if not (data.from_time and data.hours):
			return

		_to_time = get_datetime(add_to_date(data.from_time, hours=data.hours, as_datetime=True))
		if data.to_time != _to_time:
			data.to_time = _to_time

	def validate_overlap(self, data):
		settings = frappe.get_single("Projects Settings")
		self.validate_overlap_for("user", data, self.user, settings.ignore_user_time_overlap)
		self.validate_overlap_for("employee", data, self.employee, settings.ignore_employee_time_overlap)

	def set_project(self, data):
		data.project = data.project or frappe.db.get_value("Task", data.task, "project")

	def validate_project(self, data):
		if self.parent_project and self.parent_project != data.project:
			frappe.throw(
				_("Row {0}: Project must be same as the one set in the Timesheet: {1}.").format(
					data.idx, self.parent_project
				)
			)

	def validate_overlap_for(self, fieldname, args, value, ignore_validation=False):
		if not value or ignore_validation:
			return

		existing = self.get_overlap_for(fieldname, args, value)
		if existing:
			frappe.throw(
				_("Row {0}: From Time and To Time of {1} is overlapping with {2}").format(
					args.idx, self.name, existing.name
				),
				OverlapError,
			)

	def get_overlap_for(self, fieldname, args, value):
		timesheet = frappe.qb.DocType("Timesheet")
		timelog = frappe.qb.DocType("Timesheet Detail")

		from_time = get_datetime(args.from_time)
		to_time = get_datetime(args.to_time)

		existing = (
			frappe.qb.from_(timesheet)
			.join(timelog)
			.on(timelog.parent == timesheet.name)
			.select(
				timesheet.name.as_("name"), timelog.from_time.as_("from_time"), timelog.to_time.as_("to_time")
			)
			.where(
				(timelog.name != (args.name or "No Name"))
				& (timesheet.name != (args.parent or "No Name"))
				& (timesheet.docstatus < 2)
				& (timesheet[fieldname] == value)
				& (
					((from_time > timelog.from_time) & (from_time < timelog.to_time))
					| ((to_time > timelog.from_time) & (to_time < timelog.to_time))
					| ((from_time <= timelog.from_time) & (to_time >= timelog.to_time))
				)
			)
		).run(as_dict=True)

		if self.check_internal_overlap(fieldname, args):
			return self

		return existing[0] if existing else None

	def check_internal_overlap(self, fieldname, args):
		for time_log in (self.get("time_logs") or []):
			if not (time_log.from_time and time_log.to_time and args.from_time and args.to_time):
				continue

			from_time = get_datetime(time_log.from_time)
			to_time = get_datetime(time_log.to_time)
			args_from_time = get_datetime(args.from_time)
			args_to_time = get_datetime(args.to_time)

			if (
				(args.get(fieldname) == time_log.get(fieldname))
				and (args.idx != time_log.idx)
				and (
					(args_from_time > from_time and args_from_time < to_time)
					or (args_to_time > from_time and args_to_time < to_time)
					or (args_from_time <= from_time and args_to_time >= to_time)
				)
			):
				return True
		return False

	def update_cost(self):
		for data in (self.get("time_logs") or []):
			if data.activity_type or data.is_billable:
				rate = get_activity_cost(self.employee, data.activity_type)
				hours = data.billing_hours or 0
				costing_hours = data.billing_hours or data.hours or 0
				if rate:
					data.billing_rate = (
						flt(rate.get("billing_rate")) if flt(data.billing_rate) == 0 else data.billing_rate
					)
					data.costing_rate = (
						flt(rate.get("costing_rate")) if flt(data.costing_rate) == 0 else data.costing_rate
					)
					data.billing_amount = data.billing_rate * hours
					data.costing_amount = data.costing_rate * costing_hours

	def update_time_rates(self, ts_detail):
		if not ts_detail.is_billable:
			ts_detail.billing_rate = 0.0


@frappe.whitelist()
def get_projectwise_timesheet_data(project=None, parent=None, from_time=None, to_time=None):
	condition = ""
	if project:
		condition += "AND tsd.project = %(project)s "
	if parent:
		condition += "AND tsd.parent = %(parent)s "
	if from_time and to_time:
		condition += "AND CAST(tsd.from_time as DATE) BETWEEN %(from_time)s AND %(to_time)s"

	query = f"""
		SELECT
			tsd.name as name,
			tsd.parent as time_sheet,
			tsd.from_time as from_time,
			tsd.to_time as to_time,
			tsd.billing_hours as billing_hours,
			tsd.billing_amount as billing_amount,
			tsd.activity_type as activity_type,
			tsd.description as description,
			ts.currency as currency,
			tsd.project_name as project_name
		FROM `tabTimesheet Detail` tsd
			INNER JOIN `tabTimesheet` ts
			ON ts.name = tsd.parent
		WHERE
			tsd.parenttype = 'Timesheet'
			AND tsd.docstatus = 1
			AND tsd.is_billable = 1
			AND tsd.sales_invoice is NULL
			{condition}
		ORDER BY tsd.from_time ASC
	"""

	filters = {"project": project, "parent": parent, "from_time": from_time, "to_time": to_time}

	return frappe.db.sql(query, filters, as_dict=1)


@frappe.whitelist()
def get_timesheet_detail_rate(timelog, currency):
	timelog_detail = frappe.db.sql(
		f"""SELECT tsd.billing_amount as billing_amount,
		ts.currency as currency FROM `tabTimesheet Detail` tsd
		INNER JOIN `tabTimesheet` ts ON ts.name=tsd.parent
		WHERE tsd.name = '{timelog}'""",
		as_dict=1,
	)[0]

	if timelog_detail.currency:
		exchange_rate = get_exchange_rate(timelog_detail.currency, currency)

		return timelog_detail.billing_amount * exchange_rate
	return timelog_detail.billing_amount


@frappe.whitelist()
@frappe.validate_and_sanitize_search_inputs
def get_timesheet(doctype, txt, searchfield, start, page_len, filters):
	if not filters:
		filters = {}

	condition = ""
	if filters.get("project"):
		condition = "and tsd.project = %(project)s"

	return frappe.db.sql(
		f"""select distinct tsd.parent from `tabTimesheet Detail` tsd,
			`tabTimesheet` ts where
			ts.status in ('Submitted', 'Payslip') and tsd.parent = ts.name and
			tsd.docstatus = 1 and ts.total_billable_amount > 0
			and tsd.parent LIKE %(txt)s {condition}
			order by tsd.parent limit %(page_len)s offset %(start)s""",
		{
			"txt": "%" + txt + "%",
			"start": start,
			"page_len": page_len,
			"project": filters.get("project"),
		},
	)


@frappe.whitelist()
def get_timesheet_data(name, project):
	data = None
	if project and project != "":
		data = get_projectwise_timesheet_data(project, name)
	else:
		data = frappe.get_all(
			"Timesheet",
			fields=[
				"(total_billable_amount - total_billed_amount) as billing_amt",
				"total_billable_hours as billing_hours",
			],
			filters={"name": name},
		)
	return {
		"billing_hours": data[0].billing_hours if data else None,
		"billing_amount": data[0].billing_amt if data else None,
		"timesheet_detail": data[0].name if data and project and project != "" else None,
	}


@frappe.whitelist()
def make_sales_invoice(source_name, item_code=None, customer=None, currency=None):
	target = frappe.new_doc("Sales Invoice")
	timesheet = frappe.get_doc("Timesheet", source_name)

	if not timesheet.total_billable_hours:
		frappe.throw(_("Invoice can't be made for zero billing hour"))

	if timesheet.total_billable_hours == timesheet.total_billed_hours:
		frappe.throw(_("Invoice already created for all billing hours"))

	hours = flt(timesheet.total_billable_hours) - flt(timesheet.total_billed_hours)
	billing_amount = flt(timesheet.total_billable_amount) - flt(timesheet.total_billed_amount)
	billing_rate = billing_amount / hours

	target.company = timesheet.company
	target.project = timesheet.parent_project
	if customer:
		target.customer = customer
		default_price_list = frappe.get_value("Customer", customer, "default_price_list")
		if default_price_list:
			target.selling_price_list = default_price_list

	if currency:
		target.currency = currency

	if item_code:
		target.append("items", {"item_code": item_code, "qty": hours, "rate": billing_rate})

	for time_log in timesheet.time_logs:
		if time_log.is_billable:
			target.append(
				"timesheets",
				{
					"time_sheet": timesheet.name,
					"project_name": time_log.project_name,
					"from_time": time_log.from_time,
					"to_time": time_log.to_time,
					"billing_hours": time_log.billing_hours,
					"billing_amount": time_log.billing_amount,
					"timesheet_detail": time_log.name,
					"activity_type": time_log.activity_type,
					"description": time_log.description,
				},
			)

	target.run_method("calculate_billing_amount_for_timesheet")
	target.run_method("set_missing_values")

	return target


@frappe.whitelist()
def get_activity_cost(employee=None, activity_type=None, currency=None):
	base_currency = frappe.defaults.get_global_default("currency")
	rate = frappe.db.get_values(
		"Activity Cost",
		{"employee": employee, "activity_type": activity_type},
		["costing_rate", "billing_rate"],
		as_dict=True,
	)
	if not rate:
		rate = frappe.db.get_values(
			"Activity Type",
			{"activity_type": activity_type},
			["costing_rate", "billing_rate"],
			as_dict=True,
		)
		if rate and currency and currency != base_currency:
			exchange_rate = get_exchange_rate(base_currency, currency)
			rate[0]["costing_rate"] = rate[0]["costing_rate"] * exchange_rate
			rate[0]["billing_rate"] = rate[0]["billing_rate"] * exchange_rate

	return rate[0] if rate else {}


@frappe.whitelist()
def get_events(start, end, filters=None):
	"""Returns events for Gantt / Calendar view rendering.
	:param start: Start date-time.
	:param end: End date-time.
	:param filters: Filters (JSON).
	"""
	filters = json.loads(filters)
	from frappe.desk.calendar import get_event_conditions

	conditions = get_event_conditions("Timesheet", filters)

	return frappe.db.sql(
		"""select `tabTimesheet Detail`.name as name,
			`tabTimesheet Detail`.docstatus as status, `tabTimesheet Detail`.parent as parent,
			from_time as start_date, hours, activity_type,
			`tabTimesheet Detail`.project, to_time as end_date,
			CONCAT(`tabTimesheet Detail`.parent, ' (', ROUND(hours,2),' hrs)') as title
		from `tabTimesheet Detail`, `tabTimesheet`
		where `tabTimesheet Detail`.parent = `tabTimesheet`.name
			and `tabTimesheet`.docstatus < 2
			and (from_time <= %(end)s and to_time >= %(start)s) {conditions} {match_cond}
		""".format(conditions=conditions, match_cond=get_match_cond("Timesheet")),
		{"start": start, "end": end},
		as_dict=True,
		update={"allDay": 0},
	)


def get_timesheets_list(doctype, txt, filters, limit_start, limit_page_length=20, order_by="modified"):
	user = frappe.session.user
	# find customer name from contact.
	customer = ""
	timesheets = []

	contact = frappe.db.exists("Contact", {"user": user})
	if contact:
		# find customer
		contact = frappe.get_doc("Contact", contact)
		customer = contact.get_link_for("Customer")

	if customer:
		sales_invoices = [
			d.name for d in frappe.get_all("Sales Invoice", filters={"customer": customer})
		] or [None]
		projects = [d.name for d in frappe.get_all("Project", filters={"customer": customer})]
		# Return timesheet related data to web portal.
		timesheets = frappe.db.sql(
			f"""
			SELECT
				ts.name, tsd.activity_type, ts.status, ts.total_billable_hours,
				COALESCE(ts.sales_invoice, tsd.sales_invoice) AS sales_invoice, tsd.project
			FROM `tabTimesheet` ts, `tabTimesheet Detail` tsd
			WHERE tsd.parent = ts.name AND
				(
					ts.sales_invoice IN %(sales_invoices)s OR
					tsd.sales_invoice IN %(sales_invoices)s OR
					tsd.project IN %(projects)s
				)
			ORDER BY `end_date` ASC
			LIMIT {limit_page_length} offset {limit_start}
		""",
			dict(sales_invoices=sales_invoices, projects=projects),
			as_dict=True,
		)  # nosec

	return timesheets


def get_list_context(context=None):
	return {
		"show_sidebar": True,
		"show_search": True,
		"no_breadcrumbs": True,
		"title": _("Timesheets"),
		"get_list": get_timesheets_list,
		"row_template": "templates/includes/timesheet/timesheet_row.html",
	}


@frappe.whitelist()
def export_timesheet_to_excel(timesheet_name):
    timesheet = frappe.get_doc("Timesheet", timesheet_name)
    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet Export"

    # Header row
    headers = ["Employee", "Project", "Activity Type", "From Time", "To Time", "Hours", "Description"]
    ws.append(headers)

    for col in ws.iter_cols(min_row=1, max_row=1, min_col=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    # Timesheet Details Table
    for entry in timesheet.time_logs:
        row = [
            timesheet.employee,
            entry.project,
            entry.activity_type,
            entry.from_time.strftime("%Y-%m-%d %H:%M"),
            entry.to_time.strftime("%Y-%m-%d %H:%M"),
            entry.hours,
            entry.description,
        ]
        ws.append(row)

    # Adjust column widths
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    # Save the file and attach to Timesheet
    file_name = f"{timesheet_name}_Export.xlsx"
    file_path = frappe.utils.get_site_path("private", "files", file_name)
    wb.save(file_path)

    # Save as Frappe file and attach
    with open(file_path, "rb") as f:
        file_content = f.read()
    saved_file = save_file(file_name, file_content, "Timesheet", timesheet_name, is_private=1)
    frappe.msgprint(f"Excel file created: <a href='{saved_file.file_url}' target='_blank'>{file_name}</a>")

    return saved_file.file_url


@frappe.whitelist()
def export_multiple_timesheets(timesheet_names):
    if isinstance(timesheet_names, str):
        timesheet_names = json.loads(timesheet_names)
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Multiple Timesheets Export"

    # Header row
    headers = ["Timesheet", "Employee", "Project", "Activity Type", "From Time", "To Time", "Hours", "Description"]
    ws.append(headers)

    for col in ws.iter_cols(min_row=1, max_row=1, min_col=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

    # Process each timesheet
    for timesheet_name in timesheet_names:
        try:
            timesheet = frappe.get_doc("Timesheet", timesheet_name)
            
            # Add timesheet details
            for entry in timesheet.time_logs:
                row = [
                    timesheet.name,
                    timesheet.employee,
                    entry.project,
                    entry.activity_type,
                    entry.from_time.strftime("%Y-%m-%d %H:%M"),
                    entry.to_time.strftime("%Y-%m-%d %H:%M"),
                    entry.hours,
                    entry.description,
                ]
                ws.append(row)
        except Exception as e:
            frappe.log_error(f"Error processing timesheet {timesheet_name}: {str(e)}")

    # Adjust column widths
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length + 2

    # Save the file
    file_name = f"Multiple_Timesheets_Export_{frappe.utils.nowdate()}.xlsx"
    file_path = frappe.utils.get_site_path("private", "files", file_name)
    wb.save(file_path)

    # Save as Frappe file
    with open(file_path, "rb") as f:
        file_content = f.read()
    saved_file = save_file(file_name, file_content, "Timesheet", None, is_private=1)
    frappe.msgprint(f"Excel file created: <a href='{saved_file.file_url}' target='_blank'>{file_name}</a>")

    return saved_file.file_url


@frappe.whitelist()
def export_timesheet_to_formatted_excel(timesheet_name, template_file_name=None):
    """
    Export timesheet data to a formatted Excel template.
    If template_file_name is provided, it will use that template from File doctype.
    Otherwise, it will create a default formatted template.
    """
    try:
        frappe.logger().info(f"Starting formatted export for timesheet: {timesheet_name}, template: {template_file_name}")
        
        timesheet = frappe.get_doc("Timesheet", timesheet_name)
        
        if template_file_name:
            # Use existing template from File doctype
            template_file = frappe.get_doc("File", {"file_name": template_file_name})
            if not template_file:
                frappe.throw(f"Template file '{template_file_name}' not found in File doctype")
            
            frappe.logger().info(f"Using template: {template_file_name}")
            
            # Download template file
            template_path = frappe.get_site_path("private", "files", template_file.file_name)
            wb = load_workbook(template_path)
            ws = wb.active
        else:
            # Create default formatted template
            frappe.logger().info("Creating default formatted template")
            wb = Workbook()
            ws = wb.active
            ws.title = "Timesheet Export"
            
            # Create formatted header
            ws['A1'] = "TIMESHEET EXPORT"
            ws['A1'].font = Font(bold=True, size=16)
            ws.merge_cells('A1:H1')
            
            # Add timesheet info
            ws['A3'] = "Timesheet ID:"
            ws['B3'] = timesheet.name
            ws['A4'] = "Employee:"
            ws['B4'] = timesheet.employee
            ws['A5'] = "Period:"
            ws['B5'] = f"{timesheet.start_date} to {timesheet.end_date}"
            
            # Add headers starting from row 7
            headers = ["Date", "Project", "Activity Type", "From Time", "To Time", "Hours", "Description", "Billable"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=7, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Set data start row
            data_start_row = 8
        
        # Find the data section in the template
        data_start_row = None
        if template_file_name:
            # Look for common header patterns in the template
            for row in range(1, 50):  # Search first 50 rows
                for col in range(1, 10):  # Search first 10 columns
                    cell_value = ws.cell(row=row, column=col).value
                    if cell_value and any(header in str(cell_value).lower() for header in ['date', 'project', 'activity', 'time', 'hours']):
                        data_start_row = row + 1
                        break
                if data_start_row:
                    break
        
        if not data_start_row:
            data_start_row = 8  # Default if no template or no headers found
        
        frappe.logger().info(f"Data start row: {data_start_row}")
        
        # Populate data
        row_num = data_start_row
        for entry in timesheet.time_logs:
            if template_file_name:
                # Map data to template fields based on headers
                for col in range(1, ws.max_column + 1):
                    header_cell = ws.cell(row=data_start_row - 1, column=col)
                    if header_cell.value:
                        header = str(header_cell.value).lower()
                        if 'date' in header:
                            ws.cell(row=row_num, column=col, value=entry.from_time.strftime("%Y-%m-%d"))
                        elif 'project' in header:
                            ws.cell(row=row_num, column=col, value=entry.project)
                        elif 'activity' in header:
                            ws.cell(row=row_num, column=col, value=entry.activity_type)
                        elif 'from time' in header or 'start time' in header:
                            ws.cell(row=row_num, column=col, value=entry.from_time.strftime("%H:%M"))
                        elif 'to time' in header or 'end time' in header:
                            ws.cell(row=row_num, column=col, value=entry.to_time.strftime("%H:%M"))
                        elif 'hours' in header:
                            ws.cell(row=row_num, column=col, value=entry.hours)
                        elif 'description' in header:
                            ws.cell(row=row_num, column=col, value=entry.description)
                        elif 'billable' in header:
                            ws.cell(row=row_num, column=col, value="Yes" if entry.is_billable else "No")
            else:
                # Use default format
                ws.cell(row=row_num, column=1, value=entry.from_time.strftime("%Y-%m-%d"))
                ws.cell(row=row_num, column=2, value=entry.project)
                ws.cell(row=row_num, column=3, value=entry.activity_type)
                ws.cell(row=row_num, column=4, value=entry.from_time.strftime("%H:%M"))
                ws.cell(row=row_num, column=5, value=entry.to_time.strftime("%H:%M"))
                ws.cell(row=row_num, column=6, value=entry.hours)
                ws.cell(row=row_num, column=7, value=entry.description)
                ws.cell(row=row_num, column=8, value="Yes" if entry.is_billable else "No")
            
            row_num += 1
        
        # Auto-adjust column widths
        for column_cells in ws.columns:
            length = max(len(str(cell.value or "")) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 50)
        
        # Save the file
        file_name = f"{timesheet_name}_Formatted_Export.xlsx"
        file_path = frappe.utils.get_site_path("private", "files", file_name)
        wb.save(file_path)
        
        frappe.logger().info(f"File saved to: {file_path}")
        
        # Save as Frappe file and attach
        with open(file_path, "rb") as f:
            file_content = f.read()
        saved_file = save_file(file_name, file_content, "Timesheet", timesheet_name, is_private=1)
        
        frappe.logger().info(f"File saved to Frappe: {saved_file.file_url}")
        frappe.msgprint(f"Formatted Excel file created: <a href='{saved_file.file_url}' target='_blank'>{file_name}</a>")
        
        return saved_file.file_url
        
    except Exception as e:
        frappe.logger().error(f"Error in formatted export: {str(e)}")
        frappe.throw(f"Export failed: {str(e)}")


@frappe.whitelist()
def get_available_templates():
    """Get list of available Excel templates from File doctype"""
    templates = frappe.get_all(
        "File",
        filters={
            "file_name": ["like", "%.xlsx"],
            "attached_to_doctype": ["in", ["", "Timesheet"]]
        },
        fields=["file_name", "file_url"],
        order_by="creation desc"
    )
    return templates


@frappe.whitelist()
def export_timesheet_to_google_sheets(timesheet_name, spreadsheet_id=None, sheet_name="Timesheet Data"):
    """
    Export timesheet data to Google Sheets.
    If spreadsheet_id is provided, it will use that spreadsheet.
    Otherwise, it will create a new spreadsheet.
    """
    try:
        import os
        import gspread
        from google.oauth2.service_account import Credentials
        from google.auth.exceptions import GoogleAuthError
        
        frappe.logger().info(f"Starting Google Sheets export for timesheet: {timesheet_name}, spreadsheet_id: {spreadsheet_id}")
        
        timesheet = frappe.get_doc("Timesheet", timesheet_name)
        
        # Google Sheets API setup
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # Get service account credentials
        service_account_file = frappe.get_site_path("private", "service-account-key.json")
        
        if not os.path.exists(service_account_file):
            frappe.throw("Google Sheets service account key not found. Please upload the service-account-key.json file to the private folder.")
        
        credentials = Credentials.from_service_account_file(service_account_file, scopes=scope)
        gc = gspread.authorize(credentials)
        
        # Check if we have a valid spreadsheet_id
        if spreadsheet_id and spreadsheet_id.strip():
            # Use existing spreadsheet
            try:
                frappe.logger().info(f"Attempting to open existing spreadsheet with ID: {spreadsheet_id}")
                spreadsheet = gc.open_by_key(spreadsheet_id)
                frappe.logger().info(f"Successfully opened existing spreadsheet: {spreadsheet.title}")
            except Exception as e:
                frappe.throw(f"Could not access spreadsheet with ID: {spreadsheet_id}. Error: {str(e)}")
        else:
            # Try to create new spreadsheet, but handle quota error gracefully
            try:
                spreadsheet_title = f"Timesheet Export - {timesheet.name} - {frappe.utils.nowdate()}"
                frappe.logger().info(f"Attempting to create new spreadsheet: {spreadsheet_title}")
                spreadsheet = gc.create(spreadsheet_title)
                frappe.logger().info(f"Created new spreadsheet: {spreadsheet.title}")
            except Exception as e:
                if "quota" in str(e).lower() or "storage" in str(e).lower():
                    frappe.throw(
                        "Google Drive storage quota exceeded. Please provide an existing spreadsheet ID or free up some Google Drive space."
                    )
                else:
                    raise e
        
        # Get or create worksheet
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
        except:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
        
        # Prepare data
        headers = ["Timesheet ID", "Employee", "Project", "Activity Type", "Date", "From Time", "To Time", "Hours", "Description", "Billable"]
        
        data = [headers]
        
        for entry in timesheet.time_logs:
            row = [
                timesheet.name,
                timesheet.employee,
                entry.project or "",
                entry.activity_type or "",
                entry.from_time.strftime("%Y-%m-%d") if entry.from_time else "",
                entry.from_time.strftime("%H:%M") if entry.from_time else "",
                entry.to_time.strftime("%H:%M") if entry.to_time else "",
                entry.hours or 0,
                entry.description or "",
                "Yes" if entry.is_billable else "No"
            ]
            data.append(row)
        
        # Clear existing data and update
        worksheet.clear()
        worksheet.update('A1', data)
        
        # Format headers
        worksheet.format('A1:J1', {
            'backgroundColor': {'red': 0.8, 'green': 0.8, 'blue': 0.8},
            'textFormat': {'bold': True},
            'horizontalAlignment': 'CENTER'
        })
        
        # Auto-resize columns
        worksheet.columns_auto_resize(0, len(headers))
        
        # Add timesheet info at the top
        info_data = [
            ["TIMESHEET EXPORT"],
            [""],
            ["Timesheet ID:", timesheet.name],
            ["Employee:", timesheet.employee],
            ["Period:", f"{timesheet.start_date} to {timesheet.end_date}"],
            ["Total Hours:", timesheet.total_hours],
            ["Status:", timesheet.status],
            [""]
        ]
        
        # Insert info at the top
        worksheet.insert_rows(info_data, 1)
        
        # Format the title
        worksheet.format('A1:A1', {
            'textFormat': {'bold': True, 'fontSize': 16},
            'horizontalAlignment': 'CENTER'
        })
        
        # Get the spreadsheet URL
        spreadsheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}"
        
        frappe.logger().info(f"Google Sheets export completed. URL: {spreadsheet_url}")
        
        # Show success message with link
        frappe.msgprint(
            f"Timesheet exported to Google Sheets successfully! "
            f"<a href='{spreadsheet_url}' target='_blank'>Open Spreadsheet</a>"
        )
        
        return {
            'spreadsheet_id': spreadsheet.id,
            'spreadsheet_url': spreadsheet_url,
            'spreadsheet_title': spreadsheet.title
        }
        
    except ImportError:
        frappe.throw("Google Sheets libraries not installed. Please install: pip install gspread google-auth")
    except GoogleAuthError as e:
        frappe.throw(f"Google Sheets authentication failed: {str(e)}")
    except Exception as e:
        frappe.logger().error(f"Error in Google Sheets export: {str(e)}")
        frappe.throw(f"Google Sheets export failed: {str(e)}")


@frappe.whitelist()
def get_google_sheets_config():
    """Get Google Sheets configuration status"""
    try:
        import os
        service_account_file = frappe.get_site_path("private", "service-account-key.json")
        service_account_exists = os.path.exists(service_account_file)
        
        # Check if gspread is installed
        try:
            import gspread
            gspread_installed = True
        except ImportError:
            gspread_installed = False
        
        # Check if google-auth is installed
        try:
            from google.oauth2.service_account import Credentials
            google_auth_installed = True
        except ImportError:
            google_auth_installed = False
        
        return {
            'service_account_exists': service_account_exists,
            'gspread_installed': gspread_installed,
            'google_auth_installed': google_auth_installed,
            'service_account_path': service_account_file
        }
    except Exception as e:
        frappe.logger().error(f"Error checking Google Sheets config: {str(e)}")
        return {
            'service_account_exists': False,
            'gspread_installed': False,
            'google_auth_installed': False,
            'error': str(e)
	}
