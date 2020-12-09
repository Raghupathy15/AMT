from odoo import api, fields, models, _
from odoo.addons import decimal_precision as dp
from odoo.exceptions import UserError, Warning

class WizardProforma(models.Model):
	_inherit = "sale.order"


	@api.multi
	def generate_xls_report(self):
		# for rec in self:
		return self.env['report'].get_action(self, report_name="itara_amt_proforma_xlreport.proformareport.xlsx")
