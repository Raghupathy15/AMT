from datetime import datetime
from odoo.addons.report_xlsx.report.report_xlsx import ReportXlsx
import xlsxwriter
from num2words import num2words

class ProformaReport(ReportXlsx):
	def generate_xlsx_report(self, workbook, data, obj):
		active_id = self.env.context.get('active_id')
		rec = self.env['sale.order'].browse(int(active_id))
		# Formats
		heading_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': False, 'size': 15})
		sub_heading_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'bold': True, 'size': 10})
		sub_heading_left_normal = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'bold': False, 'size': 10})
		sub_heading_bold = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'size': 10})
		sub_heading_normal = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'bold': False, 'size': 10})
		normal_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': False, 'size': 10})
		normal_bold = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'size': 10})
		normal_bold_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'bold': True, 'size': 10})
		normal_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'bold': False, 'size': 10})
		normal_top_center = workbook.add_format({'align': 'center', 'valign': 'top', 'bold': False, 'size': 10})
		normal_top_left = workbook.add_format({'align': 'left', 'valign': 'top', 'bold': False, 'size': 10})
		# Adding sheet to work book
		worksheet = workbook.add_worksheet('proformareport.xlsx')
		# Setting width for column
		worksheet.set_column('A:A', 05)
		worksheet.set_column('B:B', 30)
		worksheet.set_column('D:D', 15)
		worksheet.set_column('E:E', 20)
		worksheet.set_column('F:F', 16)

		worksheet.merge_range('A1:F1', ' ', heading_format)
		worksheet.merge_range('A1:F4', ' ', heading_format)
		worksheet.merge_range('A5:F5', 'Proforma Invoice', heading_format)
		# Customer Address details
		worksheet.merge_range('A6:C6', 'BUYER AND CONSIGNEE', sub_heading_left)
		worksheet.merge_range('A7:C7', rec.partner_id.name, sub_heading_normal)

		if rec.partner_id.street:
			worksheet.merge_range('A8:C8', rec.partner_id.street, sub_heading_normal)
		else:
			worksheet.merge_range('A8:C8',' ', sub_heading_normal)

		if rec.partner_id.street2 and rec.partner_id.city:
			worksheet.merge_range('A9:C9',(rec.partner_id.street2)+',City: '+(rec.partner_id.city), sub_heading_normal)
		else:
			worksheet.merge_range('A9:C9',' ', sub_heading_normal)
		if rec.partner_id.phone and rec.partner_id.zip:
			worksheet.merge_range('A10:C10','Tel: '+''+ (rec.partner_id.phone)+ ', ZIP: ' +(rec.partner_id.zip), sub_heading_normal)
		elif rec.partner_id.phone and not rec.partner_id.zip:
			worksheet.merge_range('A10:C10','Tel: '+''+ (rec.partner_id.phone), sub_heading_normal)
		elif rec.partner_id.zip and not rec.partner_id.phone:
			worksheet.merge_range('A10:C10','ZIP: '+''+(rec.partner_id.zip), sub_heading_normal)
		else:
			worksheet.merge_range('A10:C10','',sub_heading_normal)
		if rec.partner_id.fax and rec.partner_id.trn:
			worksheet.merge_range('A11:C11','FAX: '+''+(rec.partner_id.fax)+ ', TIN: ' +(rec.partner_id.trn), sub_heading_normal)
		elif rec.partner_id.fax and not rec.partner_id.trn:
			worksheet.merge_range('A11:C11','FAX: '+''+(rec.partner_id.fax), sub_heading_normal)
		elif rec.partner_id.trn and not rec.partner_id.fax:
			worksheet.merge_range('A11:C11','TIN: '+''+(rec.partner_id.trn), sub_heading_normal)
		else:
			worksheet.merge_range('A11:C11','',sub_heading_normal)	
		import datetime	
		dt = datetime.datetime.strptime(rec.confirmation_date,'%Y-%m-%d %H:%M:%S').strftime('%d/%m/%y')
		worksheet.merge_range('D6:F6', 'PI No:'+' '+str(rec.name) +'                         '+
							  'Date:'+' '+ str(dt), sub_heading_normal)
		if rec.incoterm:
			worksheet.merge_range('D7:F7', 'Incoterms:'+' '+rec.incoterm.name, sub_heading_normal)
		else:
			worksheet.merge_range('D7:F7', 'Incoterms:', sub_heading_normal)
		if rec.pi_shippment:
			worksheet.merge_range('D8:F8', 'EST date of Shipment:'+' '+ rec.pi_shippment, sub_heading_normal)
		else:
			worksheet.merge_range('D8:F8', 'EST date of Shipment:', sub_heading_normal)
		if rec.port_loading_id:
			worksheet.merge_range('D9:F9', 'Port of loading:'+' '+ rec.port_loading_id.name, sub_heading_normal)
		else:
			worksheet.merge_range('D9:F9', 'Port of loading:', sub_heading_normal)
		if rec.port_discharge_id:
			worksheet.merge_range('D10:F10', 'Port of discharge:'+' '+rec.port_discharge_id.name, sub_heading_normal)
		else:
			worksheet.merge_range('D10:F10', 'Port of discharge:', sub_heading_normal)
		if rec.final_destination_id:
			worksheet.merge_range('D11:F11', 'Final Destination:'+' '+rec.final_destination_id.name, sub_heading_normal)
		else:
			worksheet.merge_range('D11:F11', 'Final Destination:', sub_heading_normal)	
		worksheet.merge_range('A12:F12', 'ORIGIN OF GOODS : INDIA', normal_format)
		currency = rec.pricelist_id.currency_id.name
		rows = 12
		column = 0
		worksheet.write(rows, column, 'S.No', sub_heading_bold)
		worksheet.write(rows, column+1, 'Description of Goods', sub_heading_bold)
		worksheet.write(rows, column+2, 'Quantity', sub_heading_bold)
		worksheet.write(rows, column+3, 'UOM', sub_heading_bold)
		worksheet.write(rows, column+4, 'Price'+' '+currency, sub_heading_bold)
		worksheet.write(rows, column+5, 'Amount'+' '+currency, sub_heading_bold)		
		# Line Details
		s_no = 0
		row = 13
		sub_tot = 0
		price = 0
		qty = 0
		for value in rec.order_line:
			if value:
				s_no += 1
				worksheet.write(row, column, s_no, normal_format)
				worksheet.write(row, column+1,value.product_id.name, normal_format)
				worksheet.write(row, column+2,value.product_uom_qty, normal_format)
				worksheet.write(row, column+3,value.product_id.uom_id.name, normal_format)
				worksheet.write(row, column+4,value.price_unit,normal_format)
				worksheet.write(row, column+5,value.price_subtotal,normal_format)
				row += 1
				price += value.price_unit
				qty += value.product_uom_qty
		sub_tot = price
		sub_qty = qty
		row += 1
		worksheet.write(row, column+1,'Total',normal_bold)	
		worksheet.write(row, column+2, sub_qty,normal_bold)
		worksheet.write(row, column+4, sub_tot,normal_bold)
		worksheet.write(row, column+5, rec.amount_untaxed,normal_bold)
		# Number to word convertion
		amt = num2words(rec.amount_untaxed)
		worksheet.write(row+1, column+2, 'Total value'+' '+currency +': ' +(amt),normal_format)
		# Terms & cond..
		worksheet.write(row+2, column+1, 'TERMS & CONDITIONS ', sub_heading_left_normal)
		s_no = 0
		# row = 19
		terms_conditions = rec.terms_conditions_ids
		for terms in terms_conditions:
			if terms:
				s_no += 1
				worksheet.write(row+3,column,s_no, normal_format)
				worksheet.write(row+3,column+1,terms.name, normal_left)
				row += 1				
			else:
				worksheet.merge_range('A21:F37',' ', normal_format)
		# Bank details
		worksheet.merge_range('A33:C33', 'Our Bank Details', normal_bold_left)
		if rec.our_bank_id.name:
			worksheet.merge_range('A34:C34', rec.our_bank_id.name, normal_left)
		else:
			worksheet.merge_range('A34:C34','', normal_left)
		if rec.our_bank_id.l1: 										 
			worksheet.merge_range('A35:C35', rec.our_bank_id.l1, normal_left)
		else:
			worksheet.merge_range('A35:C35', '', normal_left)
		if rec.our_bank_id.l2:
			worksheet.merge_range('A36:C36', rec.our_bank_id.l2, normal_left)
		else:
			worksheet.merge_range('A36:C36','', normal_left)							 
		if rec.our_bank_id.l3:
			worksheet.merge_range('A37:C37', rec.our_bank_id.l3, normal_left)
		else:
			worksheet.merge_range('A37:C37', '', normal_left)
		if rec.our_bank_id.l4:
			worksheet.merge_range('A38:C38', rec.our_bank_id.l4, normal_left)
		else:
			worksheet.merge_range('A38:C38','', normal_left)	
		if rec.our_bank_id.l5:
			worksheet.merge_range('A39:C39', rec.our_bank_id.l5, normal_left)
		else:
			worksheet.merge_range('A39:C39','', normal_left)	 
 
		worksheet.merge_range('D33:F39', '', normal_left)					 
		# Signatures
		worksheet.merge_range('A40:C44', 'FOR ALIA MOHD TRADING CO.(LLC)', normal_top_center)
		worksheet.merge_range('D40:F44', 'BUYERS ACCEPTANCE WITH SEAL & SIGNATURE', normal_top_center)

		worksheet.merge_range('A45:F46', '', normal_bold_left)
		# Image 
		worksheet.merge_range('A47:F52', '', normal_bold_left)
		worksheet.insert_image('A1', '/usr/lib/python2.7/dist-packages/odoo/addons/itara_amt_proforma_xlreport/images/header.PNG', {'x_scale': 0.6, 'y_scale': 0.5})
		worksheet.insert_image('B41', '/usr/lib/python2.7/dist-packages/odoo/addons/itara_amt_proforma_xlreport/images/seal.PNG', {'x_scale': 0.8, 'y_scale': 0.7})
		worksheet.insert_image('B47', '/usr/lib/python2.7/dist-packages/odoo/addons/itara_amt_proforma_xlreport/images/footer1.PNG', {'x_scale': 0.7, 'y_scale': 0.8})
		worksheet.insert_image('F47', '/usr/lib/python2.7/dist-packages/odoo/addons/itara_amt_proforma_xlreport/images/footer2.PNG', {'x_scale': 0.8, 'y_scale': 1.0})

ProformaReport('report.itara_amt_proforma_xlreport.proformareport.xlsx', 'sale.order')