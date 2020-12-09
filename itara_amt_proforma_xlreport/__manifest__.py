# -*- coding: utf-8 -*-
{
    'name': 'Proforma Report Excel',
    'version': '1.0.0',
    'category': 'Sale',
    'summary': '''
        Proforma Excel Report.
        ''',
    'author': 'Itara IT solutions Private Limited ...',
    'license': "OPL-1",
    'depends': [
        'sale','itara_amt_third_port_shipment'
    ],
    'data': [
        'report/report_view.xml',
        'views/report_wizard.xml'
    ],
    'demo': [],  
    # 'images': ['static/description/banner.png'],
    'auto_install': False,
    'installable': True,
    'application': True
}
