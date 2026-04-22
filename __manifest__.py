{
    'name': 'Rapport Cash In/Out PdV',
    'version': '18.0.1.0.0',
    'category': 'Point of Sale/Reporting',
    'summary': 'Export Excel des mouvements caisse (Cash In/Out) par Point de Vente.',
    'author': 'SOPROMER',
    'license': 'LGPL-3',
    'depends': ['point_of_sale', 'account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/pos_cashinout_report_wizard_views.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'installable': True,
    'application': False,
}
