# Copyright 2019, Jarsa Sistemas, S.A. de C.V.
# License LGPL-3.0 or later (http://www.gnu.org/licenses/lgpl.html).

{
    'name': 'Google Spreadsheets Import',
    'summary': 'Crate and Import from spreadsheets',
    'version': '15.0.1.0.0',
    'category': 'Tools',
    'author': 'Jarsa Sistemas',
    'website': 'https://www.jarsa.com.mx',
    'license': 'LGPL-3',
    'depends': [
        'base_import',
        'google_drive'
    ],
    'data': [
        'security/security.xml',
        'views/google_spreadsheet_view.xml',
        'views/google_spreadsheet_file_view.xml',
        'data/ir_config_parameter.xml',
        'data/ir_cron_data.xml',
        'security/ir.model.access.csv',
    ],
    'qweb': [
        'static/src/xml/base_import.xml'
    ],
    'installable': True,
    'external_dependencies': {
        'python': [
            'google-api-python-client',
            'ipdb',
        ],
    },
}
