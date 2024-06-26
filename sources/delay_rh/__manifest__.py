{
    'name': "Delay RH",
    'version': '1.0',
    'depends': ['base', 'web'],
    'author': "Alvaro Cruz",
    'category': 'Rh Especial/delay_rh',
    'description': """
    Description text
    """,
    "license": "LGPL-3",
    'depends': ['hr'],
    'data': [
        'security/ir.model.access.csv',
        'views/delay_rh_property.xml',
        'views/delay_rh_menu.xml',
        ],

    "application": True,
}
