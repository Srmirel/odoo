{
    'name': "Custom RH",
    'version': '1.0',
    'depends': ['base', 'web'],
    'author': "Alvaro Cruz",
    'category': 'Rh Especial/custom_rh',
    'description': """
    Description text
    """,
    "license": "LGPL-3",

    'depends': ['hr'],

    'data': [
        'security/ir.model.access.csv',
        'wizard/report_special_view.xml',
        'wizard/save_ex_report_special_wizard_view.xml',
        'data/custom_rh_custom_weekdays.xml',
        'views/custom_rh_hr_employee_views.xml',
        'views/custom_rh_hr_attendance_view.xml',
        'views/custom_rh_hr_department_view.xml',
    ],

    "application": True,
}
