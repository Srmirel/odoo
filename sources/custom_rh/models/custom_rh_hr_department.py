from odoo import models, fields, api



class CustomRhHrDepartment(models.Model):
    _inherit = "hr.department"
    workdays_id = fields.Many2many('custom.weekdays', string='Dias No Laborables')




    
