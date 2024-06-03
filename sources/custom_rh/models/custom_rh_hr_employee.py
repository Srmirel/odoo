from odoo import models, fields


class CustomRhHrEmployee(models.Model):
    _inherit = "hr.employee"
    check_in_hour = fields.Datetime(
        string="Hora de Entrada", default=False, required=True, tracking=True)
