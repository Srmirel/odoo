from odoo import models, fields, api


class CustomWeekdays(models.Model):
    _name = 'custom.weekdays'

    name = fields.Char(string="Weekday Name")
    tag = fields.Integer()
