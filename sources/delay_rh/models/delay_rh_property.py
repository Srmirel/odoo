from odoo import fields, models, api
from odoo.http import request
import datetime


class DelayRhProperty(models.Model):
    _name = 'delay.rh.property'
    _description = 'Delay RH Property'

    # --------------------------------------- Fields Declaration ----------------------------------
    # Tipo de Busqueda
    type_search = fields.Selection(selection=[
        ("all", "Todos"),
        ("dep", "Departamento"),
        ("some", "Especifico"),
    ], string="Tipo de Busqueda",required=True)
    # Inicio de la Fecha
    start_date = fields.Date(
        default=lambda self: fields.datetime.now() - datetime.timedelta(days=15))
    # Fin de la Fecha
    end_date = fields.Date(default=lambda self: fields.datetime.now())
    workdays_id = fields.Many2many(
        'custom.weekdays', string='Dias No Laborables')

    # Days
    monday = fields.Float(string='Lunes', default=8.0)
    tuesday = fields.Float(string='Martes', default=8.0)
    wednesday = fields.Float(string='Miercoles', default=8.0)
    thursday = fields.Float(string='Jueves', default=8.0)
    friday = fields.Float(string='Viernes', default=8.0)
    saturday = fields.Float(string='Sabado', default=8.0)
    sunday = fields.Float(string='Domingo', default=8.0)
    monday_bool = fields.Boolean(default=True)
    tuesday_bool = fields.Boolean(default=True)
    wednesday_bool = fields.Boolean(default=True)
    thursday_bool = fields.Boolean(default=True)
    friday_bool = fields.Boolean(default=True)
    saturday_bool = fields.Boolean(default=True)
    sunday_bool = fields.Boolean(default=True)

    employee_id = fields.Many2many(
        'hr.employee', string="Employee", compute="_compute_employee")
    department_id = fields.Many2many(
        'hr.department', string="Department", compute="_compute_department")

    # Field Test
    field_employee = fields.Char()
    field_department = fields.Char()
    
    
    # Default
    is_default = fields.Boolean(default=False)
    
    
    # ----------------------------------------- Relational ----------------------------------------
    # Dias No trabajados
    workdays_id = fields.Many2many(
        'custom.weekdays', string='Dias No Laborables')
    # ---------------------------------------- Compute methods ------------------------------------
    # Algunos Muestra y Recuperacion de Empleados

    def _compute_employee(self):
        # for record in self:
        # record.employee_id = self.env['hr.employee']
        if (self.field_employee):
            for record in self:
                # unvideomas = self.field_employee.strip("'").strip("[]")
                record.employee_id = self.env['hr.employee'].search(
                    [('id', 'in', eval(self.field_employee))])
        else:
            for record in self:
                record.employee_id = self.env['hr.employee']
        # self.employee_id = self.env['hr.employee']

    @api.onchange('employee_id')
    def _cambioempleo(self):
        self.field_employee = self.employee_id.ids
        self.field_department = ""
    # Dep Muestra departamentos
    
    def _compute_department(self):
        if (self.field_department):
            for record in self:
                record.department_id = self.env['hr.department'].search(
                    [('id', 'in', eval(self.field_department))])
        else:
            for record in self:
                record.department_id = self.env['hr.department']

    @api.onchange('department_id')
    def _cambiodepartment(self):
        self.field_department = self.department_id.ids
        self.field_employee = self.env['hr.employee'].search(
            [('department_id', 'in', eval(self.field_department))]).ids
    @api.onchange("type_search")
    def _compute_type_search(self):
        if self.type_search == "all":
            self.field_employee = self.env['hr.employee'].search([]).ids
            self.field_department = ""
            

    # ------------------------------------------ CRUD Methods -------------------------------------
    # @api.ondelete(at_uninstall=False)
    # def unlink(self):
    #     if not set(self.mapped("state")) <= {"new", "canceled"}:
    #         raise UserError("Only new and canceled properties can be deleted.")
    #     return super().unlink()

    def create(self, vals):
        return super().create(vals)

    def write(self, vals):
        return super().write(vals)

    # ----------------------------------- Constrains and Onchanges --------------------------------
    @api.onchange("workdays_id")
    def _onchange_days(self):
        workdays = self.workdays_id.ids
        self.monday_bool = 1 in workdays
        self.tuesday_bool = 2 in workdays
        self.wednesday_bool = 3 in workdays
        self.thursday_bool = 4 in workdays
        self.friday_bool = 5 in workdays
        self.saturday_bool = 6 in workdays
        self.sunday_bool = 7 in workdays
# ----------------------------------- Actions Server Runs --------------------------------
    def hello_world(self):
         # TODO Crear una funcion que me permita obtener todos los registros del dia actual que no tengan fecha asignada de salida.

         #Se importan los paquetes
         # import datetime

        
         # Se obtiene las asistencias que tengan false en el dia de hoy
        delta_horas = datetime.timedelta(hours=6)
        pepe = datetime.datetime.now() - delta_horas
        hora_inicio = datetime.datetime.combine(pepe, datetime.time(hour=0, minute=0, second=0))
        hora_final = datetime.datetime.combine(pepe, datetime.time(hour=23, minute=59, second=59))

        final_ahora_si_entrada = hora_inicio + delta_horas
        final_ahora_si_salida = hora_final + delta_horas



        fake_data = request.env["hr.attendance"].search(
        [
            '&',
            '&',
                    ('check_in', '>=', final_ahora_si_entrada),  # Condición E
                    ('check_in', '<=', final_ahora_si_salida),  # Condición F
                    ('check_out', '=', False),  # Condición F
        ],
        # limit=100,
        )
        
        
        fake_hour = datetime.datetime.combine(pepe, datetime.time(hour=22, minute=10, second=58))
        fake_hour = fake_hour + delta_horas
        for fake in fake_data:
            fake.write({'check_out': fake_hour})
        return fake_data 