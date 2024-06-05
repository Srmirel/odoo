import datetime

from odoo import fields,models

from odoo import http
from odoo.http import request
from odoo.addons.resource.models.utils import Intervals

import pytz
import base64
import io
import xlsxwriter


def utc_to_local(self,utc_dt):
    if (utc_dt == False):
        return False
    timezone = self._context.get('tz') or self.env.user.partner_id.tz or 'UTC'
    self_tz = self.with_context(tz=timezone)
    # .normalize might be unnecessary
    return fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(utc_dt)).replace(tzinfo=None)
def _get_employee_calendar(self):
    self.ensure_one()
    return self.employee_id.resource_calendar_id or self.employee_id.company_id.resource_calendar_id


def calculate_worked_hours(self,record, bandera=False):
    timezone = self._context.get('tz') or self.env.user.partner_id.tz or 'UTC'
    self_tz = self.with_context(tz=timezone)
    # # Calculate the time delta in seconds
    # time_delta = fecha_salida - fecha_entrada
    # Verify if the date is a weekend
    calendar = record._get_employee_calendar()
    resource = record.employee_id.resource_id
    check_in_tz = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(record.check_in))
    if record.check_out == False:
        check_out_tz = check_in_tz
    else:
        check_out_tz = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(record.check_out))
   
   
    lunch_intervals = calendar._attendance_intervals_batch(
                check_in_tz, check_out_tz, resource, lunch=True)
    attendance_intervals = Intervals([(check_in_tz, check_out_tz, record)]) - lunch_intervals[resource.id]
    # delta = sum((i[1] - i[0]).total_seconds() for i in attendance_intervals)
    total_delta_seconds = 0
    for interval in attendance_intervals:
        delta = (interval[1] - interval[0]).total_seconds()
        total_delta_seconds += delta
    return total_delta_seconds


def segundos_a_horas_minutos_segundos(segundos):
    horas = int(segundos / 3600)
    segundos -= horas * 3600
    minutos = int(segundos / 60)
    segundos -= minutos * 60
    return f"{int(horas):02d}:{int(minutos):02d}:{int(segundos):02d}"
def retardo(check_in, check_in_hour):
   
    if (check_in_hour == False) or (check_in == False):
        return 0
    
    #Obtenemos el tiempo de entrada y la hora de entrada
    entrada_trabajador_time = check_in.time()
    hora_entrada_time = check_in_hour.time()
    entrada_trabajador_deltatime = datetime.timedelta(hours=entrada_trabajador_time.hour, minutes=entrada_trabajador_time.minute, seconds=entrada_trabajador_time.second)
    hora_entrada_deltatime = datetime.timedelta(hours=hora_entrada_time.hour, minutes=hora_entrada_time.minute, seconds=hora_entrada_time.second)    
   
    diferencia = entrada_trabajador_deltatime - hora_entrada_deltatime
    minutos_diferencia = int(diferencia.total_seconds() / 60)
    
    if minutos_diferencia > 5:
        return 1
    return 0

def calcular_diferencia_dias(fecha_inicio, fecha_fin):
    """
    Calcula la diferencia de días entre dos fechas dadas.

    Argumentos:
      fecha_inicio (date): La fecha de inicio del rango.
      fecha_fin (date): La fecha final del rango.

    Retorna:
      int: La diferencia de días entre las fechas.
    """

    diferencia_dias = fecha_fin - fecha_inicio
    return diferencia_dias.days

def calculo_dias_no_registrados(fecha_inicio, diferencias_en_dias, registro_fechas, lista_no_laborable):
    dias_no_registrados = []
    fecha_actual = fecha_inicio.date()
    
    
    for i in range(diferencias_en_dias):
        
        if fecha_actual not in [fecha_fila[0] for fecha_fila in registro_fechas]:
            if fecha_actual.weekday() not in lista_no_laborable:
                dias_no_registrados.append(fecha_actual)
        fecha_actual += datetime.timedelta(days=1)
    return dias_no_registrados

def employee_not_here(not_employee, employee):
    
    not_employee_list = []
    employee_list = []
    absent_employee = []
    
    
    
    for item in not_employee:
        not_employee_list.append([item.id,item.name])
    for item in employee:
        employee_list.append([item.employee_id.id,item.employee_id.name])
    # print("Not Employee List:",not_employee_list)
    # print("Employee List:",employee_list)
    for item in not_employee_list:
        if item not in employee_list:
            absent_employee.append(item)       
    return absent_employee


    

def manejo_informacion(self):
    timezone = self._context.get('tz') or self.env.user.partner_id.tz or 'UTC'
    self_tz = self.with_context(tz=timezone)
  

    hora_inicio = datetime.datetime.combine(self.start_date, datetime.time(hour=0, minute=0, second=0))
    hora_final = datetime.datetime.combine(self.end_date, datetime.time(hour=23, minute=59, second=59))

    department_id = self.department_id
    id_department = department_id.id
    
    delta_horas = datetime.timedelta(hours=6)
    final_ahora_si_entrada = hora_inicio + delta_horas
    final_ahora_si_salida = hora_final + delta_horas
    # date_start = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(hora_inicio))
    # date_end = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(hora_final))
    diferencia_en_dias = (calcular_diferencia_dias(hora_inicio, hora_final) + 1)
   
   
   
    if department_id:
        empleados = request.env["hr.employee"].search([('department_id','=',id_department)])
    else:
        empleados = request.env["hr.employee"].search([])
    
    
    

    
    # Realizo la consulta de datos
    fake_data = request.env["hr.attendance"].search(
        [
            '|',
                # '&',
                '&',
                # '|',
                # 139, 143, 187
                    # ('employee_id', 'in', [
                        # 674
                        # 187,
                    #                     #    139,
                                        #    143, 
                    #                     #    187
                                        #    ]),  # Condición A
                    ('check_in', '>=', final_ahora_si_entrada),  # Condición B
                    ('check_out', '<=', final_ahora_si_salida),  # Condición C
                    # ('check_out', '=', False)
                # '&',
                '&',
                '&',
                    # ('employee_id', 'in', [
                        # 674
                #         # 187,
                #     #     # 139,
                        # 143,
                #     #     # 187   
                        # ]),  # Condición D
                    ('check_in', '>=', final_ahora_si_entrada),  # Condición E
                    ('check_in', '<=', final_ahora_si_salida),  # Condición F
                    ('check_out', '=', False),  # Condición F
        ],
        # limit=100,
        order="employee_id,check_in")
    data = []
    for item in fake_data:
        if department_id:
            if item.employee_id.department_id.id == department_id.id:
                data.append(item)
        else:
            data.append(item)   
            
    # print(f"Soy Data:{data}") 
    
    compare_employee = employee_not_here(empleados,data)
    
   
    
    
    lista = []

    acumulador = 0
    acumuladorTotal = 0
    dia_anterior_entrada = None
    dia_anterior_salida = None
    id_anterior = None
    id_anterior_name = None
    auxiliar_contable = 0
    auxiliar_bandera = True
    contador_retardos = 0
    
    registro_fechas_asistidas = []
    dias_no_registrados = []
    dias_asistencia = []
    
    for index, record in enumerate(data):
       
        
        
        
        for item in record.employee_id.department_id.workdays_id:
            dias_asistencia.append(item.tag)
        # print("Dias de Asistencia:",dias_asistencia)

        # En caso de que una fecha este vacia
        if ((record.check_in == False) or (record.check_out == False)) and index != 0:
            fecha_entrada = utc_to_local(self,record.check_in)
            fecha_salida = utc_to_local(self,record.check_in)    
        else:
            fecha_entrada = utc_to_local(self,record.check_in)
            fecha_salida = utc_to_local(self,record.check_out)    
        # Ejecucion el primer registro
        if index == 0:
            if record.check_out == False:
                fecha_salida = utc_to_local(self,record.check_in)
            else: 
                fecha_salida = utc_to_local(self,record.check_out)         
            fecha_entrada = utc_to_local(self,record.check_in)
            dia_anterior_entrada = fecha_entrada.day
            dia_anterior_salida = fecha_salida.day
            id_anterior = record.employee_id.id
            id_anterior_name = record.employee_id.name
       


        # En caso de que cambie un dia
        if fecha_entrada.day != dia_anterior_entrada and fecha_salida.day != dia_anterior_salida:
            auxiliar_bandera = True
            dia_anterior_entrada = fecha_entrada.day
            dia_anterior_salida = fecha_salida.day
            lista.append(
                ["horas_totales_unitarias", segundos_a_horas_minutos_segundos(acumulador), "hola", "hola", "hola"])
            auxiliar_contable +=1
            acumulador = 0
        if id_anterior != record.employee_id.id:
            # En este apartado debo calcular las horas totales
           
            
            dias_no_registrados = calculo_dias_no_registrados(hora_inicio, diferencia_en_dias, registro_fechas_asistidas,dias_asistencia)
            lista.append(["titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas"])
            for item in dias_no_registrados:
                lista.append(["falta",id_anterior_name, item,len(dias_no_registrados), "Falta"])
            lista.append(["horas_totales_empleado", segundos_a_horas_minutos_segundos(acumuladorTotal), id_anterior_name, len(dias_no_registrados),
                          contador_retardos])
            
            id_anterior = record.employee_id.id
            id_anterior_name = record.employee_id.name
            auxiliar_contable = 0
            acumulador = 0
            acumuladorTotal = 0
            contador_retardos = 0
            auxiliar_bandera = True 
            registro_fechas_asistidas = []
            dias_asistencia = []
        if index == len(data) - 1 and fecha_entrada == fecha_salida:
            horas_calculadas = calculate_worked_hours(self,record,False)
        else:    
            horas_calculadas = calculate_worked_hours(self,record,auxiliar_bandera)
        horas_trabajadas = segundos_a_horas_minutos_segundos(horas_calculadas)
        acumulador +=   horas_calculadas
        acumuladorTotal += horas_calculadas
        retardo_unitorio = retardo(record.check_in, record.employee_id.check_in_hour)
        if auxiliar_bandera:
            
            if record.check_out == False:
                registro_fechas_asistidas.append([fecha_entrada.date()])
                lista.append([record.employee_id.name, fecha_entrada, "False", horas_trabajadas, retardo_unitorio])
            else:        
                registro_fechas_asistidas.append([fecha_entrada.date()])
                lista.append([record.employee_id.name, fecha_entrada, fecha_salida, horas_trabajadas, retardo_unitorio])
            auxiliar_bandera = False
            if retardo_unitorio == 1:
                contador_retardos += 1
        else:
            registro_fechas_asistidas.append([fecha_entrada.date()])
            lista.append([record.employee_id.name, fecha_entrada, fecha_salida, horas_trabajadas, "N/A"])
        # En caso de que sea el ultimo registro
        if index == len(data) - 1:
            if auxiliar_contable != 0:
                lista.append(
                    ["horas_totales_unitarias", segundos_a_horas_minutos_segundos(acumulador), "auxiliar_contable", "auxiliar_contable", "auxiliar_contable"])
            
            # Aqui debo de calcular los dias que no asistio
            dias_no_registrados = calculo_dias_no_registrados(hora_inicio, diferencia_en_dias, registro_fechas_asistidas,dias_asistencia)
            lista.append(["titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas"])
            for item in dias_no_registrados:
                lista.append(["falta",id_anterior_name, item,len(dias_no_registrados), "Falta"])
                
            lista.append(
                ["horas_totales_empleado", segundos_a_horas_minutos_segundos(acumuladorTotal), record.employee_id.name,len(dias_no_registrados),
                 contador_retardos])
            acumulador = 0
            acumuladorTotal = 0
            contador_retardos = 0
            registro_fechas_asistidas = []
            dias_asistencia = []
    lista.append(["empleados_faltantes",compare_employee,"hola","hola","hola"])        
    return lista
def manejo_informacionGeneral(self,data):
    lista = []
    for index, record in enumerate(data):
        if record[0] == "horas_totales_empleado":
            lista.append([record[2],record[1],record[4],record[3],'pruebas'])
            
             
    return lista
def manejo_informacion_employee_not(self,data):
    lista = []
    for index, record in enumerate(data):
        if record[0] == "empleados_faltantes":
            for item in record[1]:
                lista.append([item[1],"N/A","N/A","N/A","N/A"])    
             
    return lista
def creacion_csv(self,data):

    # output = io.BytesIO()
    # workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    file_path = 'timesheet report' + '.xlsx'
    workbook = xlsxwriter.Workbook('/tmp/' + file_path)
    worksheet = workbook.add_worksheet('Reporte_Unitario')
    
    
    # Define formats
    
    formato_fechas_hoja_1 = workbook.add_format(
        {
        'border': 1,
        'num_format': 'yyyy-mm-dd hh:mm',
        "fg_color": "#E5F1DF",
        
        })
    formato_fechas_hoja_1.set_align('center')
    
    formato_fechas_hoja_1_faltas = workbook.add_format(
        {
        'border': 1,
        'num_format': 'yyyy-mm-dd',
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#A5D5E3",
        
        })
    
    centrado_regular = workbook.add_format()
    
    centrado_regular.set_align('center')
    # Nuevos Formatos 
    formato_registro_basico_trabajador = workbook.add_format(
        {
            # "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#e5f1df",
        }
    )
    
    pares_hoja2 = workbook.add_format(
        {
            # "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#A5D5E3",
        }
    )
    impares_hoja2 = workbook.add_format(
        {
            # "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#E5F1DF",
        }
    )
    
    formato_encabezado_hoja_1 = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#6eae4b",
            "font_color": "#ffffff",
            
        }
    )
    formato_encabezado_hoja_2 = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#FFAF61",
            "font_color": "#000000",
            
        }
    )
    formato_encabezado_hoja_1_faltas = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#82B8F3",
            "font_color": "#000000",
            
        }
    )
    horas_unitarias = workbook.add_format(
        {
            # "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#FFDB5C",
        }
    )
    horas_totales = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#C3FF93",
        }
    )
    retardos = workbook.add_format(
        {
            "bold": 1,
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#FFAF61",
        }
    )
    
    
    # Define headers
    worksheet.write(0, 0, "Empleado",formato_encabezado_hoja_1)
    worksheet.write(0, 1, "Entrada",formato_encabezado_hoja_1)
    worksheet.write(0, 2, "Salida",formato_encabezado_hoja_1)
    worksheet.write(0, 3, "Horas Trabajadas",formato_encabezado_hoja_1)
    worksheet.write(0, 4, "Retardo",formato_encabezado_hoja_1)
    row = 1  # Starting from row 2 (1-based indexing)
    
    for record in data:
        if record[0] == "empleados_faltantes":
            continue
        if record[0] == "horas_totales_unitarias":
            worksheet.write(row, 0, "Horas del Dia", horas_unitarias)
            worksheet.merge_range(row, 1, row, 4, record[1], horas_unitarias)
            row += 1
            continue
        if record[0] == "horas_totales_empleado":
            worksheet.write(row, 0, "Horas Totales", horas_totales)
            worksheet.merge_range(row, 1, row, 2, record[1], horas_totales)
            worksheet.write(row, 3, "Retardos", retardos)
            worksheet.write(row, 4, record[4], retardos)
            row += 1
            continue
        if record[0] == "titulofaltas":
            # worksheet.write(row, 0, "Faltas", formato_encabezado_hoja_1)
            worksheet.merge_range(row, 0, row, 4, "Faltas", formato_encabezado_hoja_1_faltas)
            row += 1
            continue
        if record[0] == "falta":
            worksheet.write(row, 0, record[1], formato_registro_basico_trabajador)
            worksheet.merge_range(row,1,row,4,record[2],formato_fechas_hoja_1_faltas)
            row += 1
            continue
            
        worksheet.write(row, 0, record[0], formato_registro_basico_trabajador)
        worksheet.write(row, 1, record[1], formato_fechas_hoja_1)
        worksheet.write(row, 2, record[2], formato_fechas_hoja_1)
        
        worksheet.write(row, 3, record[3], formato_registro_basico_trabajador)
        worksheet.write(row, 4, record[4], formato_registro_basico_trabajador)
        row += 1
    # Manually adjust the width of each column
    worksheet.set_column('A:A', 30)  # Adjust width for column A (Employee)
    worksheet.set_column('B:B', 20)  # Adjust width for column B (Entrada)
    worksheet.set_column('C:C', 20)  # Adjust width for column C (Salida)
    # Adjust width for column D (Horas Trabajadas)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 10)  # Adjust width for column E (Color)
    # Adjust widths based on your specific requirements
    
    # Construccion Segunda Hoja
    worksheet2 = workbook.add_worksheet('Reporte_General')
    # Style 
    title_format = workbook.add_format(
                {'border': 1, 'bold': True, 'valign': 'vcenter', 'align': 'center', 'font_size': 11, 'bg_color': '#FFAF61','font_color': '#000000'})

    formato_fechas_hoja_2 = workbook.add_format(
        {'border': 1, 'bold': True, 'valign': 'vcenter', 'align': 'center', 'font_size': 11, 'bg_color': '#FFAF61','font_color':"#000000",'num_format': 'yyyy-mm-dd'})
    formato_fechas_hoja_2.set_align('center')
    
    datos_generales = manejo_informacionGeneral(self,data)
    not_employee = manejo_informacion_employee_not(self,data)
    
    row = 7
    row2 = 7
    
    # Titulo
    # worksheet.merge_range(row, 1, row, 4, record[1], horas_unitarias)
    worksheet2.merge_range(1, 0,2,3, "Reporte de Asistencia", title_format)
    # TODO Integrar en el titulo si se selecciona un departamento para que lo integre
    worksheet2.merge_range(4,0,4,3,f"Periodo: {self.start_date} - {self.end_date} / Departamento: {self.department_id.complete_name if self.department_id else 'N/A'}",formato_fechas_hoja_2)
    
    # worksheet2.write(3,0,"Fecha Inicial")
    # worksheet2.write(3,1,self.start_date,formato_fechas_hoja_2)
    # worksheet2.write(3,2,"Fecha Inicial")
    # worksheet2.write(3,3,self.end_date,formato_fechas_hoja_2)

     # Define headers
    worksheet2.write(6, 0, "Empleado",formato_encabezado_hoja_2)
    worksheet2.write(6, 1, "Total Horas Trabajadas",formato_encabezado_hoja_2)
    worksheet2.write(6, 2, "Total de Retardos",formato_encabezado_hoja_2)
    worksheet2.write(6, 3, "Total de Faltas",formato_encabezado_hoja_2)
    
    # worksheet2.write(4, 4, "Proximamente",formato_encabezado_hoja_2)

    for index,item in enumerate(datos_generales):
        if index % 2 == 0:
            worksheet2.write(row, 0, item[0],pares_hoja2)
            worksheet2.write(row, 1, item[1],pares_hoja2)
            worksheet2.write(row, 2, item[2],pares_hoja2)
            worksheet2.write(row, 3, item[3],pares_hoja2)
            row += 1
        else:    
            worksheet2.write(row, 0, item[0], impares_hoja2)
            worksheet2.write(row, 1, item[1], impares_hoja2)
            worksheet2.write(row, 2, item[2], impares_hoja2)
            worksheet2.write(row, 3, item[3], impares_hoja2)
            # worksheet2.write(row, 4, item[4], impares_hoja2)
            row += 1
            
    worksheet2.merge_range(6, 5,6,8, "Empleados Fuera de Rango",formato_encabezado_hoja_2)
    for index,item in enumerate(not_employee):
        if index % 2 == 0:
            worksheet2.merge_range(row2, 5,row2,8, item[0],pares_hoja2)
            row2 += 1
        else:    
            worksheet2.merge_range(row2, 5,row2,8, item[0],impares_hoja2)
            # worksheet2.write(row2, 4, item[4], impares_hoja2)
            row2 += 1        
    # Manually adjust the width of each column
    worksheet2.set_column('A:A', 35)  # Adjust width for column A (Employee)
    worksheet2.set_column('B:B', 20)  # Adjust width for column B (Entrada)
    worksheet2.set_column('C:C', 20)  # Adjust width for column C (Salida)
    # Adjust width for column D (Horas Trabajadas)
    worksheet2.set_column('D:D', 15)
    # worksheet2.set_column('E:E', 10)  # Adjust width for column E (Color)

    workbook.close()
    ex_report = base64.b64encode(open('/tmp/' + file_path, 'rb+').read())

    excel_report_id = self.env['save.ex.report.special.wizard'].create({"document_frame": file_path,
                                                                "file_name": ex_report})
#https://www.colorhunt.co/palette/d9edbfffb996ffcf81fdffab
    return {
        'res_id': excel_report_id.id,
        'name': 'Descarga Archivo',
        'view_type': 'form',
        "view_mode": 'form',
        'view_id': False,
        'res_model': 'save.ex.report.special.wizard',
        'type': 'ir.actions.act_window',
        'target': 'new',
    }
    # response = request.make_response(output.getvalue(), headers=[
    #     ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
    #     ('Content-Disposition', 'attachment; filename=attendance.csv')
    # ])
    # return response


class ReportSpecialWizard(models.TransientModel):
    _name = "report.special.wizard"
    _description = "Reporte Especial"

    end_date = fields.Date(default=lambda self: fields.datetime.now())
    start_date = fields.Date(default=lambda self: fields.datetime.now() - datetime.timedelta(days=15))
    
    department_id = fields.Many2one("hr.department")
    
    
    def action_generate_csv(self):
       
        data = manejo_informacion(self)

        return creacion_csv(self,data)