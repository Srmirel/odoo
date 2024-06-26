from odoo.http import request
from odoo import http
from odoo import fields, models

import xlsxwriter
import pytz
import pandas as pd
import datetime
import base64
from odoo.addons.resource.models.utils import Intervals

def parse_hour(hour):
    return '{0:02.0f}:{1:02.0f}'.format(*divmod(hour * 60, 60))


def parse_time(self, utc_dt):
    if (utc_dt == False):
        return False
    timezone = self.env.user.partner_id.tz or 'UTC'

    self_tz = self.with_context(tz=timezone)
    # .normalize might be unnecessary
    return fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(utc_dt)).replace(tzinfo=None)


def special_parse_time(self, utc_dt):
    unvideo = parse_time(self, utc_dt)
    return unvideo.day


def verificar_numero_any(numero, lista_anidada):
    """
    Función que verifica si un número se encuentra en una lista anidada usando any.

    Args:
      numero: El número a buscar.
      lista_anidada: La lista anidada en la que se realiza la búsqueda.

    Returns:
      True si el número se encuentra en la lista, False en caso contrario.
    """
    return any(numero in sublista for sublista in lista_anidada)


def retardo(employee_id, check_in, check_in_hour):
    entrada_trabajador_time = check_in.time()
    entrada_trabajador_deltatime = datetime.timedelta(
        hours=entrada_trabajador_time.hour, minutes=entrada_trabajador_time.minute, seconds=entrada_trabajador_time.second)
    hora_entrada_deltatime = datetime.timedelta(hours=check_in_hour)
    diferencia = entrada_trabajador_deltatime - hora_entrada_deltatime
    minutos_diferencia = int(diferencia.total_seconds() / 60)

    if minutos_diferencia > 6:
        return 1
    return 0


def is_retardo(employee_id, check_in, delay_rh, delay_rh_default):
    papoi = []
    for item in delay_rh_default:
        # Si el empleado se encuentra en la lista de empleados
        if str(employee_id) in item.field_employee:
            weekday = check_in.date().weekday()
            schedule_map = {
                0: item.monday,
                1: item.tuesday,
                2: item.wednesday,
                3: item.thursday,
                4: item.friday,
                5: item.saturday,
                6: item.sunday
            }
            # if check_in.date() <= item.start_date <= check_out.date() and check_in.date() <= item.end_date <= check_out.date():
            # if check_in.date() >= item.start_date and check_in.date() <= item.end_date:
            papoi.append(retardo(employee_id, check_in, schedule_map[weekday]))
    for item in delay_rh:
        # Si el empleado se encuentra en la lista de empleados
        if str(employee_id) in item.field_employee:
            weekday = check_in.date().weekday()
            schedule_map = {
                0: item.monday,
                1: item.tuesday,
                2: item.wednesday,
                3: item.thursday,
                4: item.friday,
                5: item.saturday,
                6: item.sunday
            }
            # if check_in.date() <= item.start_date <= check_out.date() and check_in.date() <= item.end_date <= check_out.date():
            # if check_in.date() >= item.start_date and check_in.date() <= item.end_date:
            if item.start_date <= check_in.date() <= item.end_date:
                papoi.append(retardo(employee_id, check_in, schedule_map[weekday]))
    return papoi[-1] if papoi else 0


def filter_data(fake_data, department_id=None):
    """Filters fake_data based on department_id efficiently.

    Args:
        fake_data: A list of items to filter.
        department_id (int, optional): The department ID to filter by. Defaults to None (no filtering).

    Returns:
        A list of filtered items.
    """
    if department_id:
            # No filtering required, return all items efficiently
        return [item for item in fake_data if item.employee_id.department_id.id == department_id]
    return fake_data  # Efficient copy to avoid modifying original data


def employee_not_here(not_employee_list, employee_list):
    absent_employee_ids = [[emp.employee_id.id, emp.employee_id.name] for emp in employee_list]  # Set of employee IDs
    return [[emp.id, emp.name] for emp in not_employee_list if [emp.id,emp.name] not in absent_employee_ids]

def manejo_informacion_employee_not(self, data):
    # Caso etiquetas
    #   ['empleados_faltantes', [], 'hola', 'hola', 'hola']
    # ['empleados_faltantes', [[...], [...]], 'hola', 'hola', 'hola']
    for item in data:
        if item[0] == "empleados_faltantes":
            return [item[1], "N/A","N/A","N/A","N/A"]
        
    
    # return [[item[1][0][1], "N/A","N/A","N/A","N/A"] for item in data if item[0] == "empleados_faltantes"]

def calculo_etiquetado(self, fecha_anterior,fecha_entrada):
    if fecha_entrada.weekday() == 5 or fecha_entrada.weekday() == 6:
        return datetime.datetime.combine(fecha_entrada.date(), datetime.time(hour=15, minute=0, second=0))
    else:
    # TODO 1.2 Calculo de fechas
        return datetime.datetime.combine(fecha_entrada.date(),fecha_anterior.time())
    

# TODO 1.3 CALCULO DE HORAS 
def calculate_worked_hours(self,record,new_check_out):
    # CIUDAD DE MEXICO A UTC
    timezone = self._context.get('tz') or self.env.user.partner_id.tz or 'UTC'
    self_tz = self.with_context(tz=timezone)
    # # Calculate the time delta in seconds
    # time_delta = fecha_salida - fecha_entrada
    # Verify if the date is a weekend
    calendar = record._get_employee_calendar()
    resource = record.employee_id.resource_id
    check_in_tz = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(record.check_in))
    delta_horas = datetime.timedelta(hours=6)
    new_check_out_utc = new_check_out + delta_horas
    
    if new_check_out == False:
        check_out_tz = check_in_tz
    else:
        # Puse este valor ya que new_check_out ya tiene el horario de ciudad de mexico
        check_out_tz = fields.Datetime.context_timestamp(self_tz, fields.Datetime.from_string(new_check_out_utc))
    lunch_intervals = calendar._attendance_intervals_batch(
                check_in_tz, check_out_tz, resource, lunch=True)
    attendance_intervals = Intervals([(check_in_tz, check_out_tz, record)]) - lunch_intervals[resource.id]
    # delta = sum((i[1] - i[0]).total_seconds() for i in attendance_intervals)
    total_delta_seconds = 0
    for interval in attendance_intervals:
        delta = (interval[1] - interval[0]).total_seconds()
        total_delta_seconds += delta
    minutos = total_delta_seconds / 60
    horas = minutos / 60
    valor_float = float(horas)    
    return valor_float
def segundos_a_horas_minutos_segundos(segundos):
    horas = int(segundos / 3600)
    segundos -= horas * 3600
    minutos = int(segundos / 60)
    segundos -= minutos * 60
    return f"{int(horas):02d}:{int(minutos):02d}:{int(segundos):02d}"
# TODO 5.6 Calculation_of_unregister_days
def calculate_unregistered_days(start_date, end_date, delay_days, holidays,working_days):
    unregistered_days = []
    current_date = start_date.date()
    
    for i in range((end_date.date() - start_date.date()).days + 1):
        if current_date in delay_days:
            pass
        elif current_date in holidays:
            pass
        elif current_date in working_days:
            pass
        else:
            unregistered_days.append(current_date)
        current_date += datetime.timedelta(days=1)
    return unregistered_days
    
# TODO 5.5 Retardos rango de fecha
def limpiador_de_fechas_retardo(fecha_inicio, fecha_fin, dias_semana):
    dias_no_registrados = []
    fecha_actual = fecha_inicio
    diferencia_dias = fecha_fin - fecha_inicio
    diferencias_en_dias = (diferencia_dias.days + 1)
    for _ in range(diferencias_en_dias):
        if fecha_actual.weekday() == 0 and dias_semana["lunes"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 1 and dias_semana["martes"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 2 and dias_semana["miercoles"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 3 and dias_semana["jueves"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 4 and dias_semana["viernes"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 5 and dias_semana["sabado"]:
            dias_no_registrados.append(fecha_actual)
        if fecha_actual.weekday() == 6 and dias_semana["domingo"]:
            dias_no_registrados.append(fecha_actual)
        fecha_actual += datetime.timedelta(days=1)
    return dias_no_registrados
# TODO 5.4 Retardos Fechas
def retardos_dias(employee_id,fecha_inicio,fecha_final,delay_rh,delay_rh_default):
    pepe = []
    for item in delay_rh:
        # Si el empleado se encuentra en la lista de empleados
        if str(employee_id) in item.field_employee:
            # Fecha de inicio y fin del delay , monday...sunday
            dias_semana = {
                "lunes": item.monday_bool,
                "martes": item.tuesday_bool,
                "miercoles": item.wednesday_bool,
                "jueves": item.thursday_bool,
                "viernes": item.friday_bool,
                "sabado": item.saturday_bool,
                "domingo": item.sunday_bool
            }
            
            if fecha_inicio.date() <= item.start_date <= fecha_final.date() and fecha_inicio.date() <= item.end_date <= fecha_final.date():
                pepe.append(limpiador_de_fechas_retardo(item.start_date,item.end_date,dias_semana))
    for item in delay_rh_default:
        if str(employee_id) in item.field_employee:
            # Fecha de inicio y fin del delay , monday...sunday
            dias_semana = {
                "lunes": item.monday_bool,
                "martes": item.tuesday_bool,
                "miercoles": item.wednesday_bool,
                "jueves": item.thursday_bool,
                "viernes": item.friday_bool,
                "sabado": item.saturday_bool,
                "domingo": item.sunday_bool
            }
            pepe.append(limpiador_de_fechas_retardo(fecha_inicio.date(),fecha_final.date(),dias_semana))
        
    conjunto_fechas = set(fecha for sublista in pepe for fecha in sublista)
    lista_sin_repetidos = sorted(list(conjunto_fechas))
    return lista_sin_repetidos
    

# TODO 5.3 Limpieza de fechas
def limpiador_de_fechas(fecha_inicio, fecha_fin):
    dias_no_registrados = []
    fecha_actual = fecha_inicio.date()
    diferencia_dias = fecha_fin - fecha_inicio
    diferencias_en_dias = (diferencia_dias.days + 1)
    for _ in range(diferencias_en_dias):
        dias_no_registrados.append(fecha_actual)
        fecha_actual += datetime.timedelta(days=1)
    return dias_no_registrados

# TODO 5.2 Vacaciones
def vacaciones_dias(hr_holiday):
    dias_no_registrados = []
    for item in hr_holiday:
        dias_no_registrados.append(limpiador_de_fechas(item.date_from,item.date_to))
    return [val for sublist in dias_no_registrados for val in sublist]
        
# TODO 5.1 Faltas
def holidays(employee_id,hora_inicio,hora_final):
    hr_holiday = request.env['hr.leave'].search([
        # F & E & (A|B) & (C|D)
        '&',
        '&',
        '&',
        ('state', '=', 'validate'),
        ('employee_id', '=', employee_id),
        '|',
        ('date_from', '>=', hora_inicio),
        ('date_to', '>=', hora_inicio),
        '|',
        ('date_from', '<=', hora_final),
        ('date_to', '<=', hora_final), 
    ],
    order="date_from,date_to")
    if hr_holiday:
        holidays = vacaciones_dias(hr_holiday)
        return holidays
    else:
        return []
        


def manejo_informacion(self):

    # Obtener la fecha de formulario
    hora_inicio = datetime.datetime.combine(
        self.start_date, datetime.time(hour=0, minute=0, second=0))
    hora_final = datetime.datetime.combine(
        self.end_date, datetime.time(hour=23, minute=59, second=59))
    delta_horas = datetime.timedelta(hours=6)
    final_ahora_si_entrada = hora_inicio + delta_horas
    final_ahora_si_salida = hora_final + delta_horas
    department_id = self.department_id
    id_department = department_id.id

    if department_id:
        empleados = request.env["hr.employee"].search([('department_id', '=',id_department)])
    else:
        empleados = request.env["hr.employee"].search([])

    # Realizo la consulta de datos
    fake_data = request.env["hr.attendance"].search(
        [
            '|',
            '&',
            # '&',
            # ('employee_id.id', 'in', [671, 679, 717]),
            ('check_in', '>=', final_ahora_si_entrada),  # Condición B
            ('check_out', '<=', final_ahora_si_salida),  # Condición C
            '&',
            '&',
            # '&',
            # ('employee_id.id', 'in', [671, 679, 717]),
            ('check_in', '>=', final_ahora_si_entrada),  # Condición B
            ('check_in', '<=', final_ahora_si_salida),  # Condición C
            ('check_out', '=', False),  # Condición F
        ],
        order="employee_id,check_in")

    # Obtener los datos de delay_rh
    delay_rh = request.env["delay.rh.property"].search(
        [
            # ('department_id', '=', id_department),
            '&',
            '&',
            ('is_default', '=', False),
            '|',
            ('start_date', '>=', hora_inicio),
            ('end_date', '>=', hora_inicio),
            '|',
            ('start_date', '<=', hora_final),
            ('end_date', '<=', hora_final),
        ],
        order="start_date,end_date,create_date asc"
    )

    delay_rh_default = request.env["delay.rh.property"].search(
        [
            ('is_default', '=', True),
        ],
        order="start_date,end_date,create_date asc"
    )

   

    # Obtener una lista con los empleados id
    # employee_id = type(delay_rh.employee_id.ids)
    employees_id_list = [eval(item.field_employee) for item in delay_rh]
    employees_id_list_default = [eval(item.field_employee) for item in delay_rh_default]

    data = filter_data(fake_data, id_department)
    # Empleados que no estan
    compare_employee = employee_not_here(empleados, data)


    asistencias = [[asistencia.employee_id.name, asistencia.check_in,
                    asistencia.check_out, asistencia.worked_hours, '',asistencia.employee_id.id,asistencia.employee_id.department_id.name,asistencia] for asistencia in data]

    df = pd.DataFrame(asistencias, columns=[
                      "Empleado", "Entrada", "Salida", "Horas trabajadas", "Retardo","ID","DEPART","All"])

    df["Entrada"] = df["Entrada"].apply(lambda x: parse_time(self, x))
    df["Salida"] = df["Salida"].apply(lambda x: parse_time(self, x))

    df['Horas trabajadas'] = df["Horas trabajadas"]
    # .apply(parse_hour)

    # Usuario Anterior
    df['prevuser'] = df['Empleado'].shift()
    # Usuario Siguiente
    df['nextuser'] = df['Empleado'].shift(-1)

    entrada = df['Entrada'].apply(lambda x: special_parse_time(self, x))
    df['prevdate'] = entrada.shift()
    df['actual'] = entrada
    df['nextdate'] = entrada.shift(-1)
    df['salida_anterior'] = df['Salida'].shift()

    '''
    Creacion de la lista de datos con datos de interes
    '''
    lista = []
    working_days = []
    horas_unitarias = 0
    horas_totales = 0
    first_day = 0
    retardo = 0
    retardo_acumulador = 0
  

    for index, row in df.iterrows():
        #TODO 1.1 LOGICA DE ETIQUETADO
        if row['DEPART'] == "ETIQUETADO" and row['Salida'] == False:
            etiquetado = calculo_etiquetado(self,row['salida_anterior'],row['Entrada'])
            row['Salida'] = etiquetado
            #TODO 1.3 CALCULO DE HORAS
            horas_trabajadas = calculate_worked_hours(self,row['All'],etiquetado)
            row['Horas trabajadas'] = horas_trabajadas
        horas_trabajadas = row['Horas trabajadas']
        row['Horas trabajadas'] = parse_hour(row['Horas trabajadas'])
        # HACER LOGICA DE RETARDOS
        estoy_en_lista = verificar_numero_any(row['ID'], employees_id_list)
        estoy_en_lista_default = verificar_numero_any(row['ID'], employees_id_list_default)
        if estoy_en_lista:
            if first_day != 1:
                retardo = is_retardo(row['ID'], row['Entrada'],delay_rh,delay_rh_default)
        elif estoy_en_lista_default:
            if first_day != 1:
                retardo = is_retardo(row['ID'], row['Entrada'],delay_rh,delay_rh_default)
        if row['Empleado'] != row['nextuser']:
            row['Retardo'] = retardo
            retardo_acumulador += retardo
            ts = pd.Timestamp(row['Entrada'], tz=None)
            
            working_days.append(ts.to_pydatetime().date())
            lista.append(row.values)
            horas_unitarias += horas_trabajadas
            horas_totales += horas_trabajadas
            # AGREGAR HORAS TOTALES CON NUMEROS
            lista.append(['horas_totales_unitarias', parse_hour(horas_unitarias), 'Horas Unitarias',
                          'Horas Unitarias', 'Horas Unitarias'])
            # TODO 5.0 Faltas
            holidays_special = holidays(row['ID'],hora_inicio,hora_final)
            delay_days_special = retardos_dias(row['ID'],hora_inicio,hora_final,delay_rh,delay_rh_default)
            not_working_days_special =  calculate_unregistered_days(hora_inicio,hora_final,delay_days_special,holidays_special,working_days)
            lista.append(["titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas", "titulofaltas"])
            for item in not_working_days_special:
                lista.append(["falta",row['Empleado'], item,'Falta', "Falta"])
            
            working_days = []
            lista.append(['horas_totales_empleado', row['Empleado'],  parse_hour(horas_totales),
                         retardo_acumulador, len(not_working_days_special)])
            horas_unitarias = 0
            horas_totales = 0
            first_day = 0
            retardo = 0
            retardo_acumulador = 0
        elif row['actual'] != row['nextdate']:
            row['Retardo'] = retardo
            retardo_acumulador += retardo
            ts = pd.Timestamp(row['Entrada'], tz=None)
            
            working_days.append(ts.to_pydatetime().date())
            lista.append(row.values)
            horas_unitarias += horas_trabajadas
            horas_totales += horas_trabajadas
            # AGREGAR HORAS UNITARIAS CON NUMEROS
            lista.append(['horas_totales_unitarias', parse_hour(horas_unitarias), 'Horas Unitarias',
                          'Horas Unitarias', 'Horas Unitarias'])
            horas_unitarias = 0
            first_day = 0
            retardo = 0
        else:
            row['Retardo'] = retardo
            retardo_acumulador += retardo
            ts = pd.Timestamp(row['Entrada'], tz=None)
            
            working_days.append(ts.to_pydatetime().date())
            lista.append(row.values)
            horas_unitarias += horas_trabajadas
            horas_totales += horas_trabajadas
            first_day = 1
            retardo = 0
    lista.append(["empleados_faltantes", compare_employee,"hola","hola","hola"])   
    unvideomas = pd.DataFrame(lista, columns=[
                              'Empleado', 'Entrada', 'Salida', 'Horas trabajadas', 'Retardo','ID','DEPART','All', 'prevuser', 'nextuser', 'prevdate', 'actual', 'nextdate','salida_anterior'])

    # Limpieza de valores no necesario
    del unvideomas['ID']
    del unvideomas['DEPART']
    del unvideomas['All']
    del unvideomas['prevuser']
    del unvideomas['nextuser']
    del unvideomas['prevdate']
    del unvideomas['actual']
    del unvideomas['nextdate']
    del unvideomas['salida_anterior']

    return unvideomas.values.tolist()


def datos_generales(self, data):
    return [record for record in data if record[0] == "horas_totales_empleado"]


def creacion_csv(self, data):
    file_path = 'timesheet report' + '.xlsx'
    workbook = xlsxwriter.Workbook('/tmp/' + file_path)

    # Formato de celdas
    BASE_FORMAT_1 = {
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }
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
    formato_fechas_hoja_1_faltas = workbook.add_format(
        {
        'border': 1,
        'num_format': 'yyyy-mm-dd',
        "align": "center",
        "valign": "vcenter",
        "fg_color": "#A5D5E3",
        
        })

    formato_encabezado_hoja_1 = workbook.add_format(BASE_FORMAT_1)
    formato_encabezado_hoja_1.set_fg_color("#6eae4b")
    formato_encabezado_hoja_1.set_font_color("#ffffff")
    formato_encabezado_hoja_1.set_bold(1)

    formato_fechas_hoja_1 = workbook.add_format(BASE_FORMAT_1)
    formato_fechas_hoja_1.set_num_format('yyyy-mm-dd hh:mm')
    formato_fechas_hoja_1.set_fg_color("#E5F1DF")

    horas_unitarias = workbook.add_format(BASE_FORMAT_1)
    horas_unitarias.set_fg_color("#FFDB5C")

    formato_registro_basico_trabajador = workbook.add_format(BASE_FORMAT_1)
    formato_registro_basico_trabajador.set_fg_color("#E5F1DF")

    horas_totales = workbook.add_format(BASE_FORMAT_1)
    horas_totales.set_fg_color("#C3FF93")
    horas_totales.set_bold(1)

    retardos = workbook.add_format(BASE_FORMAT_1)
    retardos.set_fg_color("#FFAF61")
    retardos.set_bold(1)

    # HOJA 1 CREADA
    worksheet = workbook.add_worksheet('Reporte_Unitario')

    worksheet.write(0, 0, "Empleado", formato_encabezado_hoja_1)
    worksheet.write(0, 1, "Entrada", formato_encabezado_hoja_1)
    worksheet.write(0, 2, "Salida", formato_encabezado_hoja_1)
    worksheet.write(0, 3, "Horas Trabajadas", formato_encabezado_hoja_1)
    worksheet.write(0, 4, "Retardo", formato_encabezado_hoja_1)

    row = 1

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
            worksheet.merge_range(row, 1, row, 2, record[2], horas_totales)
            worksheet.write(row, 3, "Retardo", retardos)
            worksheet.write(row, 4, record[3], retardos)
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

    worksheet.set_column('A:A', 30)  # Adjust width for column A (Employee)
    worksheet.set_column('B:B', 20)  # Adjust width for column B (Entrada)
    worksheet.set_column('C:C', 20)  # Adjust width for column C (Salida)
    # Adjust width for column D (Horas Trabajadas)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 10)  # Adjust width for column E (Color)
    # HOJA 1 TERMINADA

    # HOJA 2 CREADA

    # FORMATO
    BASE_FORMAT_2 = {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }

    title_format = workbook.add_format(BASE_FORMAT_2)
    title_format.set_bg_color("#FFAF61")
    title_format.set_font_color("#000000")
    title_format.set_font_size(11)

    formato_fechas_hoja_2 = workbook.add_format(BASE_FORMAT_2)
    formato_fechas_hoja_2.set_bg_color("#FFAF61")
    formato_fechas_hoja_2.set_font_color("#000000")
    formato_fechas_hoja_2.set_font_size(11)
    formato_fechas_hoja_2.set_num_format('yyyy-mm-dd')

    formato_encabezado_hoja_2 = workbook.add_format(BASE_FORMAT_2)
    formato_encabezado_hoja_2.set_bg_color("#FFAF61")
    formato_encabezado_hoja_2.set_font_color("#000000")

    formato_encabezado_hoja_1 = workbook.add_format(BASE_FORMAT_2)
    formato_encabezado_hoja_1.set_bg_color("#6eae4b")
    formato_encabezado_hoja_1.set_font_color("#ffffff")

    formato_encabezado_hoja_1_faltas = workbook.add_format(BASE_FORMAT_2)
    formato_encabezado_hoja_1_faltas.set_bg_color("#82B8F3")
    formato_encabezado_hoja_1_faltas.set_font_color("#000000")

    horas_totales = workbook.add_format(BASE_FORMAT_2)
    horas_totales.set_bg_color("#C3FF93")

    retardos = workbook.add_format(BASE_FORMAT_2)
    retardos.set_bg_color("#FFAF61")

    pares_hoja2 = workbook.add_format(BASE_FORMAT_2)
    pares_hoja2.set_bg_color("#A5D5E3")
    pares_hoja2.set_bold(0)

    impares_hoja2 = workbook.add_format(BASE_FORMAT_2)
    impares_hoja2.set_bg_color("#E5F1DF")
    impares_hoja2.set_bold(0)

    horas_unitarias = workbook.add_format(BASE_FORMAT_2)
    horas_unitarias.set_bg_color("#FFDB5C")
    horas_unitarias.set_bold(0)

    worksheet2 = workbook.add_worksheet('Reporte_General')

    row = 7
    row2 = 7
    # Titulo
    worksheet2.merge_range(1, 0, 2, 3, "Reporte de Asistencia", title_format)
    # Subtitulo
    worksheet2.merge_range(
        4, 0, 4, 3, f"Periodo: {self.start_date} - {self.end_date} / Departamento: {self.department_id.complete_name if self.department_id else 'N/A'}", formato_fechas_hoja_2)

    # Encabezados
    worksheet2.write(6, 0, "Empleado", formato_encabezado_hoja_2)
    worksheet2.write(6, 1, "Total Horas Trabajadas", formato_encabezado_hoja_2)
    worksheet2.write(6, 2, "Total de Retardos", formato_encabezado_hoja_2)
    worksheet2.write(6, 3, "Total de Faltas", formato_encabezado_hoja_2)

    general_date = datos_generales(self, data)
    # NOT EMPLOYEE

    for index, item in enumerate(general_date):
        if index % 2 == 0:
            worksheet2.write(row, 0, item[1], pares_hoja2)
            worksheet2.write(row, 1, item[2], pares_hoja2)
            worksheet2.write(row, 2, item[3], pares_hoja2)
            worksheet2.write(row, 3, item[4], pares_hoja2)
            row += 1
        else:
            worksheet2.write(row, 0, item[1], impares_hoja2)
            worksheet2.write(row, 1, item[2], impares_hoja2)
            worksheet2.write(row, 2, item[3], impares_hoja2)
            worksheet2.write(row, 3, item[4], impares_hoja2)
            # worksheet2.write(row, 4, item[4], impares_hoja2)
            row += 1
    # Empleados que no estan en el reporte
    worksheet2.merge_range(6, 5, 6,8, "Empleados Fuera de Rango",formato_encabezado_hoja_2)
    not_employee = manejo_informacion_employee_not(self, data)
    for index, item in enumerate(not_employee[0]):
        if index % 2 == 0:
            worksheet2.merge_range(row2, 5, row2,8, item[1],pares_hoja2)
            row2 += 1
        else:
            worksheet2.merge_range(row2, 5, row2,8, item[1],impares_hoja2)
            # worksheet2.write(row2, 4, item[4], impares_hoja2)
            row2 += 1


    # Manually adjust the width of each column
    worksheet2.set_column('A:A', 35)  # Adjust width for column A (Employee)
    worksheet2.set_column('B:B', 20)  # Adjust width for column B (Entrada)
    worksheet2.set_column('C:C', 20)  # Adjust width for column C (Salida)
    # Adjust width for column D (Horas Trabajadas)
    worksheet2.set_column('D:D', 15)

    # HOJA 2 TERMINADA

    workbook.close()
    ex_report = base64.b64encode(open('/tmp/' + file_path, 'rb+').read())
    excel_report_id = self.env['save.ex.report.special.wizard'].create({"document_frame": file_path,
                                                                        "file_name": ex_report})
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


class ReportSpecialWizard(models.TransientModel):
    _name = "report.special.wizard"
    _description = "Reporte Especial"

    end_date = fields.Date(default=lambda self: fields.datetime.now())
    start_date = fields.Date(
        default=lambda self: fields.datetime.now() - datetime.timedelta(days=15))

    department_id = fields.Many2one("hr.department")

    def action_generate_csv(self):
        data = manejo_informacion(self)

        return creacion_csv(self, data)
