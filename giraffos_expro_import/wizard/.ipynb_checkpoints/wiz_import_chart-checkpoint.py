# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import time
from datetime import datetime
import calendar
from dateutil.rrule import rrule, MONTHLY
import tempfile
import binascii
try:
    import xlrd
    try:
        from xlrd import xlsx
    except ImportError:
        xlsx = None
except ImportError:
    xlrd = xlsx = None

from datetime import date, datetime
from odoo.exceptions import Warning, UserError
from odoo import models, fields, exceptions, api, _
import logging
_logger = logging.getLogger(__name__)
import io
import re

try:
	import csv
except ImportError:
	_logger.debug('Cannot `import csv`.')
try:
	import xlwt
except ImportError:
	_logger.debug('Cannot `import xlwt`.')
try:
	import cStringIO
except ImportError:
	_logger.debug('Cannot `import cStringIO`.')
try:
	import base64
except ImportError:
	_logger.debug('Cannot `import base64`.')

class Siniestro():
    def __init__(self,fecha_ingreso,rut,nombre,tipo_accidente,estado_calificacion,reingreso,fecha_inicio_reposo,fecha_termino_reposo,dias_licencia,centro_costo):
        self.fecha_ingreso=fecha_ingreso
        self.rut=rut
        self.nombre=nombre
        self.tipo_accidente=tipo_accidente
        self.estado_calificacion=estado_calificacion
        self.reingreso=reingreso
        self.fecha_inicio_reposo=fecha_inicio_reposo
        self.fecha_termino_reposo=fecha_termino_reposo
        self.dias_licencia=dias_licencia
        self.centro_costo=centro_costo
        self.periodo=determina_periodo(self.fecha_inicio_reposo)
        
class Arrastre():
    def __init__(self,dias,mes,periodo_id,centro_costo,index,isArrastre):
        self.dias=dias
        self.mes=mes
        self.periodo_id=periodo_id
        self.centro_costo=centro_costo
        self.index=index
        self.isArrastre=isArrastre
        
class Periodo():
    def __init__(self,id,inicio,termino,glosa):
        self.id=id
        self.inicio=inicio
        self.termino=termino
        self.glosa=glosa

def crea_siniestro(milista):
    class_member=Siniestro(milista[0],milista[1],milista[2],milista[3],milista[4],milista[5],milista[6],milista[7],milista[8],milista[9])
    return class_member

def inicializa_periodos():
    lista=[]
    ini_date_1=date(2018,7,1)
    end_date_1=date(2019,6,30)
    id=1
    glosa1="Periodo 1"
    lista.append(Periodo(id,ini_date_1,end_date_1,glosa1))
    ini_date_2=date(2019,7,1)
    end_date_2=date(2020,6,30)
    id=2
    glosa2="Periodo 2"
    lista.append(Periodo(id,ini_date_2,end_date_2,glosa2))
    ini_date_3=date(2020,7,1)
    end_date_3=date(2021,6,30)
    id=3
    glosa3="Periodo 3"
    lista.append(Periodo(id,ini_date_3,end_date_3,glosa3))
    return lista

def determina_periodo(fecha):
    lista_periodos = inicializa_periodos()
    for item in lista_periodos:
        if fecha >= item.inicio and fecha <=item.termino:
            return item.id
        
def encuentra_casos_arrastre(milista):
    indice_arrastres =[]
    i=0
    for evento in milista:
        if evento.fecha_inicio_reposo.month != evento.fecha_termino_reposo.month:
            indice_arrastres.append(i)
        i=i+1
    return indice_arrastres

def encuentra_casos_reingresos(milista):
    indice_reingresos=[]
    i=0
    for evento in milista:
        if evento.reingreso=="Si":
            indice_reingresos.append(i)
        i=i+1
    return indice_reingresos

def calcula_dp_arrastre(indexlist,accidentelist):
    rslt=[]

    for indice in indexlist:
        fecha_incio=accidentelist[indice].fecha_inicio_reposo
        fecha_termino=accidentelist[indice].fecha_termino_reposo
        fecha_termino_ficticia=date(fecha_termino.year,fecha_termino.month, calendar.monthrange(fecha_termino.year, fecha_termino.month)[1])
        centro=accidentelist[indice].centro_costo
    
        dates = [dt for dt in rrule(MONTHLY, dtstart=fecha_incio, until=fecha_termino_ficticia)]
        for i in range (0,len(dates)):
            if i==0:
                dias_restantes=calendar.monthrange(dates[i].year,dates[i].month)[1]-fecha_incio.day+1
                mes=dates[i].month
                per=determina_periodo(fecha_incio)
                elemento= Arrastre(dias_restantes,mes,per,centro,indice,True)
                rslt.append(elemento)
            if i==len(dates)-1:
                dias_restantes=fecha_termino.day
                mes=dates[i].month
                per=determina_periodo(fecha_termino)
                elemento= Arrastre(dias_restantes,mes,per,centro,indice,True)
                rslt.append(elemento)
            elif i>0 and i<len(dates)-1:
                dias_restantes=calendar.monthrange(dates[i].year,dates[i].month)[1]
                mes=dates[i].month
                per=determina_periodo(date(dates[i].year, dates[i].month,dias_restantes ))
                elemento= Arrastre(dias_restantes,mes,per,centro,indice,True)
                rslt.append(elemento)

    return rslt

def calcula_dp_normal(indexlist,accidentelist):
    rl_normal=[]
    complemento_index=[]
    for i in range(0, len(accidentelist)):
        if i not in indexlist:
            complemento_index.append(i)
    
    for indice in complemento_index:
        dias=accidentelist[indice].dias_licencia
        mes=accidentelist[indice].fecha_inicio_reposo.month
        per=determina_periodo(accidentelist[indice].fecha_ingreso)
        centro=accidentelist[indice].centro_costo
        elemento=Arrastre(dias,mes,per,centro,indice,False)
        rl_normal.append(elemento)

    return rl_normal

class ImportChartAccount(models.TransientModel):
    _name = "import.expro.siniestro"

    name = fields.Char(string='Name')
    archivo = fields.Binary(string="Seleccionar archivo")
    
    @api.multi
    def import_archivo(self):
        data = base64.b64decode(self.archivo)
        wb = xlrd.open_workbook(file_contents=data)
        sheet = wb.sheet_by_index(0)
        
        diccionario={
        "2":"RIVAS Y ASOCIADOS",
        "3":"EXPROCAP S.A.",
        "5":"EXPROCHILE S.A.",
        "700":"EST EXPROSERVICIOS S.A.",
        "800":"EST EXPROTIEMPO S.A.",
        "900":"EXPROSERVICIOS S.A.",
        "8":"Helpnet Ingeniería y Servicios de Recursos Humanos S.A",
        "9":"Acrotex Chile Comercial S.A",
        "10":"Helpnet Empresa de Servicios Transitorios S.A",
        "11":"Centro de Capacitacion de Competencias Tecnológicas S.A.",
        "31":"Back Office South América"}

        filas = sheet.nrows
        columnas = sheet.ncols
        i = 0
        lista = []
        resultado = []
        seleccionados = []
        for i in range(filas):
            j = 0
            for j in range(columnas):
                celda = sheet.cell_value(i, j)
                if j == 4 and i > 0:
                    if celda == "ACCIDENTE DE TRABAJO":
                        if sheet.cell_value(i, j + 1) == "ACEPTADO" or sheet.cell_value(i, j + 1) == "PENDIENTE":
                            if sheet.cell_value(i, j + 5) > 0:
                                seleccionados.append(i)
                    elif celda == "ENFERMEDAD PROFESIONAL":
                        if sheet.cell_value(i, j + 1) == "ACEPTADO" or sheet.cell_value(i, j + 1) == "PENDIENTE":
                            if sheet.cell_value(i, j + 5) > 0:
                                seleccionados.append(i)
                j = j + 1
            i = i + 1

        # Borrar Todo
        # self.env['expro.siniestro'].search([]).unlink()

        for item in seleccionados:

            item_dict = {}

            j = 0
            for j in range(columnas):
                celda = sheet.cell_value(item, j)
                if j == 0 and sheet.cell_type(item, j) == 3:
                    fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                    fecha = fechahora.date()
                    item_dict['fecha_ingreso'] = fecha
                    lista.append(fecha)
                elif j == 2:
                    item_dict['rut'] = celda
                    lista.append(celda)
                elif j == 3:
                    item_dict['nombre'] = celda
                    lista.append(celda)
                elif j == 4:
                    item_dict['tipo_accidente'] = celda
                    lista.append(celda)
                elif j == 5:
                    item_dict['estado_calificacion'] = celda
                    lista.append(celda)
                elif j == 6:
                    item_dict['reingreso'] = celda
                    lista.append(celda)
                elif j == 7:
                    if sheet.cell_type(item, j) == 3:
                        fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                        fecha = fechahora.date()
                        item_dict['fecha_inicio_reposo'] = fecha
                        lista.append(fecha)
                    else:
                        fecha = None
                        item_dict['fecha_inicio_reposo'] = False
                        lista.append(fecha)
                elif j == 8:
                    if sheet.cell_type(item, j) == 3:
                        fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                        fecha = fechahora.date()
                        item_dict['fecha_termino_reposo'] = fecha
                        lista.append(fecha)
                    else:
                        fecha = None
                        item_dict['fecha_termino_reposo'] = False
                        lista.append(fecha)
                elif j == 9 and i > 0:
                    item_dict['dias_licencia'] = int(celda)
                    lista.append(int(celda))
                elif j == 12 and i > 0:
                    palabra=str(celda)
                    delimitador="("
                    n=palabra.find(delimitador)
                    
                    item_dict['centro_costo'] = palabra[0:n]
                    ind=re.findall('\d+',celda).pop()
                    valor=diccionario.get(ind)
                    if valor:
                        item_dict['razon_social']=valor
                    else:
                        item_dict['razon_social']="Desconocido"
                        
                    
                    lista.append(celda)
                j = j + 1

            item_dict['archivo_procesado'] = self.name
            self.env['expro.siniestro'].create(item_dict)
            accidente = crea_siniestro(lista)
            resultado.append(accidente)
            lista = []
            
            dp_arrastre=calcula_dp_arrastre(encuentra_casos_arrastre(resultado),resultado)
            dp_normales=calcula_dp_normal(encuentra_casos_arrastre(resultado),resultado)
            
            dict_arrastre = {}
            dict_normal= {}
            
            self.env['expro.siniestro.dias.perdidos'].search([]).unlink()
            
            for item in dp_arrastre:
                dict_arrastre['dias_perdidos']=item.dias
                dict_arrastre['mes']=item.mes
                dict_arrastre['periodo']=item.periodo_id
                dict_arrastre['centro_costo']=item.centro_costo
                dict_arrastre['arrastre']=item.isArrastre
                self.env['expro.siniestro.dias.perdidos'].create(dict_arrastre)
                
            for item in dp_normales:
                dict_normal['dias_perdidos']=item.dias
                dict_normal['mes']=item.mes
                dict_normal['periodo']=item.periodo_id
                dict_normal['centro_costo']=item.centro_costo
                if item.isArrastre  !=True:
                    dict_arrastre['arrastre']="False"
                    
                self.env['expro.siniestro.dias.perdidos'].create(dict_normal)               

        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }
