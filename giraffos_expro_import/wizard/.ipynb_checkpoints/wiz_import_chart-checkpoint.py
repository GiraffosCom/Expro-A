# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import time
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

@api.multi
def determina_periodo(self,fecha):
    arreglo = self.env['expro.siniestro.periodo'].search([])
    for item in arreglo:
        if fecha >= item.fecha_inicio and fecha <=item.fecha_fin:
            return item.name


class ImportChartAccount(models.TransientModel):
    _name = "import.expro.siniestro"

    name = fields.Char(string='Name')
    archivo = fields.Binary(string="Seleccionar archivo")

    @api.multi
    def import_archivo(self):
        data = base64.b64decode(self.archivo)
        wb = xlrd.open_workbook(file_contents=data)
        sheet = wb.sheet_by_index(0)

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

        # Borrar todos los siniestros antes de crear
        self.env['expro.siniestro'].search([]).unlink()
        meses_dict={
            1:'Enero',
            2:'Febrero',
            3:'Marzo',
            4:'Abril',
            5:'Mayo',
            6:'Junio',
            7:'Julio',
            8:'Agosto',
            9:'Septiembre',
            10:'Octubre',
            11:'Noviembre',
            12:'Diciembre'}
        

        for item in seleccionados:

            item_dict = {}
            
            j = 0
            for j in range(columnas):
                celda = sheet.cell_value(item, j)
                if j == 0 and sheet.cell_type(item, j) == 3:
                    fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                    fecha = fechahora.date()
                    item_dict['fecha_ingreso'] = fecha
                    #lista.append(fecha)
                elif j == 2:
                    item_dict['rut'] = celda
                    #lista.append(celda)
                elif j == 3:
                    item_dict['nombre'] = celda
                    #lista.append(celda)
                elif j == 4:
                    item_dict['tipo_accidente'] = celda
                    #lista.append(celda)
                elif j == 5:
                    item_dict['estado_calificacion'] = celda
                    #lista.append(celda)
                elif j == 6:
                    item_dict['reingreso'] = celda
                    #lista.append(celda)
                elif j == 7:
                    if sheet.cell_type(item, j) == 3:
                        fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                        fecha = fechahora.date()
                        item_dict['fecha_inicio_reposo'] = fecha
                        #lista.append(fecha)
                    else:
                        fecha = None
                        item_dict['fecha_inicio_reposo'] = False
                        #lista.append(fecha)
                elif j == 8:
                    if sheet.cell_type(item, j) == 3:
                        fechahora = datetime(*xlrd.xldate_as_tuple(sheet.cell_value(item, j), wb.datemode))
                        fecha = fechahora.date()
                        item_dict['fecha_termino_reposo'] = fecha
                        #lista.append(fecha)
                    else:
                        fecha = None
                        item_dict['fecha_termino_reposo'] = False
                        #lista.append(fecha)
                elif j == 9 and i > 0:
                    item_dict['dias_licencia'] = int(celda)
                    #lista.append(int(celda))
                elif j == 12 and i > 0:
                    palabra=str(celda)
                    delimitador="("
                    n=palabra.find(delimitador)
                    item_dict['centro_costo'] = palabra[0:n]
                    ind=re.findall('\d+',celda).pop()
                    #item_dict['rs_id']=str(ind)
                    item_dict['rs_id']=ind
                    
                    #lista.append(celda)
                j = j + 1

            item_dict['archivo_procesado'] = self.name
            self.env['expro.siniestro'].create(item_dict)
            
        arraste_dict={}
        
        #elimina los registros de dias perdidos
        self.env['expro.siniestro.dias.perdidos'].search([]).unlink()
        
        
        #obtiene listado de siniestros procesados
        listado = self.env['expro.siniestro'].search([])
        #_logger.info('_E_: {c}:'.format(c='hola')) 
        
        
        for record in listado:
            #casos con arrastre
            if record.fecha_inicio_reposo.month != record.fecha_termino_reposo.month:
                fecha_inicio=record.fecha_inicio_reposo
                fecha_termino=record.fecha_termino_reposo
                fecha_termino_ficticia=date(fecha_termino.year,fecha_termino.month, calendar.monthrange(fecha_termino.year, fecha_termino.month)[1])
                centro=record.centro_costo
                rsoc=record.rs_des 
                dates=[dt for dt in rrule(MONTHLY, dtstart=fecha_inicio, until=fecha_termino_ficticia)]
                
                for i in range (0,len(dates)):
                    if i==0:
                        dias_restantes=calendar.monthrange(dates[i].year,dates[i].month)[1]-fecha_inicio.day+1
                        arraste_dict['dias_perdidos']=dias_restantes
                        arraste_dict['mes']=meses_dict[dates[i].month]                        
                        arraste_dict['periodo']=record.periodo
                        arraste_dict['centro_costo']=centro
                        arraste_dict['codigo_rs']=record.rs_id
                        arraste_dict['rs']=rsoc
                        arraste_dict['arrastre']='Si'                            
                    if i==len(dates)-1:
                        dias_restantes=fecha_termino.day
                        arraste_dict['dias_perdidos']=dias_restantes
                        arraste_dict['mes']=meses_dict[dates[i].month]
                        arraste_dict['periodo']=determina_periodo(self,fecha_termino)
                        arraste_dict['centro_costo']=centro
                        arraste_dict['codigo_rs']=record.rs_id
                        arraste_dict['rs']=rsoc
                        arraste_dict['arrastre']='Si'
                    elif i>0 and i<len(dates)-1:
                        dias_restantes=calendar.monthrange(dates[i].year,dates[i].month)[1]
                        arraste_dict['dias_perdidos']=dias_restantes
                        arraste_dict['mes']=meses_dict[dates[i].month]
                        arraste_dict['periodo']=determina_periodo(self,date(dates[i].year, dates[i].month,dias_restantes ))
                        arraste_dict['centro_costo']=centro
                        arraste_dict['codigo_rs']=record.rs_id
                        arraste_dict['rs']=rsoc
                        arraste_dict['arrastre']='Si'
                    self.env['expro.siniestro.dias.perdidos'].create(arraste_dict)
            #casos sin arrastre. Normales.
            else:
                arraste_dict['dias_perdidos']=record.dias_licencia
                arraste_dict['mes']=meses_dict[record.fecha_inicio_reposo.month]                
                arraste_dict['periodo']=record.periodo
                arraste_dict['centro_costo']=record.centro_costo
                arraste_dict['codigo_rs']=record.rs_id
                arraste_dict['rs']=record.rs_des
                arraste_dict['arrastre']='No'
                self.env['expro.siniestro.dias.perdidos'].create(arraste_dict)
        
        registros_dp=self.env['expro.siniestro.dias.perdidos'].search([])
        _logger.info('_AAA_: {c}:'.format(c=len(registros_dp)))
        rs_set=set()
        rs_id_set=set()
        
        for elemento in registros_dp:
            rs_id_set.add(elemento.codigo_rs)
        _logger.info('_BBB_: {c}:'.format(c=len(rs_id_set)))
            
            
                
                
                            
                            
                            
   
            
            
            #accidente = crea_siniestro(lista)
            #resultado.append(accidente)
            #lista = []

        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }
