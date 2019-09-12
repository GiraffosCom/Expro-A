# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
import logging
_logger = logging.getLogger(__name__)

class ExproSiniestro(models.Model):
    _name = 'expro.siniestro'

    fecha_ingreso = fields.Date(string='Fecha Ingreso')
    rut = fields.Char(string='RUT')
    nombre = fields.Char(string='Nombre')
    tipo_accidente = fields.Char(string='Tipo Accidente')
    estado_calificacion = fields.Char(string='Estado Calificacion')
    reingreso = fields.Char(string='Reingreso')
    fecha_inicio_reposo = fields.Date(string='Fecha Inicio Reposo')
    fecha_termino_reposo = fields.Date(string='Fecha Termino Reposo')
    dias_licencia = fields.Integer(string='Dias Licencia')
    centro_costo = fields.Char(string='Centro de costo')
    rs_id=fields.Integer(string='Id Razón Social')
    rs_des=fields.Char(comodel_name='expro.siniestro.razon.social',string='Razón Social', compute='_get_razon_social', store=True)
    periodo = fields.Integer(comodel_name='expro.siniestro.periodo', string='Periodo', compute='_get_periodo', store=True)
    archivo_procesado = fields.Char('Archivo')

    @api.multi
    @api.depends('fecha_inicio_reposo')
    def _get_periodo(self):
        for siniestro in self:
            periodo_id = self.env['expro.siniestro.periodo'].search([('fecha_inicio','<=',siniestro.fecha_inicio_reposo),('fecha_fin','>=',siniestro.fecha_inicio_reposo)], limit=1).name
            siniestro.periodo = periodo_id

    @api.multi
    @api.depends('rs_id')
    def _get_razon_social(self):
        for siniestro in self:
            var=self.env['expro.siniestro.razon.social'].search([('codigo','=',siniestro.rs_id)],limit=1).rsocial
            #_logger.info('_RS_ {c}:'.format(c=var)) 
            siniestro.rs_des =var

class ExproSiniestroPeriodo(models.Model):
    _name = 'expro.siniestro.periodo'

    name = fields.Integer('Nº Periodo')
    fecha_inicio = fields.Date('Fecha Inicio')
    fecha_fin = fields.Date('Fecha Fin')

class ExproSiniestroDP(models.Model):
    _name = 'expro.siniestro.dias.perdidos'
    
    dias_perdidos=fields.Integer(string='Días Perdidos')
    mes=fields.Char(string='Mes')
    mes_id=fields.Char(string='ID MES')
    periodo=fields.Integer(string='Periodo')
    centro_costo=fields.Char(string='Centro de costo')
    codigo_rs=fields.Integer(string='Id Razón Social')
    rs=fields.Char(string='Razón Social')
    arrastre=fields.Char(string='Arrastre')

class ExproSiniestroRazonSocial(models.Model):
    _name ='expro.siniestro.razon.social'

    codigo=fields.Integer(string='Código RS')
    rsocial=fields.Char(string='Razón Social')
    worker_num=fields.Integer(string='Nº Trabajadores')
    
class ExproSiniestroTasas(models.Model):
    _name = 'expro.siniestro.tasas'
    
    rango_inferior=fields.Integer(string='Rango Inferior')
    rango_superior=fields.Integer(string='Rango Superior')
    cotizacion_adicional=fields.Float(string='Cotización Adicional')
    tasa_total=fields.Float(string='Tasa Total')
    situacion_acutal=fields.Boolean(string='Situación Actual')
    codigo_rs=fields.Integer(string='Código RS')
    #rs_des=fields.Char(comodel_name='expro.siniestro.razon.social',string='Razón Social', compute='_get_razon_social', store=True)
    rs_des=fields.Char(string='Razón Social')

    #Sirve para que en el Form se ingrese el codigo_rs y se obtenga el rs_des en el tree view
    #@api.multi
    #@api.depends('codigo_rs')
    #def _get_razon_social(self):
        #for record in self:
            #var=self.env['expro.siniestro.razon.social'].search([('codigo','=',record.codigo_rs)],limit=1).rsocial
            #record.rs_des =var
    
class ExproSiniestroProyeccion(models.Model):
    _name = 'expro.siniestro.proyeccion'
    
    per_id=fields.Integer('Nº Periodo')
    year_inicio=fields.Integer('Año Inicio')
    year_termino=fields.Integer('Año Termino')
    tasa_siniestralidad_periodo=fields.Float('Tasa Periodo')
    num_trabajadores_promedio_mensual=fields.Integer('Año Termino')
    tst_mensual_max=fields.Float('Tasa mensual máxima')
    dp_max_mes=fields.Integer('Días perdidos máximo por mes')
    dp_max_periodo=fields.Integer('Días perdidos máximo por periodo')
    
    
    
