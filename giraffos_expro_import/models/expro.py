# -*- coding: utf-8 -*-

from odoo import models, fields, api, _

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
    razon_social=fields.Char(string='Razón Social')
    periodo = fields.Many2one(comodel_name='expro.siniestro.periodo', string='Periodo', compute='_get_periodo', store=True)
    archivo_procesado = fields.Char('Archivo')

    @api.multi
    @api.depends('fecha_inicio_reposo')
    def _get_periodo(self):
        for siniestro in self:
            periodo_id = self.env['expro.siniestro.periodo'].search([('fecha_inicio','<=',siniestro.fecha_inicio_reposo),('fecha_fin','>=',siniestro.fecha_inicio_reposo)], limit=1)
            siniestro.periodo = periodo_id.id

class ExproSiniestroPeriodo(models.Model):
    _name = 'expro.siniestro.periodo'

    name = fields.Char('Nº Periodo')
    fecha_inicio = fields.Date('Fecha Inicio')
    fecha_fin = fields.Date('Fecha Fin')

class ExproSiniestroDP(models.Model):
    _name = 'expro.siniestro.dias.perdidos'
    
    dias_perdidos=fields.Integer(string='Días Perdidos')
    mes=fields.Integer(string='Mes')
    periodo=fields.Integer(string='Periodo')
    centro_costo=fields.Char(string='Centro de costo')
    arrastre=fields.Char(string='Arrastre')
    
    
    


    
