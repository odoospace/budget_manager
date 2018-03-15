# -*- coding: utf-8 -*-
from openerp import models, fields, api

class account_analytic_account(models.Model):
    _inherit = 'account.analytic.account'

    group = fields.Selection([
        ('A', 'Gastos Secretarias'),
        ('B', 'Gastos Areas y Equipos'),
        ('C', 'Gastos Generales'),
        ('D', 'Ingresos'),
        ('E', 'Otros')
    ])