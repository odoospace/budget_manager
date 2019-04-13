# -*- coding: utf-8 -*-
from openerp import models, fields, api

class account_analytic_account(models.Model):
    _inherit = 'account.analytic.account'

    # TODO: add i18n
    group = fields.Selection([
        ('A', 'Gastos Secretarias'),
        ('B', 'Gastos Areas y Equipos'),
        ('C', 'Gastos Generales'),
        ('D', 'Ingresos'),
        ('E', 'Otros'),
        ('F', 'Gastos Extraordinarios')
    ])
    
    def first_parent(self, parent=None):
        # is parent empty?
        if not parent:
            obj = self
        else:
            obj = parent
        if not obj.parent_id:
            return obj
        else:
            return self.first_parent(obj.parent_id)

