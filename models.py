# -*- coding: utf-8 -*-
from openerp import models, fields, api
import openerp.addons.decimal_precision as dp
import calendar
from datetime import date, datetime

class crossovered_budget(models.Model):
    _inherit = 'crossovered.budget'

    budget_manager_line_ids = fields.One2many('budget_manager.line', 'crossovered_budget_id')

    def manage_crossovered_budget_lines(self):
        # remove old lines
        for line in self.crossovered_budget_line:
            line.unlink()
        
        # recreate budget lines
        groups = {}
        for line in self.budget_manager_line_ids:
            key = (line.analytic_account_id.id, line.general_budget_id.id)
            month_from = int(line.date_from.split('-')[1])
            month_to = int(line.date_to.split('-')[1])
            # TODO: add periods longer 1 month
            if month_from == month_to:
                key2 = month_from
            else:
                key2 = 0
            if not groups.has_key(key):
                groups[key] = {key2: line}
            elif not groups[key].has_key(key2):
                groups[key][key2] = line 

        # TODO: better management for periods
        for k, v in groups.items():
            months = {}
            t = 0
            for j in v:
                if j != 0:
                   months[j] = (groups[k][j], 1) # value, parts
                   t += groups[k][j].planned_amount
            if groups[k].has_key(0):
                months[0] = (groups[k][0], 13-len(groups[k])) # include 0 itself
        
            print '***', months
            # create budget lines
            for i in range(12):
                month = i + 1
                if months.has_key(month):
                    line = months[month][0]
                    div = months[month][1]
                    planned_amount = line.planned_amount
                else:
                    line = months[0][0]
                    div = months[0][1]
                    planned_amount = (line.planned_amount - t) / div
                # TODO: dynamic year
                first_day = 1
                last_day = calendar.monthrange(2018, month)[1]
                values = {
                    'crossovered_budget_id': self.id,
                    'budget_manager_line_id': line.id,
                    'analytic_account_id': line.analytic_account_id.id,
                    'general_budget_id': line.general_budget_id.id,
                    'date_from': date(2018, month, first_day),
                    'date_to':  date(2018, month, last_day),
                    'planned_amount': planned_amount
                }
                self.env['crossovered.budget.lines'].create(values)

    @api.model
    def create(self, values):
        # TODO: only one overlap by period  
        res_id = super(crossovered_budget, self).create(values)
        self.manage_crossovered_budget_lines()
        return res_id


    @api.multi
    def write(self, values):
        # TODO: only one overlap by period  
        print '>>>> write', values
        res_id = super(crossovered_budget, self).write(values)
        self.manage_crossovered_budget_lines()
        return res_id


class crossovered_budget_lines(models.Model):
    _inherit = 'crossovered.budget.lines'

    budget_manager_line_id = fields.Many2one('budget_manager.line')


class budget_manager_line(models.Model):
    _name = 'budget_manager.line'
    _description = "Budget Line Managed"
    
    crossovered_budget_id = fields.Many2one('crossovered.budget', 'Budget', ondelete='cascade', required=True)
    crossovered_budget_line_ids = fields.One2many('crossovered.budget.lines', 'budget_manager_line_id')
    analytic_account_id = fields.Many2one('account.analytic.account', 'Analytic Account')
    general_budget_id = fields.Many2one('account.budget.post', 'Budgetary Position', required=True)
    date_from = fields.Date('Start Date', required=True)
    date_to = fields.Date('End Date', required=True)
    planned_amount = fields.Float('Planned Amount', required=True) #, digits=dp.get_precision('Account')),
    company_id = fields.Many2one(related = 'crossovered_budget_id.company_id', store=True)
