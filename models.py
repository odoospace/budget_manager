# -*- coding: utf-8 -*-
from openerp import models, fields, api
import openerp.addons.decimal_precision as dp
import calendar
from datetime import date, datetime

class crossovered_budget(models.Model):
    _inherit = 'crossovered.budget'

    def _search_segment_user(self, operator, value):
        user = self.env['res.users'].browse(value)
        segment_tmpl_ids = []
        segment_ids = user.segment_ids
        for s in segment_ids:
            segment_tmpl_ids += s.segment_id.segment_tmpl_id.get_childs_ids()
        virtual_segments = self.env['analytic_segment.template'].search([('virtual', '=', True)])
        segment_tmpl_ids += [i.id for i in virtual_segments]

        segment_ids = self.env['analytic_segment.segment'].search([('segment_tmpl_id', 'in', segment_tmpl_ids)])

        return [('segment_id', 'in', [i.id for i in segment_ids])]

    @api.multi
    def _segment_user_id(self):
        # TODO: use a helper in analytic_segment if it's possible...
        if self.env.user.id == 1:
            for obj in self:
                obj.segment_user_id = self.env.uid
        else:
            # add users segments
            segment_tmpl_ids = []
            segment_ids = self.env.user.segment_ids
            for s in segment_ids:
                segment_tmpl_ids += s.segment_id.segment_tmpl_id.get_childs_ids()
            # add virtual companies segments
            virtual_segments = self.env['analytic_segment.template'].search([('virtual', '=', True)])
            segment_tmpl_ids += [i.id for i in virtual_segments]

            # mark segments with user id
            segment_ids = self.env['analytic_segment.segment'].search([('segment_tmpl_id', 'in', segment_tmpl_ids)])
            for obj in self:
                if obj.segment_id in segment_ids:
                    obj.segment_user_id = self.env.uid


    budget_manager_line_ids = fields.One2many('budget_manager.line', 'crossovered_budget_id')
    segment_id = fields.Many2one('analytic_segment.segment', required=True)
    segment = fields.Char(related='segment_id.segment', readonly=True)
    segment_user_id = fields.Many2one('res.users', compute='_segment_user_id', search=_search_segment_user)


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
        last_month = 0
        for k, v in groups.items():
            months = {}
            t = 0
            m = 0
            for j in v:
                if j != 0:
                   months[j] = groups[k][j] # value, parts
                   t += groups[k][j].planned_amount
                   m += 1
            if groups[k].has_key(0):
                line = groups[k][0]
                months[0] =  line# include 0 itself
                last_month = int(line.date_to.split('-')[1])
        
            print last_month, t, m
            # create budget lines
            for i in range(last_month):
                month = i + 1
                if months.has_key(month):
                    line = months[month]
                    planned_amount = line.planned_amount
                else:
                    line = months[0]
                    planned_amount = (line.planned_amount - t) / (last_month - m)
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
        res_id = super(crossovered_budget, self).write(values)
        self.manage_crossovered_budget_lines()
        return res_id


class crossovered_budget_lines(models.Model):
    _inherit = 'crossovered.budget.lines'

    budget_manager_line_id = fields.Many2one('budget_manager.line')
    segment_id = fields.Many2one(related='crossovered_budget_id.segment_id', store=True)
    segment = fields.Char(related='crossovered_budget_id.segment', store=True)

    def _prac_amt(self, cr, uid, ids, context=None):
        # TODO: remove old segment dependency
        res = {}
        result = 0.0
        if context is None:
            context = {}
        account_obj = self.pool.get('account.account')
        for line in self.browse(cr, uid, ids, context=context):
            acc_ids = [x.id for x in line.general_budget_id.account_ids]
            if not acc_ids:
                raise osv.except_osv(_('Error!'),_("The Budget '%s' has no accounts!") % ustr(line.general_budget_id.name))
            acc_ids = account_obj._get_children_and_consol(cr, uid, acc_ids, context=context)
            date_to = line.date_to
            date_from = line.date_from
            segment_id = line.segment_id
            # get lower segments (one level)
            segment_tmpl_ids = []
            segment_tmpl_ids += segment_id.segment_tmpl_id.get_direct_childs_ids()
            segment_ids = self.pool.get('analytic_segment.segment').search(cr, uid, [('segment_tmpl_id', 'in', segment_tmpl_ids)])

            if line.analytic_account_id.id:
                SQL = """
                SELECT SUM(amount) 
                FROM account_analytic_line as a
                LEFT JOIN account_move_line as l ON l.id = a.move_id
                LEFT JOIN account_move as m ON m.id = l.move_id
                WHERE a.account_id = %s
                    AND (a.date between to_date(%s, 'yyyy-mm-dd')
                        AND to_date(%s, 'yyyy-mm-dd')) 
                    AND a.general_account_id = ANY(%s)
                    AND m.segment_id = ANY(%s) 
                """
                cr.execute(SQL, (line.analytic_account_id.id, date_from, date_to,acc_ids, segment_ids))
                result = cr.fetchall()[0]
            if result is None:
                result = 0.00
            res[line.id] = result[0]
        return res

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
