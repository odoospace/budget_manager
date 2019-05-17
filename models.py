# -*- coding: utf-8 -*-
from openerp import models, fields, api, osv, _
import openerp.addons.decimal_precision as dp
import calendar
from datetime import date, datetime
import json

class crossovered_budget(models.Model):
    _inherit = 'crossovered.budget'

    @api.multi
    def export_xlsxwizard(self, context=None):
        # https://stackoverflow.com/questions/30180535/odoo-8-launch-a-ir-actions-act-window-from-a-python-method
        self.ensure_one()
        if context is None: context = {}
        context['default_budget_id'] = self.id # !important
        res = {
            'name':"Export to Excel",
            'view_mode': 'form',
            'view_type': 'form',
            'res_model': 'budget_manager.xlsxwizard',
            'type': 'ir.actions.act_window',
            'nodestroy': True,
            'target': 'new',
            'domain': '[]',
            'context': context
        }
        return res

    def _domain_segment(self):
        if self.env.user.id == 1:
            domain = []
        else:
            segment_by_company_open = json.loads(self.env.user.segment_by_company_open)[str(self.env.user.company_id.id)]
            domain = [('id', 'in', segment_by_company_open)]
        return domain

    def _search_segment_user(self, operator, value):
        user = self.env['res.users'].browse(self.env.context['user'])
        segment_by_company = json.loads(user.segment_by_company)[str(user.company_id.id)]
        res = [('segment_id', 'in', segment_by_company)]
        return res

    @api.multi
    def _segment_user_id(self):
        # TODO: use a helper in analytic_segment if it's possible...
        if self.env.user.id == 1:
            for obj in self:
                obj.segment_user_id = self.env.uid
            return
        else:
            for obj in self:
                segment_by_company = json.loads(self.env.user.segment_by_company)[str(self.env.user.company_id.id)]
                if obj.segment_id in segment_by_company:
                    obj.segment_user_id = self.env.uid
            return


    budget_manager_line_ids = fields.One2many('budget_manager.line', 'crossovered_budget_id')
    segment_id = fields.Many2one('analytic_segment.segment', required=True,
        domain=_domain_segment)
    segment = fields.Char(related='segment_id.segment', readonly=True)
    segment_user_id = fields.Many2one('res.users', compute='_segment_user_id', search=_search_segment_user)
    zero_incoming = fields.Boolean(default=False)
    with_children = fields.Boolean(default=False)
    category = fields.Selection([
        ('CCE', 'Estatal (CCE)'),
        ('CCA', 'Autonómico (CCA)'),
        ('CCM', 'Municipal (CCM)')
    ])
    group_ids = fields.Many2many('crossovered.budget.group', 'budget_ids')


    def manage_crossovered_budget_lines(self):
        # TODO: use datetime manipulation functions
        # remove old budget lines
        for line in self.crossovered_budget_line:
            line.unlink()
        # recreate budget lines
        # groups = {
        #     (line.analytic_account_id.id, line.general_budget_id.id): {
        #         'main_line': ,
        #         'month_from': ,
        #         'month_to':,
        #         'total_amount_acc': ,
        #          '1', '2', '3'...
        #     }
        # }
        groups = {}
        # get year from first line
        # TODO: add fiscal year or so to general_budget_id
        if self.budget_manager_line_ids:
            year = int(self.budget_manager_line_ids[0].date_from.split('-')[0])
        else:
            year = None
        for line in self.budget_manager_line_ids:
            key = (line.analytic_account_id.id, line.general_budget_id.id)
            # start with an empty dir 
            if not groups.has_key(key):
                groups[key] = {
                    'total_amount_acc': 0,
                    'total_months_acc': 0,
                    'main_line': None,
                    'months': []
                }
            month_from = int(line.date_from.split('-')[1])
            month_to = int(line.date_to.split('-')[1])
            # TODO: add more periods longer 1 month
            if month_from == month_to:
                groups[key][month_from] = line
                groups[key]['total_amount_acc'] += line.planned_amount
                groups[key]['total_months_acc'] += 1
                groups[key]['months'] = list(set(groups[key]['months']+[month_from]))

            else:
                # TODO: check for two or more months lines...
                groups[key]['main_line'] = line
                groups[key]['months'] = list(set(groups[key]['months']+range(month_from, month_to+1)))

        for v in groups.values():
            # create budget lines
            m = max(v['months']) - min(v['months']) + 1
            for month in v['months']:
                if v.has_key(month):
                    line = v[month]
                    planned_amount = line.planned_amount
                else:
                    line = v['main_line']
                    planned_amount = (line.planned_amount - v['total_amount_acc']) / (m - v['total_months_acc'])
                # TODO: dynamic year
                first_day = 1
                last_day = calendar.monthrange(year, month)[1]
                values = {
                    'crossovered_budget_id': self.id,
                    'budget_manager_line_id': line.id,
                    'analytic_account_id': line.analytic_account_id.id,
                    'general_budget_id': line.general_budget_id.id,
                    'date_from': date(year, month, first_day),
                    'date_to':  date(year, month, last_day),
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
    analytic_line_ids = fields.Many2many('account.analytic.line', compute='_analytic_line_ids', readonly=True)
    analytic_line_counter = fields.Integer(compute='_analytic_line_counter', default=0, readonly=True, string="Lines")

    @api.one
    @api.depends('analytic_line_ids')
    def _analytic_line_counter(self):
        for obj in self:
            obj.analytic_line_counter = len(obj.analytic_line_ids)
            

    @api.one
    def _analytic_line_ids(self, context=None):
        res = {}
        _acc_ids = {}
        if context is None:
            context = {}
        for obj in self:
            account_obj = self.pool.get('account.account')
            result = 0.0
            acc_ids = [x.id for x in obj.general_budget_id.account_ids]
            if not acc_ids:
                raise osv.except_osv(_('Error!'),_("The Budget '%s' has no accounts!") % str(obj.general_budget_id.name))
            if not str(acc_ids) in _acc_ids:
                acc_ids_all = account_obj._get_children_and_consol(self._cr, self._uid, acc_ids)
                _acc_ids[str(acc_ids)] = acc_ids_all
            else:
                acc_ids_all = _acc_ids[str(acc_ids)]
            
            date_to = obj.date_to
            date_from = obj.date_from
            segment_id = obj.segment_id
            segment_ids = [segment_id.id]

            if obj.analytic_account_id.id:
                SQL = """
                SELECT a.id, a.name, a.amount 
                FROM account_analytic_line as a
                INNER JOIN account_move_line as l ON l.id = a.move_id
                INNER JOIN account_move as m ON m.id = l.move_id
                WHERE a.account_id = %s
                    AND (a.date between to_date(%s, 'yyyy-mm-dd')
                        AND to_date(%s, 'yyyy-mm-dd')) 
                    AND a.general_account_id = ANY(%s)
                    AND m.segment_id = ANY(%s)
                """
                #print SQL, line.analytic_account_id.id, date_from, date_to, acc_ids, segment_ids
                # TODO: add more lower leves (childs of childs)
                #analytic_ids = self.pool.get('account.analytic.account').search(cr, uid, [('parent_id', '=', line.analytic_account_id.id)])
                #analytic_ids += [line.analytic_account_id.id]
                self.env.cr.execute(SQL, (obj.analytic_account_id.id, date_from, date_to, acc_ids_all, segment_ids))
                _result = self.env.cr.fetchall()
                #for i in _result:
                #    print  i
                if _result:
                    self.analytic_line_ids = self.env['account.analytic.line'].browse([i[0] for i in _result])
                    #self.analytic_line_counter = len(_result)
        
    
    def _prac_amt(self, cr, uid, ids, context=None):
        res = {}
        _acc_ids = {}
        if context is None:
            context = {}
        account_obj = self.pool.get('account.account')
        for line in self.browse(cr, uid, ids, context=context):
            result = 0.0
            acc_ids = [x.id for x in line.general_budget_id.account_ids]
            if not acc_ids:
                raise osv.except_osv(_('Error!'),_("The Budget '%s' has no accounts!") % ustr(line.general_budget_id.name))
            if not str(acc_ids) in _acc_ids:
                acc_ids_all = account_obj._get_children_and_consol(cr, uid, acc_ids, context=context)
                _acc_ids[str(acc_ids)] = acc_ids_all
            else:
                acc_ids_all = _acc_ids[str(acc_ids)]
            
            date_to = line.date_to
            date_from = line.date_from
            segment_id = line.segment_id
            # get lower segments (one level)
            #segment_tmpl_ids = [segment_id.id]
            #segment_tmpl_ids += segment_id.segment_tmpl_id.get_direct_childs_ids()
            #segment_ids = self.pool.get('analytic_segment.segment').search(cr, uid, [('segment_tmpl_id', 'in', segment_tmpl_ids)])
            segment_ids = [segment_id.id]

            if line.analytic_account_id.id:
                SQL = """
                SELECT a.id, a.name, a.amount 
                FROM account_analytic_line as a
                INNER JOIN account_move_line as l ON l.id = a.move_id
                INNER JOIN account_move as m ON m.id = l.move_id
                WHERE a.account_id = %s
                    AND (a.date between to_date(%s, 'yyyy-mm-dd')
                        AND to_date(%s, 'yyyy-mm-dd')) 
                    AND a.general_account_id = ANY(%s)
                    AND m.segment_id = ANY(%s)
                """
                #print SQL, line.analytic_account_id.id, date_from, date_to, acc_ids, segment_ids
                # TODO: add more lower leves (childs of childs)
                #analytic_ids = self.pool.get('account.analytic.account').search(cr, uid, [('parent_id', '=', line.analytic_account_id.id)])
                #analytic_ids += [line.analytic_account_id.id]
                cr.execute(SQL, (line.analytic_account_id.id, date_from, date_to, acc_ids_all, segment_ids))
                _result = cr.fetchall()
                #for i in _result:
                #    print  i
                result = sum([i[2] for i in _result])
                if result is None:
                    result = 0.0

            res[line.id] = result
        #print res
        return res
    
    def _prac(self, cr, uid, ids, name, args, context=None):
        res={}
        #for line in self.browse(cr, uid, ids, context=context):
        res = self._prac_amt(cr, uid, ids, context=context)
        return res

class budget_manager_line(models.Model):
    _name = 'budget_manager.line'
    _description = "Budget Line Managed"

    @api.onchange('analytic_account_id')
    def _domain_budget_line(self):
        segment_tmpl_ids = []
        res = {} 
        budget = self.crossovered_budget_id
        domain = [('state', '=', 'open'), ('company_id', '=', budget.company_id.id)]
        segment_tmpl_ids += [budget.segment_id.segment_tmpl_id.id]
        segment_tmpl_ids += budget.segment_id.segment_tmpl_id.get_childs_ids()
        virtual_segments = self.env['analytic_segment.template'].search([('virtual', '=', True)])
        segment_tmpl_ids += [i.id for i in virtual_segments]
        segment_ids = self.env['analytic_segment.segment'].search([('segment_tmpl_id', 'in', segment_tmpl_ids)])
        domain += [('segment_id', 'in', [i.id for i in segment_ids])]
        res['domain'] = {
            'analytic_account_id': domain
        }
        return res
    
    def _default_crossovered_budget_id(self):
        return self.env.context.get('budget_id', False)

    crossovered_budget_id = fields.Many2one('crossovered.budget', 'Budget', default=_default_crossovered_budget_id, ondelete='cascade', required=True)
    crossovered_budget_line_ids = fields.One2many('crossovered.budget.lines', 'budget_manager_line_id')
    analytic_account_id = fields.Many2one('account.analytic.account', 'Analytic Account')
    general_budget_id = fields.Many2one('account.budget.post', 'Budgetary Position', required=True)
    date_from = fields.Date('Start Date', required=True)
    date_to = fields.Date('End Date', required=True)
    planned_amount = fields.Float('Planned Amount', required=True) #, digits=dp.get_precision('Account')),
    company_id = fields.Many2one(related = 'crossovered_budget_id.company_id', store=True)


class crossovered_budget_group(models.Model):
    _name = 'crossovered.budget.group'

    @api.multi
    def export_xlsxwizard(self, context=None):
        # https://stackoverflow.com/questions/30180535/odoo-8-launch-a-ir-actions-act-window-from-a-python-method
        self.ensure_one()
        if context is None: context = {}
        context['default_budget_id'] = self.id # !important
        res = {
            'name':"Export to Excel",
            'view_mode': 'form',
            'view_type': 'form',
            'res_model': 'budget_manager.group.xlsxwizard',
            'type': 'ir.actions.act_window',
            'nodestroy': True,
            'target': 'new',
            'domain': '[]',
            'context': context
        }
        return res

    name = fields.Char()
    budget_ids = fields.Many2many('crossovered.budget', 'group_ids')
    