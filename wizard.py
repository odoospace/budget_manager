# -*- coding: utf-8 -*-
from openerp import models, fields, api
from pprint import pprint
from xlsxwriter.utility import xl_range, xl_rowcol_to_cell
from datetime import datetime
import StringIO
import string
import xlsxwriter
import base64
# TODO: more general use...
G = {
    False: 'No asignado',
    'A': 'Gastos Secretarias',
    'B': 'Gastos Areas y Equipos',
    'C': 'Gastos Generales',
    'D': 'Ingresos',
    'E': 'Otros'
}

class XLSXWizard(models.TransientModel):
    _name = 'budget_manager.xlsxwizard'

    budget_id = fields.Many2one('crossovered.budget', string="Budget", required=True)
    date_from = fields.Date(required=True)
    date_to = fields.Date(required=True)

    
    @api.one
    def run_export_xlsx(self):
        
        # prepare groups
        groups = {}
        X = []
        # get data from budget lines
        for line in self.budget_id.crossovered_budget_line:
            # check dates
            if line.date_from >= self.date_from and line.date_to <= self.date_to:
                group = G[line.analytic_account_id.group]
                first_parent = line.analytic_account_id.first_parent().name
                account_name = line.analytic_account_id.name
                # Matrix
                # append parents at X
                if first_parent not in X:
                    X.append(first_parent)
                # Data
                if not groups.has_key(group):
                    groups[group] = {}
                if not groups[group].has_key(account_name):
                    groups[group][account_name] = {}
                if not groups[group][account_name].has_key(first_parent):
                    groups[group][account_name][first_parent] = []
                    
                groups[group][account_name][first_parent].append((1,line))
        
        
        # get data from analytic to prepare virtual groups
        anaylitic_lines = [i.id for i in self.env['account.analytic.line'].search([
            ('date', '>=', self.date_from),
            ('date', '<=', self.date_to),
            ('company_id', '=', self.budget_id.company_id.id),
            ('segment_id', '=', self.budget_id.segment_id.id)
        ])]
        
        account_obj = self.pool.get('account.account')
        for line in self.budget_id.crossovered_budget_line:
            acc_ids = [x.id for x in line.general_budget_id.account_ids]
            if not acc_ids:
                raise osv.except_osv(_('Error!'),_("The Budget '%s' has no accounts!") % ustr(line.general_budget_id.name))
            acc_ids = account_obj._get_children_and_consol(self._cr, self._uid, acc_ids)
            date_to = line.date_to
            date_from = line.date_from
            segment_id = line.segment_id
            # get lower segments (one level)
            #segment_tmpl_ids = []
            #segment_tmpl_ids += segment_id.segment_tmpl_id.get_direct_childs_ids()
            #segment_ids = [i.id for i in self.env['analytic_segment.segment'].search([('segment_tmpl_id', 'in', segment_tmpl_ids)])]
            segment_ids = [segment_id.id]
        
            if line.analytic_account_id.id:
                SQL = """
                SELECT a.id, a.amount
                FROM ((account_analytic_line as a
                INNER JOIN account_move_line as l ON l.id = a.move_id)
                INNER JOIN account_move as m ON m.id = l.move_id)
                WHERE a.account_id = %s
                    AND (a.date between to_date(%s, 'yyyy-mm-dd')
                        AND to_date(%s, 'yyyy-mm-dd')) 
                    AND a.general_account_id = ANY(%s)
                    AND m.segment_id = ANY(%s)
                """ 
                #_z = line.analytic_account_id.id, date_from, date_to, list(acc_ids), list(segment_ids)
                self.env.cr.execute(SQL, (line.analytic_account_id.id, date_from, date_to, acc_ids, segment_ids))
                result = self.env.cr.fetchall()
                #print '>>>', result
                for res in result:
                    #print res
                    if res[0] in anaylitic_lines:
                        #print 'removed!!!'
                        anaylitic_lines.remove(res[0])
                
        #print len(anaylitic_lines)
        lines = self.env['account.analytic.line'].search([
            ('id', 'in', anaylitic_lines)
        ])
        for l in lines:
            #print '%s,"%s","%s",%s,"%s",%s,"%s","%s","%s"' % (
            #    l.id, l.date, l.name, l.amount, l.account_id.name,
            #    l.account_id.level, 
            #    l.move_id.name, l.account_id.group,
            #    l.account_id.first_parent().name
            #)
            level = l.account_id.level
            if level <= 2:
                group = G[l.account_id.group]
                first_parent = l.account_id.first_parent().name
                account_name = l.account_id.name
                # Matrix
                # append parents at X
                if first_parent not in X:
                    X.append(first_parent)
                # Data
                if not groups.has_key(group):
                    groups[group] = {}
                if not groups[group].has_key(account_name):
                    groups[group][account_name] = {}
                if not groups[group][account_name].has_key(first_parent):
                    groups[group][account_name][first_parent] = []
                #print group, account_name, first_parent    
                groups[group][account_name][first_parent].append((2, l))

        #stop
        
        # Create an new Excel file and add a worksheet
        # https://www.odoo.com/es_ES/forum/ayuda-1/question/return-an-excel-file-to-the-web-client-63980
        xlsxfile = StringIO.StringIO()
        workbook = xlsxwriter.Workbook(xlsxfile, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        _money = workbook.add_format({'num_format': '#,##0.00'})
        _porcentage = workbook.add_format({'num_format': '#,##0.00"%"'})
        _bold = workbook.add_format({'bold': True})
        _yellow = workbook.add_format({'bold': True, 'bg_color': 'yellow'})
        _gray = workbook.add_format({'bold': True, 'bg_color': 'gray'})
        _gray_money = workbook.add_format({'bold': True, 'bg_color': 'gray', 'num_format': '#,##0.00'})
        _gray_porcentage = workbook.add_format({'bold': True, 'bg_color': 'gray', 'num_format': '#,##0.00"%"'})
        _purple = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'purple', 'font_color': 'white'})
        _red = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'red', 'font_color': 'white'})
        _blue = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white'})
        _blue_money = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white', 'num_format': '#,##0.00'})
        _blue_porcentage = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white', 'num_format': '#,##0.00"%"'}) 
        
        # header
        y = 0
        x = 1
        worksheet.set_column(0, 0, 60)
        date_from = datetime.strptime(self.date_from, '%Y-%m-%d').strftime('%d/%m/%y')
        date_to = datetime.strptime(self.date_to, '%Y-%m-%d').strftime('%d/%m/%y')
        name = '%s (%s - %s)' % (self.budget_id.name, date_from, date_to)
        worksheet.write(y, 0, name, _yellow)
        for column in X:
            worksheet.set_column(x, x+1, 12)
            worksheet.merge_range(y, x, y, x+1, column, _purple)
            x += 2
        worksheet.set_column(x, x+1, 12)
        worksheet.merge_range(y, x, y, x+1, 'TOTAL', _red)
        worksheet.write(y, x+2, 'DESV.', _red)
        
        # process data
        y += 1
        for row in groups:
            worksheet.write(y, 0, row.upper(), _bold)
            y += 1
            y0 = y
            for line in groups[row]:
                worksheet.write(y, 0, line.upper())
                columns = groups[row][line]
                # do sums
                for i, column in enumerate(X):
                    planned_amount = 0
                    practical_amount = 0
                    x = i * 2 + 1
                    if groups[row][line].has_key(column):
                        for ttype, l in groups[row][line][column]:
                            if ttype == 1:
                                # from budget
                                planned_amount += l.planned_amount
                                practical_amount += l.practical_amount
                            elif ttype == 2:
                                # from analytic
                                practical_amount += l.amount
                    # TODO: add euros
                    worksheet.write(y, x, planned_amount, _money)
                    worksheet.write(y, x+1, practical_amount, _money)
                # add X totals (red)
                cell_range_planned = ''
                cell_range_practical = ''
                for i, column in enumerate(X):
                    x0 = i * 2 + 1
                    cell_range_planned += '%s+' % xl_rowcol_to_cell(y, x0)
                    cell_range_practical += '%s+' % xl_rowcol_to_cell(y, x0+1)
                worksheet.write_formula(y, x+2, '{=%s}' % cell_range_planned[:-1], _money)
                worksheet.write_formula(y, x+3, '{=%s}' % cell_range_practical[:-1], _money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x+2)
                cell_practical = xl_rowcol_to_cell(y, x+3)
                worksheet.write_formula(y, x+4, '{=%s/%s*100}' % (cell_practical, cell_planned), _porcentage)
                y += 1
            # add Y total
            worksheet.write(y, 0, 'TOTAL DE %s' % row.upper(), _gray)
            for i in range(len(X)+1): # add X total column too
                x = i * 2 + 1
                cell_range = xl_range(y0, x, y-1, x)
                worksheet.write_formula(y, x, '{=SUM(%s)}' % cell_range, _gray_money)
                cell_range = xl_range(y0, x+1, y-1, x+1)
                worksheet.write_formula(y, x+1, '{=SUM(%s)}' % cell_range, _gray_money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x)
                cell_practical = xl_rowcol_to_cell(y, x+1)
                worksheet.write_formula(y, x+2, '{=%s/%s*100}' % (cell_practical, cell_planned), _gray_porcentage)
            y += 1
            
        # total
        y += 1 # empty line
        worksheet.write(y, 0, 'TOTAL GENERAL', _blue)
        for i in range(len(X)+1):
            x = i * 2 + 1
            cell_range = xl_range(1, x, y-1, x)
            worksheet.write_formula(y, x, '{=SUM(%s)/2}' % cell_range, _blue_money)
            cell_range = xl_range(1, x+1, y-1, x+1)
            worksheet.write_formula(y, x+1, '{=SUM(%s)/2}' % cell_range, _blue_money)
            # add %
            cell_planned = xl_rowcol_to_cell(y, x)
            cell_practical = xl_rowcol_to_cell(y, x+1)
            worksheet.write_formula(y, x+2, '{=%s/%s*100}' % (cell_practical, cell_planned), _blue_porcentage)
        
        # close it 
        workbook.close()
        
        # Rewind the buffer.
        xlsxfile.seek(0)
        vals = {
            'name': 'presupuesto_%s_%s.xlsx' % (date_from, date_to),
            'datas': base64.encodestring(xlsxfile.read()),
            'datas_fname': 'presupuesto_%s_%s.xlsx' % (date_from, date_to),
            'res_model': self.budget_id._name,
            'res_id': self.budget_id.id,
            'type': 'binary'
        }
        attachment_id = self.env['ir.attachment'].create(vals)

        return True

