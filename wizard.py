# -*- coding: utf-8 -*-
from openerp import models, fields, api
from pprint import pprint
from xlsxwriter.utility import xl_range, xl_rowcol_to_cell
from datetime import datetime
from copy import copy
from calendar import monthrange
import StringIO
import string
import xlsxwriter
import base64
import collections
import time

# TODO: more general use...
G = {
    False: 'No asignado',
    'A': 'Gastos Secretarias',
    'B': 'Gastos Areas y Equipos',
    'C': 'Gastos Generales',
    'D': 'Ingresos',
    'E': 'Otros',
    'F': 'Gastos Extraordinarios'
}

class XLSXWizard(models.TransientModel):
    _name = 'budget_manager.xlsxwizard'

    budget_id = fields.Many2one('crossovered.budget', string="Budget", required=True)
    date_from = fields.Date(required=True)
    date_to = fields.Date(required=True)
    incoming_bypass = fields.Boolean(default=False)

    def process_data(self, date_from, date_to):
        # prepare groups
        _acc_ids = {}
        groups = {}
        X = []
        XX = []

        last_day = monthrange(date_to.year, date_to.month)[1]
        self.date_from = datetime(date_from.year, date_from.month, 1).strftime('%Y-%m-%d')
        self.date_to = datetime(date_to.year, date_to.month, last_day).strftime('%Y-%m-%d')

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
                    X.append(first_parent.upper().strip())
                # Data
                if not groups.has_key(group):
                    groups[group] = {}
                if not groups[group].has_key(account_name):
                    groups[group][account_name] = {}
                if not groups[group][account_name].has_key(first_parent):
                    groups[group][account_name][first_parent] = []

                groups[group][account_name][first_parent].append((1,line))

        # get data from analytic to prepare virtual groups
        _search = [
            ('date', '>=', self.date_from),
            ('date', '<=', self.date_to),
            ('company_id', '=', self.budget_id.company_id.id)
        ]
        # first segment
        segment_ids = [self.budget_id.segment_id.id]
        
        if self.budget_id.with_children:
            segment_ids += self.budget_id.segment_id.segment_tmpl_id.get_childs_ids()
        
        _search.append(('segment_id', 'in', segment_ids))

        _anaylitic_lines = self.env['account.analytic.line'].sudo().search(_search)
        analytic_lines = []
        analytic_lines_obj = []
        for line in _anaylitic_lines:
            code0 = line.general_account_id.code[0]
            code1 = line.general_account_id.code[:2]
            if (code0 in ['6', '7'] and code1 != '68') or code1 in ['20', '21']:
                analytic_lines.append(line.id)
                analytic_lines_obj.append(line)

        account_obj = self.pool.get('account.account')
        print 'lines:', len(self.budget_id.crossovered_budget_line)
        for line in self.budget_id.crossovered_budget_line:
            acc_ids = [x.id for x in line.general_budget_id.account_ids]

            if not acc_ids:
                raise osv.except_osv(_('Error!'),_("The Budget '%s' has no accounts!") % ustr(line.general_budget_id.name))
            #print 'acc_ids/1', len(acc_ids)
            if not str(acc_ids) in _acc_ids:
                acc_ids_all = account_obj._get_children_and_consol(self._cr, self._uid, acc_ids)
                _acc_ids[str(acc_ids)] = acc_ids_all
            else:
                acc_ids_all = _acc_ids[str(acc_ids)]

            date_to = line.date_to
            date_from = line.date_from
            segment_id = line.segment_id
            # get lower segments (one level)
            #segment_tmpl_ids = [segment_id.id]
            #segment_tmpl_ids += segment_id.segment_tmpl_id.get_direct_childs_ids()
            #segment_ids = [i.id for i in self.env['analytic_segment.segment'].search([('segment_tmpl_id', 'in', segment_tmpl_ids)])]
            segment_ids = [segment_id.id]

            if line.analytic_account_id.id:
                SQL = """
                SELECT a.id, a.amount
                FROM account_analytic_line as a
                INNER JOIN account_move_line as l ON l.id = a.move_id
                INNER JOIN account_move as m ON m.id = l.move_id
                WHERE a.account_id = %s
                    AND (a.date between to_date(%s, 'yyyy-mm-dd')
                        AND to_date(%s, 'yyyy-mm-dd'))
                    AND a.general_account_id = ANY(%s)
                    AND m.segment_id = ANY(%s)
                """
                
                #_z = line.analytic_account_id.id, date_from, date_to, list(acc_ids), list(segment_ids)
                #print 'params:', line.analytic_account_id.id, date_from, date_to, acc_ids, segment_ids 
                self.env.cr.execute(SQL, (line.analytic_account_id.id, date_from, date_to, acc_ids_all, segment_ids))
                
                result = self.env.cr.fetchall()

                for res in result:
                    #print res
                    if res[0] in analytic_lines:
                        #print 'removed!!!'
                        analytic_lines.remove(res[0])

        #print len(anaylitic_lines)
        lines = self.env['account.analytic.line'].sudo().search([
            ('id', 'in', analytic_lines)
        ])
        for l in lines:
            level = l.account_id.level
            if level <= 2:
                group = G[l.account_id.group]
                account_name = l.account_id.name
            elif level == 3:
                group = G[l.account_id.parent_id.group]
                account_name = l.account_id.parent_id.name
            first_parent = l.account_id.first_parent().name

            # Matrix
            # append parents at X
            if first_parent not in X:
                X.append(first_parent.upper().strip())
            # Data
            if not groups.has_key(group):
                groups[group] = {}
            if not groups[group].has_key(account_name):
                groups[group][account_name] = {}
            if not groups[group][account_name].has_key(first_parent):
                groups[group][account_name][first_parent] = []
            #print group, account_name, first_parent
            groups[group][account_name][first_parent].append((2, l))

        # reordered X
        if not self.budget_id.segment_id.is_campaign:
            refs = [
                'SALARIOS', 'MATERIALES', 'SERVICIOS EXTERNOS',
                'DESPLAZAMIENTOS', 'SUSCRIPCIONES - LICENCIAS', 'OTROS'
            ]
        else:
            refs = ['GASTOS']
    
        for item in refs:
            if item in X:
                XX.append(item)
        
        return X, XX, groups, analytic_lines, analytic_lines_obj


    @api.one
    def run_export_xlsx(self):
        # adjust dates internally
        # TODO: use odoo stuff for dates

        _date_to = self.date_to
        _date_from = self.date_from
        date_from = datetime.strptime(self.date_from, '%Y-%m-%d').date()
        date_to = datetime.strptime(self.date_to, '%Y-%m-%d').date()

        X, XX, groups, analytic_lines, analytic_lines_obj = self.process_data(date_from, date_to)

        # Create an new Excel file and add a worksheet
        # https://www.odoo.com/es_ES/forum/ayuda-1/question/return-an-excel-file-to-the-web-client-63980
        xlsxfile = StringIO.StringIO()
        workbook = xlsxwriter.Workbook(xlsxfile, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        worksheet.freeze_panes(1, 1) # freeze first column and first row

        # styles
        _money = workbook.add_format({'num_format': '#,##0.00'})
        _porcentage = workbook.add_format({'num_format': '#,##0.00"%"', 'bg_color': '#92ff96'})
        _bold = workbook.add_format({'bold': True})
        _bold_center = workbook.add_format({'bold': True, 'align': 'center'})
        _yellow = workbook.add_format({'bold': True, 'bg_color': 'yellow'})
        _silver_money = workbook.add_format({'bg_color': '#D0D0D0', 'num_format': '#,##0.00'})
        _silver_bold_center = workbook.add_format({'bold': True, 'bg_color': '#D0D0D0', 'align': 'center'})
        _gray = workbook.add_format({'bold': True, 'bg_color': 'silver'})
        _gray_money = workbook.add_format({'bold': True, 'bg_color': 'silver', 'num_format': '#,##0.00'})
        _gray_porcentage = workbook.add_format({'bold': True, 'bg_color': 'silver', 'num_format': '#,##0.00"%"'})
        _purple = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'purple', 'font_color': 'white'})
        _red = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'red', 'font_color': 'white'})
        _red_total = workbook.add_format({'bold': True, 'bg_color': 'red', 'font_color': 'white'})
        _red_money = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'red', 'font_color': 'white', 'num_format': '#,##0.00'})
        _red_porcentage = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': 'red', 'font_color': 'white', 'num_format': '#,##0.00"%"'})
        _blue = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white'})
        _blue_money = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white', 'num_format': '#,##0.00'})
        _blue_porcentage = workbook.add_format({'bold': True, 'bg_color': 'blue', 'font_color': 'white', 'num_format': '#,##0.00"%"'})

        # header
        y = 0
        x = 1
        worksheet.set_column(0, 0, 40)
        date_from = datetime.strptime(_date_from, '%Y-%m-%d').strftime('%d/%m/%y')
        date_to = datetime.strptime(_date_to, '%Y-%m-%d').strftime('%d/%m/%y')
        name = '%s (%s - %s)' % (self.budget_id.name, date_from, date_to)
        worksheet.write(y, 0, name, _yellow)
        for column in XX:
            worksheet.set_column(x, x+1, 12)
            worksheet.merge_range(y, x, y, x+1, column, _purple)
            x += 2
        x_total = x
        worksheet.set_column(x, x+1, 12)
        worksheet.merge_range(y, x, y, x+1, 'TOTAL', _red)
        worksheet.write(y, x+2, 'DESV.', _red)

        # process data
        y += 1
        #for row in groups:
        _groups = [
            'Gastos Secretarias', 'Gastos Areas y Equipos',
            'Gastos Generales', 'Gastos Extraordinarios', 'Otros', 'No asignado']
        for row in _groups:
            if not groups.has_key(row):
                continue
            worksheet.write(y, 0, row.upper(), _bold)
            x = 1
            for column in XX:
                worksheet.write(y, x, 'PRESUP.', _bold_center)
                worksheet.write(y, x+1, 'REALES', _bold_center)
                x += 2
            worksheet.write(y, x, 'PRESUP.', _bold_center)
            worksheet.write(y, x+1, 'REALES', _bold_center)
            worksheet.write(y, x+2, '%', _bold_center)
            y += 1
            y0 = y
            # budget data
            _lines = collections.OrderedDict(sorted(groups[row].items()))
            for line in _lines:
                worksheet.write(y, 0, line.upper())
                # TODO: remove columns ?
                #columns = groups[row][line]
                # do sums
                for i, column in enumerate(XX):
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
                    worksheet.write(y, x+1, practical_amount, _silver_money)
                # add X totals (red)
                cell_range_planned = ''
                cell_range_practical = ''
                for i, column in enumerate(XX):
                    x0 = i * 2 + 1
                    cell_range_planned += '%s+' % xl_rowcol_to_cell(y, x0)
                    cell_range_practical += '%s+' % xl_rowcol_to_cell(y, x0+1)
                worksheet.write_formula(y, x+2, '=%s' % cell_range_planned[:-1], _money)
                worksheet.write_formula(y, x+3, '=%s' % cell_range_practical[:-1], _silver_money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x+2)
                cell_practical = xl_rowcol_to_cell(y, x+3)
                worksheet.write_formula(y, x+4, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _porcentage)
                y += 1
            # add Y total
            worksheet.write(y, 0, 'TOTAL %s' % row.upper(), _gray)
            for i in range(len(XX)+1): # add X total column too
                x = i * 2 + 1
                cell_range = xl_range(y0, x, y-1, x)
                worksheet.write_formula(y, x, '=SUM(%s)' % cell_range, _gray_money)
                cell_range = xl_range(y0, x+1, y-1, x+1)
                worksheet.write_formula(y, x+1, '=SUM(%s)' % cell_range, _gray_money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x)
                cell_practical = xl_rowcol_to_cell(y, x+1)
                worksheet.write_formula(y, x+2, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _gray_porcentage)
            y += 1

        # total
        y += 1 # empty line
        worksheet.write(y, 0, 'TOTAL GENERAL', _blue)
        for i in range(len(XX)+1):
            x = i * 2 + 1
            cell_range = xl_range(1, x, y-1, x)
            worksheet.write_formula(y, x, '=SUM(%s)/2' % cell_range, _blue_money)
            cell_range = xl_range(1, x+1, y-1, x+1)
            worksheet.write_formula(y, x+1, '=SUM(%s)/2' % cell_range, _blue_money)
            # add %
            cell_planned = xl_rowcol_to_cell(y, x)
            cell_practical = xl_rowcol_to_cell(y, x+1)
            worksheet.write_formula(y, x+2, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _blue_porcentage)

        y_total = y
        y += 2

        print '-->', y, x_total-3, x_total-1

        # special INCOMING part
        # TODO: refactorize this!
        row = 'Ingresos'
        
        if not groups.has_key(row) and (self.incoming_bypass or self.budget_id.zero_incoming):
            groups[row] = {}
        
        worksheet.merge_range(y, x_total-3, y, x_total-1, row.upper(), _yellow )
        #worksheet.set_column(x_total-1, x_total-1, 40)
        #worksheet.write(y, x_total-1, row.upper(), _yellow)
        worksheet.merge_range(y, x_total, y, x_total+1, 'TOTAL', _red)
        worksheet.write(y, x_total+2, 'DESV.', _red)
        y += 1
        y_income = y
        # budget data
        _lines = collections.OrderedDict(sorted(groups[row].items()))
        for line in _lines:
            worksheet.merge_range(y, x_total-3, y, x_total-1, line.upper())
            # TODO: remove columns ?
            # columns = groups[row][line]
            # do sums
            column = 'INGRESOS'
            planned_amount = 0
            practical_amount = 0
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
            worksheet.write(y, x_total, planned_amount, _money)
            worksheet.write(y, x_total+1, practical_amount, _silver_money)
            # add %
            cell_planned = xl_rowcol_to_cell(y, x_total)
            cell_practical = xl_rowcol_to_cell(y, x_total+1)
            worksheet.write_formula(y, x+2, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _porcentage)
            y += 1

        # total 'Ingresos'
        y += 1 # empty line
        worksheet.merge_range(y, x_total-3, y, x_total-1, 'TOTAL %s' % row.upper(), _blue)
        cell_range = xl_range(y_income, x_total, y-1, x_total)
        worksheet.write_formula(y, x_total, '=SUM(%s)' % cell_range, _blue_money)
        cell_range = xl_range(y_income, x_total+1, y-1, x_total+1)
        worksheet.write_formula(y, x_total+1, '=SUM(%s)' % cell_range, _blue_money)
        # add %
        cell_planned = xl_rowcol_to_cell(y, x_total)
        cell_practical = xl_rowcol_to_cell(y, x_total+1)
        worksheet.write_formula(y, x_total+2, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _blue_porcentage)

        # supertotal!
        y += 2 # empty line
        worksheet.merge_range(y, x_total-3, y, x_total-1, 'REMANENTE', _red_total)
        cell_range_planned = '%s+%s' % (xl_rowcol_to_cell(y_total, x_total), xl_rowcol_to_cell(y-2, x_total))
        cell_range_practical = '%s+%s' % (xl_rowcol_to_cell(y_total, x_total+1), xl_rowcol_to_cell(y-2, x_total+1))
        worksheet.write_formula(y, x_total, '=%s' % cell_range_planned, _red_money)
        worksheet.write_formula(y, x_total+1, '=%s' % cell_range_practical, _red_money)
        # add %
        cell_planned = xl_rowcol_to_cell(y, x_total)
        cell_practical = xl_rowcol_to_cell(y, x_total+1)
        worksheet.write_formula(y, x_total+2, '=(%s/%s-1)*100' % (cell_practical, cell_planned), _red_porcentage)


        # new worksheet with analytic lines!!!
        worksheet_lines = workbook.add_worksheet()
        y = 0

        # headers
        worksheet_lines.set_column(0, 0, 12)
        worksheet_lines.write(y, 0, 'Date', _gray)
        worksheet_lines.set_column(1, 1, 30)
        worksheet_lines.write(y, 1, 'Move', _gray)
        worksheet_lines.set_column(2, 2, 30)
        worksheet_lines.write(y, 2, 'Analytic-1', _gray)
        worksheet_lines.set_column(3, 3, 30)
        worksheet_lines.write(y, 3, 'Analytic-2', _gray)
        worksheet_lines.set_column(4, 4, 30)
        worksheet_lines.write(y, 4, 'Analytic-3', _gray)
        worksheet_lines.set_column(5, 5, 15)
        worksheet_lines.write(y, 5, 'Segment', _gray)
        worksheet_lines.set_column(6, 6, 14)
        worksheet_lines.write(y, 6, 'Segment Code', _gray)
        worksheet_lines.set_column(7, 7, 40)
        worksheet_lines.write(y, 7, 'Account Code', _gray)
        worksheet_lines.set_column(8, 8, 40)
        worksheet_lines.write(y, 8, 'Account', _gray)
        worksheet_lines.set_column(9, 9, 40)
        worksheet_lines.write(y, 9, 'Partner', _gray)
        worksheet_lines.set_column(10, 10, 70)
        worksheet_lines.write(y, 10, 'Description', _gray)
        worksheet_lines.set_column(11, 11, 12)
        worksheet_lines.write(y, 11, 'Amount', _gray)


        worksheet_lines.freeze_panes(1, 0) # freeze first row
        y +=1

        # data
        for line in sorted(analytic_lines_obj, key=lambda x: x.date):
            worksheet_lines.write(y, 0, line.date)
            worksheet_lines.write(y, 1, line.move_id.move_id.name)
            if line.account_id.parent_id.parent_id: # level 3
                worksheet_lines.write(y, 2, line.account_id.parent_id.parent_id.name)
                worksheet_lines.write(y, 3, line.account_id.parent_id.name)
                worksheet_lines.write(y, 4, line.account_id.name)
            elif line.account_id.parent_id: # level 2
                worksheet_lines.write(y, 2, line.account_id.parent_id.name)
                worksheet_lines.write(y, 3, line.account_id.name)
            else: # level 1
                worksheet_lines.write(y, 2, line.account_id.name)
            worksheet_lines.write(y, 5, line.account_id.segment)
            worksheet_lines.write(y, 6, line.move_id.segment)
            worksheet_lines.write(y, 7, line.general_account_id.code)
            worksheet_lines.write(y, 8, line.general_account_id.name)
            worksheet_lines.write(y, 9, line.move_id.partner_id and line.move_id.partner_id.name or '')
            worksheet_lines.write(y, 10, line.name)
            worksheet_lines.write(y, 11, line.amount, _money)
            y += 1

        # close it
        workbook.close()

        # Rewind the buffer.
        xlsxfile.seek(0)
        name = self.budget_id.name.lower().replace(' ', '_')
        vals = {
            'name': 'presupuesto_%s_%s_%s.xlsx' % (name, _date_from, _date_to),
            'datas': base64.encodestring(xlsxfile.read()),
            'datas_fname': 'presupuesto_%s_%s_%s.xlsx' % (name, _date_from, _date_to),
            'res_model': self.budget_id._name,
            'res_id': self.budget_id.id,
            'type': 'binary'
        }
        attachment_id = self.env['ir.attachment'].create(vals)

        return True

