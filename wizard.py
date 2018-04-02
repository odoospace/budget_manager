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
        # first pass: group by account type and first parent
        groups = {}
        X = []
        Y = {}
        for line in self.budget_id.crossovered_budget_line:
            # check dates
            if line.date_from >= self.date_from and line.date_to <= self.date_to:
                group = G[line.analytic_account_id.group]
                first_parent = line.analytic_account_id.first_parent().name
                name = line.analytic_account_id.name
                # X/Y Matrix
                # init Y group
                if not Y.has_key(group):
                    Y[group] = []
                # append accounts at Y
                if not name in Y[group]:
                    Y[group].append(name)
                # append parents at X
                if first_parent not in X:
                    X.append(first_parent)
                # Data
                if not groups.has_key(group):
                    groups[group] = {}
                if not groups[group].has_key(name):
                    groups[group][name] = {}
                if not groups[group][name].has_key(first_parent):
                    groups[group][name][first_parent] = []
                    
                groups[group][name][first_parent].append(line)
        
        
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
        
        # data
        y += 1
        for row in groups:
            worksheet.write(y, 0, row.upper(), _bold)
            y += 1
            y0 = y
            for line in groups[row]:
                worksheet.write(y, 0, line.upper())
                columns = groups[row][line]
                for i, column in enumerate(X):
                    planned_amount = 0
                    practical_amount = 0
                    x = i * 2 + 1
                    if groups[row][line].has_key(column):
                        for budget_line in groups[row][line][column]:
                            planned_amount += budget_line.planned_amount
                            practical_amount += budget_line.practical_amount                        
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
            'name': 'presupuesto.xlsx',
            'datas': base64.encodestring(xlsxfile.read()),
            'datas_fname': 'presupuesto.xlsx',
            'res_model': self.budget_id._name,
            'res_id': self.budget_id.id,
            'type': 'binary'
        }
        attachment_id = self.env['ir.attachment'].create(vals)

        return True

