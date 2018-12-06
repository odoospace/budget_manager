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

# TODO: more general use...
# TODO: translations

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
    _name = 'budget_manager.group.xlsxwizard'

    group_id = fields.Many2one('crossovered.budget.group', string="Group", required=True)
    date_from = fields.Date(required=True)
    date_to = fields.Date(required=True)
    incoming_bypass = fields.Boolean(default=False)


    @api.one
    def run_export_xlsx(self):
        results = {}
        _date_to = self.date_to
        _date_from = self.date_from
        date_from = datetime.strptime(self.date_from, '%Y-%m-%d').date()
        date_to = datetime.strptime(self.date_to, '%Y-%m-%d').date()
        data = {
            'Ingresos': (0, 0),
            'Salarios': (0, 0),
        }

        # predefine columns
        COLUMNS = {
            'Gastos': [],
            'Ingresos': [],
            'Salarios': [],
            # Gastos
            'Unidades Funcionales y CCE': [],
            'Alquiler y Gastos de Oficiana': [],
            'Asignaciones Autonómicas y Municipales': [],
            'Otros': [],
            # Ingresos
            'Aportaciones GP': [],
            'Aportaciones Cargos Públicos': [],
            'Colaboraciones Adscritas': [],
            'Subvenciones': [],
            'Otros': []
        }

        # [level, topic]
        MAPPING = {
            3: {
                'INGRESOS': 'Ingresos'
            },
            2: {
                'Gastos Secretarias': {
                    '*': 'Unidades Funcionales y CCE'
                },
                'Gastos Areas y Equipos': {
                    '*': 'Unidades Funcionales y CCE'
                },
                'Gastos Generales': {
                    'ACTOS VARIOS': 'Unidades Funcionales y CCE',
                    'CONSEJO ESTATAL CCE': 'Unidades Funcionales y CCE',
                    'CONSULTAS CIUDADANAS': 'Unidades Funcionales y CCE',
                    'CONSEJO AUTONOMICO': 'Unidades Funcionales y CCE'
                },
                'Gastos Extraordinarios': {
                    'DESARROLLO PARTICIPA': 'Unidades Funcionales y CCE',
                    'ESTUDIOS DEMOSCOPICOS': 'Unidades Funcionales y CCE',
                    'FONDO ANUAL ACTIVIDADES': 'Unidades Funcionales y CCE',
                    'ORDENADORES': 'Unidades Funcionales y CCE',
                    'PROYECTOS AREAS': 'Unidades Funcionales y CCE',
                    'PROYECTOS SOCIALES': 'Unidades Funcionales y CCE'
                }
            }
        }
        
        total_planned_amount = {}
        total_practical_amount = {}

        for i in self.group_id.budget_ids:
            vals = {
                'budget_id': i.id, 
                'date_from': date_from,
                'date_to': date_to
            }
            budget_wiz = self.env['budget_manager.xlsxwizard'].create(vals)
            # get data from detail budget report
            X, XX, groups, analytic_lines, analytic_lines_obj = budget_wiz.process_data(date_from, date_to)
            results[i.id] = (X, XX, groups, analytic_lines, analytic_lines_obj)
            print i.name, groups.keys()

            # add keys
            total_planned_amount[i] = dict([item, 0] for item in COLUMNS.keys())
            total_practical_amount[i] = dict([item, 0] for item in COLUMNS.keys())
 

            for k1, l1 in groups.items(): # level 1                  
                for k2, l2 in l1.items(): # level 2
                    for k3, l3 in l2.items(): # level 3
                        print 'k1:', k1, 'k2:', k2, 'k3:', k3
                        # check mapping for level 3
                        for m in MAPPING[3]: 
                            if k3 == m:
                                column = MAPPING[3][m]
                                for ttype, v in l3:
                                    print i, column, '...'
                                    if ttype == 1:
                                        total_planned_amount[i][column] += v.planned_amount
                                        total_practical_amount[i][column] += v.practical_amount
                                    else:
                                        total_practical_amount[i][column] += v.amount
                        # check mapping for level 2
                        for m in MAPPING[2]: 
                            for n in MAPPING[2][m]:
                                print '>>>', m, n
                                if (n == '*' and k1 == m) or (k1 == m and k2 == n):
                                    column = MAPPING[2][m][n]
                                    for ttype, v in l3:
                                        print i, column, '...'
                                        if ttype == 1:
                                            total_planned_amount[i][column] += v.planned_amount
                                            total_practical_amount[i][column] += v.practical_amount
                                        else:
                                            total_practical_amount[i][column] += v.amount

                        # check mapping for level 2
                        #for m, n in MAPPINGS[2]:
                        #    for o, p in .items():
                        #        if o == '*' and :
                        #            print n, l3
            stop
            
            #UNIDADES FUNCIONALES
            GROUPS = {
                'Gastos Secretarias': [],
                'Gastos Areas y Equipos': [],
                'Gastos Generales': [
                    'ACTOS VARIOS',
                    'CONSEJO ESTATAL CCE',
                    'CONSULTAS CIUDADANAS',
                    'CONSEJO AUTONOMICO'
                ],
                'Gastos Extraordinarios': [
                    'DESARROLLO PARTICIPA',
                    'ESTUDIOS DEMOSCOPICOS',
                    'FONDO ANUAL ACTIVIDADES',
                    'ORDENADORES',
                    'PROYECTOS AREAS',
                    'PROYECTOS SOCIALES',
                ]
            }

            for k, l1 in groups.items():
                if k in GROUPS:
                    for k2, l2 in l1.items():
                        if (k2 in GROUPS[k]) or not GROUPS[k]: 
                            print k, k2
                            planned_amount = 0
                            practical_amount = 0
                            for l3 in l2.values():
                                for l in l3: 
                                    ttype = l[0]
                                    line = l[1]
                                    # print k
                                    if ttype == 1:
                                        planned_amount += line.planned_amount
                                        practical_amount += line.practical_amount
                                    else:
                                        practical_amount += line.amount
                            print 'planned_amount:', planned_amount
                            print 'practical_amount:', practical_amount
                            total_planned_amount += planned_amount
                            total_practical_amount += practical_amount
            print 'total planned_amount:', total_planned_amount
            print 'total practical_amount:', total_practical_amount


            # INGRESOS
            for j in groups['Ingresos'].values():
                # print j
                for k in j['INGRESOS']:
                    # print k
                    if k[0] == 1:
                        total_planned_amount += k[1].planned_amount
                        total_practical_amount += k[1].practical_amount
                    else:
                        total_practical_amount += k[1].amount
            print 'total planned_amount:', total_planned_amount
            print 'total practical_amount:', total_practical_amount


            #UNIDADES FUNCIONALES
            GROUPS = {
                'Gastos Secretarias': [],
                'Gastos Areas y Equipos': [],
                'Gastos Generales': [
                    'ACTOS VARIOS',
                    'CONSEJO ESTATAL CCE',
                    'CONSULTAS CIUDADANAS',
                    'CONSEJO AUTONOMICO'
                ],
                'Gastos Extraordinarios': [
                    'DESARROLLO PARTICIPA',
                    'ESTUDIOS DEMOSCOPICOS',
                    'FONDO ANUAL ACTIVIDADES',
                    'ORDENADORES',
                    'PROYECTOS AREAS',
                    'PROYECTOS SOCIALES',
                ]
            }

            total_planned_amount = 0
            total_practical_amount = 0
            for k, l1 in groups.items():
                if k in GROUPS:
                    for k2, l2 in l1.items():
                        if (k2 in GROUPS[k]) or not GROUPS[k]: 
                            print k, k2
                            planned_amount = 0
                            practical_amount = 0
                            for l3 in l2.values():
                                for l in l3: 
                                    ttype = l[0]
                                    line = l[1]
                                    # print k
                                    if ttype == 1:
                                        planned_amount += line.planned_amount
                                        practical_amount += line.practical_amount
                                    else:
                                        practical_amount += line.amount
                            print 'planned_amount:', planned_amount
                            print 'practical_amount:', practical_amount
                            total_planned_amount += planned_amount
                            total_practical_amount += practical_amount
            print 'total planned_amount:', total_planned_amount
            print 'total practical_amount:', total_practical_amount

            
            

            #SALARIOS
            # total_planned_amount = 0
            # total_practical_amount = 0
            # for j in groups.values():
            #     for k in j.values():
            #         if 'SALARIOS' in k:
            #             print k['SALARIOS']
            #             for l in k['SALARIOS']:
            #                 # print k
            #                 if l[0] == 1:
            #                     total_planned_amount += l[1].planned_amount
            #                     total_practical_amount += l[1].practical_amount
            #                 else:
            #                     total_practical_amount += l[1].amount
            # print 'total planned_amount:', total_planned_amount
            # print 'total practical_amount:', total_practical_amount

            



        # print results
        stop
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
        name = '%s (%s - %s)' % (budget_id.name, date_from, date_to)
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
                worksheet.write_formula(y, x+2, '{=%s}' % cell_range_planned[:-1], _money)
                worksheet.write_formula(y, x+3, '{=%s}' % cell_range_practical[:-1], _silver_money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x+2)
                cell_practical = xl_rowcol_to_cell(y, x+3)
                worksheet.write_formula(y, x+4, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _porcentage)
                y += 1
            # add Y total
            worksheet.write(y, 0, 'TOTAL %s' % row.upper(), _gray)
            for i in range(len(XX)+1): # add X total column too
                x = i * 2 + 1
                cell_range = xl_range(y0, x, y-1, x)
                worksheet.write_formula(y, x, '{=SUM(%s)}' % cell_range, _gray_money)
                cell_range = xl_range(y0, x+1, y-1, x+1)
                worksheet.write_formula(y, x+1, '{=SUM(%s)}' % cell_range, _gray_money)
                # add %
                cell_planned = xl_rowcol_to_cell(y, x)
                cell_practical = xl_rowcol_to_cell(y, x+1)
                worksheet.write_formula(y, x+2, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _gray_porcentage)
            y += 1

        # total
        y += 1 # empty line
        worksheet.write(y, 0, 'TOTAL GENERAL', _blue)
        for i in range(len(XX)+1):
            x = i * 2 + 1
            cell_range = xl_range(1, x, y-1, x)
            worksheet.write_formula(y, x, '{=SUM(%s)/2}' % cell_range, _blue_money)
            cell_range = xl_range(1, x+1, y-1, x+1)
            worksheet.write_formula(y, x+1, '{=SUM(%s)/2}' % cell_range, _blue_money)
            # add %
            cell_planned = xl_rowcol_to_cell(y, x)
            cell_practical = xl_rowcol_to_cell(y, x+1)
            worksheet.write_formula(y, x+2, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _blue_porcentage)

        y_total = y
        y += 2

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
            worksheet.write_formula(y, x+2, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _porcentage)
            y += 1

        # total 'Ingresos'
        y += 1 # empty line
        worksheet.merge_range(y, x_total-3, y, x_total-1, 'TOTAL %s' % row.upper(), _blue)
        cell_range = xl_range(y_income, x_total, y-1, x_total)
        worksheet.write_formula(y, x_total, '{=SUM(%s)}' % cell_range, _blue_money)
        cell_range = xl_range(y_income, x_total+1, y-1, x_total+1)
        worksheet.write_formula(y, x_total+1, '{=SUM(%s)}' % cell_range, _blue_money)
        # add %
        cell_planned = xl_rowcol_to_cell(y, x_total)
        cell_practical = xl_rowcol_to_cell(y, x_total+1)
        worksheet.write_formula(y, x_total+2, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _blue_porcentage)

        # supertotal!
        y += 2 # empty line
        worksheet.merge_range(y, x_total-3, y, x_total-1, 'REMANENTE', _red_total)
        cell_range_planned = '%s+%s' % (xl_rowcol_to_cell(y_total, x_total), xl_rowcol_to_cell(y-2, x_total))
        cell_range_practical = '%s+%s' % (xl_rowcol_to_cell(y_total, x_total+1), xl_rowcol_to_cell(y-2, x_total+1))
        worksheet.write_formula(y, x_total, '{=%s}' % cell_range_planned, _red_money)
        worksheet.write_formula(y, x_total+1, '{=%s}' % cell_range_practical, _red_money)
        # add %
        cell_planned = xl_rowcol_to_cell(y, x_total)
        cell_practical = xl_rowcol_to_cell(y, x_total+1)
        worksheet.write_formula(y, x_total+2, '{=(%s/%s-1)*100}' % (cell_practical, cell_planned), _red_porcentage)


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

