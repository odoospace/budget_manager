# -*- coding: utf-8 -*-
from openerp import models, fields, api
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
    
    @api.model
    def _default_group(self):
        res = self._context.get('active_id', False)
        return res

    group_id = fields.Many2one('crossovered.budget.group', default=_default_group, string="Group", required=True)
    date_from = fields.Date(required=True)
    date_to = fields.Date(required=True)
    incoming_bypass = fields.Boolean(default=False)

    @api.one
    def run_export_xlsx(self):
        res = {
            'CCE': {},
            'CCA': {},
            'CCM': {}
        }
        _date_to = self.date_to
        _date_from = self.date_from
        date_from = datetime.strptime(self.date_from, '%Y-%m-%d').date()
        date_to = datetime.strptime(self.date_to, '%Y-%m-%d').date()
        """
        data = {
            'Gastos': (0, 0),
            'Ingresos': (0, 0),
            'Salarios': (0, 0),
        }
        """

        # Create an new Excel file and add a worksheet
        # https://www.odoo.com/es_ES/forum/ayuda-1/question/return-an-excel-file-to-the-web-client-63980
        xlsxfile = StringIO.StringIO()
        workbook = xlsxwriter.Workbook(xlsxfile, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        #xworksheet.freeze_panes(1, 1) # freeze first column and first row

        # styles (centered by default)
        _money = workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
        _porcentage = workbook.add_format({'num_format': '#,##0.00"%"', 'bg_color': '#92ff96', 'align': 'center'})

        _yellow = workbook.add_format({'bg_color': 'yellow'})
        _superyellow = workbook.add_format({'bold': True, 'bg_color': 'yellow', 'num_format': '#,##0.00', 'align': 'center'})
        _orange = workbook.add_format({'bg_color': '#fbe6a2', 'num_format': '#,##0.00', 'align': 'center'})
        _green = workbook.add_format({'bg_color': '#cbddb9', 'num_format': '#,##0.00', 'align': 'center'})
        _red = workbook.add_format({'bg_color': '#f1cdb0', 'num_format': '#,##0.00', 'align': 'center'})
        _silver = workbook.add_format({'bold': True, 'bg_color': '#D0D0D0', 'align': 'center'})


        # predefine columns
        # each name of column have to be unique (Otros-, Otros+)
        COLUMNS = [
            ('Gastos', ['CCE', 'CCA', 'CCM'], _orange),
            ('Ingresos', ['CCE', 'CCA', 'CCM'], _green),
            ('Salarios', ['CCE', 'CCA', 'CCM'], _red),
            # Gastos
            ('Unidades Funcionales y CCE', ['CCE'], _orange),
            ('Unidades Funcionales y CCA', ['CCA'], _orange),
            ('Unidades Funcionales y CCM', ['CCM'], _orange),
            ('Alquiler y Gastos de Oficina', ['CCE', 'CCA', 'CCM'], _orange),
            ('Asignaciones Autonómicas y Municipales', ['CCE'], _orange),
            ('Asignaciones Municipales y Círculos', ['CCA'], _orange),
            # TODO: remove this one
            ('Asignaciones Círculos', ['CCM'], _orange),
            ('Otros-', ['CCE', 'CCA', 'CCM'], _orange),
            # Ingresos
            ('Aportaciones GP', ['CCE', 'CCA', 'CCM'], _green),
            ('Aportaciones Cargos P\xc3\xbablicos', ['CCE', 'CCA', 'CCM'], _green),
            ('Colaboraciones Adscritas', ['CCE', 'CCA', 'CCM'], _green),
            ('Subvenciones', ['CCE'], _green), # TODO: review CCA and CCM!!!!
            ('Estatal', ['CCA', 'CCM'], _green),
            ('Otros+', ['CCE', 'CCA'], _green),
            ('CCA', ['CCM'], _green)
        ]

        COLUMNS_SHORT = {
            'Unidades Funcionales y CCA': 'UNID. FUNC. Y CCA',
            'Unidades Funcionales y CCM': 'UNID. FUNC. Y CCM',
            'Unidades Funcionales y CCE': 'UNID. FUNC. Y CCE',
            'Alquiler y Gastos de Oficina': 'ALQ. Y GASTOS DE OFIC.',
            u'Asignaciones Autonómicas y Municipales': 'ASIGN. AUTO. Y MUN.',
            u'Asignaciones Municipales y Círculos': 'ASIGN. MUNIC. Y CIRC.',
            u'Aportaciones Cargos Públicos': 'APORTAC. CARGOS PUB.',
            'Colaboraciones Adscritas': 'COLAB. ADSCRITAS'
        }

        # [level, topic] - %s -> CCA, CCE, CCM
        MAPPING = {
            3: {
                'INGRESOS': 'Ingresos',
                'SALARIOS': 'Salarios'
            },
            2: {
                'Ingresos': {
                    'APORTACIONES GP': 'Aportaciones GP',
                    'APORTACIONES CARGOS PUB.': u'Aportaciones Cargos Públicos',
                    'CONSEJO AUTONOMICO': 'CCA', # CCA ?
                    'ESTATAL': 'Estatal', # Estatal 
                    'CROWFUNDING': 'Otros+',
                    'COLABORACIONES ADSCRITAS': 'Colaboraciones Adscritas',
                    'SUBVENCIONES': 'Subvenciones'
                },
                'Gastos Secretarias': {
                    '*': 'Unidades Funcionales y %s'
                },
                'Gastos Areas y Equipos': {
                    '*': 'Unidades Funcionales y %s'
                },
                'Gastos Generales': {
                    'ACTOS VARIOS': 'Unidades Funcionales y %s',
                    'CONSEJO ESTATAL CCE': 'Unidades Funcionales y %s',
                    'CONSULTAS CIUDADANAS': 'Unidades Funcionales y %s',
                    'CONSEJO AUTONOMICO': 'Unidades Funcionales y %s',
                    'CONSEJO MUNICIPAL': 'Unidades Funcionales y %s',
                    'CONSULTAS CIUDADANAS': 'Unidades Funcionales y %s',
                    'ALQUILER Y GASTOS OFICINAS': 'Alquiler y Gastos de Oficina',
                    'COMISIONES BANCARIAS': 'Otros-',
                    'ASIGNACIONES AUTONOMICAS': u'Asignaciones Autonómicas y Municipales',
                    'ASIGNACIONES MUNICIPALES': u'Asignaciones Municipales y Círculos',
                    'ASIGNACIONES CIRCULOS': u'Asignaciones Municipales y Círculos',
                    'PROVISION CONTINGENCIAS': 'Otros-'
                },
                'Gastos Extraordinarios': {
                    'DESARROLLO PARTICIPA': 'Unidades Funcionales y %s',
                    'ESTUDIOS DEMOSCOPICOS': 'Unidades Funcionales y %s',
                    'FONDO ANUAL ACTIVIDADES': 'Unidades Funcionales y %s',
                    'ORDENADORES': 'Unidades Funcionales y %s',
                    'PROYECTOS AREAS': 'Unidades Funcionales y %s',
                    'PROYECTOS SOCIALES': 'Unidades Funcionales y %s',
                    'SEDE CENTRAL FCO VILLAESPESA': 'Otros-'
                }
            }
        }
        
        total_planned_amount = {}
        total_practical_amount = {}
        budgets = {
            'CCE': [],
            'CCM': [],
            'CCA': []
        }

        for i in self.group_id.budget_ids:
            # compatibility with names too if category isn't defined
            if i.category:
                category = i.category
            elif 'CCE' in i.name:
                category = 'CCE'
            elif 'CCA' in i.name:
                category = 'CCA'
            elif 'CCM' in i.name:
                category = 'CCM'

            #print 'category...', category

            budgets[category].append(i)
            vals = {
                'budget_id': i.id, 
                'date_from': date_from,
                'date_to': date_to
            }
            budget_wiz = self.env['budget_manager.xlsxwizard'].create(vals)
            # get data from detail budget report
            X, XX, groups, analytic_lines, analytic_lines_obj = budget_wiz.process_data(date_from, date_to)
            #res[i.id] = (X, XX, groups, analytic_lines, analytic_lines_obj) # TODO: review to use this
            #print '===', i.name, groups.keys()

            # reset totals (planned and practical)
            total_planned_amount[i] = {}
            total_practical_amount[i] = {}
            for c in COLUMNS:
                if category in c[1]:
                    column = c[0].decode('utf-8')
                    total_planned_amount[i][column] = 0
                    total_practical_amount[i][column] = 0

            #print 'starting...'

            for k1, l1 in groups.items(): # level 1
                #print 'k1', k1    
                for k2, l2 in l1.items(): # level 2
                    #print 'k2', k2
                    for k3, l3 in l2.items(): # level 3
                        #print 'k3', k3
                        #print 'k1:', k1, 'k2:', k2, 'k3:', k3
                        # check mapping for level 3
                        for m in MAPPING[3]:
                            print m, MAPPING[3]
                            if k3 == m:
                                column = MAPPING[3][m]
                                # some name of columns are dynamic
                                if '%s' in column:
                                    column = column % category
                                # check type
                                
                                for ttype, v in l3:
                                    #print i, column, '...'
                                    if ttype == 1:
                                        #print 'ttype:', 1
                                        #print '+++ 1', m, column, v.planned_amount, v.practical_amount
                                        total_planned_amount[i][column] += v.planned_amount
                                        total_practical_amount[i][column] += v.practical_amount
                                    else:
                                        #print 'ttype:', ttype
                                        #print '+++ 2', m, column, 0, v.amount
                                        total_practical_amount[i][column] += v.amount
                        # check mapping for level 2
                        for m in MAPPING[2]:
                            for n in MAPPING[2][m]:
                                if (n == '*' and k1 == m.decode('utf-8')) or (k1 == m.decode('utf-8') and k2 == n.decode('utf-8')):
                                    column = MAPPING[2][m][n]
                                    # some name of columns are dynamic
                                    if '%s' in column:
                                        column = column % category
                                    # check type
                                    for ttype, v in l3:
                                        # TODO: refactor this (gastos)
                                        if ttype == 1:
                                            # sum again
                                            if 'Gastos' in m:
                                                total_planned_amount[i]['Gastos'] += v.planned_amount
                                                total_practical_amount[i]['Gastos'] += v.practical_amount
                                            #print '+++ 3', m, n, column, v.planned_amount, v.practical_amount
                                            try:
                                                total_planned_amount[i][column] += v.planned_amount
                                                total_practical_amount[i][column] += v.practical_amount
                                            except Exception as e:
                                                #print '>>>', e
                                                if column == 'Subvenciones':
                                                    _column = 'Otros+'
                                                    total_planned_amount[i][_column] += v.planned_amount
                                                    total_practical_amount[i][_column] += v.practical_amount
                                        else:
                                            #print '+++ 4', m, n, column, 0, v.amount
                                            try:
                                                total_practical_amount[i][column] += v.amount
                                            except Exception as e:
                                                #print '***', e
                                                if column == 'Asignaciones Municipales y Círculos':
                                                    _column = 'Asignaciones Autonómicas y Municipales'
                                                    total_practical_amount[i][column] += v.amount
                                            # sum again
                                            if 'Gastos' in m:
                                                total_practical_amount[i]['Gastos'] += v.amount
        

            # print columns
            res[category][i.name] = {}
            for c in COLUMNS:
                if category in c[1]:
                    column = c[0].decode('utf-8')
                    res[category][i.name][c[0]] = {
                        'planned': total_planned_amount[i][column],
                        'practical': total_practical_amount[i][column]
                    } 
        
        # TODO: add monthly data
        # Excel stuff

        y = 0
        worksheet.set_column(0, 0, 30)
        worksheet.set_column(1, 1, 20)
        for category in ['CCE', 'CCA', 'CCM']:
            # headers
            if res[category]:
                x = 0
                worksheet.merge_range(y, x, y+1, x, category, _silver)
                x += 1
                for c in COLUMNS:
                    if category in c[1]:
                        column = c[0].decode('utf-8')
                        if column == 'Salarios':
                            worksheet.set_column(x, x+3, 12) # 4 colums
                            worksheet.merge_range(y, x, y, x+3, column.upper(), _silver)
                        else:
                            worksheet.set_column(x, x+1, 12) # 2 columns
                            # change some column names to short ones
                            if column.startswith('Aportaciones Cargos P'):
                                # hack
                                worksheet.merge_range(y, x, y, x+1, 'APORTAC. CARGOS PUB.', _silver)
                            elif column in COLUMNS_SHORT:
                                worksheet.merge_range(y, x, y, x+1, COLUMNS_SHORT[column], _silver)
                            else:
                                worksheet.merge_range(y, x, y, x+1, column.upper(), _silver)
                        worksheet.write(y+1, x, 'PRESUP.', _silver)
                        if column == 'Salarios':
                            worksheet.write(y+1, x+1, '%', _silver)
                            x += 1
                        worksheet.write(y+1, x+1, 'REALES', _silver)
                        if column == 'Salarios':
                            worksheet.write(y+1, x+2, '%', _silver)
                            x += 1
                        x += 2
                y += 1
                y_start_total = y
                for line in res[category]:
                    y += 1
                    x = 0
                    worksheet.write(y, x, line)
                    x += 1
                    for c in COLUMNS:
                        if category in c[1]:
                            column = c[0].decode('utf-8').upper()
                            if c[2]:
                                worksheet.write(y, x, res[category][line][c[0]]['planned'], _money)
                                if column == 'SALARIOS':
                                    cell_gastos_planned = xl_rowcol_to_cell(y, 1)
                                    cell_planned = xl_rowcol_to_cell(y, x)
                                    worksheet.write_formula(y, x+1, '=(%s/%s)*100' % (cell_planned, cell_gastos_planned), _money)
                                    x += 1
                                worksheet.write(y, x+1, res[category][line][c[0]]['practical'], c[2])
                                if column == 'SALARIOS':
                                    cell_gastos_practical = xl_rowcol_to_cell(y, 2)
                                    cell_practical = xl_rowcol_to_cell(y, x+1)
                                    worksheet.write_formula(y, x+2, '=(%s/%s)*100' % (cell_practical, cell_gastos_practical), c[2])
                                    x += 1
                            else:
                                # SALARIOS have an color
                                worksheet.write(y, x, res[category][line][c[0]]['planned'], _money)
                                worksheet.write(y, x+1, res[category][line][c[0]]['practical'])
                            x += 2
                y += 1
                x = 1
                worksheet.write(y, 0, 'TOTAL', _yellow)
                for c in COLUMNS:
                    if category in c[1]:
                        column = c[0].decode('utf-8').upper()
                        cell_range = xl_range(y_start_total+1, x, y-1, x)
                        worksheet.write_formula(y, x, '=SUM(%s)' % cell_range, _superyellow)
                        if column == 'SALARIOS':
                            cell_gastos_planned = xl_rowcol_to_cell(y, 1)
                            cell_planned = xl_rowcol_to_cell(y, x)
                            worksheet.write_formula(y, x+1, '=(%s/%s)*100' % (cell_planned, cell_gastos_planned), _superyellow)
                            x += 1
                        cell_range = xl_range(y_start_total+1, x+1, y-1, x+1)
                        worksheet.write_formula(y, x+1, '=SUM(%s)' % cell_range, _superyellow)
                        if column == 'SALARIOS':
                            cell_gastos_practical = xl_rowcol_to_cell(y, 2)
                            cell_practical = xl_rowcol_to_cell(y, x+1)
                            worksheet.write_formula(y, x+2, '=(%s/%s)*100' % (cell_practical, cell_gastos_practical), _superyellow)
                            x += 1
                        x += 2
                y += 2
        
        workbook.close()
        
        # Rewind the buffer.
        xlsxfile.seek(0)
        name = self.group_id.name.lower().replace(' ', '_')
        vals = {
            'name': 'presupuesto_agrupado_%s_%s_%s.xlsx' % (name, _date_from, _date_to),
            'datas': base64.encodestring(xlsxfile.read()),
            'datas_fname': 'presupuesto_%s_%s_%s.xlsx' % (name, _date_from, _date_to),
            'res_model': self.group_id._name,
            'res_id': self.group_id.id,
            'type': 'binary'
        }
        attachment_id = self.env['ir.attachment'].create(vals)

        return True

