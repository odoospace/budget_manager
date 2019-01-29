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
        data = {
            'Ingresos': (0, 0),
            'Salarios': (0, 0),
        }

        # predefine columns
        # each name of column have to be unique (Otros-, Otros+)
        COLUMNS = [
            ('Gastos', ['CCE', 'CCA', 'CCM']),
            ('Ingresos', ['CCE', 'CCA', 'CCM']),
            ('Salarios', ['CCE', 'CCA', 'CCM']),
            # Gastos
            ('Unidades Funcionales y CCE', ['CCE']),
            ('Unidades Funcionales y CCA', ['CCA']),
            ('Unidades Funcionales y CCM', ['CCM']),
            ('Alquiler y Gastos de Oficina', ['CCE', 'CCA', 'CCM']),
            ('Asignaciones Autonómicas y Municipales', ['CCE']),
            ('Asignaciones Municipales y Círculos', ['CCA']),
            ('Asignaciones Círculos', ['CCM']),
            ('Otros-', ['CCE', 'CCA', 'CCM']),
            # Ingresos
            ('Aportaciones GP', ['CCE', 'CCA', 'CCM']),
            ('Aportaciones Cargos P\xc3\xbablicos', ['CCE', 'CCA', 'CCM']),
            ('Colaboraciones Adscritas', ['CCE', 'CCA', 'CCM']),
            ('Subvenciones', ['CCE']), # TODO: review CCA and CCM!!!!
            ('Estatal', ['CCA', 'CCM']),
            ('Otros+', ['CCE', 'CCA']),
            ('CCA', ['CCM'])
        ]

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
                    'CONSEJO AUTONOMICO': 'Otros+', # CCA ?
                    'ESTATAL': 'Otros+', # Estatal 
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
                    'ASIGNACIONES AUTONÓMICAS': u'Asignaciones Autonómicas y Municipales',
                    'ASIGNACIONES MUNICIPALES': u'Asignaciones Municipales y Círculos',
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

            print 'category...', category

            budgets[category].append(i)
            vals = {
                'budget_id': i.id, 
                'date_from': date_from,
                'date_to': date_to
            }
            budget_wiz = self.env['budget_manager.xlsxwizard'].create(vals)
            # get data from detail budget report
            print '___'
            X, XX, groups, analytic_lines, analytic_lines_obj = budget_wiz.process_data(date_from, date_to)
            #res[i.id] = (X, XX, groups, analytic_lines, analytic_lines_obj) # TODO: review to use this
            print '===', i.name, groups.keys()

            # reset totals (planed and practical)
            total_planned_amount[i] = {}
            total_practical_amount[i] = {}
            for c in COLUMNS:
                if category in c[1]:
                    column = c[0].decode('utf-8')
                    total_planned_amount[i][column] = 0
                    total_practical_amount[i][column] = 0

            print 'starting...'

            for k1, l1 in groups.items(): # level 1
                #print 'k1', k1    
                for k2, l2 in l1.items(): # level 2
                    #print 'k2', k2
                    for k3, l3 in l2.items(): # level 3
                        #print 'k3', k3
                        print 'k1:', k1, 'k2:', k2, 'k3:', k3
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
                                        print 'ttype:', 1
                                        total_planned_amount[i][column] += v.planned_amount
                                        total_practical_amount[i][column] += v.practical_amount
                                    else:
                                        print 'ttype:', ttype
                                        total_practical_amount[i][column] += v.amount
                            #stop
                        # check mapping for level 2
                        #stop
                        for m in MAPPING[2]:
                            for n in MAPPING[2][m]:
                                if (n == '*' and k1 == m.decode('utf-8')) or (k1 == m.decode('utf-8') and k2 == n.decode('utf-8')):
                                    column = MAPPING[2][m][n]
                                    # some name of columns are dynamic
                                    if '%s' in column:
                                        column = column % category
                                    # check type
                                    for ttype, v in l3:
                                        if ttype == 1:
                                            try:
                                                total_planned_amount[i][column] += v.planned_amount
                                                total_practical_amount[i][column] += v.practical_amount
                                            except Exception as e:
                                                print '***', e
                                                if column == 'Subvenciones':
                                                    _column = 'Otros+'
                                                    total_planned_amount[i][_column] += v.planned_amount
                                                    total_practical_amount[i][_column] += v.practical_amount
                                        else:
                                            total_practical_amount[i][column] += v.amount
            # print columns
            res[category][i.name] = {}
            for c in COLUMNS:
                if category in c[1]:
                    column = c[0].decode('utf-8')
                    res[category][i.name][c[0]] = {
                        'planned': total_planned_amount[i][column],
                        'practical': total_practical_amount[i][column]
                    } 
                    
        #print '>>>', res
        
        # Create an new Excel file and add a worksheet
        # https://www.odoo.com/es_ES/forum/ayuda-1/question/return-an-excel-file-to-the-web-client-63980
        xlsxfile = StringIO.StringIO()
        workbook = xlsxwriter.Workbook(xlsxfile, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        #xworksheet.freeze_panes(1, 1) # freeze first column and first row

        # styles
        _money = workbook.add_format({'num_format': '#,##0.00'})
        _porcentage = workbook.add_format({'num_format': '#,##0.00"%"', 'bg_color': '#92ff96'})
        _bold = workbook.add_format({'bold': True})
        _bold_center = workbook.add_format({'bold': True, 'align': 'center'})

        _yellow = workbook.add_format({'bg_color': '#fbe5a3', 'num_format': '#,##0.00'})
        _green = workbook.add_format({'bg_color': '#cbdeb9', 'num_format': '#,##0.00'})
        _red = workbook.add_format({'bg_color': '#f0cdb1', 'num_format': '#,##0.00'})
        _gray = workbook.add_format({'bold': True, 'bg_color': 'silver'})

        """
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
        """

        y = 0
        for category in ['CCE', 'CCA', 'CCM']:
            # headers
            if res[category]:
                x = 1
                worksheet.set_column(y, 0, 20)
                worksheet.write(y, 0, category, _gray)
                for c in COLUMNS:
                    if category in c[1]:
                        column = c[0].decode('utf-8')
                        worksheet.set_column(x, x+1, 12)
                        worksheet.merge_range(y, x, y, x+1, column, _gray)
                        x += 2
                for line in res[category]:
                    y += 1
                    x = 0
                    worksheet.write(y, x, line, _gray)
                    x += 1
                    for c in COLUMNS:
                        if category in c[1]:
                            worksheet.write(y, x, res[category][line][c[0]]['planned'])
                            worksheet.write(y, x+1, res[category][line][c[0]]['practical'])
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

