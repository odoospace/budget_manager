# -*- coding: utf-8 -*-
{
    'name': "Budget Manager",

    'summary': """
        Manage your bugdet crossover lines in a sustainable way""",

    'description': """
        Manage your bugdet crossover lines in a sustainable way
    """,

    'author': "Impulzia",
    'website': "http://impulzia.com",

    # Categories can be used to filter modules in modules listing
    # Check https://github.com/odoo/odoo/blob/master/openerp/addons/base/module/module_data.xml
    # for the full list
    'category': 'Uncategorized',
    'version': '8.0.4.1',

    # any module necessary for this one to work correctly
    'depends': ['base', 'account', 'account_budget', 'analytic_segment'],

    # always loaded
    'data': [
        'security/ir.model.access.csv',
        'templates.xml',
    ],
    'css': [
        'static/src/css/default.cssa'
    ],
    # only loaded in demonstration mode
    #'demo': [
    #    'demo.xml',
    #],
}
