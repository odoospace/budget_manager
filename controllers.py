# -*- coding: utf-8 -*-
from openerp import http

# class BudgetManager(http.Controller):
#     @http.route('/budget_manager/budget_manager/', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/budget_manager/budget_manager/objects/', auth='public')
#     def list(self, **kw):
#         return http.request.render('budget_manager.listing', {
#             'root': '/budget_manager/budget_manager',
#             'objects': http.request.env['budget_manager.budget_manager'].search([]),
#         })

#     @http.route('/budget_manager/budget_manager/objects/<model("budget_manager.budget_manager"):obj>/', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('budget_manager.object', {
#             'object': obj
#         })