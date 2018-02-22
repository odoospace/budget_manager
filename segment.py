# -*- coding: utf-8 -*-
from openerp import models, fields, api

class analytic_segment_template(models.Model):
    _inherit = 'analytic_segment.template'

    def get_direct_childs(self):
        """return a list with childrens, grandchildrens, etc."""
        res = [self]
        for obj in self.child_ids:
            res.append(obj)
        return res

    def get_direct_childs_ids(self, levels=0):
        """return a list with ids of childrens, grandchildrens, etc."""
        res = [self.id]
        for obj in self.child_ids:
            res.append(obj.id)
        return res

