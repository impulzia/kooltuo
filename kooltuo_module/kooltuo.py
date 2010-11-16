# -*- encoding: utf-8 -*-
########################################################################
#
#   OOOP, OpenObject On Python
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   $Id$
#
#   This program is free software: you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation, either version 3 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License
#   along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
########################################################################

from osv import osv
from osv import fields
from tools.translate import _
from time import time
import time


class outlook_application(osv.osv):
    _name = "kooltuo.application"
    _description = "Outlook Application"
    _columns = {
                'name' : fields.char('Name', size=128,select=True),
                'computer_id' : fields.char('Computer ID', size=128),
                'computer_name' : fields.char('Computer Name', size=128),
               }

outlook_application()

class outlook_sync_report(osv.osv):
    _name = "kooltuo.sync.report"
    _description = "Outlook Syncronization Times"
    def _get_default_user(self,cr,uid,context):
        return uid
    
    def _get_datetime(self,cr,uid,context):
        """docstring for _get_datetime"""
        return time()
        
    _columns = {
        		'outlook_application_id':fields.many2one('kooltuo.application', 'Outlook Application',required=True,select=True,states={'done':[('readonly',True)]}),
		        'last_sync_date' : fields.datetime('Last Sync. Date',required=True, states={'done':[('readonly',True)]}),
                'name' : fields.selection([('contacts', 'Contacts'),
                                          ('tasks', 'Tasks'),
                                          ('appointments', 'Appointments'),
                                         ], 'Outlook Object',states={'done':[('readonly',True)]}),
                'note' : fields.text('Description'),
                'user_id' : fields.many2one('res.users','User'),
               }
    
    _defaults = {
                'user_id' : _get_default_user,
            }
    
    

outlook_sync_report()

class outlook_unique_contacts(osv.osv):
    _name = "kooltuo.unique.contacts"
    _description = "Outlook Unique Contacts"
    _columns = {
                'outlook_application_id':fields.many2one("kooltuo.application","Outlook Application",required=True,select=True),
                'contact_entry_id' : fields.char("Outlook Contact EntryID",size=128,select=True,required=True),
                'address_id' : fields.integer("Partner Contact",required=True),
		        'outlook_modified_date':fields.datetime("Outlook Last Modified Date",readonly=True),
                'delete_flag':fields.boolean("Outlook Delete Flag"),
            }
    _defaults = {
                'delete_flag':lambda *a:False,
            }

outlook_unique_contacts()

class outlook_unique_cases(osv.osv):
    _name = "kooltuo.unique.appointments.tasks"
    _description = "Outlook Unique Cases"
    _columns = {
                'outlook_application_id':fields.many2one("kooltuo.application","Outlook Application",required=True,select=True),
                'appointment_entry_id' : fields.char("Outlook Appoitment EntryID",size=128,select=True,required=True),
                'case_id' : fields.integer("Associated Case",required=True),
		        'outlook_modified_date':fields.datetime("Outlook Last Modified Date",readonly=True),
                'delete_flag':fields.boolean("Outlook Delete Flag"),
                'otype' : fields.selection([('tasks', 'Tasks'),
                                          ('appointments', 'Appointments'),
                                         ], 'Type'),
            }
    _defaults = {
                'delete_flag':lambda *a:False,
            }

outlook_unique_cases()


