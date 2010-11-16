# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: main code to syncronize
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

from outlook_wrapper import outlook_wrapper, outlook_contacts
from conect import openerp
import conect
from dialogs.sync_dialog import syncDialog
from datetime import datetime
from dialogs.config_dialog import kooltuo_config
from common import actions, fields, sync_item_type
from time import mktime, localtime
import time
import gettext
_ = gettext.gettext

def format_outlook_time(optime):
    """get the time from openerp string and set in the fucking outlook way"""
    opformat = "%Y-%m-%d %H:%M:%S.%f"
    print optime
    return datetime(*time.strptime(optime, opformat)[0:6]).strftime("%m/%d/%y %H:%M:%S")

class syncronize():
    """syncronize outlook <-> openerp"""
    def __init__(self):
        print "[syncronize][__init__] start"
        self.openerp = conect.get_openerp()
        self.out_contacts = outlook_contacts()
        self.config = kooltuo_config()
        print "[syncronize][__init__] end"
    
    def sync_all(self):
        """docstring for sync_all"""
        #para usar el outlook como parent
        # window = wx.Window_FromHWND(None, hwnd)
        self.dlg = syncDialog()
        self.dlg.reset()
        self.dlg.max(33)
        
        self.item_sync(sync_item_type().contacts)
        self.item_sync(sync_item_type().calendar)
        self.item_sync(sync_item_type().tasks)
        
        self.dlg.close()
        self.dlg = None
    
    def item_sync(self, item_type):
        """ Sync Contacts """
        print "[syncronize][contacts_sync] start"
        print "[syncronize][contacts_sync] outlook -> openerp"
        
        # get the last sync date
        self.dlg.set_label(_("Get last sync"))
        last_sync = self.openerp.get_last_sync_datetime(item_type)
        # get outlook last modifiyed outlook contacts
        self.dlg.set_label(_("getting outlook items"))
        outlook_items = self.out_contacts.search_outlook_item_by_modification_date(last_sync, item_type)
        #get last modified openerp contacts
        self.dlg.set_label(_("getting openerp items"))
        openerp_items = self.openerp.item_sync_from_openerp(last_sync, item_type)
        #we get the existent relations
        self.dlg.set_label(_("getting the relations"))
        unique_items = self.openerp.get_unique_items(item_type)
        #we prepare the matrix
        self.dlg.set_label(_("creating sync table"))
        items = self.create_sync_table(unique_items, item_type)
        #we introduce new outlook items and fix the changes in the table
        self.dlg.set_label(_("Adjusting Outlook data"))
        items = self.fix_outlook_data(outlook_items, items, last_sync, item_type)
        #we same as before but with the openerp data
        self.dlg.set_label(_("Adjusting Openerp data"))
        items = self.fix_openerp_data(openerp_items, items, item_type)
        #remove the unmodified data
        self.dlg.set_label(_("Removind not sync items"))
        items = self.remove_unmodified_data(items)
        self.dlg.next()
        # #setting actions
        self.dlg.set_label(_("Establishing acctions"))
        items = self.setting_actions(items)
        # #syncronize objects
        self.dlg.set_label(_("Syncronize"))
        self.syncronize_items(items, item_type)
        # #New sync time
        self.dlg.set_label(_("Setting new last sync datetime"))
        self.openerp.set_new_syncronization_time(item_type)
        
        print "[syncronize][contacts_sync] end"
    
    def syncronize_items(self, items, item_type):
        """syncronization for contacts"""
        for contact in items:
            if contact[fields.action] == actions.create_outlook:
                print "creando registro en outlook"
                item = self.out_contacts.create_item(contact[fields.oContact], item_type)
                self.openerp.create_unique(item, contact[fields.oContact], item_type)
                
            elif contact[fields.action] == actions.create_openerp:
                print "creando registro en openerp"
                self.openerp.create_item(contact[fields.oItem], item_type)
                
            elif contact[fields.action] == actions.modify_openerp: 
                print "modificando openerp"
                self.openerp.modify_item(contact[fields.oContact], contact[fields.oItem], item_type)
                
            elif contact[fields.action] == actions.modify_outlook:
                print "modificando outlook"
                self.out_contacts.modify_item(contact[fields.oContact], contact[fields.oItem], item_type)
                
            elif contact[fields.action] == actions.conflict:
                if self.config.contacts_priority_choice == 0:
                    print "prioridad openerp"
                    self.out_contacts.modify_item(contact[fields.oContact], contact[fields.oItem], item_type)
                elif self.config.contacts_priority_choice == 1:
                    print "prioridad outlook"
                    self.openerp.modify_item(contact[fields.oContact], contact[fields.oItem], item_type)
                print "conflicto"
    
    #preparation functions
    def setting_actions(self, contacts):
        """ set actions to the valid records for the sync """
        for i in contacts:
            if  i[fields.action] == actions.none:
                if i[fields.outlook_modified] and i[fields.openerp_modified]:
                    i[fields.action] = actions.conflict
                elif i[fields.openerp_modified]:
                    i[fields.action] = actions.modify_outlook
                elif i[fields.outlook_modified]:
                    i[fields.action] = actions.modify_openerp
                else:
                    i[fields.action] = actions.conflict
        return contacts
    
    def remove_unmodified_data(self, contacts):
        """we remove the unmodified datd from the sync table"""
        remove = []
        for i in contacts:
            if not i[fields.outlook_modified] and not i[fields.openerp_modified]:
                remove.append(i)
                
        for i in remove:
            contacts.remove(i)
            
        return contacts
    
    def fix_openerp_data(self, openerp_items, items, item_type):
        """set the new fields in openerp and fix the information in table"""
        new_items = []
        for i in openerp_items:
            found = False
            for j in items:
                if j[fields.address_id] == i[0]:
                    j[fields.openerp_modified] = True
                    j[fields.oContact] = i[1]
                    if item_type == sync_item_type().contacts:
                        j[fields.oItem] = self.out_contacts.get_contact_from_id(j[fields.entry_id])
                    elif item_type == sync_item_type().calendar:
                        j[fields.oItem] = self.out_contacts.get_appointment_from_id(j[fields.entry_id])
                    elif item_type == sync_item_type().tasks:
                        j[fields.oItem] = self.out_contacts.get_task_from_id(j[fields.entry_id])
                    found = True
            if not found:
                #the task option has no a type in openerp there's no creation task erp->outlook
                if item_type <> sync_item_type().tasks:
                    if item_type == sync_item_type().calendar:
                        if not self.openerp.is_task(i[0]):
                            new_items.append([None, i[0], False, True, None, i[1], actions.create_outlook])
                    else:
                        # new record to outlook
                        new_items.append([None, i[0], False, True, None, i[1], actions.create_outlook])
                    
        for i in new_items:
            items.append(i)
            
        return items
    
    def fix_outlook_data(self, outlook_items, items, last_sync, item_type):
        """docstring for fix_outlook_data"""
        new_items = []
        for i in outlook_items:
            found = False
            for j in items:
                if j[fields.entry_id] == i[0]:
                    found = True
                    if  str(i[1].LastModificationTime) > format_outlook_time(last_sync):
                        j[fields.outlook_modified] = True
                        j[fields.oItem] = i[1]
            if not found:
                new_items.append([i[1].EntryID, None, True, False, i[1], None, actions.create_openerp])
        
        for i in new_items:
            items.append(i)
            
        return items
    
    def create_sync_table(self, unique_items, item_type):
        """docstring for create_sync_table"""
        items = []
        for i in  unique_items:
            """contact_entry_id, address_id, outlook_modified, openerp_modified, action
                actions (create, modify, conflict)"""
            if item_type == sync_item_type().contacts:
                _id = i.address_id
                _entry_id = i.contact_entry_id
            elif item_type == sync_item_type().calendar:
                _id = i.case_id
                _entry_id = i.appointment_entry_id
            elif item_type == sync_item_type().tasks:
                _id = i.case_id
                _entry_id = i.appointment_entry_id
                
            items.append([_entry_id, _id, False, False, None, 
                self.openerp.get_item_by_id(_id, item_type), actions.none])
                
        return items
    
    def contacts_add_new(self, item):
        """direct adding contact outlook -> openerp"""
        print "[syncronize][contacts_add_new] start"
        self.openerp.contact_validate(item)
        print "[syncronize][contacts_add_new] end"
    


    


