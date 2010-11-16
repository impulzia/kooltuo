# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: wrapper for outlook
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


from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents, Dispatch, constants
from conect import openerp
from datetime import datetime
from common import sync_item_type
import time
import conect
import wx
import winerror
import pythoncom
import sys
import win32con


class outlook_types:
    none, appointment, contact = range(3)

class outlook_wrapper():
    """Encapsulate the most used functions assosiated with the outlook"""
    def __init__(self):
        print "[outlook_wrapper] __init__"
        self.outlook = Dispatch("Outlook.Application")
        print "[outlook_wrapper] end __init__"
    
    def instance(self):
        """the com outlook instance"""
        print "[outlook_wrapper] get instance"
        return self.outlook
        print "[outlook_wrapper] end get instance"
    
    def namespace(self):
        """getting the namespace"""
        print "[outlook_wrapper] get namespace"
        return self.outlook.GetNamespace("MAPI")
        print "[outlook_wrapper] end get namespace"
    
    def create_item(self, _id):
        """create item by id 1=appointment, 2=contact"""
        return self.outlook.CreateItem(_id)
    
    


class outlook_contacts():
    """wrapper for contacts"""
    def __init__(self):
        print "[outlook_contacts] init"
        self.outlook = outlook_wrapper()
        print "[outlook_contacts] end init"
    
    def contact_list(self):
        """list of all contacts in the outlook"""
        print "[outlook_contacts] contact list"
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderContacts).Items
        print "[outlook_contacts] end contact list"
    
    def calendar_list(self):
        """list of all contacts in the outlook"""
        print "[outlook_contacts] contact list"
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderCalendar).Items
        print "[outlook_contacts] end contact list"
    
    def task_list(self):
        """list of all contacts in the outlook"""
        print "[outlook_contacts] contact list"
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderTasks).Items
        print "[outlook_contacts] end contact list"
    
    def get_store_id_contact(self):
        """docstring for get_store_id_contact"""
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderContacts).StoreID
    
    def get_store_id_calendar(self):
        """docstring for get_store_id_contact"""
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderCalendar).StoreID

    def get_store_id_task(self):
        """docstring for get_store_id_contact"""
        return self.outlook.namespace().GetDefaultFolder(constants.olFolderTasks).StoreID
    
    def get_appointment_from_id(self, entry_id):
        """obtain a contact from EntryID"""
        return self.outlook.namespace().GetItemFromID(entry_id, self.get_store_id_calendar())
    
    def get_contact_from_id(self, entry_id):
        """obtain a contact from EntryID"""
        return self.outlook.namespace().GetItemFromID(entry_id, self.get_store_id_contact())
    
    def get_task_from_id(self, entry_id):
        """obtain a contact from EntryID"""
        return self.outlook.namespace().GetItemFromID(entry_id, self.get_store_id_task())
    
    def search_outlook_item_by_modification_date(self, date, item_type):
        """you can get the list of the contacts modify after a date"""
        print "[outlook_contacts] search by modification date"
        restriction = None
        if date:
            fch = datetime(*time.strptime(date, "%Y-%m-%d %H:%M:%S.%f")[0:6])
            restriction = "[LastModificationTime] >= '%s'" % fch.strftime('%m/%d/%Y %H:%M')
        print "restriction %s" % restriction
        
        if restriction:
            if item_type == sync_item_type().contacts:
                items = self.contact_list().Restrict(restriction)
            elif item_type == sync_item_type().calendar:
                items = self.calendar_list().Restrict(restriction)
            elif item_type == sync_item_type().tasks:
                items = self.task_list().Restrict(restriction)
        else:
            if item_type == sync_item_type().contacts:
                items = self.contact_list()
            elif item_type == sync_item_type().calendar:
                items = self.calendar_list()
            elif item_type == sync_item_type().tasks:
                items = self.task_list()
                
        outlook_items = []
        for i in range(1, len(items) + 1):
            item = items.Item(i) 
            outlook_items.append([item.EntryID,  item])
        
        print "[outlook_contacts] end search by modification date"
        return outlook_items
    
    def validate_field(self, field):
        """docstring for validate_field"""
        return (field and field <> "")
    
    def modify_contact(self, contact, item):
        """change the data for teh contact"""
        if self.validate_field(contact.name):
            item.FullName = contact.name
        elif contact.partner_id:
            item.FullName = contact.partner_id.name
            

        if self.validate_field(contact.street):
            item.HomeAddressStreet  = contact.street

        if self.validate_field(contact.city):
            item.HomeAddressCity = contact.city

        if self.validate_field(contact.zip):
            item.HomeAddressPostalCode = contact.zip

        #contact.title
        #contact.street2
        #contact.state_id

        if self.validate_field(contact.mobile):
            item.MobileTelephoneNumber = contact.mobile

        if self.validate_field(contact.phone):
            item.HomeTelephoneNumber = contact.phone
        #item.HomeFaxNumber = contact.fax
        if self.validate_field(contact.email):
            item.Email1Address = contact.email

        if contact.partner_id:
            item.CompanyName = contact.partner_id.name
        #Trabajar con el list
        # item.HomeAddressCountry = contact.country
        #Hay que manipular la fecha
        #item.Birthday = contact.birthdate 

        item.Save()
        return item
    
    def modify_appointment(self, erpItem, oItem):
        """change the data for teh contact"""
        oItem.Start = erpItem.date
        oItem.Subject = erpItem.name
        oItem.Body = erpItem.description
        oItem.Duration = erpItem.duration
        #item.Location = 'Gooseberry Mesa'
        oItem.Save()

        return oItem
    
    def modify_task(self, erpItem, oItem):
        """change the data for teh contact"""
        oItem.StartDate = erpItem.date
        oItem.Subject = erpItem.name
        oItem.Body = erpItem.description
        oItem.Duration = erpItem.duration
        #item.Location = 'Gooseberry Mesa'
        oItem.Save()

        return oItem

        # erpItem.active = True
        # # print erpItem.date
        # # print oItem.Start
        # # print format_openerp_time(str(oItem.Start))
        # erpItem.date = format_openerp_time(str(oItem.StartDate))
        # erpItem.name = oItem.Subject
        # erpItem.description = oItem.Body
        # # erpItem.state
        # # erpItem.duration = oItem.Duration 
        # #getting section
        # erpItem.save()
        # 
        # return oItem
    
    
    #global functions
    def create_item(self, erpItem, item_type):
        """docstring for create_item"""
        item = None
        if item_type == sync_item_type().contacts:
            return self.modify_contact(erpItem, self.outlook.create_item(outlook_types().contact))
        elif item_type == sync_item_type().calendar:
            return self.modify_appointment(erpItem, self.outlook.create_item(outlook_types().appointment))
        elif item_type == sync_item_type().tasks:
            pass
    
    def modify_item(self, erpItem, oItem, item_type):
        """docstring for create_item"""
        item = None
        if item_type == sync_item_type().contacts:
            return self.modify_contact(erpItem, oItem)
        elif item_type == sync_item_type().calendar:
            return self.modify_appointment(erpItem, oItem)
        elif item_type == sync_item_type().tasks:
            return self.modify_task(erpItem, oItem)
            
    
            

  
