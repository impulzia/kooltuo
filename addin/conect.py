# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: middle connector with ooop to openerp
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

from ooop import OOOP
import os
from dialogs.config_dialog import kooltuo_config
from datetime import datetime, timedelta
from time import time
import time
from common import sync_item_type


def format_openerp_time(optime):
    """get the time from openerp string and set in the fucking outlook way"""
    opformat = "%m/%d/%y %H:%M:%S"
    return datetime(*time.strptime(optime, opformat)[0:6]).strftime("%Y-%m-%d %H:%M:%S")


def machinename():
    """get the machinename"""
    return os.getenv('COMPUTERNAME')

    
def username():
    """get the username"""
    return os.getenv('USERNAME')

    
def record_name():
    """docstring for record_name"""
    return "%s %s" % (username(), machinename())


def item_type_str(item_type):
    if item_type == sync_item_type().contacts:
        ittype = "contacts"
    elif item_type == sync_item_type().calendar:
        ittype = "appointments"
    elif item_type == sync_item_type().tasks:
        ittype = "tasks"
    
    return ittype


class openerp():
    """docstring for openerp"""
    def __init__(self, dbname="openerp", user="admin", pwd="admin", uri="http://localhost"):
        print "Starting openerp connector"
        self.ooop = OOOP(dbname=dbname, user=user, pwd=pwd, uri=uri)
        rec = self.ooop.KooltuoApplication.filter(name=record_name())
        print "\t buscando: %s" % record_name()
        if not rec:
            print "\t %s registro no encontrado" 
            rec = self.ooop.KooltuoApplication.new()
            rec.computer_name = machinename()
            rec.name = record_name()
            rec.save()
            print "registro creado"
        else:
            rec = rec[0]
        self.oapp_id = rec
        print "end Starting openerp connector"
    
    def get_unique_items(self, item_type):
        """docstring for fname"""
        if item_type == sync_item_type().contacts:
            return self.ooop.KooltuoUniqueContacts.filter(outlook_application_id=self.oapp_id.name)
        elif item_type == sync_item_type().calendar:
            return self.ooop.KooltuoUniqueAppointmentsTasks.filter(outlook_application_id=self.oapp_id.name, otype="appointments")
        elif item_type == sync_item_type().tasks:
            return self.ooop.KooltuoUniqueAppointmentsTasks.filter(outlook_application_id=self.oapp_id.name, otype="tasks")

    
    def is_task(self, _id):
        data = self.ooop.KooltuoUniqueAppointmentsTasks.filter(case_id=_id, otype="tasks")
        if data:
            return True
        else:
            return False
        
        
    
    #contacts
    def get_item_by_id(self, _id, item_type):
        """Get the contact object by id"""
        if item_type == sync_item_type().contacts:
            return self.ooop.ResPartnerAddress.get(_id)
        elif item_type == sync_item_type().calendar:
            return self.ooop.CrmMeeting.get(_id)
        elif item_type == sync_item_type().tasks:
            return self.ooop.CrmMeeting.get(_id)
    
    def contact_validate(self, item):
        """validates existence of the contact in the openerp"""
        print "[openerp][contact_validate] start"
        rec = self.ooop.KooltuoUniqueContacts.filter(contact_entry_id=item.EntryID)
        if not rec:
            print "\t[openerp][contact_validate] don't exists"
            if item.CompanyName:
                print "\t\t[openerp][contact_validate] create partner with the CompanyName"
                part = self.partner_find_or_create(item.CompanyName)
            else: 
                print "\t\t[openerp][contact_validate] create partner with the FullName (No CompanyName)"
                part = self.partner_find_or_create(item.FullName) 
            self.contact_create(item, part)
        else:
            print "\t[openerp][contact_validate] exists"
            contact = self.ooop.ResPartnerAddress.get(rec[0].address_id)
            if contact:
                #chnage value
                contact = self.contact_set_fields(contact, item)
                contact.save()
            # cuando se elimina/modifica un registro en el openerp
            # aun hay que ver como gestionar el contacto en 
            # 
            # if rec.delete_flag =True:
        print "[openerp][contact_validate] end"
        
    
    def contact_create(self, item, part):
        """docstring for contact_create"""
        print "begin contact_create"
        contact = self.ooop.ResPartnerAddress.new()
        contact.partner_id = part
        contact = self.contact_set_fields(contact, item)
        print "begin unique contact create"
        contact = self.create_unique(item, contact, sync_item_type().contacts)
        print "end contact_create"
    
    def contact_set_fields(self, contact, item):
        """set the contacts item fields"""
        print "[openerp][contact_set_fields] Start"
        contact.active = True
        contact.type = 0
        
        # if the partner and the contact are the same
        if contact.name <> "" and contact.name <> item.FullName and contact.partner_id:
            if contact.partner_id.name == contact.name:
                contact.partner_id.name = item.FullName
                if item.CompanyName <> contact.partner_id.name:
                    item.CompanyName = contact.partner_id.name
                    item.Save()

        contact.name = item.FullName
        contact.street = item.HomeAddressStreet 
        contact.city = item.HomeAddressCity
        contact.zip = item.HomeAddressPostalCode    
        contact.mobile = item.MobileTelephoneNumber
        contact.phone = item.HomeTelephoneNumber
        contact.email = item.Email1Address

        #contact.title
        #contact.street2
        #contact.state_id
        # contact.fax = item.HomeFaxNumber
        
        #Trabajar con el list
        # contact.country_id = item.HomeAddressCountry
        #Hay que manipular la fecha
        # contact.birthdate = item.Birthday
        # existen muchos campos que requeririan de usar mas de una direccion pero eso cambia por completo 
        # el concepto con el que estamos trabajando y con el que trabaja normalmente axelor y openerp
        print "[openerp][contact_set_fields] End"
        contact.save_all()
        return contact
    
    def create_unique(self, oItem, erpItem, item_type):
        """ Create record data for outlook_unique_XXXXXXX outside """
        unique_item = None
        if item_type == sync_item_type().contacts:
            unique_item = self.ooop.KooltuoUniqueContacts.new()
            unique_item.address_id = erpItem._ref
            unique_item.contact_entry_id = oItem.EntryID
        elif item_type == sync_item_type().calendar:
            unique_item = self.ooop.KooltuoUniqueAppointmentsTasks.new()
            unique_item.case_id = erpItem._ref
            unique_item.appointment_entry_id = oItem.EntryID
            unique_item.otype = 'appointments'
        elif item_type == sync_item_type().tasks:
            unique_item = self.ooop.KooltuoUniqueAppointmentsTasks.new()
            unique_item.case_id = erpItem._ref
            unique_item.otype = 'tasks'
            unique_item.appointment_entry_id = oItem.EntryID
            
        if unique_item:
            unique_item.outlook_application_id = self.oapp_id
            unique_item.save()
            print unique_item
            
        return unique_item
    
    def has_unique_contact(self, contact_ref):
        """find is exists the unique row for sync in OutlookUniqueContacts"""
        print "[openerp][has_unique_contact] start %s, %s" % (contact_ref, self.oapp_id._ref)
        return self.ooop.KooltuoUniqueContacts.filter(address_id=contact_ref, outlook_application_id=self.oapp_id.name)
        print "[openerp][has_unique_contact] end"
    
    ### partner
    def partner_find_or_create(self, name):
        """docstring for partner_findcreate"""
        print "start partner_findcreate"
        part = self.partner_find(name)
        if not part:
            part = self.partner_create(name)
        else:
            part = part[0]
        print "end partner_findcreate"
        return part
    
    def partner_find(self, name):
        """docstring for find_partner"""
        print "partner find %s" % name
        return self.ooop.ResPartner.filter(name=name)
        print "end partner_find"
    
    def partner_create(self, name):
        """docstring for create_partner"""
        print "creating partner %s" % name
        part = self.ooop.ResPartner.new()
        part.name = name
        part.active = True
        part.save()
        print "partner created %s" % name
        return part
    
    ### sync from openrp
    def item_sync_from_openerp(self, last_sync, _item_type):
        """docstring for contact_sync_from_openerp"""
        print "[openerp][item_sync_from_openerp] Start"
        ret = create_items = mod_items = []
        
        if _item_type == sync_item_type().contacts:
            if last_sync:
                mod_items = self.ooop.ResPartnerAddress.filter(write_date__gt=last_sync, create_date__lt=last_sync)
                create_items = self.ooop.ResPartnerAddress.filter(create_date__gt=last_sync)
            else:
                mod_items = self.ooop.ResPartnerAddress.all()
        elif _item_type == sync_item_type().calendar:
            if last_sync:
                mod_items = self.ooop.CrmMeeting.filter(write_date__gt=last_sync, create_date__lt=last_sync)
                create_items = self.ooop.CrmMeeting.filter(create_date__gt=last_sync)
            else:
                mod_items = self.ooop.CrmMeeting.all()
        elif _item_type == sync_item_type().tasks:
            #We need the modifyed items (all of them) in order to set the sync
            if last_sync:
                mod_items = self.ooop.CrmMeeting.filter(write_date__gt=last_sync, create_date__lt=last_sync)
            else:
                mod_items = self.ooop.CrmMeeting.all()
        
        for i in create_items:
            ret.append([i._ref, i])

        for i in mod_items:
            ret.append([i._ref, i])
            
        print "[openerp][item_sync_from_openerp] End"
        
        return ret
        #for contact in contacts:
    
    #appointent
    def create_appointment(self, oItem):
        erpItem = self.ooop.CrmMeeting.new()
        return self.modify_appointment(erpItem, oItem)
    
    def modify_appointment(self, erpItem, oItem):
        """change the data for teh contact"""
        erpItem.active = True
        print erpItem.date
        print oItem.Start
        print format_openerp_time(str(oItem.Start))
        erpItem.date = format_openerp_time(str(oItem.Start))
        erpItem.name = oItem.Subject
        erpItem.description = oItem.Body
        # erpItem.state
        # erpItem.duration = oItem.Duration 
        #getting section
        erpItem.save()

        return erpItem
    
    def create_task(self, oItem):
        erpItem = self.ooop.CrmMeeting.new()
        return self.modify_task(erpItem, oItem)
    
    def modify_task(self, erpItem, oItem):
        """change the data for teh contact"""
        erpItem.active = True
        # print erpItem.date
        # print oItem.Start
        # print format_openerp_time(str(oItem.Start))
        erpItem.date = format_openerp_time(str(oItem.StartDate))
        erpItem.name = oItem.Subject
        erpItem.description = oItem.Body
        # erpItem.state
        # erpItem.duration = oItem.Duration 
        #getting section
        erpItem.save()

        return erpItem
    
    #Sync Date
    def set_new_syncronization_time(self, item_type):
        """create a new record with the sync last date"""
        sync_rec = self.ooop.KooltuoSyncReport.new()
        sync_rec.outlook_application_id = self.oapp_id
        dt = datetime.now()+ timedelta(seconds=2) #hack 4 outlook
        sync_rec.last_sync_date = str(dt)
        sync_rec.name = item_type_str(item_type)
        sync_rec.state = "done"
        sync_rec.save()
        #guardar en el cfg
        config = kooltuo_config()
        config.last_sync = dt
        # config.last_sync = sync_rec.last_sync_date
        config.save()
        return True
    
    def get_last_sync_datetime(self, item_type):
        """Get the last sync datetime for a specific type basis
            on _name (contacts, tasks, appointments)"""
        bsync = self.ooop.KooltuoSyncReport.filter(outlook_application_id=self.oapp_id.name, name=item_type_str(item_type))
        if len(bsync) == 0:
            return None
        num = max([k._ref for k in bsync])
        last = self.ooop.KooltuoSyncReport.get(num)
        return last.last_sync_date
    
    ### global item functions
    def create_item(self, oItem, item_type):
        """ Global function to save diferents types of items"""
        if item_type == sync_item_type().contacts:
            self.contact_validate(oItem)

        elif item_type == sync_item_type().calendar:
            print "create appointment item"
            erpItem = self.create_appointment(oItem)
            self.create_unique(oItem, erpItem, item_type)
            print "create appointment item end"

        elif item_type == sync_item_type().tasks:
            print "create task item"
            erpItem = self.create_task(oItem)
            self.create_unique(oItem, erpItem, item_type)
    
    def modify_item(self, erpItem, oItem, item_type):
        """ Global function to modify diferents types of items """
        if item_type == sync_item_type().contacts:
            self.contact_set_fields(erpItem, oItem)

        elif item_type == sync_item_type().calendar:
            print "modify appointment item"
            self.modify_appointment(erpItem, oItem)

        elif item_type == sync_item_type().tasks:
            print "modify task item"
            self.modify_task(erpItem, oItem)
    
            

openerp_instance = None

def get_openerp():
    """openerp_instance"""
    global openerp_instance
    if not openerp_instance:
        config = kooltuo_config()
        openerp_instance = openerp(dbname=config.dbname, uri=config.uri, user=config.user, pwd=config.password)
    return openerp_instance
