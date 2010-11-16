# -*- encoding: utf-8 -*-
try:
    import win32com
except:
    win32com= False

if win32com:
    from conect import openerp
    import conect
    from syncronize import syncronize
    from outlook_wrapper import outlook_wrapper, outlook_contacts
    from win32com import universal
    from win32com.server.exception import COMException
    from win32com.client import gencache, DispatchWithEvents, Dispatch, constants
    from common import sync_item_type

from ooop import OOOP
import time
from datetime import datetime
    

# o = OOOP(dbname="syncdata")
# x = o.ResPartner.filter(id__not_in=[1,2])


# sync = syncronize()
# sync.item_sync(sync_item_type().calendar)
# sync.item_sync(sync_item_type().tasks)
# sync.item_sync(sync_item_type().contacts)

# oc = outlook_contacts()
# tl = oc.contact_list()
# 
# print tl
# print type(tl)
# for i in range(1, len(tl) + 1):
#     item = tl.Item(i)    
#     print item.ClassName

oc = outlook_contacts()
outlook_items = oc.search_outlook_item_by_modification_date(None, sync_item_type().contacts) #, "contacts"
for x in outlook_items:
    x[1].Delete()
outlook_items = oc.search_outlook_item_by_modification_date(None, sync_item_type().calendar) #, "contacts"
for x in outlook_items:
    x[1].Delete()
outlook_items = oc.search_outlook_item_by_modification_date(None, sync_item_type().tasks) #, "contacts"
for x in outlook_items:
    x[1].Delete()


# sync = syncronize()
# sync.item_sync(sync_item_type().calendar)


#test manual 
# oc = outlook_contacts()
# store = oc.get_store_id_contact()
# ow = outlook_wrapper()
# 
# print ow.namespace().GetItemFromID('00000000AC5CDF7E7BBE8E48B7C165096FFC2FD7E40D2000', store).FullName;




#obteniendo la fecha desde el openerp
# class actions:
#     none, create, modify_openerp, modify_outlook, conflict = range(5)
# 
# class fields:
#     """ Memory table to make the syncronization """
#     entry_id, address_id, outlook_modified, openerp_modified, oItem, oContact, action = range(7)
# 
# 
# #Obtenemos los registros de Outlook modificados.

# openerp = conect.get_openerp()
# last_sync = openerp.get_last_sync_datetime("contacts")
# print "ultima sync:%s\n" %last_sync
# reg = openerp.contact_sync_from_openerp(last_sync)
# 
# 
# 
# ow = outlook_wrapper()
# contacts = ow.namespace().GetDefaultFolder(constants.olFolderContacts).Items
# contacts = contacts.Restrict("[LastModificationTime] > '10/23/10 01:50' ")
# for i in range(1, len(contacts) + 1):
#     item = contacts.Item(i)    
#     print "%s %s" % (item.FullName, item.LastModificationTime)
# 
# 
# 
# #contactos del outlook
# oc = outlook_contacts()
# print oc.get_store_id_calendar()


# outlook_items = oc.search_outlook_item_by_modification_date(last_sync, "contacts")
# outlook_items = oc.search_outlook_item_by_modification_date(None, sync_item_type().calendar) #, "contacts"
# for x in outlook_items:
#     print "%s , %s" % (x[0], x[1].LastModificationTime)
#     #x[1].Delete()
# 
# 
# 
# sync = syncronize()
# unique_contacts = openerp.get_unique_contacts()

# opencontacts = openerp.contact_sync_from_openerp(last_sync)
# contacts = sync.fix_openerp_data(opencontacts, sync.create_sync_table(unique_contacts))

# cnt= sync.create_sync_table(unique_contacts)
# for x in cnt:
#     print x
# print "\n"
# contacts = sync.fix_outlook_data(outlook_items, cnt, last_sync)
# 
# for x in contacts:
#     print x

# 
# 
# #Buscamos las relaciones existentes.
# unique_contacts = openerp.get_unique_contacts()
# 
# #Generamos la matriz necesaria para la sincronización los campos estan definidos en la clase fields
# contacts = []
# for i in  unique_contacts:
#     """contact_entry_id, address_id, outlook_modified, openerp_modified, action
#         actions (create, modify, conflict)"""
#     contacts.append([i.contact_entry_id, i.address_id, False, False, None, 
#         openerp.get_contact_address_by_id(i.address_id), actions.none])
# 
# # rellenamos los datos que no estan y ajustamos los que estan desde el outlook
# new_contacts = []
# for i in outlook_items:
#     found = False
#     for j in contacts:
#         if j[fields.entry_id] == i[0]:
#             j[fields.outlook_modified] = True
#             j[fields.oItem] = i[1]
#             found = True
#     if not found:
#         # nuevo registro para creacion de outlook
#         print i
#         new_contacts.append([i[1].EntryID, None, True, False, i[1], None, actions.create])
#         
# #Obtenemos los registros de openerp para establecer la misma actuacion
# openerp_items = openerp.contact_sync_from_openerp(last_sync)
# 
# for i in openerp_items:
#     found = False
#     for j in contacts:
#         if j[fields.address_id] == i[0]:
#             j[fields.openerp_modified] = True
#             j[fields.oItem] = i[1]
#             found = True
#     if not found:
#         # nuevo registro para creacion de outlook
#         new_contacts.append([None, i[0], False, True, i[1], None, actions.create])
#             
#     
# 
# #añadimos los nuevos contactos
# for i in new_contacts:
#     contacts.append(i)
#     
# #eliminamos los que no han sido modificados
# remove = []
# for i in contacts:
#     if not i[fields.outlook_modified] and not i[fields.openerp_modified]:
#         remove.append(i)
#         
# for i in remove:
#     contacts.remove(i)
# 
# #organizamos los que tienen no tienen accion definida
# for i in contacts:
#     if  i[fields.action] == actions.none:
#         if i[fields.outlook_modified] and i[fields.openerp_modified]:
#             i[fields.action] = actions.conflict
#         elif i[fields.openerp_modified]:
#             i[fields.action] = actions.modify_outlook
#         elif i[fields.outlook_modified]:
#             i[fields.action] = actions.modify_openerp
#         else:
#             i[fields.action] = actions.conflict
# 
# # imprimiendo datos de comprobacion
# print "a sincronizar"
# for x in contacts:
#     print x
