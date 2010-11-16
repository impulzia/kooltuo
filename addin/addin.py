# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: Addin for Outlook syncronization with openerp
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

from win32com import universal
from win32com.server.exception import COMException
from win32com.client import gencache, DispatchWithEvents, Dispatch, constants, CastTo
from wx import xrc
from syncronize import syncronize
from dialogs.config_dialog import KooltuoConfig, kooltuo_config, ErrorValidatingConection
from common import sync_item_type
import winerror
import pythoncom
import sys
import win32con
import wx

# Support for COM objects we use.
gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 0, bForDemand=True) # Outlook 9
gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 1, bForDemand=True) # Office 9

# The TLB defiining the interfaces we implement
universal.RegisterInterfaces('{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}', 0, 1, 0, ["_IDTExtensibility2"])

####################
# Event
####################

class AboutEvent:
    def OnClick(self, button, cancel):
        print "[AboutEvent][OnClick] start"
        app = KooltuoConfig(False)
        app.MainLoop()        
        
        print "[AboutEvent][OnClick] end"
        
        return cancel
    


class ButtonEvent:
    def OnClick(self, button, cancel):
        import win32ui # Possible, but not necessary, to use a Pythonwin GUI
        
        print "[ButtonEvent][OnClick] start"
        
        config = kooltuo_config()
        try:
            if config.validate_conection():
                sync = syncronize()
                sync.sync_all()
        except ErrorValidatingConection, (instance):
            win32ui.MessageBox(instance.parameter)
        
        config = None    
        print "[ButtonEvent][OnClick] end"
        return cancel
    


class FolderEvent:
    def OnItemAdd(self, item):
        try:
            print "An item was added to the inbox with subject:", item.Subject
        except AttributeError:
            print "An item was added to the inbox, but it has no subject! - ", repr(item)
    


class ContactsEvent:
    def OnItemAdd(self, item):
        print "[ContactsEvent][OnItemAdd] start"
        # sync = syncronize()
        # sync.contacts_add_new(item)
        print "[ContactsEvent][OnItemAdd] end"
    


class OutlookAddin:
    _com_interfaces_ = ['_IDTExtensibility2']
    _public_methods_ = []
    _reg_clsctx_ = pythoncom.CLSCTX_INPROC_SERVER
    _reg_clsid_ = "{0F47D9F3-598B-4d24-B7E3-92AC15ED27E2}"
    _reg_progid_ = "Python.Test.OutlookAddin"
    _reg_policy_spec_ = "win32com.server.policy.EventHandlerPolicy"
    
    def OnConnection(self, application, connectMode, addin, custom):
        print "OnConnection", application, connectMode, addin, custom
        # ActiveExplorer may be none when started without a UI (eg, WinCE synchronisation)
        activeExplorer = application.ActiveExplorer()
        if activeExplorer is not None:
            bars = activeExplorer.CommandBars
            toolbar = bars.Item("Menu Bar")
            
            item_lista = toolbar.Controls.Add(Type=constants.msoControlButton, Temporary=True)
            item_lista = self.menubarButton = DispatchWithEvents(item_lista, AboutEvent)
            item_lista.Caption="Impulzia"
            item_lista.Enabled = True
            
            #TODO: Create a Openerp Bar
            toolbar = bars.Item("Standard")
            
            # Sync button in toolbar
            item = toolbar.Controls.Add(Type=constants.msoControlButton, Temporary=True)
            item = self.toolbarButton = DispatchWithEvents(item, ButtonEvent)
            item.Caption = "Syncronization"
            item.TooltipText = "Click to sync with openerp"
            item.Enabled = True
            
            
        # Getting Messages and contacts events
        inbox = application.Session.GetDefaultFolder(constants.olFolderInbox)
        self.inboxItems = DispatchWithEvents(inbox.Items, FolderEvent)
        
        contacts = application.Session.GetDefaultFolder(constants.olFolderContacts)
        self.contactsItems = DispatchWithEvents(contacts.Items, ContactsEvent)        
        
    
    def OnDisconnection(self, mode, custom):
        print "OnDisconnection"
    
    def OnAddInsUpdate(self, custom):
        print "OnAddInsUpdate", custom
    
    def OnStartupComplete(self, custom):
        print "OnStartupComplete", custom
    
    def OnBeginShutdown(self, custom):
        print "OnBeginShutdown", custom
    


def RegisterAddin(klass):
    import _winreg
    key = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Outlook\\Addins")
    subkey = _winreg.CreateKey(key, klass._reg_progid_)
    _winreg.SetValueEx(subkey, "CommandLineSafe", 0, _winreg.REG_DWORD, 0)
    _winreg.SetValueEx(subkey, "LoadBehavior", 0, _winreg.REG_DWORD, 3)
    _winreg.SetValueEx(subkey, "Description", 0, _winreg.REG_SZ, klass._reg_progid_)
    _winreg.SetValueEx(subkey, "FriendlyName", 0, _winreg.REG_SZ, klass._reg_progid_)


def UnregisterAddin(klass):
    import _winreg
    try:
        _winreg.DeleteKey(_winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Office\\Outlook\\Addins\\" + klass._reg_progid_)
    except WindowsError:
        pass


if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(OutlookAddin)
    if "--unregister" in sys.argv:
        UnregisterAddin(OutlookAddin)
    else:
        RegisterAddin(OutlookAddin)

