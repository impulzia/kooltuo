# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: configuration dialog made in wx with resources
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

import wx
from wx import xrc
import sys, io, os, ConfigParser
from ooop import OOOP
import gettext
_ = gettext.gettext
import time
from datetime import datetime

class ErrorValidatingConection(Exception):
       def __init__(self, value):
           self.parameter = value
       def __str__(self):
           return repr(self.parameter)



class kooltuo_config():
    """Managment of the config file extract and save the information"""
    def __init__(self):
        self.uri = "http://localhost"
        self.dbname = "openerp"
        self.user = "admin"
        self.password = "admin"
        self.last_sync = None
        self.contacts_priority_choice = 0
        
        self.dformat = "%d/%m/%Y %H:%M:%S"
        self.config_file = '%s\\config.cfg' % sys.path[0]
        print self.config_file
        self.config = ConfigParser.RawConfigParser()
        if not os.path.exists(self.config_file):
            self.save()
        self.load()
    
    def load(self):
        """load the info from de configfile"""
        self.config.read(self.config_file)
        self.user = self.config.get('openerp', 'user')
        self.password = self.config.get('openerp', 'pass')
        self.uri = self.config.get('openerp', 'uri')
        self.dbname = self.config.get('openerp', 'dbname')
        if self.config.has_option('openerp', 'last_sync'):
            self.last_sync = time.strptime(self.config.get('openerp', 'last_sync'), self.dformat)
        if self.config.has_option('openerp', 'contacts_priority_choice'):
            self.contacts_priority_choice = self.config.getint('openerp', 'contacts_priority_choice')
    
    def save(self):
        """save the info"""
        if not self.config.has_section('openerp'):
            self.config.add_section('openerp')
        self.config.set('openerp', 'user', self.user)
        self.config.set('openerp', 'pass', self.password)
        self.config.set('openerp', 'uri', self.uri)
        self.config.set('openerp', 'dbname', self.dbname)
        if type(self.last_sync) == type(datetime.now()):
            self.config.set('openerp', 'last_sync', self.last_sync.strftime(self.dformat))
        else:
            if self.last_sync:
                self.config.set('openerp', 'last_sync', datetime(*self.last_sync[0:6]).strftime(self.dformat))
            else:
                self.config.set('openerp', 'last_sync', datetime.now().strftime(self.dformat))
        self.config.set('openerp', 'contacts_priority_choice', self.contacts_priority_choice)
        with open(self.config_file, 'wb') as configfile:
            self.config.write(configfile)
    
    
    def validate_conection(self):
        """docstring for validate"""
        return self.validate_params(dbname=self.dbname, user=self.user, pwd=self.password, uri=self.uri)
        
    def validate_params(self, dbname, user, pwd, uri):
        """docstring for validate_conection"""
        try:
            o = OOOP(dbname=dbname, user=user, pwd=pwd, uri=uri)
        except:
            raise ErrorValidatingConection("Can't connect with openerp")
            
        module = o.IrModuleModule.filter(name='kooltuo_module')
        if not module or module[0].state <> 'installed':
            raise ErrorValidatingConection("The kooltuo module isn't installed")
        else:
            return True
    



class KooltuoConfig(wx.App):
    def OnInit(self):
        self.res = xrc.XmlResource('%s\\dialogs\\gui.xrc' % sys.path[0])
        # self.res = xrc.XmlResource('gui.xrc' )
        #we load the configuration
        self.config = kooltuo_config()
        self.init_frame()
        return True
    
    def init_frame(self):
        self.frame = self.res.LoadFrame(None, 'MainFrame')
        #text fields
        self.url_text = xrc.XRCCTRL(self.frame, 'url_text')
        self.dbname_text = xrc.XRCCTRL(self.frame, 'dbname_text')
        self.user_text = xrc.XRCCTRL(self.frame, 'user_text')
        self.password_text = xrc.XRCCTRL(self.frame, 'password_text')
        self.last_sync_text = xrc.XRCCTRL(self.frame, 'date_stext')
        self.contacts_priority_choice_text = xrc.XRCCTRL(self.frame, 'contacts_priority_choice')
         
        #events
        self.frame.Bind(wx.EVT_BUTTON, self.OnSubmitCheck, id=xrc.XRCID('check_button'))
        self.frame.Bind(wx.EVT_BUTTON, self.OnSubmitSave, id=xrc.XRCID('save_button'))
        self.frame.Bind(wx.EVT_BUTTON, self.OnSubmitSave, id=xrc.XRCID('contacts_save_button'))
        
        self.frame.Show()
        
        self.url_text.SetValue(self.config.uri)
        self.dbname_text.SetValue(self.config.dbname)
        self.user_text.SetValue(self.config.user)
        self.password_text.SetValue(self.config.password) 
        self.contacts_priority_choice_text.SetSelection(self.config.contacts_priority_choice)
        if self.config.last_sync:
            self.last_sync_text.SetLabel(datetime(*self.config.last_sync[0:6]).strftime(self.last_sync_text.GetLabel()))   
    
    def set_config(self):
        self.config.uri = self.url_text.GetValue()
        self.config.dbname = self.dbname_text.GetValue()
        self.config.user = self.user_text.GetValue()
        self.config.password = self.password_text.GetValue()
        self.config.contacts_priority_choice = self.contacts_priority_choice_text.GetSelection()
        self.config.save()
    
    def OnSubmitCheck(self, evt):
        # try:
        #     o = OOOP(dbname=self.dbname_text.GetValue(), user=self.user_text.GetValue(), pwd=self.password_text.GetValue(), uri=self.url_text.GetValue())
        #     module = o.IrModuleModule.filter(name='kooltuo_module')
        #     if not module or module[0].state <> 'installed':
        #         wx.MessageBox(_('Conection Done but Outlook Module isn\'t installed'))
        #     else:
        #         wx.MessageBox(_('Conection Done'))
        # except:
            # wx.MessageBox(_("Error making the connection with those parameters"))
        try:
            if self.config.validate_params(dbname=self.dbname_text.GetValue(), user=self.user_text.GetValue(), pwd=self.password_text.GetValue(), uri=self.url_text.GetValue()):
                wx.MessageBox(_('Conection Done'))
        except ErrorValidatingConection, (instance):
            wx.MessageBox(instance.parameter)
                
        
    
    def OnSubmitSave(self, evt):
        self.set_config()
        wx.MessageBox(_("Configuration saved"))
    


            
if __name__ == '__main__':
    app = KooltuoConfig(False)
    app.MainLoop()