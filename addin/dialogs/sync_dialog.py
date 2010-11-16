# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: syncronization screen
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
import sys

class syncDialog():
    def __init__(self):
        self.sync = sync_dialog(False)
    
    def next(self):
        self.sync.progress_gauge.SetValue( self.sync.progress_gauge.GetValue() + 1)
    
    def reset(self):
        self.sync.progress_gauge.SetValue( 0 )
    
    def reset(self):
        self.sync.progress_gauge.SetValue( 0 )
    
    def max(self, _max):
        self.sync.progress_gauge.SetRange( _max )
    
    def set_label(self, _label, next=True):
        self.sync.m_aviso.SetLabel(_label)
        if next:
            self.next()
    
    def close(self):
        self.sync.close()
        self.sync = None
    

        
class sync_dialog(wx.App):
    def OnInit(self):
        if __name__ == '__main__':
            self.res = xrc.XmlResource('gui.xrc')
        else:    
            self.res = xrc.XmlResource('%s\\dialogs\\gui.xrc' % sys.path[0])
        self.init_frame()
        
        return True
    
    def init_frame(self):
        self.frame = self.res.LoadDialog(None, 'sync_dialog')
        self.cancel_button = xrc.XRCCTRL(self.frame, 'cancel_button')
        self.progress_gauge = xrc.XRCCTRL(self.frame, 'progress_gauge')
        self.m_aviso = xrc.XRCCTRL(self.frame, 'm_aviso')
        self.frame.Bind(wx.EVT_BUTTON, self.OnSubmit, id=xrc.XRCID('cancel_button'))
        self.frame.Show()
    
    def close(self):
        wx.MessageBox("Sincronized")
        self.frame.Close()
    
    def OnSubmit(self, evt):
        self.frame.Close()
    


if __name__ == '__main__':
    app = sync_dialog(False)
