# -*- encoding: utf-8 -*-
########################################################################
#
#   Copyright (C) 2010 Impulzia S.L. All Rights Reserved.
#   Gamaliel Toro <argami@impulzia.com>
#   Description: common 
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

class actions:
    none, create_outlook, create_openerp, modify_openerp, modify_outlook, conflict = range(6)

class fields:
    """ Memory table to make the syncronization """
    entry_id, address_id, outlook_modified, openerp_modified, oItem, oContact, action = range(7)
    
class sync_item_type:
    """Sync item type"""
    contacts, calendar, tasks = range(3)
