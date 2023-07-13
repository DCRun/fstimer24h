#fsTimer - free, open source software for race timing.
#Copyright 2012-15 Ben Letham

#This program is free software: you can redistribute it and/or modify
#it under the terms of the GNU General Public License as published by
#the Free Software Foundation, either version 3 of the License, or
#(at your option) any later version.

#This program is distributed in the hope that it will be useful,
#but WITHOUT ANY WARRANTY; without even the implied warranty of
#MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#GNU General Public License for more details.

#You should have received a copy of the GNU General Public License
#along with this program.  If not, see <http://www.gnu.org/licenses/>.

#The author/copyright holder can be contacted at bletham@gmail.com
'''Handles the root window of the application'''

import gi
gi.require_version('Gtk', '3.0')
from gi.repository import Gtk
import fstimer.gui
import webbrowser, os
from fstimer.gui.util_classes import GtkStockButton
from fstimer.gui.util_classes import MenuItemIcon
import gettext

class RootWin(Gtk.Window):
    '''Handles the root window of the application'''

    def __init__(self, path, show_about_cb, importprereg_cb,
                 prereg_cb, compreg_cb, pretime_cb, edit_cb):
        '''Creates the root window with choices for the tasks'''
        super(RootWin, self).__init__(Gtk.WindowType.TOPLEVEL)
        self.modify_bg(Gtk.StateType.NORMAL, fstimer.gui.bgcolor)
        fname = os.path.abspath(
            os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                '../data/icon.png'))
        self.set_icon_from_file(fname)
        self.set_title('fsTimer - ' + os.path.basename(path))
        self.set_position(Gtk.WindowPosition.CENTER)
        self.connect('delete_event', Gtk.main_quit)
        self.set_border_width(0)
        # Generate the menubar
        mb = Gtk.MenuBar()
        helpmenu = Gtk.Menu()
        helpm = Gtk.MenuItem(_('Menu'))
        helpm.set_submenu(helpmenu)
        menuedit = MenuItemIcon('edit', _('Edit project settings'), edit_cb, True)
        helpmenu.append(menuedit)
        menuhelp = MenuItemIcon('help', _('Documentation'), lambda x: webbrowser.open_new('http://fstimer.org/documentation/documentation_sec2.htm'))
        helpmenu.append(menuhelp)
        menuabout = MenuItemIcon('about', _('About'), show_about_cb, self)
        helpmenu.append(menuabout)
        mb.append(helpm)
        ### Frame
        rootframe = Gtk.Frame(label='al')
        rootframe_label = Gtk.Label(label='')
        rootframe_label.set_markup('<b>fsTimer - ' + os.path.basename(path) + '</b>')
        rootframe.set_label_widget(rootframe_label)
        rootframe.set_border_width(20)
        #And now fill the frame with a table
        roottable = Gtk.Table(4, 2, False)
        roottable.set_row_spacings(20)
        roottable.set_col_spacings(20)
        roottable.set_border_width(10)
        #And internal buttons
        rootbtnPREREG = Gtk.Button('Import')
        rootbtnPREREG.connect('clicked', importprereg_cb)
        rootlabelPREREG = Gtk.Label(label='')
        rootlabelPREREG.set_alignment(0, 0.5)
        rootlabelPREREG.set_markup(_('Import registration info from spreadsheet.'))
        rootbtnREG = Gtk.Button(_('Register'))
        rootbtnREG.connect('clicked', prereg_cb)
        rootlabelREG = Gtk.Label(label='')
        rootlabelREG.set_alignment(0, 0.5)
        rootlabelREG.set_markup(_('Register racer information and assign ID numbers.'))
        rootbtnCOMP = Gtk.Button(_('Compile'))
        rootbtnCOMP.connect('clicked', compreg_cb)
        rootlabelCOMP = Gtk.Label(label='')
        rootlabelCOMP.set_alignment(0, 0.5)
        rootlabelCOMP.set_markup(_('Compile registration file(s)'))
        rootbtnTIME = Gtk.Button(_('Time'))
        rootbtnTIME.connect('clicked', pretime_cb)
        rootlabelTIME = Gtk.Label(label='')
        rootlabelTIME.set_alignment(0, 0.5)
        rootlabelTIME.set_markup(_('Record race times on the day of the race.'))
        roottable.attach(rootbtnPREREG, 0, 1, 0, 1)
        roottable.attach(rootlabelPREREG, 1, 2, 0, 1)
        roottable.attach(rootbtnREG, 0, 1, 1, 2)
        roottable.attach(rootlabelREG, 1, 2, 1, 2)
        roottable.attach(rootbtnCOMP, 0, 1, 2, 3)
        roottable.attach(rootlabelCOMP, 1, 2, 2, 3)
        roottable.attach(rootbtnTIME, 0, 1, 3, 4)
        roottable.attach(rootlabelTIME, 1, 2, 3, 4)
        rootframe.add(roottable)
        ### Buttons
        roothbox = Gtk.HBox(True, 0)
        rootbtnQUIT = GtkStockButton('close',_("Quit"))
        rootbtnQUIT.connect('clicked', Gtk.main_quit)
        roothbox.pack_start(rootbtnQUIT, False, False, 5)
        #Vbox
        rootvbox = Gtk.VBox(False, 0)
        btnhalign = Gtk.Alignment.new(1, 0, 0, 0)
        btnhalign.add(roothbox)
        rootvbox.pack_start(mb, False, False, 0)
        rootvbox.pack_start(rootframe, True, True, 0)
        rootvbox.pack_start(btnhalign, False, False, 5)
        self.add(rootvbox)
        self.show_all()
