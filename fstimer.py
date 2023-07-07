#!/usr/bin/env python3

import fstimer.fslogger

import fstimer.timer
from gi.repository import Gtk
import gettext

gettext.install('base', localedir='./fstimer/locales')

def main():
    pytimer = fstimer.timer.PyTimer()
    Gtk.main()

if __name__ == '__main__':
    main()
