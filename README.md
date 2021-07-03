fstimer24h
==========

Fork supporting 24H donation runs.

See fstimer.org for more information and original version

fsTimer is written in Python3 and uses GTK3+ via PyGObject.

This fork allows 

* Excel import of registrants and Excel export to times
* displaying runner names when tracking
* send track information to REST API to allow online tracking
* Calculates donation sum

Dependencies (MacOS with python3 via brew):
* gi (``brew install pygobject3 --with-python3``)
* xlsxwriter (``pip3 install xlsxwriter``)
* openpyxl (``pip3 install openpyxl``) - https://openpyxl.readthedocs.io

Bug fixes
* Importing first excel sheet (instead of hard-coded sheet name)