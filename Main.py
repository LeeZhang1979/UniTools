#!/usr/bin/env python3
# coding=utf-8

import os
import sys

from PyQt5 import QtCore

from PyQt5.QtWidgets import QApplication

BASE_DIR= os.path.dirname(os.path.abspath(__file__) )  
sys.path.append( BASE_DIR  ) 
from src.MainWindow import MainWindow


def main(): 
    
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv) 
     
    mwin= MainWindow()
    mwin.show()
   
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
