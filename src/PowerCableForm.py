#!/usr/bin/env python3
# coding=utf-8

'''动力电缆计算器'''

import os 
import sys 
from PyQt5 import QtCore, QtGui, QtWidgets 
from PyQt5.QtGui import QPalette, QPixmap, QIcon
from PyQt5.QtWidgets import QMainWindow,QMessageBox,QTableWidgetItem,QFileDialog

from PyQt5 import QtSql
from PyQt5.QtSql import QSqlQuery

from openpyxl import Workbook 

import warnings
warnings.filterwarnings('ignore')

BASE_DIR= os.path.dirname(os.path.dirname(os.path.abspath(__file__) ) )
sys.path.append( BASE_DIR  )    
from ui.Ui_PowerCableForm import Ui_PowerCableForm 

class PowerCableForm(QMainWindow,Ui_PowerCableForm):
    
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        #self.setupUiEx()
        #self.addConnect()
    
    def closeEvent(self, QCloseEvent):
        return 

    def setupUiEx(self):
        palette = QPalette()
        icon = QIcon()
        appPath=os.path.join(BASE_DIR,u'res\\imgs\\powercable.ico')
        icon.addPixmap(QPixmap(appPath))
        self.setWindowIcon(icon)
        

    def addConnect(self):
        #self.btnUpdate.clicked.connect(self.updateClick)
        return
    
    def initdata(self): 
        return