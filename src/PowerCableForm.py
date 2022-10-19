#!/usr/bin/env python3
# coding=utf-8

'''动力电缆计算器'''

import os 
import sys 
from PyQt5 import QtCore, QtGui, QtWidgets 
from PyQt5.QtGui import QPalette, QPixmap, QIcon, QImage
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
        self.setupUiEx()
        self.addConnect()
    
    def closeEvent(self, QCloseEvent):
        return 

    def setupUiEx(self):
        palette = QPalette()
        icon = QIcon()
        appPath=os.path.join(BASE_DIR,u'res\\imgs\\powercable.ico')
        icon.addPixmap(QPixmap(appPath))
        self.setWindowIcon(icon)

        self.cmbCrossType.setIconSize(QtCore.QSize(53, 64))
        self.cmbCrossType.setItemIcon(0, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\multi_2core.png')))
        self.cmbCrossType.setItemIcon(1, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\multi_3core.png')))
        self.cmbCrossType.setItemIcon(2, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\single_2core_t.png')))
        self.cmbCrossType.setItemIcon(3, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\single_3core_t.png')))
        self.cmbCrossType.setItemIcon(4, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\signle_3core_ft.png')))
        self.cmbCrossType.setItemIcon(5, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\single_3core_hs.png')))
        self.cmbCrossType.setItemIcon(6, QIcon(os.path.join(BASE_DIR,u'res\\imgs\\single_3core_vs.png')))

        self.refreshLayingPic()

    def refreshLayingPic(self):
        filename=''
        if self.cmbLayingType.currentIndex() == 0:
            if self.rbTouch.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\tray_t_h.png') 
            elif self.rbSpace.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\tray_s_h.png') 
        elif self.cmbLayingType.currentIndex() == 1:
            if self.rbTouch.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\tray_t_v.png') 
            elif self.rbSpace.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\tray_s_v.png') 
        elif self.cmbLayingType.currentIndex() == 2:
            if self.rbTouch.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\lorc_t.png') 
            elif self.rbSpace.isChecked():
                filename=os.path.join(BASE_DIR,u'res\\imgs\\lorc_s.png')  
        qimage = QImage(filename)
        self.lblLayingImg.setPixmap(QPixmap.fromImage(qimage))                
        self.lblLayingImg.setScaledContents(True)

    def addConnect(self):

        self.rbTouch.clicked.connect(self.refreshLayingPic)
        self.rbSpace.clicked.connect(self.refreshLayingPic)
        self.cmbLayingType.currentIndexChanged.connect(self.refreshLayingPic)
        
        
        return
    
    def initdata(self): 
        return