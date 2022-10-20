#!/usr/bin/env python3
# coding=utf-8

'''动力电缆计算器'''

import os 
import sys 
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPalette, QPixmap, QIcon, QImage, QIntValidator, QDoubleValidator, QRegExpValidator
from PyQt5.QtWidgets import QMainWindow,QMessageBox,QTableWidgetItem,QFileDialog
from PyQt5.QtCore import QRegExp

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
        self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName(os.path.join(BASE_DIR,'db\\mdm.db'))
        if not self.db.isOpen():
            if not self.db.open():               
                QMessageBox.critical(self, 'MDM', self.db.lastError().text())
                return
        self.initData()

    def closeEvent(self, QCloseEvent):   
        if self.db.isOpen:
            self.db.close()            
    
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
        #发动机单绕组电流(A) 5位以内>0整数
        intReg = QRegExp('^[1-9][0-9]{1,4}')
        regExpValidator = QRegExpValidator(intReg)
        self.lineEEC.setValidator(regExpValidator)
        #绕组数 3位以内>0整数
        intReg = QRegExp('^[1-9][0-9]{1,2}')
        regExpValidator = QRegExpValidator(intReg) 
        self.lineEWings.setValidator(regExpValidator) 
        #标称截面积(㎜²) 5位以内>=0浮点数 +2位小数
        floatReg = QRegExp('^([0]|[1-9][0-9]{0,4})(?:\.\d{1,2})?$|(^\t?$)')
        regExpValidator = QRegExpValidator(floatReg)
        self.lineECS.setValidator(regExpValidator) 
        #环温(°C)  4位以内>=0整数
        intReg = QRegExp('^([0]|[1-9][0-9]{1,3})')
        regExpValidator = QRegExpValidator(intReg)
        self.lineEAmbientT.setValidator(regExpValidator)  
        #护套耐温(°C) 4位以内>=0整数
        intReg = QRegExp('^([0]|[1-9][0-9]{1,3})')
        regExpValidator = QRegExpValidator(intReg)  
        self.lineESTR.setValidator(regExpValidator)   
        #托盘/梯架数 1,2,3
        intReg = QRegExp('[1-3]?')
        regExpValidator = QRegExpValidator(intReg)  
        self.lineENumber.setValidator(regExpValidator)  
        #三相回路数 1,2,3
        intReg = QRegExp('[1-3]?')
        regExpValidator = QRegExpValidator(intReg)  
        self.lineECircuits.setValidator(regExpValidator)   

        self.rbTouch.clicked.connect(self.refreshLayingPic)
        self.rbSpace.clicked.connect(self.refreshLayingPic)
        self.cmbLayingType.currentIndexChanged.connect(self.refreshLayingPic)
        self.btnCalculation.clicked.connect(self.btnCalculationClick)
    
    def initData(self): 
        self.cleanResult()

    def cleanResult(self):
        #载流量值
        self.lblOECNum.setText('')
        #敷设系数
        self.lblOACF.setText('')
        #折算系数
        self.lblOCFnum.setText('')
        #单绕组电缆根数
        self.lblOSNum.setText('')
        #向上取整
        self.lblOSNumUp.setText('')   
        #单相电缆根数
        self.lblOSDNum.setText('')     
        #单绕组电流余量
        self.lblOECLeftNum.setText('')
        #余量百分比
        self.lblOECPer.setText('')

    def btnCalculationClick(self):
        self.cleanResult()
        
        #电缆类型
        self.cmbPCType.currentText()
        #发动机单绕组电流(A)
        self.lineEEC.text() 
        #绕组数
        self.lineEWings.text()
        #导体
        self.cmbConductor.currentText()
        #标称截面积(㎜²)
        self.lineECS.text()
        #绝缘材料
        self.cmbMaterial.currentText()
        #环温(°C)
        self.lineEAmbientT.text()
        #护套耐温(°C)
        self.lineESTR.text()
        #电缆芯数及排列
        self.cmbCrossType.currentText()
        #敷设方式
        self.cmbLayingType.currentText()
        #相互接触
        self.rbTouch.isChecked()
        self.rbTouch.text()
        #有间距
        self.rbSpace.isChecked()
        self.rbSpace.text()        
        #托盘/梯架数
        self.lineENumber.text()
        #三相回路数
        self.lineECircuits.text()

        #列表
        self.tblList.items().clear()


