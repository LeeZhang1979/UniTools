#!/usr/bin/env python3
# coding=utf-8

'''MainWindow:程序启动主窗口 ''' 
 
import os
import subprocess
import sys 

from PyQt5.QtWidgets import QApplication,QMainWindow,QMessageBox
from PyQt5.QtGui import QPalette, QBrush, QPixmap, QIcon
from src.MDMForm import MDMForm 
from src.PowerCableForm import PowerCableForm
from src.WireConvertForm import WireConvertForm
BASE_DIR= os.path.dirname(os.path.dirname(os.path.abspath(__file__) ) )
sys.path.append( BASE_DIR  )   
from ui.Ui_MainWindow import Ui_MainWindow 
from conf.AppConfigure import AppConfigure,Appconfig

class MainWindow(Ui_MainWindow, QMainWindow):
    appEvent = AppConfigure()
        
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setUiEx()
        self.addEvent() 
    
    def setUiEx(self):
        palette = QPalette()
        appPath=os.path.join(BASE_DIR,u'res\\imgs\\small.png')
        palette.setBrush(QPalette.Background, QBrush(QPixmap(appPath)))        
        self.setPalette(palette)
        icon = QIcon()
        appPath=os.path.join(BASE_DIR,u'res\\icon\\UniTools.ico')
        icon.addPixmap(QPixmap(appPath))
        self.setWindowIcon(icon)
        self.btnTool01.setStyleSheet('QPushButton{ background:orange;color:white; \
                                    font-size:10px;border-radius: 16px;font-family: Microsoft YaHei UI;} \
                                    QPushButton:pressed{ background:black; }')
        self.btnMDMConf.setStyleSheet('QPushButton{ background:orange;color:white; \
                                    font-size:10px;border-radius: 16px;font-family: Microsoft YaHei UI;} \
                                    QPushButton:pressed{ background:black; }')

    def addEvent(self):
        #conffile = os.path.join(BASE_DIR,'data\\AppConfigs.xml')
        #self.appEvent.loadConf(conffile)
        self.btnTool01.clicked.connect(self.tool01Click)
        self.btnMDMConf.clicked.connect(self.menuConfigure)
        self.actionConfigure.triggered.connect(self.menuConfigure)
        self.actionPowerCableCal.triggered.connect(self.menuPowerCableCal)
        self.actionWireConvert.triggered.connect(self.menuWireConvert)
        self.actionPowerConvert.triggered.connect(self.menuPowerConvert)
        self.actionCableMSTOptimizer.triggered.connect(self.menuCableMSTOptimizer)

    def tool01Click(self):

        appPath=os.path.join(BASE_DIR,u'res\\风力发电机组短路电流计算.xlsx')
        #subprocess.run(appPath)
        os.system('start ' + appPath)
        #os.startfile(appPath)
        #QMessageBox.information(self,"提示框","复制成功")
    
    def menuCableMSTOptimizer(self):
        try:
            appPath=os.path.join(BASE_DIR,u'CableMSTOptimizer.exe')
            subprocess.Popen(appPath)
            #os.system('start ' + appPath)         
        except PermissionError as reason : 
            QMessageBox.critical(self,'调用外部程序失败',str(reason))  
        except : 
            QMessageBox.critical(self,'调用外部程序失败','调用外部程序[CableMSTOptimizer.exe]失败!')  


    def menuPowerConvert(self):
        try:
            appPath=os.path.join(BASE_DIR,u'PowerConverter.exe')
            subprocess.Popen(appPath)
            #os.system('start ' + appPath)         
        except PermissionError as reason : 
            QMessageBox.critical(self,'调用外部程序失败',str(reason))  
        except : 
            QMessageBox.critical(self,'调用外部程序失败','调用外部程序[PowerConverter.exe]失败!')  

    def menuConfigure(self):
        self.mdmWin = MDMForm()
        self.mdmWin.show()

    def menuPowerCableCal(self):
        self.pcWin = PowerCableForm()       
        self.pcWin.show()
    
    def menuWireConvert(self):
        self.wcForm = WireConvertForm()
        self.wcForm.show()