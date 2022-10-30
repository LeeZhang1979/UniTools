#!/usr/bin/env python3
# coding=utf-8

'''线束表转化'''

import os
from pyexpat.errors import XML_ERROR_BAD_CHAR_REF 
import sys 
import math
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPalette,QPixmap, QIcon
from PyQt5.QtWidgets import QMainWindow,QMessageBox,QFileDialog

from PyQt5 import QtSql
from PyQt5.QtSql import QSqlQuery

from openpyxl import load_workbook,Workbook 
from openpyxl.styles import Alignment  
from openpyxl.utils import get_column_letter

import warnings
warnings.filterwarnings('ignore')

BASE_DIR= os.path.dirname(os.path.dirname(os.path.abspath(__file__) ) )
sys.path.append( BASE_DIR  )    
from ui.Ui_WireConvertForm import Ui_WireConvertForm 

class WireConvertForm(QMainWindow,Ui_WireConvertForm):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setupUiEx()
        self.addConnect()
        
        self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
        self.db.setDatabaseName(os.path.join(BASE_DIR,'db\\mdm.db'))

        self.initData()

    def closeEvent(self, QCloseEvent):   
        if self.db.isOpen:
            self.db.close()            

    def __wireaccSql(self,classtype):
        sql = 'select classtype, \
            spec, \
            itemno, \
            itemname, \
            unit \
            from wireacc \
            where classtype=\''       
        sql += classtype 
        sql += '\'' 
        return sql

    def __oricablelistSql(self,classtype):
        sql = 'select classtype, \
            spec, \
            itemno, \
            itemname, \
            voltagelevel, \
            voltagelevel2, \
            remark \
            from oricablelist \
            where classtype=\''       
        sql += classtype 
        sql += '\'' 
        return sql


    def setupUiEx(self):
        icon = QIcon()
        appPath=os.path.join(BASE_DIR,u'res\\imgs\\WireConvert.ico')
        icon.addPixmap(QPixmap(appPath))
        self.setWindowIcon(icon)

    def addConnect(self):
        self.btnIFile.clicked.connect(self.btnIFileClick)
        self.btnOFile.clicked.connect(self.btnOFileClick)
        self.btnStart.clicked.connect(self.btnStartClick)

    def initData(self): 
        if not self.db.isOpen():
            if not self.db.open():               
                QMessageBox.critical(self, '线束表转化', self.db.lastError().text())
                return

    def btnIFileClick(self):
        fNames= QFileDialog.getOpenFileName(self,'线束表转化', '/','Excel File (*.xlsx)')
        if not fNames[0]:
            return
        if self.lineEOFile.text() == fNames[0]:
            QMessageBox.critical(self,'线束表转化', '输入与输出文件不能为同一文件') 
            self.lineEIFile.setFocus()
            return
        self.lineEIFile.setText(fNames[0])
    
    def btnOFileClick(self):
        fNames= QFileDialog.getSaveFileName(self,'线束表转化', '/','Excel File (*.xlsx)')
        if not fNames[0]:
            return
        
        if self.lineEIFile.text() == fNames[0]:
            QMessageBox.critical(self,'线束表转化', '输入与输出文件不能为同一文件') 
            self.lineEOFile.setFocus()
            return
        self.lineEOFile.setText(fNames[0])
    
    def btnStartClick(self):

        self.lineEIFile.setText('C:\\Users\\lezha\\Desktop\\线束表转化.xlsx')
        self.lineEOFile.setText('C:\\Users\\lezha\\Desktop\\TEST.xlsx')
        

        if self.lineEIFile.text() == self.lineEOFile.text() :
            QMessageBox.critical(self,'线束表转化', '输入与输出文件不能为同一文件') 
            self.lineEOFile.setFocus()
            return
        self.convert()

    def convert(self):
         
        wib = load_workbook(self.lineEIFile.text(),True) 
        wb = Workbook()    
         
        try:                                 
            sheetname=u'线束表转化'
            if not (wib.sheetnames.index(sheetname) >= 0):
                QMessageBox.warning(self,'线束表转化', '选择的文件:' + self.lineEIFile.text() + ',未包含配置指定的Sheet[' +sheetname + ']')
                wib.close()
                return                    
            wis=wib[sheetname]    
         
            iStarRow = self.spbStart.value()
            if iStarRow < 1:
                iStarRow = 1
             
            iEndRow = self.spbEnd.value()             
            if iEndRow ==0 or wis.max_row < iEndRow:
                iEndRow = wis.max_row

            if iStarRow > iEndRow:
                QMessageBox.warning(self,'线束表转化', '选择的文件:' + self.lineEIFile.text() + '及配置无需要转换的数据')
                wib.close()
                return     
            if not self.db.isOpen():
                if not self.db.open():               
                    QMessageBox.critical(self, '线束表转化', self.db.lastError().text())
                    wib.close
                    wb.close
                    return
      
            query = QSqlQuery() 
             
            ws = wb.create_sheet(u'线束辅料')      
            #管型预绝缘端子            
            if not query.exec(self.__wireaccSql(u'管型预绝缘端子')):
                QMessageBox.critical(self,'线束表转化', query.lastError().text())
                wib.close
                wb.close
                return
            iRow = 1 
            ws.cell(iRow,1).value = '类别'  
            ws.cell(iRow,2).value = '规格'
            ws.cell(iRow,3).value = '物料号'   
            ws.cell(iRow,4).value = '名称'  
            ws.cell(iRow,5).value = '单位'
            iGXStart = 2
            while query.next():    
                iRow += 1         
                ws.cell(iRow,1).value = str(query.value('classtype'))  
                ws.cell(iRow,2).value = str(query.value('spec'))   
                ws.cell(iRow,3).value = str(query.value('itemno'))   
                ws.cell(iRow,4).value = str(query.value('itemname'))  
                ws.cell(iRow,5).value = str(query.value('unit'))  
            iGXEnd = iRow
            #热缩管
            if not query.exec(self.__wireaccSql(u'热缩管')):
                QMessageBox.critical(self,'线束表转化', query.lastError().text())
                return
            iRow = iGXEnd 
            iYSGStart = iGXEnd +1
            while query.next():   
                iRow += 1                    
                ws.cell(iRow,1).value = str(query.value('classtype'))  
                ws.cell(iRow,2).value = str(query.value('spec'))   
                ws.cell(iRow,3).value = str(query.value('itemno'))   
                ws.cell(iRow,4).value = str(query.value('itemname'))  
                ws.cell(iRow,5).value = str(query.value('unit'))  
            iYSGEnd = iRow        
            #线标
            #OT端子
            #DT端子
            
            #原缆清单
            ws = wb.create_sheet(u'原缆清单')      
            strType = self.cmbType.currentText()
            if not query.exec(self.__oricablelistSql(strType)):
                QMessageBox.critical(self,'线束表转化', query.lastError().text())
                wib.close
                wb.close
                return
            iRow = 1
            
            ws.cell(iRow,1).value = '类别'  
            ws.cell(iRow,2).value = '规格'
            ws.cell(iRow,3).value = '物料号'   
            ws.cell(iRow,4).value = '名称'  
            ws.cell(iRow,5).value = '电压' 
            ws.cell(iRow,6).value = '电压图纸' 
            ws.cell(iRow,7).value = '备注'
            while query.next():    
                iRow += 1         
                ws.cell(iRow,1).value = str(query.value('classtype'))  
                ws.cell(iRow,2).value = str(query.value('spec'))   
                ws.cell(iRow,3).value = str(query.value('itemno'))   
                ws.cell(iRow,4).value = str(query.value('itemname'))  
                ws.cell(iRow,5).value = str(query.value('voltagelevel'))  
                ws.cell(iRow,6).value = str(query.value('voltagelevel2'))  
                ws.cell(iRow,7).value = str(query.value('remark'))  

            #开始转换
            ws = wb.create_sheet(u'线束表转化')      
            strCabledt = ''
            ws.column_dimensions['A'].width=5
            ws.column_dimensions['B'].width=10
            ws.column_dimensions['C'].width=5
            ws.column_dimensions['D'].width=20
            ws.column_dimensions['E'].width=5
            ws.column_dimensions['F'].width=5
            for iRow in range(iStarRow,iEndRow+1):     
                strTemp=''           
                if strCabledt == str(ws['A' + str(iRow)].value):   #电缆号
                    continue
                if ws['A' + str(iRow)].value is None:
                    strCabledt = ''
                else:
                    strCabledt = str(ws['A' + str(iRow)].value)      #电缆号
                if ws['F' + str(iRow)].value is None:
                    strTemp =''
                else:
                    strTemp = ws['F' + str(iRow)].value

                oCurRow = (iRow-iStarRow) * 14 + 1
                ws.cell(oCurRow,1).value = '1'               #序号 第一行
                #ws.cell(oCurRow,2).value = ''               #第二列、无物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oCurRow,4).value = u'电缆组件_' + strCabledt + '_' + strTemp  #第四列
                ws.cell(oCurRow,5).value = '1'               #第五列 
                ws.cell(oCurRow,6).value = 'PC'               #第六列 



                
           
            
            
            wb.save(self.lineEOFile.text())
            wb.close
            wib.close
            QMessageBox.information(self,'线束表转化','导出数据完成，文件名：' + self.lineEOFile.text())    
        except (NameError,ZeroDivisionError):
            wib.close
            wb.close
            QMessageBox.critical(self, '线束表转化', '变量名错误或除数为0')
        except OSError as reason:
            wib.close
            wb.close
            QMessageBox.critical(self, '线束表转化', str(reason))
        except TypeError as reason:
            wib.close
            wb.close
            QMessageBox.critical(self, '线束表转化', str(reason))
        except :
            wib.close
            wb.close
            QMessageBox.information(self,'线束表转化','导出数据文件失败')     