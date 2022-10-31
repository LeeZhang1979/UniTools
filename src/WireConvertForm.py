#!/usr/bin/env python3
# coding=utf-8

'''线束表转化'''

from asyncio.format_helpers import _format_args_and_kwargs
import os
from pyexpat.errors import XML_ERROR_BAD_CHAR_REF 
import sys 
import math
from turtle import bgcolor
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPalette,QPixmap, QIcon
from PyQt5.QtWidgets import QMainWindow,QMessageBox,QFileDialog

from PyQt5 import QtSql
from PyQt5.QtSql import QSqlQuery

from openpyxl import load_workbook,Workbook 
from openpyxl.styles import Alignment,PatternFill,Color,Border,Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formatting.formatting import ConditionalFormattingList


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

        #开发测试用，发布时需要删除
        self.lineEIFile.setText('C:\\Users\\lezha\\Desktop\\线束表转化.xlsx')
        self.lineEOFile.setText('C:\\Users\\lezha\\Desktop\\TEST.xlsx')
        #开发测试用，发布时需要删除

        if self.lineEIFile.text() == self.lineEOFile.text() :
            QMessageBox.critical(self,'线束表转化', '输入与输出文件不能为同一文件') 
            self.lineEOFile.setFocus()
            return
        self.convert()

    def __getCoresFromProperty(self,s):
        if len(s) == 0:
            return ''
        if s.find("G") >= 0 :
            return s[0:s.find("G")]
        elif s.find("x") >= 0 :            
            return s[0:s.find("x")]
        return ''

    def __getRCode(self,s):
        if len(s) == 0:
            return ''
        if s.find("-") >= 0 :
            return s[s.find("-")+1:]
        return ''

    def __rTrimUnit(self,s):
        if len(s) == 0:
            return ''
        s = s.replace('mm²','')
        if s.find(' ') >= 0 :
            s = s[0:s.find(' ')] 
        if len(s)>2:
            if s[len(s)-2:len(s)] =='.0':
                s = s[0:len(s)-2]
        return s.strip()         

    def isNumber(self,s):
        if len(s) == 0:
            return False
        if s.count(".")==1:  #小数的判断
            if s[0] == "-":
                s=s[1:]
            if s[0] == ".":
                return False
            s=s.replace(".","")
            for i in s:
                if i not in "0123456789":
                    return False
            return True
        elif s.count(".")==0:  #整数的判断
            if s[0]=="-":
                s=s[1:]
            for i in s:
                if i not in "0123456789":
                    return False
            return True
        else:
            return False

    def convert(self):
         
        wib = load_workbook(filename=self.lineEIFile.text(),read_only=True,data_only=True) 
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
            ws = wb.active 
            ws.title = u'线束表转化' 
            
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
            strGXDataRange = '=线束辅料!$D$' + str(iGXStart) +':$D$' + str(iGXEnd)
            dvGX = DataValidation(type='list',formula1=strGXDataRange,allowBlank=True,prompt=u'选择绝缘端子')
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
            strYSGataRange = '=线束辅料!$D$' + str(iYSGStart) +':$D$' + str(iYSGEnd)
            dvYSG = DataValidation(type='list',formula1=strYSGataRange,allowBlank=True,prompt=u'选择热缩管')
            #线标
            if not query.exec(self.__wireaccSql(u'线标')):
                QMessageBox.critical(self,'线束表转化', query.lastError().text())
                return
            iRow = iYSGEnd 
            iXBStart = iYSGEnd + 1
            while query.next():   
                iRow += 1                    
                ws.cell(iRow,1).value = str(query.value('classtype'))  
                ws.cell(iRow,2).value = str(query.value('spec'))   
                ws.cell(iRow,3).value = str(query.value('itemno'))   
                ws.cell(iRow,4).value = str(query.value('itemname'))  
                ws.cell(iRow,5).value = str(query.value('unit'))  
            iXBEnd = iRow                    
            strXBDataRange = '=线束辅料!$D$' + str(iXBStart) +':$D$' + str(iXBEnd)
            dvXB = DataValidation(type='list',formula1=strXBDataRange,allowBlank=True,prompt=u'选择线标')
            #扎带
            if not query.exec(self.__wireaccSql(u'扎带')):
                QMessageBox.critical(self,'线束表转化', query.lastError().text())
                return
            iRow = iXBEnd 
            iZDStart = iYSGEnd + 1
            while query.next():   
                iRow += 1                    
                ws.cell(iRow,1).value = str(query.value('classtype'))  
                ws.cell(iRow,2).value = str(query.value('spec'))   
                ws.cell(iRow,3).value = str(query.value('itemno'))   
                ws.cell(iRow,4).value = str(query.value('itemname'))  
                ws.cell(iRow,5).value = str(query.value('unit'))  
            iZDEnd = iRow                    
            strZDDataRange = '=线束辅料!$D$' + str(iZDStart) +':$D$' + str(iZDEnd)
            dvZD = DataValidation(type='list',formula1=strZDDataRange,allowBlank=True,prompt=u'选择扎带')
            
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
            ws = wb[u'线束表转化']
            ws.column_dimensions['A'].width=5
            ws.column_dimensions['B'].width=20
            ws.column_dimensions['C'].width=5
            ws.column_dimensions['D'].width=50
            ws.column_dimensions['E'].width=20
            ws.column_dimensions['F'].width=10 

            # 红色 #E72512
            # 黑色 #0A0A0D
            # 白色 #FFFFFF
            # 黄色 #F5FF00
            black = PatternFill(fill_type='solid',bgColor='0a0a0d') 
            white=PatternFill(fill_type='solid',bgColor='ffffff') 
            red=PatternFill(fill_type='solid',bgColor='e72512') 
            yellow=PatternFill(fill_type='solid',bgColor='f5ff00') 
            fontwhite= Font(name=u'等线',size=9,color='ffffff')
            fontblack=Font(name=u'等线',size=9,color='0a0a0d') 
            strCabledt = ''
            iCount = 0
            iYLQDDV = 1
            for iRow in range(iStarRow,iEndRow+1):    
                if strCabledt == str(wis['A' + str(iRow)].value):   #电缆号
                    continue
                iCount += 1
                if wis['A' + str(iRow)].value is None:
                    strCabledt = ''
                else:
                    strCabledt = str(wis['A' + str(iRow)].value)                         
                if wis['B' + str(iRow)].value is None:        #芯线数与线径
                    strProperty =''
                else:
                    strProperty = wis['B' + str(iRow)].value
                if wis['D' + str(iRow)].value is None:        #外径
                    strLength =''
                else:
                    strLength = wis['D' + str(iRow)].value
                if wis['F' + str(iRow)].value is None:        #起始元件功能
                    strStarDevFun =''
                else:
                    strStarDevFun = wis['F' + str(iRow)].value                
                if wis['G' + str(iRow)].value is None:        #起始元件
                    strStarDevName =''
                else:
                    strStarDevName = wis['G' + str(iRow)].value
                  
                oCurRow = (iCount - 1) * 14 + 1          # 第一行  电缆组件
                ws.cell(oCurRow,1).value = '1'               #级别号
                #ws.cell(oCurRow,2).value = ''               #第二列、无物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oCurRow,4).value = u'电缆组件_' + strCabledt + '_' + strStarDevFun  #第四列
                ws.cell(oCurRow,5).value = '1'               #第五列 
                ws.cell(oCurRow,6).value = 'PC'               #第六列 
                ws.cell(oCurRow,7).value = ''                #第七列 
                 
                oCurRow += 1         
                oDupRow = oCurRow + 12                       # 第二行 and 第十四行   绝缘端子
                ws.cell(oCurRow,1).value = '2'               #级别号
                ws.cell(oDupRow,1).value = '2'                               
                ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iGXStart) +':$D$' + str(iGXEnd) +',线束辅料!$C$' + str(iGXStart) +':$C$' + str(iGXEnd) + '),2,FALSE)'               #第二列、无物料号
                ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iGXStart) +':$D$' + str(iGXEnd) +',线束辅料!$C$' + str(iGXStart) +':$C$' + str(iGXEnd) + '),2,FALSE)'               #第二列、无物料号
                #ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                #ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oDupRow,3).value = 'L'               #第三列 
                dvGX.add(ws.cell(oCurRow,4))
                dvGX.add(ws.cell(oDupRow,4))
                #ws.cell(oCurRow,4).value =  ''  #第四列
                #ws.cell(oDupRow,4).value =  ''  #第四列
                strTemp = self.__getCoresFromProperty(strProperty)
                ws.cell(oCurRow,5).value = strTemp               #第五列 
                ws.cell(oDupRow,5).value = strTemp              #第五列 
                ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iGXStart) +':$D$' + str(iGXEnd) +',线束辅料!$E$' + str(iGXStart) +':$E$' + str(iGXEnd) + '),2,FALSE)'               #第六列 
                ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iGXStart) +':$D$' + str(iGXEnd) +',线束辅料!$E$' + str(iGXStart) +':$E$' + str(iGXEnd) + '),2,FALSE)'               #第六列 
                #ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                #ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                ws.cell(oCurRow,7).value = ''                #第七列 
                ws.cell(oDupRow,7).value = ''                #第七列 
                 
                oCurRow += 1                                # 第三行 and 第十三行   热缩管
                oDupRow = oCurRow + 10
                ws.cell(oCurRow,1).value = '2'               #级别号
                ws.cell(oDupRow,1).value = '2'               #级别号
                ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iYSGStart) +':$D$' + str(iYSGEnd) +',线束辅料!$C$' + str(iYSGStart) +':$C$' + str(iYSGEnd) + '),2,FALSE)'               #第二列、无物料号
                ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iYSGStart) +':$D$' + str(iYSGEnd) +',线束辅料!$C$' + str(iYSGStart) +':$C$' + str(iYSGEnd) + '),2,FALSE)'               #第二列、无物料号
                #ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                #ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oDupRow,3).value = 'L'               #第三列 
                dvYSG.add(ws.cell(oCurRow,4))
                dvYSG.add(ws.cell(oDupRow,4))
                #ws.cell(oCurRow,4).value =  ''  #第四列
                #ws.cell(oDupRow,4).value =  ''  #第四列
                ws.cell(oCurRow,5).value = strLength               #第五列 
                ws.cell(oDupRow,5).value = strLength               #第五列 
                ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iYSGStart) +':$D$' + str(iYSGEnd) +',线束辅料!$E$' + str(iYSGStart) +':$E$' + str(iYSGEnd) + '),2,FALSE)'               #第六列 
                ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iYSGStart) +':$D$' + str(iYSGEnd) +',线束辅料!$E$' + str(iYSGStart) +':$E$' + str(iYSGEnd) + '),2,FALSE)'               #第六列 
                #ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                #ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                ws.cell(oCurRow,7).value = ''                #第七列 
                ws.cell(oDupRow,7).value = ''                #第七列 


                oCurRow += 1                            # 第四行 and 第十行  线标 标签 
                oDupRow = oCurRow + 6
                ws.cell(oCurRow,1).value = '2'               #级别号
                ws.cell(oDupRow,1).value = '2'               #级别号
                ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iXBStart) +':$D$' + str(iXBEnd) +',线束辅料!$C$' + str(iXBStart) +':$C$' + str(iXBEnd) + '),2,FALSE)'               #第二列、无物料号
                ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iXBStart) +':$D$' + str(iXBEnd) +',线束辅料!$C$' + str(iXBStart) +':$C$' + str(iXBEnd) + '),2,FALSE)'               #第二列、无物料号
                #ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                #ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oDupRow,3).value = 'L'               #第三列 
                dvXB.add(ws.cell(oCurRow,4))
                dvXB.add(ws.cell(oDupRow,4))
                ws.cell(oCurRow,4).value =  ''  #第四列
                ws.cell(oDupRow,4).value =  ''  #第四列
                ws.cell(oCurRow,5).value = '1'               #第五列 
                ws.cell(oDupRow,5).value = '1'               #第五列 
                ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iXBStart) +':$D$' + str(iXBEnd) +',线束辅料!$E$' + str(iXBStart) +':$E$' + str(iXBEnd) + '),2,FALSE)'               #第六列 
                ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iXBStart) +':$D$' + str(iXBEnd) +',线束辅料!$E$' + str(iXBStart) +':$E$' + str(iXBEnd) + '),2,FALSE)'               #第六列 
                #ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                #ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                ws.cell(oCurRow,7).value = ''                #第七列 
                ws.cell(oDupRow,7).value = ''                #第七列 
                
                #第五、六， 十一，十二 为标签
                oCurRow += 1                            # 第五、六， 十一，十二 为标签
                oDupRow = oCurRow + 6                
                ws.merge_cells(start_row=oCurRow,start_column=4,end_row=oCurRow+1,end_column=4)
                ws.cell(oCurRow, 4).value = strCabledt + '\n' + self.__getRCode(strStarDevName) + '\n' + strStarDevFun  #第四列                
                ws.cell(oCurRow, 4).alignment = Alignment(horizontal='center',vertical='center',wrapText=True)
                rule1  = FormulaRule(formula=['$B' + str(oCurRow -1) +'="DFB00052144"'],fill=black,stopIfTrue=True,font=fontwhite)
                rule2  = FormulaRule(formula=['$B' + str(oCurRow -1) +'="DFB00052145"'],fill=white,stopIfTrue=True,font=fontblack)
                rule3  = FormulaRule(formula=['$B' + str(oCurRow -1) +'="DFB00052146"'],fill=red,stopIfTrue=True,font=fontblack)
                rule4  = FormulaRule(formula=['$B' + str(oCurRow -1) +'<>""'],fill=yellow,stopIfTrue=True,font=fontblack)
                
                ws.conditional_formatting.add('$D$' + str(oCurRow),rule1)                
                ws.conditional_formatting.add('$D$' + str(oCurRow),rule2)
                ws.conditional_formatting.add('$D$' + str(oCurRow),rule3)
                ws.conditional_formatting.add('$D$' + str(oCurRow),rule4)

                ws.merge_cells(start_row=oDupRow,start_column=4,end_row=oDupRow+1,end_column=4)
                ws.cell(oDupRow, 4).value = strCabledt + '\n' + self.__getRCode(strStarDevName) + '\n' + strStarDevFun  #第四列   
                ws.cell(oDupRow, 4).alignment = Alignment(horizontal='center',vertical='center',wrapText=True)
                rule1  = FormulaRule(formula=['$B' + str(oDupRow -1) +'="DFB00052144"'],fill=black,stopIfTrue=True,font=fontwhite)
                rule2  = FormulaRule(formula=['$B' + str(oDupRow -1) +'="DFB00052145"'],fill=white,stopIfTrue=True,font=fontblack)
                rule3  = FormulaRule(formula=['$B' + str(oDupRow -1) +'="DFB00052146"'],fill=red,stopIfTrue=True,font=fontblack)
                rule4  = FormulaRule(formula=['$B' + str(oDupRow -1) +'<>""'],fill=yellow,stopIfTrue=True,font=fontblack)
                ws.conditional_formatting.add('$D$' + str(oDupRow),rule1)
                ws.conditional_formatting.add('$D$' + str(oDupRow),rule2)
                ws.conditional_formatting.add('$D$' + str(oDupRow),rule3)
                ws.conditional_formatting.add('$D$' + str(oDupRow),rule4)

                oCurRow += 2                            # #第七，第九  扎带 
                oDupRow = oCurRow + 2
                ws.cell(oCurRow,1).value = '2'               #级别号
                ws.cell(oDupRow,1).value = '2'               #级别号
                ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iZDStart) +':$D$' + str(iZDEnd) +',线束辅料!$C$' + str(iZDStart) +':$C$' + str(iZDEnd) + '),2,FALSE)'               #第二列、无物料号
                ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iZDStart) +':$D$' + str(iZDEnd) +',线束辅料!$C$' + str(iZDStart) +':$C$' + str(iZDEnd) + '),2,FALSE)'               #第二列、无物料号
                #ws.cell(oCurRow,2).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                #ws.cell(oDupRow,2).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$C:$C),2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列 
                ws.cell(oDupRow,3).value = 'L'               #第三列 
                dvZD.add(ws.cell(oCurRow,4))
                dvZD.add(ws.cell(oDupRow,4)) 
                ws.cell(oCurRow,4).value = u'扎带L188φ48H1.3D4.8'       #第四列
                ws.cell(oDupRow,4).value= u'扎带L188φ48H1.3D4.8'        #第四列      
                strTemp = self.__rTrimUnit(strLength)        
                if self.isNumber(strTemp): 
                    if float(strTemp) <= 25.0:     
                        ws.cell(oCurRow,4).value =  u'电缆扎带 Cable tie 2.5×100 MM'
                        ws.cell(oDupRow,4).value =  u'电缆扎带 Cable tie 2.5×100 MM'  #第四列 
                ws.cell(oCurRow,5).value = '1'               #第五列 
                ws.cell(oDupRow,5).value = '1'               #第五列 
                ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D$' + str(iZDStart) +':$D$' + str(iZDEnd) +',线束辅料!$E$' + str(iZDStart) +':$E$' + str(iZDEnd) + '),2,FALSE)'               #第六列 
                ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D$' + str(iZDStart) +':$D$' + str(iZDEnd) +',线束辅料!$E$' + str(iZDStart) +':$E$' + str(iZDEnd) + '),2,FALSE)'               #第六列 
                #ws.cell(oCurRow,6).value = '=VLOOKUP(D' + str(oCurRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                #ws.cell(oDupRow,6).value = '=VLOOKUP(D' + str(oDupRow) +',IF({1,0},线束辅料!$D:$D,线束辅料!$E:$E),2,FALSE)'               #单位
                ws.cell(oCurRow,7).value = ''                #第七列 
                ws.cell(oDupRow,7).value = ''                #第七列  
                
                oCurRow += 1            #第八行
                ws.cell(oCurRow,1).value = '2'               #级别号
                ws.cell(oCurRow,2).value = ''               #第二列、物料号
                ws.cell(oCurRow,3).value = 'L'               #第三列                 
                #ws.cell(oCurRow,4).value = '=VLOOKUP(B' + str(oCurRow) +',IF({1,0},原缆清单!$C:$C,原缆清单!$D:$D),2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,4).value = '=VLOOKUP(B' + str(oCurRow) +',原缆清单!$C:$D,2,FALSE)'               #第二列、物料号
                ws.cell(oCurRow,5).value = '1'               #第五列 
                ws.cell(oCurRow,6).value = 'M'               #第六列 
                ws.cell(oCurRow,7).value = strLength                #第七列 

                sql = self.__oricablelistSql(strType)
                sql += ' and spec like \'' + self.__rTrimUnit(strProperty) + '%\''                
                if not query.exec(sql):
                    QMessageBox.critical(self,'线束表转化', query.lastError().text())
                    wib.close
                    wb.close
                    return
                iTemp = 0 
                strFormula=''
                while query.next():   
                    if len(strFormula) > 0:
                        strFormula +=','
                    else:
                        ws.cell(oCurRow,2).value = str(query.value('itemno'))
                    strFormula += '\'' 
                    strFormula += str(query.value('itemno'))

                if strFormula.find(',')>0:
                    strFormula = '"' + strFormula + '"'
                    dvYL = DataValidation(type='list',formula1=strFormula,allowBlank=True,prompt=u'原缆清单')
                    dvYL.add(ws.cell(oCurRow,2))
                    ws.add_data_validation(dvYL)
                
            ws.add_data_validation(dvGX)
            ws.add_data_validation(dvYSG)
            ws.add_data_validation(dvXB)
            ws.add_data_validation(dvZD)

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