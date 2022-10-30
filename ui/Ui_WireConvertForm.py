# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'c:\Source\UniTools\ui\WireConvertForm.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_WireConvertForm(object):
    def setupUi(self, WireConvertForm):
        WireConvertForm.setObjectName("WireConvertForm")
        WireConvertForm.setWindowModality(QtCore.Qt.NonModal)
        WireConvertForm.resize(320, 160)
        WireConvertForm.setMinimumSize(QtCore.QSize(320, 160))
        WireConvertForm.setMaximumSize(QtCore.QSize(640, 320))
        WireConvertForm.setBaseSize(QtCore.QSize(320, 160))
        self.label01 = QtWidgets.QLabel(WireConvertForm)
        self.label01.setGeometry(QtCore.QRect(10, 10, 61, 20))
        self.label01.setObjectName("label01")
        self.lineEIFile = QtWidgets.QLineEdit(WireConvertForm)
        self.lineEIFile.setGeometry(QtCore.QRect(79, 10, 201, 20))
        self.lineEIFile.setMaxLength(256)
        self.lineEIFile.setObjectName("lineEIFile")
        self.btnStart = QtWidgets.QPushButton(WireConvertForm)
        self.btnStart.setGeometry(QtCore.QRect(20, 130, 75, 20))
        self.btnStart.setObjectName("btnStart")
        self.btnCancel = QtWidgets.QPushButton(WireConvertForm)
        self.btnCancel.setGeometry(QtCore.QRect(220, 130, 75, 20))
        self.btnCancel.setObjectName("btnCancel")
        self.btnIFile = QtWidgets.QToolButton(WireConvertForm)
        self.btnIFile.setGeometry(QtCore.QRect(280, 10, 30, 20))
        self.btnIFile.setObjectName("btnIFile")
        self.spbStart = QtWidgets.QSpinBox(WireConvertForm)
        self.spbStart.setGeometry(QtCore.QRect(140, 40, 41, 22))
        self.spbStart.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.spbStart.setProperty("value", 2)
        self.spbStart.setObjectName("spbStart")
        self.label02 = QtWidgets.QLabel(WireConvertForm)
        self.label02.setGeometry(QtCore.QRect(70, 40, 61, 22))
        self.label02.setObjectName("label02")
        self.label03 = QtWidgets.QLabel(WireConvertForm)
        self.label03.setGeometry(QtCore.QRect(180, 40, 91, 22))
        self.label03.setObjectName("label03")
        self.spbEnd = QtWidgets.QSpinBox(WireConvertForm)
        self.spbEnd.setGeometry(QtCore.QRect(270, 40, 41, 22))
        self.spbEnd.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.spbEnd.setObjectName("spbEnd")
        self.label04 = QtWidgets.QLabel(WireConvertForm)
        self.label04.setGeometry(QtCore.QRect(10, 100, 81, 20))
        self.label04.setObjectName("label04")
        self.lblOTotal = QtWidgets.QLabel(WireConvertForm)
        self.lblOTotal.setGeometry(QtCore.QRect(90, 100, 71, 20))
        self.lblOTotal.setAlignment(QtCore.Qt.AlignCenter)
        self.lblOTotal.setObjectName("lblOTotal")
        self.lblONow = QtWidgets.QLabel(WireConvertForm)
        self.lblONow.setGeometry(QtCore.QRect(240, 100, 51, 20))
        self.lblONow.setAlignment(QtCore.Qt.AlignCenter)
        self.lblONow.setObjectName("lblONow")
        self.label05 = QtWidgets.QLabel(WireConvertForm)
        self.label05.setGeometry(QtCore.QRect(180, 100, 61, 20))
        self.label05.setObjectName("label05")
        self.btnPause = QtWidgets.QPushButton(WireConvertForm)
        self.btnPause.setGeometry(QtCore.QRect(120, 130, 75, 20))
        self.btnPause.setObjectName("btnPause")
        self.label06 = QtWidgets.QLabel(WireConvertForm)
        self.label06.setGeometry(QtCore.QRect(10, 70, 61, 20))
        self.label06.setObjectName("label06")
        self.lineEOFile = QtWidgets.QLineEdit(WireConvertForm)
        self.lineEOFile.setGeometry(QtCore.QRect(79, 70, 201, 20))
        self.lineEOFile.setMaxLength(256)
        self.lineEOFile.setObjectName("lineEOFile")
        self.btnOFile = QtWidgets.QToolButton(WireConvertForm)
        self.btnOFile.setGeometry(QtCore.QRect(280, 70, 30, 20))
        self.btnOFile.setObjectName("btnOFile")
        self.cmbType = QtWidgets.QComboBox(WireConvertForm)
        self.cmbType.setGeometry(QtCore.QRect(20, 40, 41, 22))
        self.cmbType.setObjectName("cmbType")
        self.cmbType.addItem("")
        self.cmbType.addItem("")
        self.label01.setBuddy(self.lineEIFile)
        self.label02.setBuddy(self.spbStart)
        self.label03.setBuddy(self.spbEnd)
        self.label06.setBuddy(self.lineEOFile)

        self.retranslateUi(WireConvertForm)
        QtCore.QMetaObject.connectSlotsByName(WireConvertForm)
        WireConvertForm.setTabOrder(self.lineEIFile, self.btnIFile)
        WireConvertForm.setTabOrder(self.btnIFile, self.spbStart)
        WireConvertForm.setTabOrder(self.spbStart, self.spbEnd)
        WireConvertForm.setTabOrder(self.spbEnd, self.lineEOFile)
        WireConvertForm.setTabOrder(self.lineEOFile, self.btnOFile)
        WireConvertForm.setTabOrder(self.btnOFile, self.btnStart)
        WireConvertForm.setTabOrder(self.btnStart, self.btnPause)
        WireConvertForm.setTabOrder(self.btnPause, self.btnCancel)

    def retranslateUi(self, WireConvertForm):
        _translate = QtCore.QCoreApplication.translate
        WireConvertForm.setWindowTitle(_translate("WireConvertForm", "线束表转化"))
        self.label01.setText(_translate("WireConvertForm", "EPLAN文件："))
        self.btnStart.setText(_translate("WireConvertForm", "开始"))
        self.btnCancel.setText(_translate("WireConvertForm", "取消"))
        self.btnIFile.setText(_translate("WireConvertForm", "..."))
        self.label02.setText(_translate("WireConvertForm", "数据起始行:"))
        self.label03.setText(_translate("WireConvertForm", "结束行(0所有行):"))
        self.label04.setText(_translate("WireConvertForm", "需要转换行数:"))
        self.lblOTotal.setText(_translate("WireConvertForm", "2 - 100"))
        self.lblONow.setText(_translate("WireConvertForm", "0"))
        self.label05.setText(_translate("WireConvertForm", "正在转换行："))
        self.btnPause.setText(_translate("WireConvertForm", "暂停"))
        self.label06.setText(_translate("WireConvertForm", "线束BOM"))
        self.btnOFile.setText(_translate("WireConvertForm", "..."))
        self.cmbType.setItemText(0, _translate("WireConvertForm", "陆"))
        self.cmbType.setItemText(1, _translate("WireConvertForm", "海"))
