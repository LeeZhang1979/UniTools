# Form implementation generated from reading ui file 'c:\Source\UniTools\ui\MainWindow.ui'
#
# Created by: PyQt6 UI code generator 6.6.1
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1366, 768)
        MainWindow.setMinimumSize(QtCore.QSize(1366, 768))
        MainWindow.setMaximumSize(QtCore.QSize(1920, 1080))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        MainWindow.setFont(font)
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.btnMDMConf = QtWidgets.QPushButton(parent=self.centralwidget)
        self.btnMDMConf.setGeometry(QtCore.QRect(1070, 270, 91, 23))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.btnMDMConf.setFont(font)
        self.btnMDMConf.setLayoutDirection(QtCore.Qt.LayoutDirection.LeftToRight)
        self.btnMDMConf.setObjectName("btnMDMConf")
        self.btnTool01 = QtWidgets.QPushButton(parent=self.centralwidget)
        self.btnTool01.setGeometry(QtCore.QRect(60, 350, 171, 21))
        font = QtGui.QFont()
        font.setFamily("Microsoft YaHei UI")
        font.setPointSize(10)
        self.btnTool01.setFont(font)
        self.btnTool01.setObjectName("btnTool01")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1366, 23))
        self.menubar.setObjectName("menubar")
        self.menuTool = QtWidgets.QMenu(parent=self.menubar)
        self.menuTool.setObjectName("menuTool")
        self.menu = QtWidgets.QMenu(parent=self.menubar)
        self.menu.setObjectName("menu")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.MenuMDM = QtGui.QAction(parent=MainWindow)
        self.MenuMDM.setObjectName("MenuMDM")
        self.actionConfigure = QtGui.QAction(parent=MainWindow)
        self.actionConfigure.setObjectName("actionConfigure")
        self.actionPowerCableCal = QtGui.QAction(parent=MainWindow)
        self.actionPowerCableCal.setObjectName("actionPowerCableCal")
        self.actionWireConvert = QtGui.QAction(parent=MainWindow)
        self.actionWireConvert.setObjectName("actionWireConvert")
        self.actionPowerConvert = QtGui.QAction(parent=MainWindow)
        self.actionPowerConvert.setObjectName("actionPowerConvert")
        self.actionCableMSTOptimizer = QtGui.QAction(parent=MainWindow)
        self.actionCableMSTOptimizer.setObjectName("actionCableMSTOptimizer")
        self.menuTool.addAction(self.actionConfigure)
        self.menu.addAction(self.actionPowerCableCal)
        self.menu.addAction(self.actionWireConvert)
        self.menu.addAction(self.actionPowerConvert)
        self.menu.addAction(self.actionCableMSTOptimizer)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menuTool.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.btnTool01, self.btnMDMConf)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "工具集"))
        self.btnMDMConf.setText(_translate("MainWindow", "基础数据配置"))
        self.btnTool01.setText(_translate("MainWindow", "风力发电机组短路电流计算"))
        self.menuTool.setTitle(_translate("MainWindow", "系统配置"))
        self.menu.setTitle(_translate("MainWindow", "通用工具"))
        self.MenuMDM.setText(_translate("MainWindow", "基础数据配置"))
        self.actionConfigure.setText(_translate("MainWindow", "基础数据配置"))
        self.actionPowerCableCal.setText(_translate("MainWindow", "动力电缆计算"))
        self.actionWireConvert.setText(_translate("MainWindow", "线束表转换"))
        self.actionPowerConvert.setText(_translate("MainWindow", "变压器计算"))
        self.actionCableMSTOptimizer.setText(_translate("MainWindow", "海缆拓扑优化"))
