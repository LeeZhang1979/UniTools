#!/usr/bin/env python3
# coding=utf-8

import os
import sys 
import socket
import configparser
import subprocess 
import ctypes

from PyQt5 import QtCore

from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QApplication,QMainWindow,QMessageBox

BASE_DIR= os.path.dirname(os.path.abspath(__file__) )  
sys.path.append( BASE_DIR  ) 
from src.MainWindow import MainWindow

def is_company_network(companynetwork):
    try: 
        s=socket.socket(socket.AF_INET,socket.SOCK_DGRAM)
        address = ("114.114.114.114",80)
        s.connect(address)
        socketName= s.getsockname()
        ip = socketName[0]
        port = socketName[1]
    finally:
        s.close()
    if ip.startswith(companynetwork):
        return True    
    return False

def check_Upgrade():
    config = configparser.ConfigParser()
    config.read(os.path.join(BASE_DIR,u'conf\\App.ini'))
    localVersion = config["Application"]["Version"] 
    companyNetwork = config["Server"]["CompanyNetwork"] 
    serverAddress = config["Server"]["Address"] 

    if not is_company_network(companyNetwork):
        return False    

    config.read(os.path.join(serverAddress,u'conf\\App.ini'))
    serverVersion = config["Application"]["Version"] 
    foraceUpgrade = config["Application"]["ForceUpgrade"]
    if localVersion >= serverVersion:
        return False
    '''
    #暂未启用，需要导入消息框类库或转移升级程序到MainWindow.py
    strMsg = '服务端版本:[' + serverVersion + '],本地版本:[' + localVersion + ']'
    if foraceUpgrade == '1':
        strMsg = strMsg + ',本次版本强制升级!'
        buttons = QMessageBox.Yes
    else:
        strMsg = strMsg + ',是否需要升级?'
        buttons = QMessageBox.Yes|QMessageBox.No

    if QMessageBox.show(0, '版本检查', strMsg,buttons,QMessageBox.Yes) != QMessageBox.Yes:
        return False 
    '''
    bat_file = open('upgrade.bat', 'w')
    # 关闭bat脚本的输出
    upgrade_bat = 'echo off\n'
    # 3秒后删除旧程序(3秒后程序己运行结束;不延时的话,会提示被占用,无法删除)
    upgrade_bat += 'timeout /t 3\n'
    #　copy新版本并覆盖旧版本
    upgrade_bat += f'XCOPY {serverAddress} {BASE_DIR} /S /Y\n'
    print (f'XCOPY \\\\{serverAddress} {BASE_DIR} /S /Y\n')
    # 启动新程序
    app=os.path.join(BASE_DIR,u'UniTools.exe')
    upgrade_bat += fr'start {app}' 
    print(fr'start {app}' )
    bat_file.write(upgrade_bat)
    bat_file.close()

    ##config.set("Application","Version",serverVersion)
    return True
 
def main(): 
   

    if not ctypes.windll.shell32.IsUserAnAdmin(): 
        app=os.path.join(BASE_DIR,u'UniTools.exe')
        if sys.version_info[0] == 3:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, app, None, 1) 
        else:   
            ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), unicode(app), None, 1)

    if check_Upgrade():
        subprocess.Popen("upgrade.bat")
        sys.exit()

    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv) 
    mwin= MainWindow()
    
    mwin.show()
   
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
