1. 下载安装Python 3.12.2版本 https://www.python.org/downloads/
    不要用最新版本,很多Python的第三方库及案例代码都采用的较老版本
    https://www.python.org/ftp/python/3.12.2/python-3.12.2-amd64.exe
    安装时建议不用安装到默认目录(当前用户的目录),建议在C或D盘的根目录(或Program Files目录)建立Python312目录
    如： C:\Program Files\Python312

2. 下载安装SQLiteStudio https://sqlitestudio.pl/
    编辑修改MDM.db数据库
3. 先安装Visual Studio Code https://code.visualstudio.com/
    下载安装最新版本即可

4. 安装VS Code完成后,使用管理员打开VS Code,通过扩展界面安装以下扩展
    --以下部分尽量装微软的官方扩展
    Chinese (Simplified) (简体中文) Language Pack for Visual Studio Code
    Python
    vscode-icons 

--其它扩展
    SQLite: alexcvzz的版本,安装后可以在VSCode中 按F1进入命令窗口,输入SQLite调出所有SQLite命令,可以选择打开数据库命令 数据库查看MDM数据库内表的内容
    PYQT Integration (后面安装pyqt6后,再进行以下配置)
        VS Code左下角的 管理->设置, 在设置界面查询里面输入 pyqt-integration, 选择查询出来的 pyqt-integration 扩展
        在配置项目里调整:
        Pyqt-integration \ Qtdesigner: Path 的路径到安装的QTdesigner后的 designer.exe 所在全路径(包含designer.exe)
        "C:\Program Files\Python312\Lib\site-packages\qt6_applications\Qt\bin"

5. 安装Python 扩展库,在下方的终端窗口(如无:选择VS Code的查看菜单-》终端),以此安装以下扩展库,如果出错或下载慢,建议百度搜索一下国内替代的下载地址
    以下内容安装过程,如需要升级pip,注意升级版本
    pip install sip
    pip install openpyxl
    pip install configparser
    pip install PyQt6 
    pip install pyqt6-tools / pip install pySide6
    pip install pyinstaller

    pip install xlwings
    pip install wxPython
    pip install python-docx

    pip install pandas
    pip install matplotlib