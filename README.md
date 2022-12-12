1. 下载安装Python 3.8.10版本 https://www.python.org/downloads/
    不要用最新版本，很多Python的第三方库及案例代码都采用的较老版本
    https://www.python.org/downloads/release/python-3810/
    安装时建议不用安装到默认目录（当前用户的目录），建议在C或D盘的根目录建立Python380目录
    如： C:\Python38

2. 下载安装SQLiteStudio https://sqlitestudio.pl/
    编辑修改MDM.db数据库
3. 先安装Visual Studio Code https://code.visualstudio.com/
    下载安装最新版本即可

4. 安装VS Code完成后，打开VS Code，通过扩展界面安装以下扩展
    --以下部分尽量装微软的官方扩展
    Chinese (Simplified) (简体中文) Language Pack for Visual Studio Code
    Python
    vscode-icons

    --还有更多适合Python开发的扩展，可自行搜索安装，Pylance 等不建议安装（这个强制代码书写规范），部分AI的扩展也不要装，影响性能

--其它扩展
    SQLite: alexcvzz的版本，安装后可以在VSCode中 按F1进入命令窗口，输入SQLite调出所有SQLite命令，可以选择打开数据库命令 数据库查看MDM数据库内表的内容
    PYQT Integration （后面安装pyqt5后，再进行以下配置）
        VS Code左下角的 管理->设置, 在设置界面查询里面输入 pyqt-integration, 选择查询出来的 pyqt-integration 扩展
        在配置项目里调整:
        Pyqt-integration › Qtdesigner: Path 的路径到安装的Pyqt5designer后的 designer.exe 所在全路径（包含designer.exe）

5. 安装Python 扩展库，在下方的终端窗口（如无：选择VS Code的查看菜单-》终端），以此安装以下扩展库，如果出错或下载慢，建议百度搜索一下国内替代的下载地址
    以下内容安装过程，如需要升级pip，注意升级版本
    pip install sip
    pip install openpyxl
    pip install configparser
    pip install PyQt5 
    pip install pyqt5designer 
    pip install pyinstaller

    pip install xlwings
    pip install wxPython
    pip install python-docx

    pip install pandas
    pip install matplotlib