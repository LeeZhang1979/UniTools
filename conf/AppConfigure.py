import xml.dom.minidom as xmldom
import os

class Appconfig(object):
    AppCode="" 
    #AppType="" 
    Icon="" 
    Version="" 
    PathType="" 
    Path="" 
    Arguments="" 
    AppStartupType=""

    def __init__(self,appCode,icon,version,pathType,path,arguments,appStartupType):
        self.AppCode=appCode
        self.Icon=icon 
        self.Version=version 
        self.PathType=pathType
        self.Path=path
        self.Arguments=arguments 
        self.AppStartupType=appStartupType
    
   

class AppConfigure(object):
    Appconfigs = []
    def __init__(self):
        super().__init__()
        self.Appconfigs.clear()

    def loadConf(self,filepath):
        self.Appconfigs.clear() 
        if not os.path.isfile(filepath):
            return
        domTree = xmldom.parse(filepath)
        rootNode = domTree.documentElement
        apps=rootNode.getElementsByTagName("AppConfig")
        for app in apps: 
            if app.hasAttribute("AppCode"):
                appCode = app.getAttribute("AppCode")
                icon = app.getAttribute("Icon")
                version = app.getAttribute("Version")
                pathType = app.getAttribute("PathType")
                path = app.getAttribute("Path")
                arguments = app.getAttribute("Arguments")
                appStartupType = app.getAttribute("AppStartupType")                
                self.Appconfigs.append(Appconfig(appCode,icon,version,pathType,path,arguments,appStartupType))      


