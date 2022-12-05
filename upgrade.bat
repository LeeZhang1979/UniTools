echo off
timeout /t 3
XCOPY C:\Source\UniTools\dist\UniTools c:\Source\UniTools /S /Y
start c:/Source/UniTools/Main.py