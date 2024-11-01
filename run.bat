@echo off
copy source.xlsx target.xlsx >nul
cscript.exe /nologo xlHeaders.vbs target.xlsx 
