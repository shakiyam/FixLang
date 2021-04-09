@echo off
powershell Invoke-ScriptAnalyzer FixLang.ps1
powershell -ExecutionPolicy Bypass -command "&ps2exe.ps1 -inputFile FixLang.ps1 -outputFile FixLang.exe"
del FixLang.exe.config
