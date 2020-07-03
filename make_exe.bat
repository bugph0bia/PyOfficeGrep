@echo off
pyinstaller office_grep.spec --onefile
copy setting.ini .\dist
pause
