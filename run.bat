@echo off
echo "処理を開始します"
pause

C:\Windows\SysWOW64\CScript //nologo bin\sql.vbs "DSN=AXEDA;UID=CRMBT;PWD=CRMBT" list.prm

echo "処理を終了します"
pause
