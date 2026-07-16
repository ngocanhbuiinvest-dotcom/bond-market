@echo off
REM Wrapper chay cap nhat TPDN hang ngay (goi boi Task Scheduler hoac chay tay).
setlocal
cd /d "E:\NGOC ANH_BACK UP 20220921\6. AI\CLAUDE\3. Bond Market"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
echo ================ %date% %time% ================ >> update_log.txt
python update_daily.py >> update_log.txt 2>&1
echo (exit %errorlevel%) >> update_log.txt
endlocal
