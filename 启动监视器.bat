@echo off
chcp 65001 >nul
title 菌种分布图自动更新监视器
echo.
echo ================================================
echo   菌种分布图 HTML 自动更新监视器
echo ================================================
echo.
echo   检查间隔：每 1 小时自动检查一次 细菌库.xlsx
echo   用法说明：
echo     1. 在金山文档里修改菌种数据
echo     2. 导出为 xlsx，覆盖本目录下的 细菌库.xlsx
echo     3. 本程序会自动检测到变化并重新生成 HTML
echo.
echo   按 Ctrl+C 可随时停止。
echo ================================================
echo.
"C:\Users\Administrator\anaconda3\python.exe" watcher.py --interval 3600
echo.
echo 监视器已退出。
pause