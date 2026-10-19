@echo off
chcp 65001 >nul
title 立即重新生成菌种分布图
echo.
echo 正在重新生成菌种分布图 HTML ...
echo.
"C:\Users\Administrator\anaconda3\python.exe" generate_dashboard.py
echo.
echo 完成。输出文件：
echo   菌种分布图（当天日期）.html
echo   （日期每次按系统当前时间自动更新）
echo.
pause