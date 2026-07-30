@echo off
chcp 65001 >nul
title AutoCAD 清理脚本 - 请以管理员身份运行
echo ========================================
echo    AutoCAD 旧版本清理脚本
echo    请确认以管理员身份运行
echo ========================================
echo.

echo [1/3] 正在删除 AutoCAD 2016 注册表信息...
reg delete "HKEY_LOCAL_MACHINE\Software\Autodesk\AutoCAD\R22.0" /f
if %errorlevel% equ 0 (echo   ✓ 成功) else (echo   ✗ 失败)
echo.

echo [2/3] 正在删除 AutoCAD 2021 注册表信息...
reg delete "HKEY_LOCAL_MACHINE\Software\Autodesk\AutoCAD\R23.1" /f
if %errorlevel% equ 0 (echo   ✓ 成功) else (echo   ✗ 失败)
echo.

echo [3/3] 正在移除 Autodesk Access 开机自启动...
reg delete "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run" /v "Autodesk Access" /f
if %errorlevel% equ 0 (echo   ✓ 成功) else (echo   ✗ 失败)
echo.

echo ========================================
echo    清理完成！
echo    请重启 AutoCAD 2024 验证效果
echo ========================================
pause
