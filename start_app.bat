@echo off
title XYZ Electronic Appliances - Stock Management Portal
color 0B

echo ======================================================================
echo           XYZ ELECTRONIC APPLIANCES - STOCK MANAGEMENT PORTAL
echo ======================================================================
echo.
echo   DEMO INTERVIEW CREDENTIALS:
echo   ------------------------------------------------------------------
echo   Role               Email                      Password   Permissions
echo   ------------------------------------------------------------------
echo   👑 Admin1        : admin@xyzappliances.com    admin123   Full Access
echo   🏬 Store Manager : manager@xyzappliances.com  manager123 Stock, Add, POs
echo   📦 Godown Staff  : godown@xyzappliances.com   godown123  Adjust, Barcodes
echo   ------------------------------------------------------------------
echo.
echo   Web Portal URL : http://127.0.0.1:8000
echo   Excel Database : data\XYZ_Stock_Inventory.xlsx
echo.

if exist "XYZ_Stock_Portal.exe" (
    echo [*] Running Standalone Zero-Dependency App (XYZ_Stock_Portal.exe)...
    echo [*] Browser will open automatically in 2 seconds...
    start "" XYZ_Stock_Portal.exe
    exit /b 0
)

echo [*] Standalone binary not found, checking Python environment...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Neither XYZ_Stock_Portal.exe nor Python was found on your system!
    pause
    exit /b 1
)

echo [*] Starting Python backend server...
start "" http://127.0.0.1:8000
python server.py
pause
