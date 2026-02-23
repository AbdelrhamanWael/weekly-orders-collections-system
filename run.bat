@echo off
chcp 65001 >nul
color 0A
mode con: cols=62 lines=35
title  Orders & Collections System

cls
echo.
echo  ╔════════════════════════════════════════════════════════════╗
echo  ║                                                            ║
echo  ║        Orders ^& Collections Linking System                ║
echo  ║              Weekly Report Generator                       ║
echo  ║                                                            ║
echo  ╚════════════════════════════════════════════════════════════╝
echo.

REM ── Check Python ──────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo  ╔════════════════════════════════════════════════════════════╗
    echo  ║  [!]  Python is not installed on this machine              ║
    echo  ║                                                            ║
    echo  ║  Please install Python 3.10 or newer from:                ║
    echo  ║  https://www.python.org/downloads/                         ║
    echo  ╚════════════════════════════════════════════════════════════╝
    echo.
    pause
    exit /b 1
)

REM ── Step 1: Check Libraries ────────────────────────────────────
echo  [ 1 / 4 ]  Checking required libraries...
python -m pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ ... ]  Installing required libraries...
    python -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo  [ ... ]  Retrying installation with --user flag...
        python -m pip install --user -r requirements.txt
        if %errorlevel% neq 0 (
            color 0C
            echo.
            echo  [  X  ]  Failed to install libraries!
            echo           Please check errors above.
            echo.
            pause
            exit /b 1
        )
    )
    echo  [ OK  ]  Libraries installed successfully
) else (
    echo  [ OK  ]  Libraries already installed
)

REM ── Step 2: Create Folders ─────────────────────────────────────
echo.
echo  [ 2 / 4 ]  Creating required folders...
if not exist "samples"  mkdir samples
if not exist "reports"  mkdir reports
if not exist "data"     mkdir data
echo  [ OK  ]  Folders are ready

REM ── Step 3: Initialize Database ───────────────────────────────
echo.
echo  [ 3 / 4 ]  Initializing database...
python init_db.py
echo  [ OK  ]  Database is ready

REM ── Step 4: Launch Application ────────────────────────────────
echo.
echo  [ 4 / 4 ]  Starting the application...
echo.
echo  ╔════════════════════════════════════════════════════════════╗
echo  ║                                                            ║
echo  ║   System is running!                                       ║
echo  ║                                                            ║
echo  ║   Open your browser and go to:                            ║
echo  ║   http://127.0.0.1:5000                                    ║
echo  ║                                                            ║
echo  ║   To stop the server: press  Ctrl + C                     ║
echo  ║                                                            ║
echo  ╚════════════════════════════════════════════════════════════╝
echo.

python app.py

pause
