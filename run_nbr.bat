@echo off
title NBR Return Verification Setup
echo Installing dependencies...
pip install -r requirements.txt
echo.
echo Running NBR Verification Automation...
python nbr_script.py
pause
