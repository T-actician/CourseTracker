@echo off
title Building Course Tracker
echo Installing requirements...
python -m pip install --upgrade customtkinter python-docx pyinstaller
echo Building CourseTracker.exe ...
python -m PyInstaller --noconsole --onefile --clean -y --name CourseTracker --collect-all customtkinter course_tracker.py
echo.
echo Done. Your app is here: dist\CourseTracker.exe
pause
