@echo off
cd /d "%~dp0"
call cpd\Scripts\activate
start /B streamlit run app.py