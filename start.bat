@echo off
cd /d "%~dp0"
call cpd\Scripts\activate
start "" /min cmd /c "streamlit run app.py --server.address 0.0.0.0 --server.port 8501"
exit