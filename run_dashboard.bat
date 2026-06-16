@echo off
cd /d "%~dp0"
echo starting Streamlit dashboard %date% %time%>> streamlit_stdout.log
".venv\Scripts\python.exe" -m streamlit run app.py --server.headless=true --server.port=8501 --server.fileWatcherType=none --browser.gatherUsageStats=false >> streamlit_stdout.log 2>> streamlit_stderr.log
echo Streamlit exited with code %errorlevel% %date% %time%>> streamlit_stdout.log
