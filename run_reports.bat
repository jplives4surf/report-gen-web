@echo off
echo Starting Report Generator...

REM Try to install streamlit if it's not already installed
echo Checking if streamlit is installed...
py -m pip show streamlit >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing streamlit and dependencies...
    py -m pip install streamlit pandas python-docx openpyxl docx2txt
)

REM Run the Streamlit app using the Python launcher
echo Starting Streamlit app...
py -m streamlit run streamlit_app.py --server.port 8800

REM Keep the window open if there was an error
if %errorlevel% neq 0 (
    echo Error running Streamlit app. Press any key to exit.
    pause
)
