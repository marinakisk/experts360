#!/bin/bash
cd ~/Documents/Experts360_App
if pgrep -f "streamlit run app.py" > /dev/null; then
    open http://localhost:8501
else
    /Library/Frameworks/Python.framework/Versions/3.14/bin/streamlit run app.py &
    sleep 3
    open http://localhost:8501
fi
