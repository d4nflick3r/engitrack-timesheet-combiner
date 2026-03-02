import sys
import os

sys.setrecursionlimit(5000)

if getattr(sys, "frozen", False):
    base = sys._MEIPASS
    os.chdir(base)

from streamlit_desktop_app import start_desktop_app

start_desktop_app(
    "app.py",
    title="EngiTrack Timesheet Combiner",
    width=1400,
    height=900,
    min_width=900,
    min_height=600,
)
