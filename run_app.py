# run_app.py (v7.3 launcher)
import sys, webbrowser, threading, time
from streamlit.web import cli as stcli
def open_browser_later():
    time.sleep(2)
    try:
        webbrowser.open("http://localhost:8501")
    except Exception:
        pass
if __name__ == "__main__":
    threading.Thread(target=open_browser_later, daemon=True).start()
    sys.argv = ["streamlit", "run", "mo32_one_button_app.py",
                "--server.address=127.0.0.1",
                "--server.headless=false",
                "--browser.gatherUsageStats=false"]
    stcli.main()
