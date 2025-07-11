import webbrowser
import threading
import time
import sys
from main import app

def open_browser():
    time.sleep(1.5)
    webbrowser.open('http://127.0.0.1:5000/')

if __name__ == '__main__':
    # Only open browser if not running as a PyInstaller subprocess
    if not getattr(sys, 'frozen', False) or hasattr(sys, '_MEIPASS'):
        threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
    app.run(port=5000)
