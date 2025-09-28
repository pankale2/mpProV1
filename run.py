# run.py - Development server with controlled browser opening
import sys
import webbrowser
import threading
import time
from main import app

def open_browser():
    """Open default browser to app URL after a short delay."""
    time.sleep(1)
    try:
        webbrowser.open('http://127.0.0.1:5000/')
    except Exception:
        pass  # non-fatal if browser cannot be opened

if __name__ == "__main__":
    print("🚀 Starting RID+PID Processor...")
    print("🌐 Server will be available at: http://127.0.0.1:5000/")

    # Always attempt to open the browser when this script is launched (script or EXE)
    threading.Thread(target=open_browser, daemon=True).start()

    is_exe = getattr(sys, "frozen", False)

    # Run the Flask app; disable the reloader to avoid duplicate starts
    app.run(
        host="127.0.0.1",
        port=5000,
        debug=not is_exe,
        use_reloader=False,
        threaded=True
    )
    