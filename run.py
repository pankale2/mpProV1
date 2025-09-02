# run.py - Development server with controlled browser opening
import os
import sys
import webbrowser
import threading
import time
import tempfile
import datetime
from main import app

def open_browser():
    """Opens the default web browser to the application's URL after a short delay."""
    time.sleep(1.0)  # Slightly longer delay for development
    webbrowser.open('http://127.0.0.1:5000/')

def should_open_browser():
    """Check if browser should open (only once per development session)."""
    # Use temp directory for flag file
    temp_dir = tempfile.gettempdir()
    flag_file = os.path.join(temp_dir, 'ridpid_dev_browser.flag')
    
    if os.path.exists(flag_file):
        return False  # Already opened in this session
    else:
        # Create flag file to prevent future opens
        try:
            with open(flag_file, 'w') as f:
                f.write(f"Browser opened at: {datetime.datetime.now()}")
        except:
            pass  # Ignore errors creating flag file
        return True

def cleanup_flag_file():
    """Clean up browser flag file on normal exit."""
    try:
        temp_dir = tempfile.gettempdir()
        flag_file = os.path.join(temp_dir, 'ridpid_dev_browser.flag')
        if os.path.exists(flag_file):
            os.remove(flag_file)
    except:
        pass  # Ignore cleanup errors

if __name__ == "__main__":
    import atexit
    
    # Register cleanup function for normal exit
    atexit.register(cleanup_flag_file)
    
    print("🚀 Starting RID+PID Processor in development mode...")
    print("📂 Use this for development to avoid repeated browser opening")
    print("🌐 Server will be available at: http://127.0.0.1:5000/")
    
    # Open browser only once per development session
    if should_open_browser():
        print("🔗 Opening browser...")
        threading.Thread(target=open_browser, daemon=True).start()
    else:
        print("🔗 Browser already opened in this session")
    
    # Run Flask in debug mode for development
    app.run(
        host="127.0.0.1", 
        port=5000, 
        debug=True,
        use_reloader=True,
        threaded=True
    )
