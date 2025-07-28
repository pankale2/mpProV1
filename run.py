import webbrowser
import threading
import time
from main import app

def open_browser():
    """
    Opens the default web browser to the application's URL after a short delay.
    """
    time.sleep(1.5)
    webbrowser.open('http://127.0.0.1:5000/')

if __name__ == '__main__':
    # Set up the thread to open the browser.
    # The 'daemon=True' flag ensures the thread will exit when the main program exits.
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Run the Flask application.
    # debug=False is important for production environments.
    app.run(host="127.0.0.1", port=5000, debug=False)