import os
import sys
import webbrowser
import threading
import time
import uvicorn
from server import app, EXCEL_FILE_PATH

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8000")

if __name__ == "__main__":
    print("\n" + "="*70)
    print("  XYZ Electronic Appliances - Stock Management Portal")
    print("  Standalone Portable Distribution (Zero Dependencies Required)")
    print("  Excel Database: " + EXCEL_FILE_PATH)
    print("  Web Portal URL: http://127.0.0.1:8000")
    print("="*70 + "\n")
    
    # Launch browser automatically
    threading.Thread(target=open_browser, daemon=True).start()
    
    # Run uvicorn server directly with app object
    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
