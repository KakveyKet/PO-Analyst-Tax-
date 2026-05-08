import os
import sys
import streamlit.web.cli as stcli

# --- CRITICAL FIX: FORCE PYINSTALLER TO BUNDLE DEPENDENCIES ---
# We must import these here so PyInstaller packs them into the .exe
import pdfplumber
import pandas
import openpyxl
import sqlite3
import styles # import your extracted logic!
import streamlit.runtime.scriptrunner.magic_funcs  # <-- FIXES THE MISSING MODULE ERROR

def resolve_path(path):
    # Check if the app is running as a compiled PyInstaller executable
    if getattr(sys, 'frozen', False):
        # Look in the folder where the .exe is located (e.g., inside 'dist/')
        base_dir = os.path.dirname(sys.executable)
    else:
        # Look in the current folder
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    target_path = os.path.join(base_dir, path)
    if os.path.exists(target_path):
        return target_path
    
    # Fallback 1: Check the current working directory
    cwd_path = os.path.join(os.getcwd(), path)
    if os.path.exists(cwd_path):
        return cwd_path
        
    # Fallback 2: Check the parent directory 
    # (This perfectly handles the case where the .exe is inside the 'dist' folder!)
    if getattr(sys, 'frozen', False):
        parent_dir = os.path.dirname(base_dir)
        parent_path = os.path.join(parent_dir, path)
        if os.path.exists(parent_path):
            return parent_path
            
    return target_path

if __name__ == "__main__":
    script_path = resolve_path("app.py")
    
    # Safety Check: Keep terminal open to show the error if app.py is missing!
    if not os.path.exists(script_path):
        print("======================================================")
        print(" ERROR: MISSING FILE")
        print("======================================================")
        print(f" Could not find the file: {script_path}")
        print(" Please make sure 'app.py' is in the EXACT same folder as this .exe!")
        print(" Or place the .exe in the 'dist' folder and 'app.py' in the parent folder.")
        print("======================================================")
        input(" Press Enter to exit...")
        sys.exit(1)
    
    # --- NEW FIX: ALIGN THE CURRENT WORKING DIRECTORY ---
    # This forces the .exe to change its active folder to wherever app.py is located.
    # Now it will perfectly find POTemplate.xlsx sitting right next to it!
    os.chdir(os.path.dirname(script_path))
    
    # Explicitly configure the port and force the browser to open
    sys.argv = [
        "streamlit", 
        "run", 
        script_path, 
        "--server.port=8501", 
        "--server.headless=false", 
        "--global.developmentMode=false"
    ]
    
    sys.exit(stcli.main())