import sys
import os
import importlib.util
from PyQt6.QtWidgets import QApplication, QMessageBox

# --- DUMMY IMPORTS FOR PYINSTALLER ---
# We must import everything your Logic file uses so PyInstaller
# packages the libraries into the exe.
import pandas
import openpyxl
import json
import re
import urllib.request
import PyQt6.QtCore
import PyQt6.QtGui
import PyQt6.QtWidgets
# -------------------------------------

SCRIPT_NAME = "schedule_app_logic.py"

def run():
    # 1. Determine where the exe is located
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    script_path = os.path.join(base_path, SCRIPT_NAME)

    # 2. Check if the logic file exists
    if not os.path.exists(script_path):
        # We need a temporary app instance to show the error
        app = QApplication(sys.argv)
        QMessageBox.critical(None, "Fatal Error",
                             f"Could not find '{SCRIPT_NAME}'\n\n"
                             "Please ensure the logic file is in the same folder as the application.")
        return

    # 3. Dynamically load and execute the logic file
    try:
        # Load the module from the file path
        spec = importlib.util.spec_from_file_location("logic_module", script_path)
        module = importlib.util.module_from_spec(spec)
        sys.modules["logic_module"] = module
        spec.loader.exec_module(module)

        # Run the main function we defined in step 1C
        if hasattr(module, 'main'):
            module.main()
        else:
            # Fallback if main() isn't found, though it should be there
            pass

    except Exception as e:
        app = QApplication.instance() or QApplication(sys.argv)
        import traceback
        error_msg = traceback.format_exc()
        QMessageBox.critical(None, "Application Crash", f"An error occurred:\n{e}\n\n{error_msg}")

if __name__ == "__main__":
    run()