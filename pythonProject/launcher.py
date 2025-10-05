import os
import sys
import subprocess

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    main_path = resource_path("Main_menu.py")
    os.chdir(os.path.dirname(main_path))
    subprocess.run(["streamlit", "run", main_path])
