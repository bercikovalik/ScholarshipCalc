import os
import sys
import subprocess

def run():
    # Path to your main app
    script = os.path.join(os.path.dirname(__file__), "Main_menu.py")

    # Run: streamlit run app.py
    subprocess.Popen([sys.executable, "-m", "streamlit", "run", script])

if __name__ == "__main__":
    run()