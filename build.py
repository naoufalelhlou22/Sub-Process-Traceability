import os
import subprocess
import sys
import shutil

def run_build():
    print("="*60)
    print(" BUILDING EXECUTABLES ".center(60, "="))
    print("="*60)
    
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
    print("\n--- Building Main Application ---")
    if os.path.exists("Sub-Process Traceability.spec"):
        subprocess.check_call([sys.executable, "-m", "PyInstaller", "Sub-Process Traceability.spec", "--clean"])
    else:
        print("Spec file 'Sub-Process Traceability.spec' not found.")
        
    print("\n--- Building Quality Application ---")
    if os.path.exists("Quality App.spec"):
        subprocess.check_call([sys.executable, "-m", "PyInstaller", "Quality App.spec", "--clean"])
    else:
        print("Spec file 'Quality App.spec' not found.")
        
    print("\n" + "="*60)
    print(" BUILD COMPLETE ".center(60, "="))
    print("="*60)
    print("\nYou can find the compiled .exe files inside the 'dist' folder.")

if __name__ == "__main__":
    run_build()
