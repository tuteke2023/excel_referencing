#!/usr/bin/env python3
"""
Build standalone executable using PyInstaller
"""

import subprocess
import sys
import os

def build_executable():
    """Build standalone executable for distribution."""
    
    # Install PyInstaller if not available
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # PyInstaller command
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name=TB-GL-Linker",
        "--icon=icon.ico",  # Add icon if available
        "tb_gl_linker.py"
    ]
    
    print("Building executable...")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("✅ Executable built successfully!")
        print("📁 Find your executable in the 'dist' folder")
    else:
        print("❌ Build failed:")
        print(result.stderr)

if __name__ == "__main__":
    build_executable()