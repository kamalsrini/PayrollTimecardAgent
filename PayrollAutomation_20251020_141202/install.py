#!/usr/bin/env python3
"""
Payroll Automation Installer
"""

import os
import sys
import subprocess
import platform

def install_dependencies():
    """Install required dependencies"""
    print("📦 Installing dependencies...")
    
    packages = [
        'pandas>=2.0.0',
        'openpyxl>=3.1.0', 
        'xlrd>=2.0.0',
        'pytesseract>=0.3.10',
        'opencv-python>=4.8.0',
        'Pillow>=10.0.0'
    ]
    
    for package in packages:
        try:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            print(f"✅ {package} installed")
        except subprocess.CalledProcessError:
            print(f"❌ Failed to install {package}")
            return False
    
    return True

def install_tesseract():
    """Install Tesseract OCR"""
    print("🔍 Installing Tesseract OCR...")
    
    try:
        # Import and run the Tesseract installer
        from install_tesseract import TesseractInstaller
        installer = TesseractInstaller()
        return installer.install()
    except ImportError:
        print("❌ Tesseract installer not found")
        return False

def main():
    """Main installation function"""
    print("🚀 Payroll Automation Installer")
    print("=" * 40)
    
    # Install Python dependencies
    if not install_dependencies():
        print("❌ Failed to install Python dependencies")
        return
    
    # Install Tesseract
    install_tesseract()
    
    print("\n✅ Installation completed!")
    print("\nNext steps:")
    print("1. Add your timesheet files to the 'input_files' folder")
    print("2. Run: python run_payroll_automation.py")

if __name__ == "__main__":
    main()
