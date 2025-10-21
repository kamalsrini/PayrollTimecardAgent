#!/usr/bin/env python3
"""
Tesseract OCR Installer
Automatically installs Tesseract OCR for different operating systems
"""

import os
import sys
import platform
import subprocess
import urllib.request
import zipfile
import shutil
from pathlib import Path

class TesseractInstaller:
    """Handles Tesseract OCR installation for different platforms"""
    
    def __init__(self):
        self.system = platform.system().lower()
        self.arch = platform.machine().lower()
        self.tesseract_path = None
    
    def check_tesseract_installed(self) -> bool:
        """Check if Tesseract is already installed"""
        try:
            result = subprocess.run(['tesseract', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print("✅ Tesseract OCR is already installed")
                print(f"   Version: {result.stdout.split()[1] if len(result.stdout.split()) > 1 else 'Unknown'}")
                return True
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        
        return False
    
    def install_macos(self) -> bool:
        """Install Tesseract on macOS"""
        print("🍎 Installing Tesseract on macOS...")
        
        # Check if Homebrew is installed
        try:
            subprocess.run(['brew', '--version'], check=True, capture_output=True)
            print("✅ Homebrew found")
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("❌ Homebrew not found. Installing Homebrew first...")
            if not self._install_homebrew():
                return False
        
        # Install Tesseract using Homebrew
        try:
            print("📦 Installing Tesseract via Homebrew...")
            subprocess.run(['brew', 'install', 'tesseract'], check=True)
            print("✅ Tesseract installed successfully via Homebrew")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install Tesseract via Homebrew: {e}")
            return False
    
    def install_linux(self) -> bool:
        """Install Tesseract on Linux"""
        print("🐧 Installing Tesseract on Linux...")
        
        # Try different package managers
        package_managers = [
            (['apt-get', 'update'], ['apt-get', 'install', '-y', 'tesseract-ocr']),
            (['yum', 'update'], ['yum', 'install', '-y', 'tesseract']),
            (['dnf', 'update'], ['dnf', 'install', '-y', 'tesseract']),
            (['pacman', '-Sy'], ['pacman', '-S', '--noconfirm', 'tesseract']),
        ]
        
        for update_cmd, install_cmd in package_managers:
            try:
                print(f"🔄 Trying {install_cmd[0]}...")
                subprocess.run(update_cmd, check=True, capture_output=True)
                subprocess.run(install_cmd, check=True, capture_output=True)
                print(f"✅ Tesseract installed successfully via {install_cmd[0]}")
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        print("❌ Failed to install Tesseract via package managers")
        print("Please install manually: sudo apt-get install tesseract-ocr")
        return False
    
    def install_windows(self) -> bool:
        """Install Tesseract on Windows"""
        print("🪟 Installing Tesseract on Windows...")
        
        # Download Tesseract installer
        tesseract_url = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.0.20221214/tesseract-ocr-w64-setup-5.3.0.20221214.exe"
        installer_path = "tesseract_installer.exe"
        
        try:
            print("📥 Downloading Tesseract installer...")
            urllib.request.urlretrieve(tesseract_url, installer_path)
            print("✅ Download completed")
            
            print("🚀 Running Tesseract installer...")
            print("   Please follow the installation wizard.")
            print("   Make sure to add Tesseract to PATH during installation.")
            
            # Run the installer
            subprocess.run([installer_path], check=True)
            
            # Clean up
            os.remove(installer_path)
            
            print("✅ Tesseract installation completed")
            return True
            
        except Exception as e:
            print(f"❌ Failed to install Tesseract on Windows: {e}")
            print("Please download and install manually from:")
            print("https://github.com/UB-Mannheim/tesseract/wiki")
            return False
    
    def _install_homebrew(self) -> bool:
        """Install Homebrew on macOS"""
        try:
            print("📦 Installing Homebrew...")
            install_script = '/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
            subprocess.run(install_script, shell=True, check=True)
            print("✅ Homebrew installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("❌ Failed to install Homebrew")
            return False
    
    def install(self) -> bool:
        """Install Tesseract based on the operating system"""
        print("🔍 Checking Tesseract OCR installation...")
        
        if self.check_tesseract_installed():
            return True
        
        print(f"📋 Installing Tesseract for {self.system} ({self.arch})")
        
        if self.system == 'darwin':
            return self.install_macos()
        elif self.system == 'linux':
            return self.install_linux()
        elif self.system == 'windows':
            return self.install_windows()
        else:
            print(f"❌ Unsupported operating system: {self.system}")
            print("Please install Tesseract manually:")
            print("https://tesseract-ocr.github.io/tessdoc/Installation.html")
            return False
    
    def verify_installation(self) -> bool:
        """Verify that Tesseract is working correctly"""
        print("🔍 Verifying Tesseract installation...")
        
        try:
            result = subprocess.run(['tesseract', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print("✅ Tesseract is working correctly")
                return True
            else:
                print("❌ Tesseract installation verification failed")
                return False
        except Exception as e:
            print(f"❌ Error verifying Tesseract: {e}")
            return False

def main():
    """Main installation function"""
    print("🚀 Tesseract OCR Installer")
    print("=" * 40)
    
    installer = TesseractInstaller()
    
    if installer.install():
        if installer.verify_installation():
            print("\n🎉 Tesseract OCR installation completed successfully!")
            print("✅ OCR functionality is now available")
        else:
            print("\n⚠️  Tesseract installed but verification failed")
            print("Please restart your terminal and try again")
    else:
        print("\n❌ Tesseract installation failed")
        print("Please install manually and try again")

if __name__ == "__main__":
    main()
