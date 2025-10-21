#!/usr/bin/env python3
"""
Payroll Automation Launcher
Simple launcher script for the unified payroll processor
"""

import os
import sys
import subprocess
from pathlib import Path

def check_dependencies():
    """Check if required dependencies are installed"""
    print("🔍 Checking dependencies...")
    
    required_packages = [
        'pandas', 'openpyxl', 'xlrd', 'pytesseract', 'opencv-python', 'Pillow'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"  ✅ {package}")
        except ImportError:
            print(f"  ❌ {package}")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n⚠️  Missing packages: {', '.join(missing_packages)}")
        print("Installing missing packages...")
        
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing_packages)
            print("✅ Dependencies installed successfully!")
        except subprocess.CalledProcessError:
            print("❌ Failed to install dependencies. Please install manually:")
            print(f"   pip install {' '.join(missing_packages)}")
            return False
    
    return True

def setup_folders():
    """Setup required folders"""
    print("📁 Setting up folders...")
    
    folders = ['input_files', 'output']
    
    for folder in folders:
        Path(folder).mkdir(exist_ok=True)
        print(f"  ✅ {folder}/")
    
    return True

def show_instructions():
    """Show usage instructions"""
    print("\n" + "=" * 60)
    print("📋 PAYROLL AUTOMATION INSTRUCTIONS")
    print("=" * 60)
    print("1. 📁 Add your timesheet files to the 'input_files' folder:")
    print("   - Excel files (.xlsx, .xls) containing 'time' in filename")
    print("   - Image files (.png, .jpg, .jpeg) containing 'time' in filename")
    print()
    print("2. 🚀 Run the automation:")
    print("   python run_payroll_automation.py")
    print()
    print("3. 📊 Check results in the 'output' folder:")
    print("   - consolidated_payroll_data.csv")
    print("   - processing_report_*.txt")
    print("   - Updated Payroll Totals file")
    print()
    print("4. 📈 The automation will:")
    print("   - Extract data from Excel files")
    print("   - Extract data from images using OCR")
    print("   - Consolidate all employee data")
    print("   - Update your Payroll Totals file")
    print("   - Generate a processing report")

def main():
    """Main function"""
    print("🚀 PAYROLL AUTOMATION LAUNCHER")
    print("=" * 40)
    
    # Check dependencies
    if not check_dependencies():
        print("❌ Dependency check failed. Please install required packages.")
        return
    
    # Setup folders
    if not setup_folders():
        print("❌ Folder setup failed.")
        return
    
    # Show instructions
    show_instructions()
    
    # Check if there are files to process
    input_folder = Path('input_files')
    if not any(input_folder.iterdir()):
        print(f"\n⚠️  No files found in {input_folder}/")
        print("Please add your timesheet files and run again.")
        return
    
    # Run the unified processor
    print(f"\n🚀 Starting payroll processing...")
    try:
        from unified_payroll_processor import UnifiedPayrollProcessor
        processor = UnifiedPayrollProcessor()
        success = processor.run_full_processing()
        
        if success:
            print("\n🎉 Automation completed successfully!")
        else:
            print("\n❌ Automation failed. Check the error messages above.")
            
    except Exception as e:
        print(f"\n❌ Error running automation: {e}")
        print("Please check that all required files are present.")

if __name__ == "__main__":
    main()
