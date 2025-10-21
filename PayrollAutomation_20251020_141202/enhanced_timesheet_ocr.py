#!/usr/bin/env python3
"""
Enhanced Timesheet OCR with Automatic Tesseract Installation
"""

import cv2
import pytesseract
from PIL import Image
import re
import os
import subprocess
import platform
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from datetime import datetime

@dataclass
class TimesheetData:
    """Data class for timesheet information extracted from image"""
    employee_name: str
    week_period: str
    daily_hours: Dict[str, int]  # {"Mon 6": 8, "Tue 7": 8, ...}
    total_hours: int
    task_breakdown: Dict[str, int]  # {"KTLO": 19, "Enhancement": 7, ...}
    image_path: str
    extraction_date: str

class TesseractManager:
    """Manages Tesseract OCR installation and configuration"""
    
    def __init__(self):
        self.system = platform.system().lower()
        self.tesseract_installed = False
    
    def check_tesseract_installed(self) -> bool:
        """Check if Tesseract is already installed"""
        try:
            result = subprocess.run(['tesseract', '--version'], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                print("✅ Tesseract OCR is already installed")
                self.tesseract_installed = True
                return True
        except (subprocess.TimeoutExpired, FileNotFoundError):
            pass
        
        return False
    
    def install_tesseract(self) -> bool:
        """Install Tesseract based on the operating system"""
        print("🔍 Tesseract not found. Installing...")
        
        if self.system == 'darwin':  # macOS
            return self._install_macos()
        elif self.system == 'linux':
            return self._install_linux()
        elif self.system == 'windows':
            return self._install_windows()
        else:
            print(f"❌ Unsupported operating system: {self.system}")
            return False
    
    def _install_macos(self) -> bool:
        """Install Tesseract on macOS using Homebrew"""
        try:
            # Check if Homebrew is installed
            subprocess.run(['brew', '--version'], check=True, capture_output=True)
            print("🍎 Installing Tesseract via Homebrew...")
            subprocess.run(['brew', 'install', 'tesseract'], check=True)
            print("✅ Tesseract installed successfully")
            self.tesseract_installed = True
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            print("❌ Homebrew not found. Please install Homebrew first:")
            print("   /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"")
            print("   Then run: brew install tesseract")
            return False
    
    def _install_linux(self) -> bool:
        """Install Tesseract on Linux"""
        package_managers = [
            (['apt-get', 'update'], ['apt-get', 'install', '-y', 'tesseract-ocr']),
            (['yum', 'update'], ['yum', 'install', '-y', 'tesseract']),
            (['dnf', 'update'], ['dnf', 'install', '-y', 'tesseract']),
        ]
        
        for update_cmd, install_cmd in package_managers:
            try:
                print(f"🐧 Installing Tesseract via {install_cmd[0]}...")
                subprocess.run(update_cmd, check=True, capture_output=True)
                subprocess.run(install_cmd, check=True, capture_output=True)
                print("✅ Tesseract installed successfully")
                self.tesseract_installed = True
                return True
            except (subprocess.CalledProcessError, FileNotFoundError):
                continue
        
        print("❌ Failed to install Tesseract. Please install manually:")
        print("   sudo apt-get install tesseract-ocr")
        return False
    
    def _install_windows(self) -> bool:
        """Install Tesseract on Windows"""
        print("🪟 Windows detected. Please install Tesseract manually:")
        print("   1. Download from: https://github.com/UB-Mannheim/tesseract/wiki")
        print("   2. Run the installer")
        print("   3. Make sure to add Tesseract to PATH")
        return False
    
    def ensure_tesseract_available(self) -> bool:
        """Ensure Tesseract is available, install if needed"""
        if self.check_tesseract_installed():
            return True
        
        return self.install_tesseract()

class EnhancedTimesheetOCR:
    """Enhanced OCR-based timesheet parser with automatic Tesseract installation"""
    
    def __init__(self):
        self.tesseract_manager = TesseractManager()
        self.setup_tesseract()
    
    def setup_tesseract(self):
        """Setup Tesseract OCR configuration"""
        if not self.tesseract_manager.ensure_tesseract_available():
            print("⚠️  Tesseract OCR not available. OCR functionality will be limited.")
            return
        
        # Configure Tesseract for better text recognition
        custom_config = r'--oem 3 --psm 6'
        pytesseract.pytesseract.tesseract_cmd = self._find_tesseract_path()
    
    def _find_tesseract_path(self) -> str:
        """Find Tesseract executable path"""
        # Common paths for different operating systems
        possible_paths = [
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract',
            '/opt/homebrew/bin/tesseract',  # macOS with Homebrew
            'tesseract'  # If in PATH
        ]
        
        for path in possible_paths:
            if os.path.exists(path) or path == 'tesseract':
                return path
        
        return 'tesseract'  # Default fallback
    
    def extract_text_from_image(self, image_path: str) -> str:
        """Extract all text from image using OCR"""
        if not self.tesseract_manager.tesseract_installed:
            print("❌ Tesseract not available. Cannot process images.")
            return ""
        
        try:
            # Load and preprocess image
            image = Image.open(image_path)
            
            # Convert to RGB if needed
            if image.mode != 'RGB':
                image = image.convert('RGB')
            
            # Use pytesseract to extract text
            text = pytesseract.image_to_string(image, config='--oem 3 --psm 6')
            return text
        except Exception as e:
            print(f"Error extracting text from image: {e}")
            return ""
    
    def preprocess_image(self, image_path: str) -> str:
        """Preprocess image for better OCR results"""
        try:
            # Load image with OpenCV
            img = cv2.imread(image_path)
            
            # Convert to grayscale
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            
            # Apply threshold to get binary image
            _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Save processed image
            processed_path = image_path.replace('.', '_processed.')
            cv2.imwrite(processed_path, thresh)
            
            return processed_path
        except Exception as e:
            print(f"Error preprocessing image: {e}")
            return image_path
    
    def parse_timesheet_data(self, text: str, image_path: str) -> TimesheetData:
        """Parse timesheet data from extracted text"""
        lines = text.split('\n')
        
        # Extract employee name
        name = self._extract_employee_name(lines)
        
        # Extract week period
        week_period = self._extract_week_period(lines)
        
        # Extract daily hours
        daily_hours = self._extract_daily_hours(lines)
        
        # Extract total hours
        total_hours = self._extract_total_hours(lines)
        
        # Extract task breakdown
        task_breakdown = self._extract_task_breakdown(lines)
        
        return TimesheetData(
            employee_name=name,
            week_period=week_period,
            daily_hours=daily_hours,
            total_hours=total_hours,
            task_breakdown=task_breakdown,
            image_path=image_path,
            extraction_date=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        )
    
    def _extract_employee_name(self, lines: List[str]) -> str:
        """Extract employee name from text lines"""
        # Look for name patterns in the top right area
        name_patterns = [
            r'([A-Z][a-z]+\s+[A-Z][a-z]+)',  # First Last
            r'Name:\s*([A-Z][a-z]+\s+[A-Z][a-z]+)',  # Name: First Last
            r'Employee:\s*([A-Z][a-z]+\s+[A-Z][a-z]+)'  # Employee: First Last
        ]
        
        for line in lines:
            for pattern in name_patterns:
                match = re.search(pattern, line)
                if match:
                    return match.group(1).strip()
        
        return "Unknown"
    
    def _extract_week_period(self, lines: List[str]) -> str:
        """Extract week period from text lines"""
        # Look for date patterns like "5 - 11 October 2025"
        week_patterns = [
            r'(\d+\s*-\s*\d+\s+\w+\s+\d{4})',  # 5 - 11 October 2025
            r'Week\s+of\s+([^,\n]+)',  # Week of ...
            r'(\w+\s+\d+,\s+\d{4})'  # October 5, 2025
        ]
        
        for line in lines:
            for pattern in week_patterns:
                match = re.search(pattern, line)
                if match:
                    return match.group(1).strip()
        
        return "Unknown Week"
    
    def _extract_daily_hours(self, lines: List[str]) -> Dict[str, int]:
        """Extract daily hours from text lines"""
        daily_hours = {}
        
        # Look for patterns like "Mon 6: 8 Hrs", "Tue 7: 8 Hrs", etc.
        daily_patterns = [
            r'(Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+(\d+):\s*(\d+)\s*Hrs?',
            r'(Sunday|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday)\s+(\d+):\s*(\d+)\s*Hrs?',
            r'(\d+/\d+):\s*(\d+)\s*Hrs?'  # Date format like 10/6: 8 Hrs
        ]
        
        for line in lines:
            for pattern in daily_patterns:
                matches = re.findall(pattern, line)
                for match in matches:
                    if len(match) == 3:
                        day, day_num, hours = match
                        daily_hours[f"{day} {day_num}"] = int(hours)
                    elif len(match) == 2:
                        date, hours = match
                        daily_hours[date] = int(hours)
        
        return daily_hours
    
    def _extract_total_hours(self, lines: List[str]) -> int:
        """Extract total hours from text lines"""
        # Look for total hours patterns
        total_patterns = [
            r'Total:\s*(\d+)\s*hrs?',
            r'(\d+)\s*hrs?\s*Total',
            r'Time\s+Sheet\s+breakdown.*?(\d+)\s*hrs?',
            r'(\d+)\s*hrs?\s*breakdown'
        ]
        
        for line in lines:
            for pattern in total_patterns:
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    return int(match.group(1))
        
        return 0
    
    def _extract_task_breakdown(self, lines: List[str]) -> Dict[str, int]:
        """Extract task breakdown from text lines"""
        task_breakdown = {}
        
        # Look for task patterns with hours
        task_pattern = r'([A-Za-z\s]+):\s*Total:\s*(\d+)\s*hours?'
        
        for line in lines:
            match = re.search(task_pattern, line)
            if match:
                task_name = match.group(1).strip()
                hours = int(match.group(2))
                task_breakdown[task_name] = hours
        
        return task_breakdown
    
    def process_timesheet_image(self, image_path: str) -> Optional[TimesheetData]:
        """Main method to process timesheet image and extract data"""
        print(f"Processing timesheet image: {image_path}")
        
        if not self.tesseract_manager.tesseract_installed:
            print("❌ Tesseract OCR not available. Cannot process image.")
            return None
        
        # Preprocess image for better OCR
        processed_path = self.preprocess_image(image_path)
        
        # Extract text
        text = self.extract_text_from_image(processed_path)
        
        if not text.strip():
            print("No text extracted from image")
            return None
        
        # Parse the extracted text
        timesheet_data = self.parse_timesheet_data(text, image_path)
        
        # Clean up processed image
        if processed_path != image_path and os.path.exists(processed_path):
            os.remove(processed_path)
        
        return timesheet_data
    
    def print_extraction_results(self, timesheet_data: TimesheetData):
        """Print formatted extraction results"""
        if not timesheet_data:
            print("No data extracted")
            return
        
        print("\n" + "="*50)
        print("TIMESHEET DATA EXTRACTION RESULTS")
        print("="*50)
        print(f"Employee Name: {timesheet_data.employee_name}")
        print(f"Week Period: {timesheet_data.week_period}")
        print(f"Total Hours: {timesheet_data.total_hours}")
        print(f"Extraction Date: {timesheet_data.extraction_date}")
        
        print("\nDaily Hours:")
        for day, hours in timesheet_data.daily_hours.items():
            print(f"  {day}: {hours} hours")
        
        if timesheet_data.task_breakdown:
            print("\nTask Breakdown:")
            for task, hours in timesheet_data.task_breakdown.items():
                print(f"  {task}: {hours} hours")
        
        print("="*50)

def main():
    """Test the enhanced OCR timesheet parser"""
    # Test with the timesheet image
    image_path = "Timesheet.png"
    
    if not os.path.exists(image_path):
        print(f"Image file {image_path} not found. Please save your timesheet image with this name.")
        return
    
    # Create enhanced OCR parser
    ocr_parser = EnhancedTimesheetOCR()
    
    # Process the image
    timesheet_data = ocr_parser.process_timesheet_image(image_path)
    
    # Print results
    ocr_parser.print_extraction_results(timesheet_data)
    
    # Calculate total from daily hours for verification
    if timesheet_data and timesheet_data.daily_hours:
        calculated_total = sum(timesheet_data.daily_hours.values())
        print(f"\nVerification - Calculated total from daily hours: {calculated_total} hours")
        print(f"Extracted total hours: {timesheet_data.total_hours} hours")
        print(f"Match: {'✓' if calculated_total == timesheet_data.total_hours else '✗'}")

if __name__ == "__main__":
    main()
