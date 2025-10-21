#!/usr/bin/env python3
"""
Unified Payroll Processor
Combines Excel file processing and OCR image processing into a single automation
"""

import os
import sys
import time
import json
import shutil
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path

# Import our existing modules
from extract_timesheet_data import TimesheetExtractor
from enhanced_timesheet_ocr import EnhancedTimesheetOCR, TimesheetData
from update_payroll_totals import create_new_sheet_with_timesheet_data

class UnifiedPayrollProcessor:
    """Unified processor that handles both Excel files and images"""
    
    def __init__(self, input_folder: str = "input_files", output_folder: str = "output"):
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.excel_extractor = TimesheetExtractor()
        self.ocr_parser = EnhancedTimesheetOCR()
        self.processed_files = []
        self.all_employee_data = []
        
        # Create folders if they don't exist
        self.input_folder.mkdir(exist_ok=True)
        self.output_folder.mkdir(exist_ok=True)
    
    def scan_input_folder(self) -> Dict[str, List[str]]:
        """Scan input folder for Excel and image files"""
        print("🔍 Scanning input folder for files...")
        
        files = {
            'excel': [],
            'images': [],
            'other': []
        }
        
        if not self.input_folder.exists():
            print(f"❌ Input folder '{self.input_folder}' not found")
            return files
        
        for file_path in self.input_folder.iterdir():
            if file_path.is_file():
                file_ext = file_path.suffix.lower()
                file_name = file_path.name.lower()
                
                # Check for Excel files
                if file_ext in ['.xlsx', '.xls'] and 'time' in file_name:
                    files['excel'].append(str(file_path))
                
                # Check for image files
                elif file_ext in ['.png', '.jpg', '.jpeg', '.bmp', '.tiff'] and 'time' in file_name:
                    files['images'].append(str(file_path))
                
                else:
                    files['other'].append(str(file_path))
        
        print(f"📊 Found files:")
        print(f"  📈 Excel files: {len(files['excel'])}")
        print(f"  🖼️  Image files: {len(files['images'])}")
        print(f"  📄 Other files: {len(files['other'])}")
        
        return files
    
    def process_excel_files(self, excel_files: List[str]) -> List[Dict]:
        """Process Excel timesheet files"""
        print("\n📈 Processing Excel files...")
        
        employee_data = []
        
        for excel_file in excel_files:
            print(f"  Processing: {Path(excel_file).name}")
            
            try:
                # Extract data from Excel file
                employees = self.excel_extractor.extract_from_file(excel_file)
                
                for employee in employees:
                    employee_dict = {
                        'source_type': 'excel',
                        'source_file': Path(excel_file).name,
                        'employee_name': employee.employee_name,
                        'total_hours': employee.total_hours,
                        'period': employee.period,
                        'title': employee.title,
                        'project_manager': employee.project_manager,
                        'client': employee.client,
                        'firm': employee.firm,
                        'extraction_date': employee.extraction_date
                    }
                    employee_data.append(employee_dict)
                    print(f"    ✓ {employee.employee_name}: {employee.total_hours} hours")
                
                self.processed_files.append(excel_file)
                
            except Exception as e:
                print(f"    ❌ Error processing {excel_file}: {e}")
        
        return employee_data
    
    def process_image_files(self, image_files: List[str]) -> List[Dict]:
        """Process image timesheet files using OCR"""
        print("\n🖼️  Processing image files...")
        
        employee_data = []
        
        for image_file in image_files:
            print(f"  Processing: {Path(image_file).name}")
            
            try:
                # Extract data from image using OCR
                timesheet_data = self.ocr_parser.process_timesheet_image(image_file)
                
                if timesheet_data:
                    employee_dict = {
                        'source_type': 'image',
                        'source_file': Path(image_file).name,
                        'employee_name': timesheet_data.employee_name,
                        'total_hours': timesheet_data.total_hours,
                        'period': timesheet_data.week_period,
                        'title': 'Credentialing Specialist',  # Default title
                        'project_manager': '',
                        'client': '',
                        'firm': '',
                        'extraction_date': timesheet_data.extraction_date
                    }
                    employee_data.append(employee_dict)
                    print(f"    ✓ {timesheet_data.employee_name}: {timesheet_data.total_hours} hours")
                else:
                    print(f"    ❌ No data extracted from {image_file}")
                
                self.processed_files.append(image_file)
                
            except Exception as e:
                print(f"    ❌ Error processing {image_file}: {e}")
        
        return employee_data
    
    def consolidate_employee_data(self, excel_data: List[Dict], image_data: List[Dict]) -> List[Dict]:
        """Consolidate data from Excel and image sources"""
        print("\n📊 Consolidating employee data...")
        
        all_data = excel_data + image_data
        consolidated = {}
        
        # Group by employee name and take the most recent entry
        for employee in all_data:
            name = employee['employee_name']
            if name not in consolidated:
                consolidated[name] = employee
            else:
                # Keep the one with more recent extraction date
                if employee['extraction_date'] > consolidated[name]['extraction_date']:
                    consolidated[name] = employee
        
        consolidated_list = list(consolidated.values())
        
        print(f"📈 Consolidated {len(all_data)} entries into {len(consolidated_list)} unique employees")
        
        return consolidated_list
    
    def save_results(self, employee_data: List[Dict]) -> str:
        """Save consolidated results to CSV"""
        print("\n💾 Saving results...")
        
        import pandas as pd
        
        # Create DataFrame
        df = pd.DataFrame(employee_data)
        
        # Save to CSV
        csv_file = self.output_folder / "consolidated_payroll_data.csv"
        df.to_csv(csv_file, index=False)
        
        print(f"✅ Results saved to: {csv_file}")
        
        return str(csv_file)
    
    def update_payroll_totals(self, employee_data: List[Dict]) -> bool:
        """Update the Payroll Totals file"""
        print("\n📈 Updating Payroll Totals file...")
        
        try:
            # Save data to temporary CSV for the update function
            import pandas as pd
            df = pd.DataFrame(employee_data)
            temp_csv = "temp_payroll_summary.csv"
            df.to_csv(temp_csv, index=False)
            
            # Use existing update function
            success = create_new_sheet_with_timesheet_data()
            
            # Clean up temp file
            if os.path.exists(temp_csv):
                os.remove(temp_csv)
            
            if success:
                print("✅ Payroll Totals file updated successfully!")
                return True
            else:
                print("❌ Failed to update Payroll Totals file")
                return False
                
        except Exception as e:
            print(f"❌ Error updating Payroll Totals: {e}")
            return False
    
    def generate_report(self, employee_data: List[Dict]) -> str:
        """Generate a processing report"""
        print("\n📋 Generating processing report...")
        
        report_file = self.output_folder / f"processing_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        with open(report_file, 'w') as f:
            f.write("PAYROLL PROCESSING REPORT\n")
            f.write("=" * 50 + "\n")
            f.write(f"Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Input Folder: {self.input_folder}\n")
            f.write(f"Output Folder: {self.output_folder}\n\n")
            
            f.write("PROCESSED FILES:\n")
            f.write("-" * 20 + "\n")
            for file_path in self.processed_files:
                f.write(f"✓ {Path(file_path).name}\n")
            
            f.write(f"\nEMPLOYEE DATA:\n")
            f.write("-" * 20 + "\n")
            total_hours = 0
            for employee in employee_data:
                f.write(f"{employee['employee_name']}: {employee['total_hours']} hours ({employee['source_type']})\n")
                total_hours += employee['total_hours']
            
            f.write(f"\nSUMMARY:\n")
            f.write("-" * 20 + "\n")
            f.write(f"Total Employees: {len(employee_data)}\n")
            f.write(f"Total Hours: {total_hours}\n")
            f.write(f"Average Hours: {total_hours / len(employee_data) if employee_data else 0:.1f}\n")
        
        print(f"✅ Report saved to: {report_file}")
        return str(report_file)
    
    def run_full_processing(self) -> bool:
        """Run the complete processing pipeline"""
        print("🚀 Starting Unified Payroll Processing")
        print("=" * 60)
        
        try:
            # Step 1: Scan for files
            files = self.scan_input_folder()
            
            if not files['excel'] and not files['images']:
                print("❌ No timesheet files found in input folder")
                print(f"   Please add Excel (.xlsx, .xls) or image (.png, .jpg) files to: {self.input_folder}")
                return False
            
            # Step 2: Process Excel files
            excel_data = []
            if files['excel']:
                excel_data = self.process_excel_files(files['excel'])
            
            # Step 3: Process image files
            image_data = []
            if files['images']:
                image_data = self.process_image_files(files['images'])
            
            # Step 4: Consolidate data
            consolidated_data = self.consolidate_employee_data(excel_data, image_data)
            
            if not consolidated_data:
                print("❌ No employee data extracted")
                return False
            
            # Step 5: Save results
            csv_file = self.save_results(consolidated_data)
            
            # Step 6: Update Payroll Totals
            payroll_updated = self.update_payroll_totals(consolidated_data)
            
            # Step 7: Generate report
            report_file = self.generate_report(consolidated_data)
            
            # Final summary
            print("\n" + "=" * 60)
            print("🎉 PROCESSING COMPLETED SUCCESSFULLY!")
            print("=" * 60)
            print(f"📊 Processed {len(self.processed_files)} files")
            print(f"👥 Found {len(consolidated_data)} employees")
            print(f"⏰ Total hours: {sum(e['total_hours'] for e in consolidated_data)}")
            print(f"📁 Results saved to: {self.output_folder}")
            print(f"📋 Report: {report_file}")
            
            if payroll_updated:
                print("✅ Payroll Totals file updated")
            else:
                print("⚠️  Payroll Totals file not updated (check for errors)")
            
            return True
            
        except Exception as e:
            print(f"❌ Processing failed: {e}")
            return False

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Unified Payroll Processor')
    parser.add_argument('--input', '-i', default='input_files', 
                       help='Input folder containing timesheet files (default: input_files)')
    parser.add_argument('--output', '-o', default='output', 
                       help='Output folder for results (default: output)')
    
    args = parser.parse_args()
    
    # Create processor
    processor = UnifiedPayrollProcessor(args.input, args.output)
    
    # Run processing
    success = processor.run_full_processing()
    
    if success:
        print("\n✅ All done! Check the output folder for results.")
        sys.exit(0)
    else:
        print("\n❌ Processing failed. Check the error messages above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
