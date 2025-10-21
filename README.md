# PayrollTimecardAgent

A comprehensive payroll automation system that processes timesheet data from both Excel files and image files using OCR technology.

## 🚀 Features

- **Excel Processing**: Automatically extracts employee names and hours from Excel timesheet files
- **OCR Processing**: Uses Tesseract OCR to extract data from timesheet images
- **Automatic Installation**: Self-installing package with automatic dependency management
- **Cross-Platform**: Works on macOS, Linux, and Windows
- **Payroll Integration**: Updates payroll totals files automatically
- **Unified Processing**: Handles mixed file types in a single workflow

## 📦 Quick Start

### Option 1: Use the Pre-built Package (Recommended)

1. Download the latest release: `PayrollAutomation_20251020_141202.zip`
2. Extract the ZIP file
3. Run the installer: `python install.py`
4. Place your timesheet files in the `input_files` folder
5. Run the automation: `python run_payroll_automation.py`

### Option 2: Manual Setup

1. Clone this repository
2. Install Python dependencies: `pip install -r requirements.txt`
3. Install Tesseract OCR (see installation guide below)
4. Run the automation: `python unified_payroll_processor.py`

## 🔧 Installation

### Automatic Installation (Recommended)
The package includes an automatic installer that handles everything:

```bash
python install.py
```

This will:
- Install all Python dependencies
- Install Tesseract OCR automatically
- Set up the required folder structure
- Create example files

### Manual Installation

#### Python Dependencies
```bash
pip install pandas openpyxl xlrd pytesseract opencv-python Pillow
```

#### Tesseract OCR

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

**Linux (CentOS/RHEL):**
```bash
sudo yum install tesseract
```

**Windows:**
Download from: https://github.com/UB-Mannheim/tesseract/wiki

## 📁 File Structure

```
PayrollTimecardAgent/
├── PayrollAutomation_20251020_141202.zip    # Complete package
├── PayrollAutomation_20251020_141202/       # Extracted package
│   ├── unified_payroll_processor.py         # Main processor
│   ├── enhanced_timesheet_ocr.py            # OCR functionality
│   ├── extract_timesheet_data.py            # Excel processor
│   ├── install_tesseract.py                 # Tesseract installer
│   ├── run_payroll_automation.py            # Simple launcher
│   ├── install.py                           # Package installer
│   ├── requirements.txt                     # Python dependencies
│   ├── README.md                            # Documentation
│   ├── USER_GUIDE.md                        # User instructions
│   ├── input_files/                         # Place timesheets here
│   ├── output/                              # Results saved here
│   └── examples/                            # Example files
├── 3 - Payroll Totals.xlsx                  # Updated payroll file
├── Timesheet_Summary.xlsx                   # Summary report
└── extracted_timesheet_data.csv             # Raw extracted data
```

## 🎯 Usage

### Basic Usage

1. **Prepare Files**: Place your timesheet files (Excel or images) in the `input_files` folder
2. **Run Processing**: Execute `python run_payroll_automation.py`
3. **Check Results**: View results in the `output` folder

### Supported File Types

- **Excel Files**: `.xlsx`, `.xls` timesheet files
- **Image Files**: `.png`, `.jpg`, `.jpeg` timesheet images
- **Mixed Processing**: Handles both types in a single run

### Output Files

- `consolidated_payroll_data.csv`: All extracted data
- `processing_report_[timestamp].txt`: Detailed processing log
- `3 - Payroll Totals.xlsx`: Updated payroll file with new sheet
- `Timesheet_Summary.xlsx`: Clean summary of employee hours

## 🔍 OCR Capabilities

The system can extract the following data from timesheet images:

- **Employee Name**: Automatically detected from the timesheet
- **Week Period**: Date range for the timesheet
- **Daily Hours**: Hours worked each day
- **Total Hours**: Sum of all hours
- **Task Breakdown**: Hours by task/project (if available)

## 📊 Example Output

```
==================================================
TIMESHEET DATA EXTRACTION RESULTS
==================================================
Employee Name: Chad Baker
Week Period: 5-14 October 2025
Total Hours: 40
Extraction Date: 2025-10-20 14:11:10

Daily Hours:
  Mon 6: 8 hours
  Tue 7: 8 hours
  Wed 8: 8 hours
  Thu 9: 8 hours
  Fri 10: 8 hours

Task Breakdown:
  KTLO: 19 hours
  Enhancement: 7 hours
  Support: 14 hours
==================================================
```

## 🛠️ Technical Details

### Core Components

- **UnifiedPayrollProcessor**: Main orchestrator for all processing
- **EnhancedTimesheetOCR**: Advanced OCR with automatic Tesseract installation
- **TimesheetExtractor**: Excel file processing engine
- **TesseractManager**: Cross-platform Tesseract installation and management

### Dependencies

- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file processing
- **xlrd**: Legacy Excel file support
- **pytesseract**: Python wrapper for Tesseract OCR
- **opencv-python**: Image preprocessing
- **Pillow**: Image handling

## 🚨 Troubleshooting

### Common Issues

1. **Tesseract not found**: Run `python install.py` to install automatically
2. **Permission errors**: Ensure you have write permissions in the output folder
3. **Excel file errors**: Make sure Excel files are not open in another application
4. **OCR accuracy**: Ensure timesheet images are clear and well-lit

### Getting Help

- Check the `USER_GUIDE.md` for detailed instructions
- Review the processing report in the `output` folder
- Ensure all dependencies are properly installed

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📞 Support

For issues and questions:
- Create an issue in this repository
- Check the troubleshooting section above
- Review the user guide for detailed instructions

---

**Last Updated**: October 20, 2025  
**Version**: 2.0 (Enhanced OCR Package)  
**Author**: Rainy City Coder