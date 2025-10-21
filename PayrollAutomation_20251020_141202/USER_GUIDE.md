# Payroll Automation - User Guide

## Quick Start

1. **Install Dependencies**
   ```bash
   python install.py
   ```

2. **Add Your Files**
   - Place timesheet Excel files (.xlsx, .xls) in `input_files/`
   - Place timesheet images (.png, .jpg) in `input_files/`
   - Files should contain 'time' in the filename

3. **Run Automation**
   ```bash
   python run_payroll_automation.py
   ```

4. **Check Results**
   - Results will be in the `output/` folder
   - Payroll Totals file will be updated automatically

## Supported File Types

### Excel Files
- **Formats**: .xlsx, .xls
- **Requirements**: Must contain 'time' in filename
- **Data Extracted**: Employee name, total hours, period, title

### Image Files
- **Formats**: .png, .jpg, .jpeg, .bmp, .tiff
- **Requirements**: Must contain 'time' in filename
- **Data Extracted**: Employee name, total hours, week period (using OCR)

## Output Files

- `consolidated_payroll_data.csv` - All extracted employee data
- `processing_report_*.txt` - Detailed processing report
- `3 - Payroll Totals.xlsx` - Updated payroll file with new sheet
- `Timesheet_Summary.xlsx` - Simple summary of hours

## Troubleshooting

### Common Issues

1. **"No files found"**
   - Make sure files are in the `input_files/` folder
   - Check that filenames contain 'time'

2. **OCR not working**
   - Install Tesseract OCR: `brew install tesseract` (macOS)
   - Make sure images are clear and readable

3. **Excel processing fails**
   - Check that Excel files are not corrupted
   - Ensure files are not password protected

### Getting Help

- Check the processing report for detailed error information
- Ensure all dependencies are installed correctly
- Verify file formats and naming conventions

## Advanced Usage

### Command Line Options
```bash
python unified_payroll_processor.py --input custom_input_folder --output custom_output_folder
```

### Batch Processing
- Add multiple files to the input folder
- The automation will process all files in one run
- Results will be consolidated automatically

## Security Notes

- Keep your credentials secure
- Don't share the package with sensitive information
- Use environment variables for production deployments
