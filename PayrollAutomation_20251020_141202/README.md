# Payroll Automation with Browser Automation

This automation tool streamlines your payroll process by using browser automation to:

1. **Extract employee hours from emails** - Automatically logs into your email and searches for hours-related emails
2. **Extract hours from timesheet images** - Uses OCR to extract data from timesheet screenshots/images
3. **Enter data into Excel/Google Sheets** - Populates your spreadsheet with the extracted hours data
4. **Enter data into ADP** - Automatically fills in payroll information in your ADP portal

## Features

- **Browser-based automation** - No API access required, works with any web interface
- **OCR timesheet extraction** - Extract data from timesheet images using optical character recognition
- **Multi-email provider support** - Works with Gmail, Outlook, and other email services
- **Flexible Excel integration** - Supports Google Sheets and Excel Online
- **ADP integration** - Automatically enters payroll data into ADP
- **Smart data extraction** - Uses regex patterns to extract employee names and hours from emails and images
- **Configurable** - Easy to customize for your specific setup

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Install Playwright browsers:
```bash
playwright install
```

3. Install Tesseract OCR:
```bash
# macOS
brew install tesseract

# Ubuntu/Debian
sudo apt-get install tesseract-ocr

# Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

4. Configure your settings in `config.json`:
```json
{
  "email": {
    "email_url": "https://gmail.com",
    "username": "your_email@gmail.com",
    "password": "your_password"
  },
  "excel": {
    "type": "google_sheets",
    "username": "your_email@gmail.com", 
    "password": "your_password",
    "sheet_url": "https://docs.google.com/spreadsheets/d/your_sheet_id/edit"
  },
  "adp": {
    "login_url": "https://your-adp-portal.com",
    "username": "your_adp_username",
    "password": "your_adp_password"
  }
}
```

## Usage

### Option 1: Email-based automation
Run the automation with email extraction:
```bash
python payroll_automation.py
```

### Option 2: OCR-based automation
Run the automation with timesheet image extraction:
```bash
python payroll_automation.py --ocr timesheet_image.png
```

### Test OCR extraction
Test OCR functionality on a timesheet image:
```bash
python test_ocr_extraction.py
```

The script will:
1. **Email mode**: Open a browser window, log into your email, search for hours-related emails, extract employee names and hours
2. **OCR mode**: Extract data from timesheet image using optical character recognition
3. Open your Excel/Google Sheets and enter the data
4. Log into ADP and enter the payroll information

## Customization

### Email Search Terms
The script searches for emails containing these terms:
- "hours"
- "timesheet"
- "payroll"
- "overtime"
- "time worked"

You can modify the `search_terms` list in the `EmailExtractor.search_hours_emails()` method.

### Data Extraction Patterns
The script uses regex patterns to extract:
- Employee names from email content
- Hours worked from email content

You can customize these patterns in the `EmailExtractor.extract_hours_from_email()` method.

### ADP Interface
ADP interfaces vary significantly. You may need to customize the selectors in the `ADPManager` class based on your specific ADP setup.

## Security Notes

- Store your credentials securely
- Consider using environment variables for sensitive information
- The script runs in non-headless mode by default so you can monitor the automation
- Set `headless=True` in the `run_automation()` method for unattended operation

## Troubleshooting

1. **Login issues**: Make sure your credentials are correct and 2FA is disabled or handled appropriately
2. **Element not found**: The web interfaces may have changed. You may need to update the selectors
3. **Rate limiting**: The script includes delays between actions to avoid being blocked
4. **ADP customization**: ADP interfaces vary - you'll likely need to customize the ADP selectors for your specific setup

## Support

This automation is designed to be flexible and customizable. You may need to adjust the selectors and patterns based on your specific email provider, Excel setup, and ADP configuration.
