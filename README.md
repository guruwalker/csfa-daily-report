# CSFA Report Automation

Automated daily reporting system for Tintas Berger CSFA (Customer Sales Force Automation) data. Fetches sales visits, orders, and product details, generates Excel reports, and sends them via email.

## 📋 Features

- ✅ Automated data fetching from CSFA API
- ✅ Professional Excel report generation with multiple sheets
- ✅ Email delivery with HTML summary table
- ✅ Configurable via environment variables
- ✅ Comprehensive logging
- ✅ Error handling and retry logic
- ✅ Support for multiple sales representatives
- ✅ Automatic date handling (defaults to yesterday)

## 📁 Project Structure

```
csfa-report-automation/
├── .env                          # Configuration (create from .env.example)
├── .gitignore                    # Git ignore rules
├── requirements.txt              # Python dependencies
├── README.md                     # This file
├── main.py                       # Main orchestration script
├── api_client.py                 # API interaction layer
├── generate_detailed_report.py   # Report generation logic
├── send_report.py                # Email sending module
└── report_generation.log         # Log file (auto-generated)
```

## 🚀 Quick Start

### 1. Prerequisites

- Python 3.8 or higher
- pip (Python package manager)
- Access to Tintas Berger CSFA system

### 2. Installation

```bash
# Clone the repository
git clone <repository-url>
cd csfa-report-automation

# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 3. Configuration

Create a `.env` file in the project root:

```bash
cp .env.example .env
```

Edit `.env` and fill in your credentials:

```bash
# Required: API Authentication
ACCESS_TOKEN=your_access_token_here
LARAVEL_TOKEN=your_laravel_token_here
SAT_SESSION=your_session_token_here
XSRF_TOKEN=your_xsrf_token_here

# Required: Email Configuration
SENDER_EMAIL=your.email@company.com
EMAIL_PASSWORD=your_password_here
EMAIL_TO=recipient@company.com
EMAIL_CC=cc1@company.com,cc2@company.com

# Optional: Customize dates (leave empty for yesterday)
ORDER_DATE=
ORDER_DATE_RANGE=
```

### 4. Run the Report

```bash
# Generate report and send email
python main.py

# Or run individual modules
python generate_detailed_report.py  # Just generate report
python send_report.py               # Just send existing report
```

## 📧 Email Configuration

### Gmail Setup

If using Gmail SMTP, you need an **App Password**:

1. Enable 2-Step Verification on your Google Account
2. Go to: https://myaccount.google.com/apppasswords
3. Create an app password for "Mail"
4. Use this password in `EMAIL_PASSWORD`

### Custom SMTP

Update these in `.env`:

```bash
SMTP_SERVER=smtp.your-server.com
SMTP_PORT=587
SENDER_EMAIL=your.email@domain.com
EMAIL_PASSWORD=your_password
```

## 🔄 Automated Scheduling

### Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., Daily at 7:00 AM)
4. Action: Start a program
5. Program: `C:\path\to\venv\Scripts\python.exe`
6. Arguments: `C:\path\to\main.py`
7. Start in: `C:\path\to\project`

### Linux/Mac Cron Job

```bash
# Edit crontab
crontab -e

# Add line (runs daily at 7:00 AM)
0 7 * * * cd /path/to/project && /path/to/venv/bin/python main.py >> /path/to/logs/cron.log 2>&1
```

### Using Python Schedule Library

Create `scheduler.py`:

```python
import schedule
import time
from main import generate_and_send_report

# Schedule for 7:00 AM daily
schedule.every().day.at("07:00").do(generate_and_send_report)

print("Scheduler started. Waiting for scheduled time...")
while True:
    schedule.run_pending()
    time.sleep(60)
```

Run with:
```bash
python scheduler.py
```

## 🔧 Configuration Options

### Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `ACCESS_TOKEN` | ✅ | - | API access token |
| `LARAVEL_TOKEN` | ✅ | - | Laravel session token |
| `SAT_SESSION` | ✅ | - | SAT session token |
| `XSRF_TOKEN` | ✅ | - | XSRF token |
| `EMAIL_PASSWORD` | ✅ | - | Email account password |
| `SENDER_EMAIL` | ✅ | - | Sender email address |
| `EMAIL_TO` | ✅ | - | Recipient emails (comma-separated) |
| `EMAIL_CC` | ❌ | - | CC emails (comma-separated) |
| `EMAIL_BCC` | ❌ | - | BCC emails (comma-separated) |
| `SEND_EMAIL` | ❌ | `true` | Enable/disable email sending |
| `ORDER_DATE` | ❌ | Yesterday | Report date (URL encoded) |
| `OUTPUT_FILE` | ❌ | `Daily_CSFA_Report.xlsx` | Output filename |
| `LOG_LEVEL` | ❌ | `INFO` | Logging level |
| `DEBUG` | ❌ | `false` | Enable debug mode |

### Date Formats

- `ORDER_DATE`: URL-encoded format (e.g., `Thu+Dec+11+2025`)
- `ORDER_DATE_RANGE`: Date range format (e.g., `2025-12-11 - 2025-12-11`)
- Leave empty to automatically use yesterday's date

## 📊 Generated Reports

### Excel File Structure

1. **Summary Sheet**: Overview of all sales reps
   - Salesperson name
   - Customers visited
   - Order value from visits
   - Customers called
   - Order value from calls

2. **Individual Rep Sheets**: Detailed breakdown per sales rep
   - Customer name
   - Visit type (Visited/Called/Both)
   - Time spent
   - Product details
   - Order values

### Email Content

- Professional HTML formatted email
- Embedded summary table
- Excel file attachment
- Automatic date in subject line

## 🐛 Troubleshooting

### Common Issues

**Authentication Error**
```
❌ SMTP Authentication failed
```
**Solution**: Check `EMAIL_PASSWORD` in `.env`. For Gmail, use an App Password.

---

**Missing Access Token**
```
ValueError: ACCESS_TOKEN not found in environment variables
```
**Solution**: Ensure `.env` file exists and contains all required tokens.

---

**Excel File Not Found**
```
FileNotFoundError: Daily_CSFA_Report.xlsx not found
```
**Solution**: Run report generation first: `python main.py`

---

**Connection Timeout**
```
Connection Error: Cannot connect to smtp.example.com
```
**Solution**: Check SMTP server and port. Ensure firewall allows outbound connections.

### Debug Mode

Enable detailed logging:

```bash
# In .env file
DEBUG=true
LOG_LEVEL=DEBUG
SAVE_HTML_PREVIEW=true
```

This will:
- Show SMTP communication details
- Save HTML email preview
- Log all API calls
- Show detailed error traces

## 📝 Logging

Logs are saved to `report_generation.log` with rotation.

View logs:
```bash
# View entire log
cat report_generation.log

# View last 50 lines
tail -n 50 report_generation.log

# Follow log in real-time
tail -f report_generation.log
```

## 🔒 Security Best Practices

1. **Never commit `.env` file** to version control
2. Add `.env` to `.gitignore`
3. Use environment-specific `.env` files
4. Rotate credentials regularly
5. Use app-specific passwords for email
6. Restrict file permissions on `.env`:
   ```bash
   chmod 600 .env  # Read/write for owner only
   ```

## 🧪 Testing

Test individual components:

```python
# Test API connection
python -c "from api_client import get_orders; print('API OK')"

# Test report generation
python generate_detailed_report.py

# Test email sending
python send_report.py
```

## 📦 Dependencies

See `requirements.txt` for full list:

- `pandas`: Data manipulation
- `openpyxl`: Excel file generation
- `requests`: HTTP API calls
- `python-dotenv`: Environment variable management
- `urllib3`: HTTP retry logic
- `dataframe_image`: (Optional) Summary image export

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📄 License

[Your License Here]

## 👥 Authors

- Innocent Maina - Initial work

## 📞 Support

For issues or questions:
- Email: innocent.maina@robbialac.co.mz
- Create an issue in the repository

## 🔄 Version History

- **v1.0.0** (2024-12-17)
  - Initial release
  - Automated report generation
  - Email integration
  - Multi-rep support

## 🎯 Roadmap

- [ ] Web dashboard for viewing reports
- [ ] Support for multiple report formats (PDF, CSV)
- [ ] Advanced analytics and charts
- [ ] Mobile app integration
- [ ] Real-time notifications
- [ ] Database storage for historical data

---

**Made with ❤️ for Tintas Berger**
