"""
CSFA Report Web Interface
Flask application for generating and previewing reports with date selection.
"""

from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from datetime import datetime, timedelta
import os
import logging
from pathlib import Path
import traceback

# Import local modules
from generate_detailed_report import generate_detailed_report, ReportConfig
from main import Config as MainConfig, fetch_timesheet_data, fetch_orders_data, format_date_for_api, format_date_range
from holiday_checker import is_holiday

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "csfa-report-secret-key-change-in-production")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configure upload folder for generated reports
REPORTS_FOLDER = Path("generated_reports")
REPORTS_FOLDER.mkdir(exist_ok=True)


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def validate_date(date_str: str) -> tuple[bool, str, datetime]:
    """
    Validate date string and return parsed datetime.

    Returns:
        Tuple of (is_valid, error_message, datetime_object)
    """
    if not date_str:
        return False, "Date is required", None

    try:
        report_date = datetime.strptime(date_str, "%Y-%m-%d")

        # Check if date is not in the future
        if report_date.date() > datetime.now().date():
            return False, "Cannot generate reports for future dates", None

        # Check if date is a weekend
        if report_date.weekday() >= 5:
            return False, "Cannot generate reports for weekends (Saturday/Sunday)", None

        return True, "", report_date

    except ValueError:
        return False, "Invalid date format. Use YYYY-MM-DD", None


def get_report_summary(report_file: str) -> dict:
    """
    Extract summary information from generated report.

    Returns:
        Dictionary with report stats
    """
    import pandas as pd
    from openpyxl import load_workbook

    try:
        if not os.path.exists(report_file):
            return {"error": "Report file not found"}

        # Read Day Summary sheet
        df_summary = pd.read_excel(report_file, sheet_name="Day Summary")

        # Calculate stats
        total_visited = int(df_summary["CUSTOMERS VISITED"].sum())
        active_reps = int(df_summary[df_summary["APP USAGE"] == "Used App"].shape[0])
        total_reps = len(df_summary)

        # Get sheet names
        wb = load_workbook(report_file, read_only=True)
        sheets = wb.sheetnames
        wb.close()

        return {
            "total_customers_visited": total_visited,
            "active_salespeople": f"{active_reps}/{total_reps}",
            "sheets": sheets,
            "file_size": f"{os.path.getsize(report_file) / 1024:.1f} KB"
        }

    except Exception as e:
        logger.error(f"Error reading report: {e}")
        return {"error": str(e)}


# ============================================================================
# ROUTES
# ============================================================================

@app.route('/')
def index():
    """Home page with date selection form."""
    # Get default date (today)
    default_date = datetime.now().strftime("%Y-%m-%d")

    # Get last 30 business days for quick selection
    today = datetime.now()
    recent_dates = []

    for i in range(30):
        date = today - timedelta(days=i)
        if date.weekday() < 5:  # Weekday
            recent_dates.append({
                "date": date.strftime("%Y-%m-%d"),
                "display": date.strftime("%A, %B %d, %Y")
            })

    return render_template('index.html',
                          default_date=default_date,
                          recent_dates=recent_dates)


@app.route('/generate', methods=['POST'])
def generate_report():
    """Generate report for selected date."""
    try:
        # Get form data
        report_date_str = request.form.get('report_date')
        send_email = request.form.get('send_email') == 'on'

        # Validate date
        is_valid, error_msg, report_date = validate_date(report_date_str)
        if not is_valid:
            flash(error_msg, 'error')
            return redirect(url_for('index'))

        logger.info(f"Generating report for {report_date_str}")

        # Check if holiday
        if is_holiday(report_date):
            flash(f"{report_date_str} is a public holiday. Generating holiday report...", 'info')

        # Validate API credentials
        MainConfig.validate()

        # Calculate API dates
        order_date = format_date_for_api(report_date)
        order_date_range = format_date_range(report_date)

        # Fetch data
        logger.info("Fetching data from API...")
        flash("Fetching data from API...", 'info')

        orders_data = fetch_orders_data(order_date, order_date_range)
        visits_data = fetch_timesheet_data(order_date_range)

        # Generate report
        logger.info("Generating Excel report...")
        flash("Generating Excel report...", 'info')

        report_config = ReportConfig.from_env()

        # Save to dated filename
        report_filename = f"CSFA_Report_{report_date_str}.xlsx"
        report_path = REPORTS_FOLDER / report_filename
        report_config.output_file = str(report_path)

        generate_detailed_report(visits_data, orders_data, report_date, report_config)

        # Send email if requested
        if send_email:
            logger.info("Sending email report...")
            from send_report import send_report
            email_success = send_report(
                excel_file=str(report_path),
                summary_sheet="Day Summary",  # Use the new sheet name
                date_str=report_date_str
            )

            if email_success:
                flash("Report generated and email sent successfully!", 'success')
            else:
                flash("Report generated but email failed to send", 'warning')
        else:
            flash("Report generated successfully!", 'success')

        # Get report summary
        summary = get_report_summary(str(report_path))

        return render_template('result.html',
                             report_date=report_date_str,
                             report_file=report_filename,
                             summary=summary,
                             email_sent=send_email)

    except Exception as e:
        logger.error(f"Error generating report: {e}")
        logger.error(traceback.format_exc())
        flash(f"Error generating report: {str(e)}", 'error')
        return redirect(url_for('index'))


@app.route('/download/<filename>')
def download_report(filename):
    """Download generated report."""
    try:
        report_path = REPORTS_FOLDER / filename

        if not report_path.exists():
            flash("Report file not found", 'error')
            return redirect(url_for('index'))

        return send_file(
            report_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        logger.error(f"Error downloading report: {e}")
        flash(f"Error downloading report: {str(e)}", 'error')
        return redirect(url_for('index'))


@app.route('/preview/<filename>')
def preview_report(filename):
    """Preview report summary without downloading."""
    try:
        report_path = REPORTS_FOLDER / filename

        if not report_path.exists():
            return jsonify({"error": "Report file not found"}), 404

        summary = get_report_summary(str(report_path))

        # Read first few rows of Day Summary for preview
        import pandas as pd
        df = pd.read_excel(report_path, sheet_name="Day Summary")

        preview_html = df.head(10).to_html(classes='table table-striped', index=False)

        return jsonify({
            "summary": summary,
            "preview": preview_html
        })

    except Exception as e:
        logger.error(f"Error previewing report: {e}")
        return jsonify({"error": str(e)}), 500


@app.route('/health')
def health():
    """Health check endpoint for monitoring."""
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "reports_folder": str(REPORTS_FOLDER),
        "reports_count": len(list(REPORTS_FOLDER.glob("*.xlsx")))
    })


# ============================================================================
# ERROR HANDLERS
# ============================================================================

@app.errorhandler(404)
def not_found(e):
    return render_template('404.html'), 404


@app.errorhandler(500)
def server_error(e):
    return render_template('500.html'), 500


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    # Check for required environment variables
    try:
        MainConfig.validate()
        logger.info("✅ Configuration validated")
    except Exception as e:
        logger.error(f"❌ Configuration error: {e}")
        logger.warning("App will start but report generation may fail")

    # Run app
    port = int(os.getenv('PORT', 3001))
    debug = os.getenv('FLASK_DEBUG', 'False').lower() == 'true'

    logger.info(f"🚀 Starting CSFA Report Web Interface on port {port}")
    logger.info(f"📁 Reports will be saved to: {REPORTS_FOLDER.absolute()}")

    app.run(host='0.0.0.0', port=port, debug=debug)
