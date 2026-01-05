"""
CSFA Report Automation - Unified Main Script
Orchestrates data fetching, report generation, and email sending.
Automatically runs at 7pm on weekdays (Monday-Friday).
"""

import os
import sys
import logging
import time
import schedule
from datetime import datetime, timedelta
from typing import Dict, List, Tuple
from dotenv import load_dotenv

# Import local modules
from api_client import get_orders, get_timesheet, get_order_details
from generate_detailed_report import generate_detailed_report, ReportConfig
from send_report import send_report

# Load environment variables
load_dotenv()

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================

def setup_logging() -> logging.Logger:
    """Configure logging with both file and console handlers."""
    log_level = os.getenv("LOG_LEVEL", "INFO")
    log_file = os.getenv("LOG_FILE", "report_generation.log")

    # Create logger
    logger = logging.getLogger("csfa_report")
    logger.setLevel(log_level)

    # Clear existing handlers
    logger.handlers.clear()

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_level)
    console_format = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    console_handler.setFormatter(console_format)

    # File handler
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(log_level)
    file_format = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    file_handler.setFormatter(file_format)

    # Add handlers
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger

logger = setup_logging()


# ============================================================================
# CONFIGURATION
# ============================================================================

class Config:
    """Configuration class for API credentials and parameters."""

    # Authentication
    ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")
    LARAVEL_TOKEN = os.getenv("LARAVEL_TOKEN")
    SAT_SESSION = os.getenv("SAT_SESSION")
    SAT_USER_ID = os.getenv("SAT_USER_ID", "57")
    XSRF_TOKEN = os.getenv("XSRF_TOKEN")

    # Date configuration
    ORDER_DATE = os.getenv("ORDER_DATE")
    ORDER_DATE_RANGE = os.getenv("ORDER_DATE_RANGE")

    # Other config
    COUNTRY_ID = int(os.getenv("COUNTRY_ID", "149"))
    HOST = os.getenv("HOST", "tintasberger.solutechlabs.com")

    # Email config
    SEND_EMAIL = os.getenv("SEND_EMAIL", "true").lower() == "true"

    # Scheduling config
    RUN_MODE = os.getenv("RUN_MODE", "scheduled")  # "scheduled" or "once"
    REPORT_TIME = os.getenv("REPORT_TIME", "19:00")  # 7pm in 24-hour format

    @classmethod
    def validate(cls) -> None:
        """Validate that all required config is present."""
        required = ["ACCESS_TOKEN", "LARAVEL_TOKEN", "SAT_SESSION", "XSRF_TOKEN"]
        missing = [key for key in required if not getattr(cls, key)]
        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    @classmethod
    def get_dates(cls) -> Tuple[str, str, str]:
        """
        Get dates for report.
        Returns: (order_date, order_date_range, display_date)
        """
        today = datetime.now()  # Changed from yesterday to today
        date_str = today.strftime("%Y-%m-%d")

        order_date = cls.ORDER_DATE or today.strftime("%a+%b+%d+%Y")
        order_date_range = cls.ORDER_DATE_RANGE or f"{date_str} - {date_str}"
        display_date = date_str

        return order_date, order_date_range, display_date


# ============================================================================
# API PARAMETER BUILDERS
# ============================================================================

def build_orders_query_string(date: str, country_id: int = 149) -> str:
    """Build query string for orders API."""
    return (
        f"?start_date={date}"
        f"&end_date={date}"
        f"&country_id[]={country_id}"
        f"&stage=0"
        f"&page=1"
        f"&per_page=25"
        f"&orderWorkflowId=1"
    )


def build_timesheet_headers() -> Dict:
    """Build headers for timesheet API."""
    return {
        "Host": Config.HOST,
        "Referer": f"https://{Config.HOST}/timesheet",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "sat_user_id": Config.SAT_USER_ID,
        "laravel_token": Config.ACCESS_TOKEN,
        "XSRF-TOKEN": Config.XSRF_TOKEN,
        "sat_session": Config.SAT_SESSION,
    }


def build_timesheet_cookies() -> Dict:
    """Build cookies for timesheet API."""
    return {
        "sat_user_id": Config.SAT_USER_ID,
        "laravel_token": Config.ACCESS_TOKEN,
        "XSRF-TOKEN": Config.XSRF_TOKEN,
        "sat_session": Config.SAT_SESSION,
    }


def build_timesheet_params(date_range: str) -> Dict:
    """Build parameters for timesheet API."""
    return {
        "group_by": "",
        "survey_id": "",
        "rep_id": "",
        "customer_id": "",
        "product_category": "",
        "product_name": "",
        "reportparameter": "",
        "distributorid": 0,
        "stageid": 1,
        "sales_rep_id": "",
        "inventorytype": "virtual",
        "mtd": 1,
        "daterange": date_range,
        "groupdate": "all",
        "relationship": "",
        "status": "",
        "maincategoryselect": "",
        "timesheet_updated": False,
        "search_timesheet": "",
        "draw": 1,
        "columns[0][data]": "timesheet_id",
        "columns[0][name]": "timesheet_id",
        "columns[0][searchable]": True,
        "columns[0][orderable]": True,
        "order[0][column]": 0,
        "order[0][dir]": "desc",
        "start": 0,
        "length": 25,
        "search[value]": "",
        "search[regex]": False,
    }


# ============================================================================
# DATA FETCHING
# ============================================================================

def fetch_orders_data(order_date: str) -> List[Dict]:
    """
    Fetch orders data from API.

    Args:
        order_date: Date string for orders

    Returns:
        List of order dictionaries
    """
    logger.info("📦 Fetching orders...")
    try:
        query_string = build_orders_query_string(order_date, Config.COUNTRY_ID)
        orders_response = get_orders(Config.ACCESS_TOKEN, query_string)
        orders_data = orders_response.get('data', [])
        logger.info(f"✅ Found {len(orders_data)} orders")
        return orders_data
    except Exception as e:
        logger.error(f"❌ Failed to fetch orders: {e}", exc_info=True)
        raise


def fetch_timesheet_data(order_date_range: str) -> List[Dict]:
    """
    Fetch timesheet/visits data from API.

    Args:
        order_date_range: Date range string for timesheet

    Returns:
        List of visit dictionaries
    """
    logger.info("⏰ Fetching timesheet data...")
    try:
        timesheet_headers = build_timesheet_headers()
        timesheet_cookies = build_timesheet_cookies()
        timesheet_params = build_timesheet_params(order_date_range)

        timesheet_response = get_timesheet(
            timesheet_headers,
            timesheet_cookies,
            timesheet_params
        )
        visits_data = timesheet_response.get('data', [])
        logger.info(f"✅ Found {len(visits_data)} visits")
        return visits_data
    except Exception as e:
        logger.error(f"❌ Failed to fetch timesheet: {e}", exc_info=True)
        raise


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def generate_and_send_report() -> bool:
    """
    Main workflow: fetch data, generate report, send email.

    Returns:
        True if successful, False otherwise
    """
    start_time = datetime.now()

    try:
        # Validate configuration
        logger.info("🔍 Validating configuration...")
        Config.validate()

        # Get dates
        order_date, order_date_range, display_date = Config.get_dates()
        logger.info(f"📅 Generating report for: {display_date}")
        logger.info("=" * 70)

        # Fetch data
        orders_data = fetch_orders_data(order_date)
        visits_data = fetch_timesheet_data(order_date_range)

        # Generate report
        logger.info("\n📊 Generating detailed Excel report...")
        report_config = ReportConfig.from_env()
        generate_detailed_report(visits_data, orders_data, report_config)

        # Send email (if configured)
        if Config.SEND_EMAIL:
            logger.info("\n📧 Sending email report...")
            email_success = send_report(date_str=display_date)
            if not email_success:
                logger.warning("⚠️ Email sending failed, but report was generated")
        else:
            logger.info("📧 Email sending disabled (SEND_EMAIL=false)")
            email_success = True  # Don't fail if email is disabled

        # Calculate execution time
        duration = (datetime.now() - start_time).total_seconds()

        logger.info("\n" + "=" * 70)
        logger.info(f"✅ Report generation complete!")
        logger.info(f"⏱️  Total execution time: {duration:.2f} seconds")
        logger.info(f"📁 Report saved: {report_config.output_file}")

        if Config.SEND_EMAIL and email_success:
            logger.info(f"📧 Email sent successfully")

        return True

    except Exception as e:
        logger.error(f"\n❌ Fatal error: {e}", exc_info=True)
        logger.error("=" * 70)
        logger.error("Report generation failed!")
        return False


def is_weekday() -> bool:
    """Check if today is a weekday (Monday-Friday)."""
    return datetime.now().weekday() < 5  # 0=Monday, 4=Friday


def scheduled_job() -> None:
    """Job that runs on schedule - only executes on weekdays."""
    current_time = datetime.now()
    day_name = current_time.strftime("%A")

    if is_weekday():
        logger.info(f"\n🕐 Scheduled job triggered on {day_name} at {current_time.strftime('%H:%M:%S')}")
        logger.info("=" * 70)
        logger.info("CSFA REPORT AUTOMATION - SCHEDULED RUN")
        logger.info("=" * 70)

        generate_and_send_report()
    else:
        logger.info(f"⏭️  Skipping report generation - {day_name} is not a weekday")


def run_scheduler() -> None:
    """Set up and run the scheduler for weekday reports at 7pm."""
    report_time = Config.REPORT_TIME

    logger.info("\n" + "=" * 70)
    logger.info("CSFA REPORT AUTOMATION - SCHEDULER MODE")
    logger.info("=" * 70)
    logger.info(f"📅 Schedule: Monday-Friday at {report_time}")
    logger.info(f"⏰ Current time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70)

    # Schedule for each weekday
    schedule.every().monday.at(report_time).do(scheduled_job)
    schedule.every().tuesday.at(report_time).do(scheduled_job)
    schedule.every().wednesday.at(report_time).do(scheduled_job)
    schedule.every().thursday.at(report_time).do(scheduled_job)
    schedule.every().friday.at(report_time).do(scheduled_job)

    logger.info("\n✅ Scheduler started. Waiting for scheduled times...")
    logger.info("💡 Press Ctrl+C to stop the scheduler\n")

    # Show next scheduled run
    next_run = schedule.next_run()
    if next_run:
        logger.info(f"📌 Next scheduled run: {next_run.strftime('%A, %Y-%m-%d at %H:%M:%S')}\n")

    # Keep the scheduler running
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # Check every minute
    except KeyboardInterrupt:
        logger.info("\n⚠️ Scheduler stopped by user")


def main() -> int:
    """
    Main entry point.

    Returns:
        Exit code (0 for success, 1 for failure)
    """
    try:
        run_mode = Config.RUN_MODE.lower()

        if run_mode == "once":
            # Run once immediately
            logger.info("\n" + "=" * 70)
            logger.info("CSFA REPORT AUTOMATION - SINGLE RUN MODE")
            logger.info("=" * 70)
            logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            logger.info("=" * 70 + "\n")

            success = generate_and_send_report()
            return 0 if success else 1

        elif run_mode == "scheduled":
            # Run in scheduler mode
            run_scheduler()
            return 0

        else:
            logger.error(f"❌ Invalid RUN_MODE: {run_mode}. Must be 'once' or 'scheduled'")
            return 1

    except KeyboardInterrupt:
        logger.warning("\n⚠️ Process interrupted by user")
        return 130  # Standard exit code for SIGINT
    except Exception as e:
        logger.error(f"\n❌ Unexpected error: {e}", exc_info=True)
        return 1
    finally:
        logger.info(f"\nFinished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info("=" * 70)


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
