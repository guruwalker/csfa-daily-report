"""
CSFA Report Automation - Unified Main Script
Orchestrates data fetching, report generation, and email sending.
"""

import os
import sys
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Tuple
from dotenv import load_dotenv

# Import local modules
from api_client import get_orders, get_timesheet, get_order_details
from generate_detailed_report import generate_detailed_report, ReportConfig
from send_report import send_report
from holiday_checker import is_holiday

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
# DATE UTILITIES
# ============================================================================

def get_report_date() -> datetime:
    """
    Get the report date from environment or use today.

    Since the report runs at 7 PM EAT, we report on the same business day.
    For example:
    - Monday 7 PM → Monday's report
    - Friday 7 PM → Friday's report

    Returns:
        datetime object for the report date
    """
    date_str = os.getenv("REPORT_DATE", "")

    if date_str:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            logger.warning(f"⚠️ Invalid REPORT_DATE format: {date_str}, using today")

    # Default to today (same business day at 7 PM)
    return datetime.now()


def format_date_for_api(dt: datetime) -> str:
    """
    Format date for API order endpoint.
    Example: Mon+Dec+15+2025

    Args:
        dt: datetime object

    Returns:
        Formatted date string like "Mon+Dec+15+2025"
    """
    return dt.strftime("%a+%b+%d+%Y")


def format_date_range(dt: datetime) -> str:
    """
    Format date range for API (single day).
    Example: 2025-12-15 - 2025-12-15

    Args:
        dt: datetime object

    Returns:
        Date range string like "2025-12-15 - 2025-12-15"
    """
    date_str = dt.strftime("%Y-%m-%d")
    return f"{date_str} - {date_str}"


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

    # Other config
    COUNTRY_ID = int(os.getenv("COUNTRY_ID", "149"))
    HOST = os.getenv("HOST", "tintasberger.solutechlabs.com")

    # Email config
    SEND_EMAIL = os.getenv("SEND_EMAIL", "true").lower() == "true"

    @classmethod
    def validate(cls) -> None:
        """Validate that all required config is present."""
        required = ["ACCESS_TOKEN", "LARAVEL_TOKEN", "SAT_SESSION", "XSRF_TOKEN"]
        missing = [key for key in required if not getattr(cls, key)]
        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

        # Additional validation for token format
        if cls.ACCESS_TOKEN:
            # Check if token is masked (common in CI/CD)
            if cls.ACCESS_TOKEN == '***' or cls.ACCESS_TOKEN.startswith('***'):
                raise ValueError(
                    "ACCESS_TOKEN appears to be masked. "
                    "In GitHub Actions, ensure you're using secrets correctly: "
                    "${{ secrets.ACCESS_TOKEN }}"
                )

            # Check minimum length
            if len(cls.ACCESS_TOKEN) < 50:
                logger.warning(
                    f"⚠️ ACCESS_TOKEN seems short ({len(cls.ACCESS_TOKEN)} chars). "
                    "Verify it's the complete token."
                )

    @classmethod
    def get_dates(cls) -> Tuple[str, str, str, datetime]:
        """
        Get dates for report - AUTOMATICALLY calculated or from environment.

        Returns:
            Tuple of (order_date, order_date_range, display_date, report_date)
        """
        # Get report date (today by default, or from REPORT_DATE env var)
        report_date = get_report_date()

        # Allow manual override via environment variables (for testing/debugging)
        manual_order_date = os.getenv("ORDER_DATE")
        manual_date_range = os.getenv("ORDER_DATE_RANGE")

        if manual_order_date and manual_date_range:
            logger.info("📝 Using manual date override from environment variables")
            order_date = manual_order_date
            order_date_range = manual_date_range
            try:
                display_date = manual_date_range.split(" - ")[0].strip()
            except:
                display_date = report_date.strftime("%Y-%m-%d")
        else:
            # AUTOMATIC: Calculate dates from report_date
            logger.info("🤖 Automatically calculating dates for today's report")
            order_date = format_date_for_api(report_date)
            order_date_range = format_date_range(report_date)
            display_date = report_date.strftime("%Y-%m-%d")

        logger.info(f"📅 Report Date: {display_date}")
        logger.info(f"📅 Order Date (API): {order_date}")
        logger.info(f"📅 Date Range: {order_date_range}")

        return order_date, order_date_range, display_date, report_date


# ============================================================================
# API PARAMETER BUILDERS
# ============================================================================

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


def build_orders_headers() -> Dict:
    """Build headers for orders API."""
    return {
        "Host": Config.HOST,
        "Referer": f"https://{Config.HOST}/orders",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "sat_user_id": Config.SAT_USER_ID,
        "laravel_token": Config.ACCESS_TOKEN,
        "XSRF-TOKEN": Config.XSRF_TOKEN,
        "sat_session": Config.SAT_SESSION,
    }


def build_orders_cookies() -> Dict:
    """Build cookies for orders API."""
    return {
        "sat_user_id": Config.SAT_USER_ID,
        "laravel_token": Config.ACCESS_TOKEN,
        "XSRF-TOKEN": Config.XSRF_TOKEN,
        "sat_session": Config.SAT_SESSION,
    }


def build_orders_params(order_date: str, order_date_range: str) -> Dict:
    """Build parameters for orders API."""
    return {
        "start_date": order_date,
        "end_date": order_date,
        "country_id[]": str(Config.COUNTRY_ID),
        "stage": "0",
        "page": "1",
        "per_page": "25",
        "orderWorkflowId": "1",
    }


# ============================================================================
# DATA FETCHING
# ============================================================================

def fetch_timesheet_data(order_date_range: str) -> List[Dict]:
    """
    Fetch timesheet/visits data from API.

    Args:
        order_date_range: Date range string for timesheet

    Returns:
        List of visit dictionaries (empty list if no data or error)
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

        if not visits_data:
            logger.warning("⚠️  No visits found for this date (holiday or no activity)")
        else:
            logger.info(f"✅ Found {len(visits_data)} visits")

        return visits_data

    except Exception as e:
        logger.error(f"❌ Failed to fetch timesheet: {e}", exc_info=True)
        logger.warning("⚠️  Continuing with empty visits data")
        return []


def fetch_orders_data(order_date: str, order_date_range: str) -> List[Dict]:
    """
    Fetch orders data from API.

    Args:
        order_date: Single date for orders
        order_date_range: Date range for orders

    Returns:
        List of order dictionaries (empty list if no data or error)
    """
    logger.info("📦 Fetching orders data...")
    try:
        orders_headers = build_orders_headers()
        orders_cookies = build_orders_cookies()
        orders_params = build_orders_params(order_date, order_date_range)

        orders_response = get_orders(
            orders_headers,
            orders_cookies,
            orders_params
        )
        orders_data = orders_response.get('data', [])

        if not orders_data:
            logger.warning("⚠️  No orders found for this date (holiday or no activity)")
        else:
            logger.info(f"✅ Found {len(orders_data)} orders")

        return orders_data

    except Exception as e:
        logger.error(f"❌ Failed to fetch orders: {e}", exc_info=True)
        logger.warning("⚠️  Continuing with empty orders data")
        return []


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

        # Get dates (automatically calculated)
        order_date, order_date_range, display_date, report_date = Config.get_dates()
        logger.info("=" * 70)

        # Check if this is a holiday
        if is_holiday(report_date):
            logger.info(" Holiday detected - skipping data fetch")
            logger.info("📝 Generating holiday report...")

            # Generate holiday report (no data needed)
            report_config = ReportConfig.from_env()
            generate_detailed_report([], [], report_date, report_config)

            # Send email (if configured)
            if Config.SEND_EMAIL:
                logger.info("\n📧 Sending holiday report email...")
                email_success = send_report(date_str=display_date)
                if not email_success:
                    logger.warning("⚠️ Email sending failed, but report was generated")

            # Calculate execution time
            duration = (datetime.now() - start_time).total_seconds()

            logger.info("\n" + "=" * 70)
            logger.info(f"✅ Holiday report generation complete!")
            logger.info(f"⏱️  Total execution time: {duration:.2f} seconds")
            logger.info(f"📁 Report saved: {report_config.output_file}")

            return True

        # Not a holiday - proceed with normal data fetch
        logger.info("📊 Regular business day - fetching data...")

        # Fetch data
        orders_data = fetch_orders_data(order_date, order_date_range)
        visits_data = fetch_timesheet_data(order_date_range)

        # Generate report
        logger.info("\n📊 Generating detailed Excel report...")
        report_config = ReportConfig.from_env()

        # Pass both report_date (datetime) and report_config to generate_detailed_report
        generate_detailed_report(visits_data, orders_data, report_date, report_config)

        # Send email (if configured)
        if Config.SEND_EMAIL:
            logger.info("\n📧 Sending email report...")
            email_success = send_report(date_str=display_date)
            if not email_success:
                logger.warning("⚠️ Email sending failed, but report was generated")
        else:
            logger.info("📧 Email sending disabled (SEND_EMAIL=false)")
            email_success = True

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


def main() -> int:
    """
    Main entry point.

    Returns:
        Exit code (0 for success, 1 for failure)
    """
    logger.info("\n" + "=" * 70)
    logger.info("CSFA REPORT AUTOMATION")
    logger.info("=" * 70)
    logger.info(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("=" * 70 + "\n")

    try:
        success = generate_and_send_report()
        return 0 if success else 1
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Process interrupted by user")
        return 130
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
