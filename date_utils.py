"""
Date Utilities for CSFA Report Generation
Handles date calculations for daily reports (uses previous day's data)
"""

from datetime import datetime, timedelta
import os
import logging

logger = logging.getLogger(__name__)


def get_report_date() -> datetime:
    """
    Get the date for which the report should be generated.

    Rules:
    - If REPORT_DATE env var is set, use that
    - Otherwise, use TODAY's date
    - This is because reports run at end of business day (7 PM EAT) for today's data

    Returns:
        datetime object for the report date
    """
    # Check if date is specified in environment
    report_date_str = os.getenv("REPORT_DATE", "").strip()

    if report_date_str:
        try:
            # Parse the specified date
            report_date = datetime.strptime(report_date_str, "%Y-%m-%d")
            logger.info(f"📅 Using specified report date: {report_date.strftime('%Y-%m-%d')}")
            return report_date
        except ValueError as e:
            logger.warning(f"Invalid REPORT_DATE format '{report_date_str}': {e}")
            logger.info("Falling back to today's date")

    # Default: Use today's date (end of business day report)
    today = datetime.now()
    logger.info(f"📅 Using today's date for report: {today.strftime('%Y-%m-%d')}")

    return today


def format_order_date(date: datetime) -> str:
    """
    Format date for ORDER_DATE parameter.

    Args:
        date: datetime object

    Returns:
        String in format "Mon+Feb+02+2026"

    Example:
        datetime(2026, 2, 2) -> "Mon+Feb+02+2026"
    """
    # Format: DayName+MonthName+DD+YYYY with + separators
    formatted = date.strftime("%a+%b+%d+%Y")
    logger.debug(f"Formatted ORDER_DATE: {formatted}")
    return formatted


def format_order_date_range(date: datetime) -> str:
    """
    Format date range for ORDER_DATE_RANGE parameter.
    For single-day reports, start and end are the same.

    Args:
        date: datetime object

    Returns:
        String in format "2026-02-02 - 2026-02-02"

    Example:
        datetime(2026, 2, 2) -> "2026-02-02 - 2026-02-02"
    """
    # Format: YYYY-MM-DD - YYYY-MM-DD
    date_str = date.strftime("%Y-%m-%d")
    formatted = f"{date_str} - {date_str}"
    logger.debug(f"Formatted ORDER_DATE_RANGE: {formatted}")
    return formatted


def get_previous_working_day(date: datetime = None) -> datetime:
    """
    Get the previous working day (Monday-Friday).
    Skips weekends.

    Args:
        date: Starting date (defaults to today)

    Returns:
        datetime object for previous working day

    Examples:
        Monday -> Previous Friday
        Tuesday -> Previous Monday
        Saturday -> Previous Friday
        Sunday -> Previous Friday
    """
    if date is None:
        date = datetime.now()

    # Go back one day
    previous_day = date - timedelta(days=1)

    # If it's Saturday (5) or Sunday (6), go back to Friday
    while previous_day.weekday() >= 5:
        previous_day -= timedelta(days=1)

    return previous_day


def should_generate_report(date: datetime = None) -> bool:
    """
    Check if a report should be generated for the given date.
    Only generates reports for Monday-Friday.

    Args:
        date: Date to check (defaults to today)

    Returns:
        True if report should be generated, False otherwise
    """
    if date is None:
        date = datetime.now()

    # Check if it's a weekday (Monday=0 to Friday=4)
    is_weekday = date.weekday() < 5

    if not is_weekday:
        logger.info(f"⏭️  Skipping report - {date.strftime('%A')} is not a working day")

    return is_weekday
