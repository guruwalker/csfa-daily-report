"""
Attendance Tracker Module
Manages daily attendance tracking in attendance.json
Tracks absent days for each salesperson and generates monthly summaries.
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, List, Set
from pathlib import Path

logger = logging.getLogger(__name__)


# ============================================================================
# ATTENDANCE DATA STRUCTURE
# ============================================================================

class AttendanceTracker:
    """
    Manages attendance tracking in attendance.json

    JSON Structure:
    {
        "2026": {
            "02": {
                "John Doe": [2, 5, 13, 18],
                "Jane Smith": [],
                "Mike Johnson": [1, 2, 3, 15, 20]
            },
            "03": {
                ...
            }
        }
    }
    """

    def __init__(self, json_file: str = "attendance.json"):
        """
        Initialize attendance tracker.

        Args:
            json_file: Path to JSON file for storing attendance data
        """
        self.json_file = json_file
        self.data = self._load_data()

    def _load_data(self) -> Dict:
        """Load attendance data from JSON file."""
        if not os.path.exists(self.json_file):
            logger.info(f"Creating new attendance file: {self.json_file}")
            return {}

        try:
            with open(self.json_file, 'r') as f:
                data = json.load(f)
            logger.info(f"Loaded attendance data from {self.json_file}")
            return data
        except json.JSONDecodeError as e:
            logger.error(f"Error reading {self.json_file}: {e}")
            logger.warning("Starting with empty attendance data")
            return {}
        except Exception as e:
            logger.error(f"Unexpected error loading attendance: {e}")
            return {}

    def _save_data(self) -> None:
        """Save attendance data to JSON file."""
        try:
            with open(self.json_file, 'w') as f:
                json.dump(self.data, f, indent=2)
            logger.info(f"Saved attendance data to {self.json_file}")
        except Exception as e:
            logger.error(f"Error saving attendance data: {e}")

    def mark_absent(self, salesperson: str, date: datetime) -> None:
        """
        Mark a salesperson as absent on a specific date.

        Args:
            salesperson: Name of the salesperson
            date: Date object for the absent day
        """
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        # Initialize nested structure if needed
        if year not in self.data:
            self.data[year] = {}
        if month not in self.data[year]:
            self.data[year][month] = {}
        if salesperson not in self.data[year][month]:
            self.data[year][month][salesperson] = []

        # Add day if not already present
        if day not in self.data[year][month][salesperson]:
            self.data[year][month][salesperson].append(day)
            self.data[year][month][salesperson].sort()  # Keep sorted
            logger.info(f"Marked {salesperson} absent on {date.strftime('%Y-%m-%d')}")
            self._save_data()
        else:
            logger.debug(f"{salesperson} already marked absent on {date.strftime('%Y-%m-%d')}")

    def mark_present(self, salesperson: str, date: datetime) -> None:
        """
        Mark a salesperson as present (remove from absent list if exists).

        Args:
            salesperson: Name of the salesperson
            date: Date object for the present day
        """
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        # Check if marked as absent
        if (year in self.data and
            month in self.data[year] and
            salesperson in self.data[year][month] and
            day in self.data[year][month][salesperson]):

            self.data[year][month][salesperson].remove(day)
            logger.info(f"Marked {salesperson} present on {date.strftime('%Y-%m-%d')} (removed from absent list)")
            self._save_data()

    def get_absent_days(self, salesperson: str, year: int, month: int) -> List[int]:
        """
        Get list of absent days for a salesperson in a specific month.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            List of day numbers when salesperson was absent
        """
        year_str = str(year)
        month_str = f"{month:02d}"

        if (year_str in self.data and
            month_str in self.data[year_str] and
            salesperson in self.data[year_str][month_str]):
            return self.data[year_str][month_str][salesperson].copy()

        return []

    def get_monthly_summary(self, year: int, month: int) -> Dict[str, List[int]]:
        """
        Get attendance summary for all salespeople in a specific month.

        Args:
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            Dictionary mapping salesperson names to their absent days
        """
        year_str = str(year)
        month_str = f"{month:02d}"

        if year_str in self.data and month_str in self.data[year_str]:
            return self.data[year_str][month_str].copy()

        return {}

    def ensure_salesperson_exists(self, salesperson: str, year: int, month: int) -> None:
        """
        Ensure a salesperson exists in the tracking for a given month.
        Initializes with empty absent days list if they don't exist.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)
        """
        year_str = str(year)
        month_str = f"{month:02d}"

        if year_str not in self.data:
            self.data[year_str] = {}
        if month_str not in self.data[year_str]:
            self.data[year_str][month_str] = {}
        if salesperson not in self.data[year_str][month_str]:
            self.data[year_str][month_str][salesperson] = []
            self._save_data()

    def is_absent(self, salesperson: str, date: datetime) -> bool:
        """
        Check if a salesperson is marked as absent on a specific date.

        Args:
            salesperson: Name of the salesperson
            date: Date to check

        Returns:
            True if marked absent, False otherwise
        """
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        return (year in self.data and
                month in self.data[year] and
                salesperson in self.data[year][month] and
                day in self.data[year][month][salesperson])

    def get_total_absent_days(self, salesperson: str, year: int, month: int) -> int:
        """
        Get total number of absent days for a salesperson in a month.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            Number of absent days
        """
        absent_days = self.get_absent_days(salesperson, year, month)
        return len(absent_days)

    def format_attendance_status(self, salesperson: str, year: int, month: int) -> str:
        """
        Format attendance status as a human-readable string.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            Formatted string like "Perfect attendance" or "No attendance on date: 2, 13, 18"
        """
        absent_days = self.get_absent_days(salesperson, year, month)

        if not absent_days:
            return "Perfect attendance ✓"

        # Format the days
        days_str = ", ".join(map(str, absent_days))
        count = len(absent_days)

        if count == 1:
            return f"No attendance on date: {days_str}"
        else:
            return f"No attendance on dates: {days_str} ({count} days)"


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def update_attendance_for_date(
    tracker: AttendanceTracker,
    salespeople_present: List[str],
    all_salespeople: List[str],
    date: datetime
) -> None:
    """
    Update attendance for a specific date.

    Args:
        tracker: AttendanceTracker instance
        salespeople_present: List of salespeople who used the app
        all_salespeople: Complete list of all salespeople
        date: Date to update attendance for
    """
    salespeople_present_set = set(salespeople_present)

    for salesperson in all_salespeople:
        # Ensure salesperson exists in tracking
        tracker.ensure_salesperson_exists(salesperson, date.year, date.month)

        if salesperson in salespeople_present_set:
            # Mark present (remove from absent list if mistakenly added)
            tracker.mark_present(salesperson, date)
        else:
            # Mark absent
            tracker.mark_absent(salesperson, date)

    logger.info(f"Updated attendance for {date.strftime('%Y-%m-%d')}")
    logger.info(f"  Present: {len(salespeople_present_set)} salespeople")
    logger.info(f"  Absent: {len(all_salespeople) - len(salespeople_present_set)} salespeople")


def generate_monthly_attendance_summary(
    tracker: AttendanceTracker,
    all_salespeople: List[str],
    year: int = None,
    month: int = None
) -> List[Dict[str, str]]:
    """
    Generate monthly attendance summary for email/report.

    Args:
        tracker: AttendanceTracker instance
        all_salespeople: Complete list of all salespeople
        year: Year (defaults to current year)
        month: Month (defaults to current month)

    Returns:
        List of dictionaries with salesperson name and attendance status
    """
    now = datetime.now()
    year = year or now.year
    month = month or now.month

    summary = []

    for salesperson in sorted(all_salespeople):
        attendance_status = tracker.format_attendance_status(salesperson, year, month)

        summary.append({
            "Salesperson": salesperson,
            "Attendance": attendance_status
        })

    return summary


# ============================================================================
# EXAMPLE USAGE
# ============================================================================

if __name__ == "__main__":
    # Example usage
    logging.basicConfig(level=logging.INFO)

    tracker = AttendanceTracker()

    # Example: Mark some people absent
    date = datetime(2026, 2, 13)
    tracker.mark_absent("John Doe", date)
    tracker.mark_absent("Jane Smith", datetime(2026, 2, 5))
    tracker.mark_absent("John Doe", datetime(2026, 2, 18))

    # Get monthly summary
    summary = generate_monthly_attendance_summary(
        tracker,
        ["John Doe", "Jane Smith", "Mike Johnson"],
        year=2026,
        month=2
    )

    for entry in summary:
        print(f"{entry['Salesperson']}: {entry['Attendance']}")
