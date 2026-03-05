"""
Attendance Tracker Module
Manages daily attendance tracking in attendance.json
Tracks absent days for each salesperson and generates monthly summaries.
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, List
import pandas as pd

logger = logging.getLogger(__name__)


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
            }
        }
    }
    """

    def __init__(self, json_file: str = "attendance.json"):
        """Initialize attendance tracker."""
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
        """Mark a salesperson as absent on a specific date."""
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        if year not in self.data:
            self.data[year] = {}
        if month not in self.data[year]:
            self.data[year][month] = {}
        if salesperson not in self.data[year][month]:
            self.data[year][month][salesperson] = []

        if day not in self.data[year][month][salesperson]:
            self.data[year][month][salesperson].append(day)
            self.data[year][month][salesperson].sort()
            logger.info(f"Marked {salesperson} absent on {date.strftime('%Y-%m-%d')}")
            self._save_data()

    def mark_present(self, salesperson: str, date: datetime) -> None:
        """Mark a salesperson as present (remove from absent list if exists)."""
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        if (year in self.data and
                month in self.data[year] and
                salesperson in self.data[year][month] and
                day in self.data[year][month][salesperson]):
            self.data[year][month][salesperson].remove(day)
            logger.info(f"Marked {salesperson} present on {date.strftime('%Y-%m-%d')}")
            self._save_data()

    def get_absent_days(self, salesperson: str, year: int, month: int) -> List[int]:
        """Get list of absent days for a salesperson in a specific month."""
        year_str = str(year)
        month_str = f"{month:02d}"

        if (year_str in self.data and
                month_str in self.data[year_str] and
                salesperson in self.data[year_str][month_str]):
            return self.data[year_str][month_str][salesperson].copy()

        return []

    def get_monthly_summary(self, year: int, month: int) -> Dict[str, List[int]]:
        """Get attendance summary for all salespeople in a specific month."""
        year_str = str(year)
        month_str = f"{month:02d}"

        if year_str in self.data and month_str in self.data[year_str]:
            return self.data[year_str][month_str].copy()

        return {}

    def ensure_salesperson_exists(self, salesperson: str, year: int, month: int) -> None:
        """Ensure a salesperson exists in the tracking for a given month."""
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
        """Check if a salesperson is marked as absent on a specific date."""
        year = str(date.year)
        month = f"{date.month:02d}"
        day = date.day

        return (year in self.data and
                month in self.data[year] and
                salesperson in self.data[year][month] and
                day in self.data[year][month][salesperson])

    def get_total_absent_days(self, salesperson: str, year: int, month: int) -> int:
        """Get total number of absent days for a salesperson in a month."""
        return len(self.get_absent_days(salesperson, year, month))

    def format_attendance_status(
        self,
        salesperson: str,
        year: int,
        month: int,
        leave_checker: 'LeaveChecker' = None
    ) -> str:
        """
        Format attendance status as a human-readable string.

        "Not active - On leave" is only shown if the person is CURRENTLY on leave
        (i.e. their leave extends up to or beyond today). Past leave within the
        month does not affect the status label — only the absent-day list matters
        for those days.
        """
        absent_days = self.get_absent_days(salesperson, year, month)

        # Only flag as currently on leave if today is still a leave day
        currently_on_leave = (
            leave_checker is not None and
            leave_checker.is_on_leave(salesperson, datetime.now())
        )

        if not absent_days:
            if currently_on_leave:
                return "Not active - On leave"
            return "Perfect attendance"

        days_str = ", ".join(map(str, absent_days))
        count = len(absent_days)

        if count == 1:
            base_msg = f"No attendance on date: {days_str}"
        else:
            base_msg = f"No attendance on dates: {days_str} ({count} days)"

        if currently_on_leave:
            base_msg += " | Not active - On leave"

        return base_msg


def update_attendance_for_date(
    tracker: AttendanceTracker,
    salespeople_present: List[str],
    all_salespeople: List[str],
    date: datetime,
    leave_checker: 'LeaveChecker' = None
) -> None:
    """
    Update attendance for a specific date, excluding people on leave.

    Only skips FUTURE dates (strictly after today).
    Today is always processed since reports run at end of business day.
    """
    current_date = datetime.now().date()

    if date.date() > current_date:
        logger.info(f"⏭️  Skipping attendance update for {date.strftime('%Y-%m-%d')} - date is in the future")
        return

    salespeople_present_set = set(salespeople_present)

    for salesperson in all_salespeople:
        tracker.ensure_salesperson_exists(salesperson, date.year, date.month)

        if leave_checker and leave_checker.is_on_leave(salesperson, date):
            logger.info(f"{salesperson} on leave on {date.strftime('%Y-%m-%d')} - not marking absent")
            continue

        if salesperson in salespeople_present_set:
            tracker.mark_present(salesperson, date)
        else:
            tracker.mark_absent(salesperson, date)

    logger.info(f"✅ Updated attendance for {date.strftime('%Y-%m-%d')}")


def generate_monthly_attendance_grid(
    tracker: AttendanceTracker,
    all_salespeople: List[str],
    year: int,
    month: int,
    leave_checker: 'LeaveChecker' = None,
    current_date: datetime = None
) -> pd.DataFrame:
    """
    Generate daily attendance grid for Excel (salesperson vs dates).
    Only shows ✓/X/L for days that have already passed or are today.
    Future days are left blank.
    """
    import calendar

    if current_date is None:
        current_date = datetime.now()

    num_days = calendar.monthrange(year, month)[1]

    weekdays = [
        day for day in range(1, num_days + 1)
        if datetime(year, month, day).weekday() < 5
    ]

    grid_data = []

    for salesperson in sorted(all_salespeople):
        row = {"Salesperson": salesperson}

        absent_days = tracker.get_absent_days(salesperson, year, month)
        leave_days = leave_checker.get_leave_days(salesperson, year, month) if leave_checker else []

        present_count = 0
        total_past_weekdays = 0

        for day in weekdays:
            date_to_check = datetime(year, month, day)

            if date_to_check.date() > current_date.date():
                row[str(day)] = ""
            else:
                total_past_weekdays += 1

                if day in leave_days:
                    row[str(day)] = "L"
                elif day in absent_days:
                    row[str(day)] = "X"
                else:
                    row[str(day)] = "✓"
                    present_count += 1

        if total_past_weekdays == 0:
            row["Summary"] = "No days yet"
        elif present_count == total_past_weekdays:
            row["Summary"] = f"{present_count}/{total_past_weekdays} - Perfect attendance"
        else:
            row["Summary"] = f"{present_count}/{total_past_weekdays}"

        grid_data.append(row)

    return pd.DataFrame(grid_data)


def generate_monthly_attendance_summary(
    tracker: AttendanceTracker,
    all_salespeople: List[str],
    year: int = None,
    month: int = None,
    leave_checker: 'LeaveChecker' = None
) -> List[Dict[str, str]]:
    """Generate monthly attendance summary for email/report."""
    now = datetime.now()
    year = year or now.year
    month = month or now.month

    return [
        {
            "Salesperson": salesperson,
            "Attendance": tracker.format_attendance_status(salesperson, year, month, leave_checker)
        }
        for salesperson in sorted(all_salespeople)
    ]
