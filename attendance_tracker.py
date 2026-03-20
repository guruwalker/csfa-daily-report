"""
Attendance Tracker Module
Manages daily attendance tracking in attendance.json
Tracks absent days for each salesperson and generates monthly summaries.

Updated: Supports suspended salespeople (SUS marker, frozen records) and
         new starters (blank cells before their start date).
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, List, Optional
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
        leave_checker=None,
        suspension_checker=None,
        new_starter_checker=None,
    ) -> str:
        """
        Format attendance status as a human-readable string.

        Priority order for status labels:
          1. Suspended (overrides everything — record is frozen)
          2. New starter (note their start date)
          3. On leave (currently)
          4. Normal absent-day listing or perfect attendance
        """
        # 1. Suspended?
        if suspension_checker is not None:
            suspension_date = suspension_checker.get_suspension_date(salesperson)
            if suspension_date:
                absent_days = self.get_absent_days(salesperson, year, month)
                if absent_days:
                    days_str = ", ".join(map(str, absent_days))
                    count = len(absent_days)
                    base = (
                        f"No attendance on date: {days_str}"
                        if count == 1
                        else f"No attendance on dates: {days_str} ({count} days)"
                    )
                    return f"{base} | Suspended from {suspension_date}"
                return f"Suspended from {suspension_date}"

        # 2. New starter — note their start date if within this month
        if new_starter_checker is not None:
            start_date_str = new_starter_checker.get_start_date(salesperson)
            if start_date_str:
                try:
                    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
                    if start_dt.year == year and start_dt.month == month:
                        # Only flag within the month they started
                        absent_days = self.get_absent_days(salesperson, year, month)
                        if absent_days:
                            days_str = ", ".join(map(str, absent_days))
                            count = len(absent_days)
                            base = (
                                f"No attendance on date: {days_str}"
                                if count == 1
                                else f"No attendance on dates: {days_str} ({count} days)"
                            )
                            return f"{base} | New starter from {start_date_str}"
                        return f"New starter from {start_date_str}"
                except ValueError:
                    pass

        absent_days = self.get_absent_days(salesperson, year, month)

        # 3. Currently on leave?
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
    leave_checker=None,
    suspension_checker=None,
    new_starter_checker=None,
) -> None:
    """
    Update attendance for a specific date.

    Skips:
      - Future dates
      - Salespeople on leave
      - Suspended salespeople (record is frozen — no further updates)
      - New starters on days before their start date
    """
    current_date = datetime.now().date()

    if date.date() > current_date:
        logger.info(f"⏭️  Skipping attendance update for {date.strftime('%Y-%m-%d')} — date is in the future")
        return

    salespeople_present_set = set(salespeople_present)

    for salesperson in all_salespeople:
        # Suspended — freeze record entirely
        if suspension_checker and suspension_checker.is_suspended(salesperson, date):
            logger.info(f"⏸️  {salesperson} is suspended — skipping attendance update")
            continue

        # Before start date — skip (don't mark absent, don't mark present)
        if new_starter_checker and new_starter_checker.is_before_start(salesperson, date):
            logger.info(f"🆕 {salesperson} hasn't started yet on {date.strftime('%Y-%m-%d')} — skipping")
            continue

        tracker.ensure_salesperson_exists(salesperson, date.year, date.month)

        # On leave — skip
        if leave_checker and leave_checker.is_on_leave(salesperson, date):
            logger.info(f"🏖️  {salesperson} on leave on {date.strftime('%Y-%m-%d')} — not marking absent")
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
    leave_checker=None,
    current_date: datetime = None,
    suspension_checker=None,
    new_starter_checker=None,
) -> pd.DataFrame:
    """
    Generate daily attendance grid for Excel (salesperson vs dates).

    Cell values:
      ✓   — present
      X   — absent
      L   — on leave
      SUS — suspended (all days from suspension date onwards)
      (blank) — future day, OR day before new starter's start date
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

        # Get suspension date (day number within this month) if applicable
        suspension_day_start = None
        if suspension_checker:
            susp_date_str = suspension_checker.get_suspension_date(salesperson)
            if susp_date_str:
                try:
                    susp_dt = datetime.strptime(susp_date_str, "%Y-%m-%d")
                    if susp_dt.year == year and susp_dt.month == month:
                        suspension_day_start = susp_dt.day
                    elif susp_dt.year < year or (susp_dt.year == year and susp_dt.month < month):
                        # Suspended before this month — all days are SUS
                        suspension_day_start = 1
                    # If suspension is in a future month, no SUS days this month
                except ValueError:
                    pass

        # Get new-starter start day within this month (if applicable)
        starter_day_start = None
        if new_starter_checker:
            start_date_str = new_starter_checker.get_start_date(salesperson)
            if start_date_str:
                try:
                    start_dt = datetime.strptime(start_date_str, "%Y-%m-%d")
                    if start_dt.year == year and start_dt.month == month:
                        starter_day_start = start_dt.day
                    elif start_dt.year > year or (start_dt.year == year and start_dt.month > month):
                        # Starts after this month — all days blank
                        starter_day_start = num_days + 1  # sentinel: all blank
                    # If started before this month — no blanks needed
                except ValueError:
                    pass

        present_count = 0
        total_past_weekdays = 0

        for day in weekdays:
            date_to_check = datetime(year, month, day)

            # Before new starter's start date → blank
            if starter_day_start is not None and day < starter_day_start:
                row[str(day)] = ""
                continue

            # Future day → blank
            if date_to_check.date() > current_date.date():
                row[str(day)] = ""
                continue

            total_past_weekdays += 1

            # Suspended from this day onwards → SUS
            if suspension_day_start is not None and day >= suspension_day_start:
                row[str(day)] = "SUS"
                continue

            if day in leave_days:
                row[str(day)] = "L"
            elif day in absent_days:
                row[str(day)] = "X"
            else:
                row[str(day)] = "✓"
                present_count += 1

        # Summary — exclude SUS days from denominator
        # Count days where the person was actually expected to work
        effective_total = sum(
            1 for day in weekdays
            if (
                # Not before start date
                (starter_day_start is None or day >= starter_day_start) and
                # Not future
                datetime(year, month, day).date() <= current_date.date() and
                # Not suspended
                (suspension_day_start is None or day < suspension_day_start)
            )
        )

        susp_date_str = suspension_checker.get_suspension_date(salesperson) if suspension_checker else None

        if effective_total == 0:
            if susp_date_str:
                row["Summary"] = f"Suspended from {susp_date_str}"
            elif starter_day_start is not None and starter_day_start > max(weekdays, default=0):
                row["Summary"] = "Not started yet"
            else:
                row["Summary"] = "No days yet"
        elif susp_date_str:
            susp_dt = datetime.strptime(susp_date_str, "%Y-%m-%d") if susp_date_str else None
            if susp_dt and (susp_dt.year < year or (susp_dt.year == year and susp_dt.month <= month)):
                row["Summary"] = (
                    f"{present_count}/{effective_total} | Suspended from {susp_date_str}"
                )
            else:
                row["Summary"] = f"{present_count}/{effective_total}"
        elif present_count == effective_total:
            row["Summary"] = f"{present_count}/{effective_total} - Perfect attendance"
        else:
            row["Summary"] = f"{present_count}/{effective_total}"

        grid_data.append(row)

    return pd.DataFrame(grid_data)


def generate_monthly_attendance_summary(
    tracker: AttendanceTracker,
    all_salespeople: List[str],
    year: int = None,
    month: int = None,
    leave_checker=None,
    suspension_checker=None,
    new_starter_checker=None,
) -> List[Dict[str, str]]:
    """Generate monthly attendance summary for email/report."""
    now = datetime.now()
    year = year or now.year
    month = month or now.month

    return [
        {
            "Salesperson": salesperson,
            "Attendance": tracker.format_attendance_status(
                salesperson,
                year,
                month,
                leave_checker=leave_checker,
                suspension_checker=suspension_checker,
                new_starter_checker=new_starter_checker,
            )
        }
        for salesperson in sorted(all_salespeople)
    ]
