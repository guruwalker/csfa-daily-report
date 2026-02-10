"""
Leave Checker Module
Checks if a salesperson is on leave on a given date based on leave_days.json
"""

import json
import os
import logging
from datetime import datetime
from typing import List, Dict

logger = logging.getLogger(__name__)


class LeaveChecker:
    """Manages leave checking from leave_days.json file."""

    def __init__(self, json_file: str = "leave_days.json"):
        """Initialize leave checker."""
        self.json_file = json_file
        self.leave_days = self._load_leave_days()

    def _load_leave_days(self) -> Dict[str, List[str]]:
        """Load leave days from JSON file."""
        if not os.path.exists(self.json_file):
            logger.warning(f"Leave days file not found: {self.json_file}")
            logger.info("Creating empty leave_days.json file")
            self._create_empty_file()
            return {}

        try:
            with open(self.json_file, 'r') as f:
                data = json.load(f)
            leave_days = data.get('leave_days', {})
            total_leaves = sum(len(dates) for dates in leave_days.values())
            logger.info(f"Loaded leave data for {len(leave_days)} people ({total_leaves} total leave days)")
            return leave_days
        except json.JSONDecodeError as e:
            logger.error(f"Error reading {self.json_file}: {e}")
            logger.warning("Using empty leave days")
            return {}
        except Exception as e:
            logger.error(f"Unexpected error loading leave days: {e}")
            return {}

    def _create_empty_file(self) -> None:
        """Create an empty leave_days.json file."""
        try:
            with open(self.json_file, 'w') as f:
                json.dump({"leave_days": {}}, f, indent=2)
            logger.info(f"Created {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to create {self.json_file}: {e}")

    def is_on_leave(self, salesperson: str, date: datetime) -> bool:
        """
        Check if a salesperson is on leave on a given date.

        Args:
            salesperson: Name of the salesperson
            date: datetime object to check

        Returns:
            True if on leave, False otherwise
        """
        date_str = date.strftime("%Y-%m-%d")

        if salesperson in self.leave_days:
            is_leave = date_str in self.leave_days[salesperson]
            if is_leave:
                logger.info(f"📅 {salesperson} is on leave on {date_str}")
            return is_leave

        return False

    def get_leave_days(self, salesperson: str, year: int, month: int) -> List[int]:
        """
        Get list of leave day numbers for a salesperson in a specific month.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            List of day numbers when salesperson was on leave
        """
        if salesperson not in self.leave_days:
            return []

        leave_days_in_month = []
        for date_str in self.leave_days[salesperson]:
            try:
                leave_date = datetime.strptime(date_str, "%Y-%m-%d")
                if leave_date.year == year and leave_date.month == month:
                    leave_days_in_month.append(leave_date.day)
            except ValueError:
                logger.warning(f"Invalid date format in leave_days: {date_str}")

        return sorted(leave_days_in_month)

    def get_total_leave_days(self, salesperson: str, year: int, month: int) -> int:
        """Get total number of leave days for a salesperson in a month."""
        return len(self.get_leave_days(salesperson, year, month))

    def format_leave_info(self, salesperson: str, year: int, month: int) -> str:
        """
        Format leave information as a human-readable string.

        Args:
            salesperson: Name of the salesperson
            year: Year (e.g., 2026)
            month: Month (1-12)

        Returns:
            Formatted string like "On leave: 10, 11" or empty string if no leave
        """
        leave_days = self.get_leave_days(salesperson, year, month)

        if not leave_days:
            return ""

        days_str = ", ".join(map(str, leave_days))
        count = len(leave_days)

        if count == 1:
            return f"On leave: {days_str}"
        else:
            return f"On leave: {days_str} ({count} days)"

    def add_leave(self, salesperson: str, date: datetime) -> None:
        """Add a leave day for a salesperson."""
        date_str = date.strftime("%Y-%m-%d")

        if salesperson not in self.leave_days:
            self.leave_days[salesperson] = []

        if date_str not in self.leave_days[salesperson]:
            self.leave_days[salesperson].append(date_str)
            self.leave_days[salesperson].sort()
            self._save_leave_days()
            logger.info(f"Added leave day for {salesperson}: {date_str}")

    def _save_leave_days(self) -> None:
        """Save leave days to JSON file."""
        try:
            with open(self.json_file, 'w') as f:
                json.dump({"leave_days": self.leave_days}, f, indent=2)
            logger.info(f"Saved leave days to {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to save leave days: {e}")


def is_on_leave(salesperson: str, date: datetime, json_file: str = "leave_days.json") -> bool:
    """
    Quick check if a salesperson is on leave.

    Args:
        salesperson: Name of the salesperson
        date: datetime object to check
        json_file: Path to leave days JSON file

    Returns:
        True if on leave, False otherwise
    """
    checker = LeaveChecker(json_file)
    return checker.is_on_leave(salesperson, date)
