"""
Holiday Checker Module
Checks if a given date is a holiday based on holidays.json
"""

import json
import os
import logging
from datetime import datetime
from typing import List

logger = logging.getLogger(__name__)


class HolidayChecker:
    """Manages holiday checking from holidays.json file."""

    def __init__(self, json_file: str = "holidays.json"):
        """
        Initialize holiday checker.

        Args:
            json_file: Path to holidays JSON file
        """
        self.json_file = json_file
        self.holidays = self._load_holidays()

    def _load_holidays(self) -> List[str]:
        """Load holidays from JSON file."""
        if not os.path.exists(self.json_file):
            logger.warning(f"Holidays file not found: {self.json_file}")
            logger.info("Creating empty holidays.json file")
            self._create_empty_file()
            return []

        try:
            with open(self.json_file, 'r') as f:
                data = json.load(f)
            holidays = data.get('holidays', [])
            logger.info(f"Loaded {len(holidays)} holidays from {self.json_file}")
            return holidays
        except json.JSONDecodeError as e:
            logger.error(f"Error reading {self.json_file}: {e}")
            logger.warning("Using empty holidays list")
            return []
        except Exception as e:
            logger.error(f"Unexpected error loading holidays: {e}")
            return []

    def _create_empty_file(self) -> None:
        """Create an empty holidays.json file."""
        try:
            with open(self.json_file, 'w') as f:
                json.dump({"holidays": []}, f, indent=2)
            logger.info(f"Created {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to create {self.json_file}: {e}")

    def is_holiday(self, date: datetime) -> bool:
        """
        Check if a given date is a holiday.

        Args:
            date: datetime object to check

        Returns:
            True if date is a holiday, False otherwise
        """
        date_str = date.strftime("%Y-%m-%d")
        is_hol = date_str in self.holidays

        if is_hol:
            logger.info(f"📅 {date_str} is a registered holiday")

        return is_hol

    def get_all_holidays(self) -> List[str]:
        """Get list of all registered holidays."""
        return self.holidays.copy()

    def add_holiday(self, date: datetime) -> None:
        """
        Add a holiday to the list and save to file.

        Args:
            date: datetime object for the holiday
        """
        date_str = date.strftime("%Y-%m-%d")

        if date_str not in self.holidays:
            self.holidays.append(date_str)
            self.holidays.sort()
            self._save_holidays()
            logger.info(f"Added holiday: {date_str}")
        else:
            logger.info(f"Holiday already exists: {date_str}")

    def remove_holiday(self, date: datetime) -> None:
        """
        Remove a holiday from the list and save to file.

        Args:
            date: datetime object for the holiday to remove
        """
        date_str = date.strftime("%Y-%m-%d")

        if date_str in self.holidays:
            self.holidays.remove(date_str)
            self._save_holidays()
            logger.info(f"Removed holiday: {date_str}")
        else:
            logger.info(f"Holiday not found: {date_str}")

    def _save_holidays(self) -> None:
        """Save holidays list to JSON file."""
        try:
            with open(self.json_file, 'w') as f:
                json.dump({"holidays": self.holidays}, f, indent=2)
            logger.info(f"Saved holidays to {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to save holidays: {e}")


def is_holiday(date: datetime, json_file: str = "holidays.json") -> bool:
    """
    Quick check if a date is a holiday.

    Args:
        date: datetime object to check
        json_file: Path to holidays JSON file

    Returns:
        True if date is a holiday, False otherwise
    """
    checker = HolidayChecker(json_file)
    return checker.is_holiday(date)
