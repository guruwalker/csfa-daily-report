"""
New Starter Checker Module
Tracks the start date of new salespeople, based on new_starters.json.

JSON structure:
{
    "new_starters": {
        "JOHN DOE": "2026-03-20"
    }
}

Days before a salesperson's start date are left blank in the attendance grid
and excluded from attendance calculations entirely.
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, Optional

logger = logging.getLogger(__name__)


class NewStarterChecker:
    """Manages new starter start dates from new_starters.json."""

    def __init__(self, json_file: str = "new_starters.json"):
        self.json_file = json_file
        self.new_starters = self._load()

    def _load(self) -> Dict[str, str]:
        """
        Load new starter data from JSON file.

        Accepts two formats:
          Dict (correct):  {"new_starters": {"NAME": "YYYY-MM-DD"}}
          List (legacy):   {"new_starters": ["NAME", ...]}

        If a list is found, today's date is used as the start date and the file
        is rewritten in the correct dict format automatically.
        """
        if not os.path.exists(self.json_file):
            logger.warning(f"New starters file not found: {self.json_file}")
            self._create_empty_file()
            return {}

        try:
            with open(self.json_file, "r") as f:
                data = json.load(f)

            raw = data.get("new_starters", {})

            if isinstance(raw, list):
                today = datetime.now().strftime("%Y-%m-%d")
                new_starters = {name: today for name in raw if isinstance(name, str)}
                logger.warning(
                    f"new_starters.json contained a list — converted to dict with "
                    f"today ({today}) as start date. File has been rewritten in the "
                    f'correct format: {{"NAME": "YYYY-MM-DD"}}.'
                )
                self._rewrite(new_starters)
            else:
                new_starters = raw

            logger.info(f"Loaded {len(new_starters)} new starter(s) from {self.json_file}")
            return new_starters

        except json.JSONDecodeError as e:
            logger.error(f"Error reading {self.json_file}: {e}")
            return {}
        except Exception as e:
            logger.error(f"Unexpected error loading new starters: {e}")
            return {}

    def _rewrite(self, new_starters: Dict[str, str]) -> None:
        """Overwrite the file with the correct dict format."""
        try:
            with open(self.json_file, "w") as f:
                json.dump({"new_starters": new_starters}, f, indent=2)
            logger.info(f"Rewrote {self.json_file} in dict format")
        except Exception as e:
            logger.error(f"Failed to rewrite {self.json_file}: {e}")

    def _create_empty_file(self) -> None:
        try:
            with open(self.json_file, "w") as f:
                json.dump({"new_starters": {}}, f, indent=2)
            logger.info(f"Created {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to create {self.json_file}: {e}")

    def _save(self) -> None:
        try:
            with open(self.json_file, "w") as f:
                json.dump({"new_starters": self.new_starters}, f, indent=2)
            logger.info(f"Saved new starters to {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to save new starters: {e}")

    def is_before_start(self, salesperson: str, date: datetime) -> bool:
        """
        Return True if the given date is strictly before the salesperson's start date.
        Salespeople NOT in the file are assumed to have always been active (return False).
        """
        if salesperson not in self.new_starters:
            return False

        try:
            start_date = datetime.strptime(self.new_starters[salesperson], "%Y-%m-%d")
            return date.date() < start_date.date()
        except ValueError:
            logger.warning(f"Invalid start date for {salesperson}: {self.new_starters[salesperson]}")
            return False

    def get_start_date(self, salesperson: str) -> Optional[str]:
        """Return the start date string (YYYY-MM-DD) or None if not a new starter."""
        return self.new_starters.get(salesperson)

    def get_start_date_as_datetime(self, salesperson: str) -> Optional[datetime]:
        """Return the start date as a datetime object, or None."""
        date_str = self.new_starters.get(salesperson)
        if not date_str:
            return None
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            return None

    def add_new_starter(self, salesperson: str, start_date: datetime) -> None:
        """Register a new starter with their start date."""
        date_str = start_date.strftime("%Y-%m-%d")
        self.new_starters[salesperson] = date_str
        self._save()
        logger.info(f"Registered new starter {salesperson} from {date_str}")

    def remove_new_starter(self, salesperson: str) -> None:
        """Remove a new starter entry (e.g. they are now fully established)."""
        if salesperson in self.new_starters:
            del self.new_starters[salesperson]
            self._save()
            logger.info(f"Removed new starter entry for {salesperson}")

    def get_all_new_starters(self) -> Dict[str, str]:
        """Return all new starters and their start dates."""
        return self.new_starters.copy()


def is_before_start(salesperson: str, date: datetime, json_file: str = "new_starters.json") -> bool:
    """Quick check: is this date before the salesperson's start date?"""
    checker = NewStarterChecker(json_file)
    return checker.is_before_start(salesperson, date)
