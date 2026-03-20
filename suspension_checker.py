"""
Suspension Checker Module
Checks whether a salesperson is suspended on a given date,
based on suspended_salespeople.json.

JSON structure:
{
    "suspended": {
        "JOHN DOE": "2026-03-15"
    }
}

A salesperson is considered suspended on any date >= their suspension date.
Suspension is permanent (no reinstatement logic — once suspended, record is frozen).
"""

import json
import os
import logging
from datetime import datetime
from typing import Dict, Optional

logger = logging.getLogger(__name__)


class SuspensionChecker:
    """Manages suspension checking from suspended_salespeople.json."""

    def __init__(self, json_file: str = "suspended_salespeople.json"):
        self.json_file = json_file
        self.suspended = self._load()

    def _load(self) -> Dict[str, str]:
        """
        Load suspension data from JSON file.

        Accepts two formats:
          Dict (correct):  {"suspended": {"NAME": "YYYY-MM-DD"}}
          List (legacy):   {"suspended": ["NAME", ...]}

        If a list is found, today's date is used as the suspension date and the
        file is rewritten in the correct dict format automatically.
        """
        if not os.path.exists(self.json_file):
            logger.warning(f"Suspension file not found: {self.json_file}")
            self._create_empty_file()
            return {}

        try:
            with open(self.json_file, "r") as f:
                data = json.load(f)

            raw = data.get("suspended", {})

            if isinstance(raw, list):
                today = datetime.now().strftime("%Y-%m-%d")
                suspended = {name: today for name in raw if isinstance(name, str)}
                logger.warning(
                    f"suspended_salespeople.json contained a list — converted to dict "
                    f"with today ({today}) as suspension date. File has been rewritten "
                    f'in the correct format: {{"NAME": "YYYY-MM-DD"}}.'
                )
                self._rewrite(suspended)
            else:
                suspended = raw

            logger.info(f"Loaded {len(suspended)} suspended salesperson(s) from {self.json_file}")
            return suspended

        except json.JSONDecodeError as e:
            logger.error(f"Error reading {self.json_file}: {e}")
            return {}
        except Exception as e:
            logger.error(f"Unexpected error loading suspensions: {e}")
            return {}

    def _rewrite(self, suspended: Dict[str, str]) -> None:
        """Overwrite the file with the correct dict format."""
        try:
            with open(self.json_file, "w") as f:
                json.dump({"suspended": suspended}, f, indent=2)
            logger.info(f"Rewrote {self.json_file} in dict format")
        except Exception as e:
            logger.error(f"Failed to rewrite {self.json_file}: {e}")

    def _create_empty_file(self) -> None:
        try:
            with open(self.json_file, "w") as f:
                json.dump({"suspended": {}}, f, indent=2)
            logger.info(f"Created {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to create {self.json_file}: {e}")

    def _save(self) -> None:
        try:
            with open(self.json_file, "w") as f:
                json.dump({"suspended": self.suspended}, f, indent=2)
            logger.info(f"Saved suspensions to {self.json_file}")
        except Exception as e:
            logger.error(f"Failed to save suspensions: {e}")

    def is_suspended(self, salesperson: str, date: datetime) -> bool:
        """
        Return True if the salesperson is suspended on the given date.
        Suspended means: date >= suspension_start_date.
        """
        if salesperson not in self.suspended:
            return False

        try:
            suspension_date = datetime.strptime(self.suspended[salesperson], "%Y-%m-%d")
            result = date.date() >= suspension_date.date()
            if result:
                logger.debug(f"{salesperson} is suspended as of {self.suspended[salesperson]}")
            return result
        except ValueError:
            logger.warning(f"Invalid suspension date for {salesperson}: {self.suspended[salesperson]}")
            return False

    def get_suspension_date(self, salesperson: str) -> Optional[str]:
        """Return the suspension date string (YYYY-MM-DD) or None."""
        return self.suspended.get(salesperson)

    def add_suspension(self, salesperson: str, date: datetime) -> None:
        """Record a suspension starting from the given date."""
        date_str = date.strftime("%Y-%m-%d")
        self.suspended[salesperson] = date_str
        self._save()
        logger.info(f"Suspended {salesperson} from {date_str}")

    def remove_suspension(self, salesperson: str) -> None:
        """Remove a suspension (reinstate salesperson)."""
        if salesperson in self.suspended:
            del self.suspended[salesperson]
            self._save()
            logger.info(f"Removed suspension for {salesperson}")

    def get_all_suspended(self) -> Dict[str, str]:
        """Return all suspended salespeople and their suspension dates."""
        return self.suspended.copy()


def is_suspended(salesperson: str, date: datetime, json_file: str = "suspended_salespeople.json") -> bool:
    """Quick check: is this salesperson suspended on this date?"""
    checker = SuspensionChecker(json_file)
    return checker.is_suspended(salesperson, date)
