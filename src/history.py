"""
History management for storing and retrieving previous month data.
"""

import json
from pathlib import Path
from datetime import datetime


class HistoryManager:
    """Manages saving and loading historical report data."""

    def __init__(self, history_dir="output/history"):
        """
        Initialize the history manager.

        Args:
            history_dir: Directory to store history files
        """
        self.history_dir = Path(history_dir)
        self.history_dir.mkdir(parents=True, exist_ok=True)

    def save_data(self, data, month_str=None):
        """
        Save data to history.

        Args:
            data: Dictionary of extracted data
            month_str: Optional month string (YYYY-MM). If not provided, uses current date.

        Returns:
            Path to the saved history file
        """
        if month_str is None:
            # Use current date
            month_str = datetime.now().strftime("%Y-%m")

        # Filter out table data (only save scalar values)
        scalar_data = {k: v for k, v in data.items() if not isinstance(v, list)}

        filename = f"{month_str}.json"
        filepath = self.history_dir / filename

        with open(filepath, 'w') as f:
            json.dump(scalar_data, f, indent=2, default=str)

        return filepath

    def load_previous_month(self, current_month_str=None):
        """
        Load previous month data from history.

        Args:
            current_month_str: Optional current month string (YYYY-MM).
                             If not provided, uses current date.

        Returns:
            Dictionary of previous month data, or None if not found
        """
        if current_month_str is None:
            current_month_str = datetime.now().strftime("%Y-%m")

        # Parse current month
        try:
            current_date = datetime.strptime(current_month_str, "%Y-%m")
        except ValueError:
            return None

        # Calculate previous month
        if current_date.month == 1:
            prev_month = datetime(current_date.year - 1, 12, 1)
        else:
            prev_month = datetime(current_date.year, current_date.month - 1, 1)

        prev_month_str = prev_month.strftime("%Y-%m")
        filepath = self.history_dir / f"{prev_month_str}.json"

        if not filepath.exists():
            return None

        with open(filepath, 'r') as f:
            return json.load(f)

    def get_month_from_data(self, data):
        """
        Extract month string from data (expects a 'month' field).

        Args:
            data: Dictionary containing a 'month' field

        Returns:
            Month string in YYYY-MM format, or None if can't parse
        """
        month_field = data.get('month')
        if not month_field:
            return None

        try:
            # Parse "November 2025" format
            date = datetime.strptime(month_field, "%B %Y")
            return date.strftime("%Y-%m")
        except:
            return None
