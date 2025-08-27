# -*- coding: utf-8 -*-
"""
Base class for all bank processors
Defines common interface and shared utilities
"""

from abc import ABC, abstractmethod
from datetime import datetime, date
import pandas as pd
import re
from typing import Tuple, Optional


class BaseBankProcessor(ABC):
    """Abstract base class for bank-specific email processors"""
    
    def __init__(self, bank_name: str):
        self.bank_name = bank_name
        self.DATE_DMY = r"(?:0[1-9]|[12]\d|3[01])/(?:0[1-9]|1[0-2])/(?:19|20)\d\d"
    
    @abstractmethod
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """
        Parse bank-specific email format
        Returns: (forward_df, spot_df, central_df)
        """
        pass
    
    def _to_int(self, s) -> Optional[int]:
        """Convert string to int, handling commas"""
        return int(str(s).replace(',', '')) if s is not None and str(s).strip() else None
    
    def _to_date(self, dmy: str) -> date:
        """Convert DD/MM/YYYY string to date"""
        return datetime.strptime(dmy, "%d/%m/%Y").date()
    
    def _days(self, d1: date, d2: date) -> int:
        """Calculate days between two dates"""
        return (d2 - d1).days
    
    def _yearfrac_30360_us(self, d1: date, d2: date) -> float:
        """Excel YEARFRAC basis=0 approximation for Term(lookup)"""
        d1d = 30 if d1.day == 31 else d1.day
        d2d = 30 if (d2.day == 31 and d1d in (30, 31)) else d2.day
        months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
        days = d2d - d1d
        return (months * 30 + days) / 360.0
    
    def _first_date(self, text: str) -> Optional[str]:
        """Extract first date from text"""
        m = re.search(self.DATE_DMY, text)
        return m.group(0) if m else None
    
    def get_standard_columns(self) -> dict:
        """Return standard column definitions"""
        return {
            'forward': [
                "No.", "Bid/Ask", "Bank", "Quoting date", "Trading date", "Value date",
                "Spot Exchange rate", "Gap(%)", "Forward Exchange rate",
                "Term (days)", "% forward (cal)", "Diff.", "Term (lookup)"
            ],
            'spot': [
                "No.", "Bid/Ask", "Bank", "Quoting date",
                "Lowest rate of preceeding week",
                "Highest rate of preceeding week", 
                "Closing rate of Friday (last week)"
            ],
            'central': [
                "No.", "Quoting date", "Central Bank Rate"
            ]
        }
    
    def create_empty_dataframes(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Create empty DataFrames with standard columns"""
        cols = self.get_standard_columns()
        return (
            pd.DataFrame(columns=cols['forward']),
            pd.DataFrame(columns=cols['spot']),
            pd.DataFrame(columns=cols['central'])
        )