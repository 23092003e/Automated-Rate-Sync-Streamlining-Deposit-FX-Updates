# -*- coding: utf-8 -*-
"""
ACB (Asia Commercial Bank) Email Parser
Refactored from original ACB.py to work with multi-bank system
"""

import re
from datetime import datetime, date
import pandas as pd
from typing import Tuple

from ..base import BaseBankProcessor


class ACBProcessor(BaseBankProcessor):
    """ACB-specific email processor"""
    
    def __init__(self):
        super().__init__("ACB")
    
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Parse ACB email format"""
        df_forward = self._parse_forwards(email_text)
        df_spot = self._parse_spot(email_text)  
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse Spot Exchange Rates table"""
        out_cols = self.get_standard_columns()['spot']
        
        parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        forward_parts = re.split(r"(?i)forward\s+exchange\s+rates", spot_section, maxsplit=1)
        spot_only = forward_parts[0]
        
        all_nums = re.findall(r"\b\d{5}\b", spot_only)
        all_nums = list(dict.fromkeys(all_nums))
        
        low = None
        high = None
        bid_close = None
        ask_close = None
        
        if len(all_nums) >= 2:
            bid_close = self._to_int(all_nums[0])  # 26370
            ask_close = self._to_int(all_nums[1])  # 26380
        elif len(all_nums) >= 1:
            bid_close = self._to_int(all_nums[0])
            ask_close = self._to_int(all_nums[0])
        
        quoting_date = self._first_date(email_text) or ""
        
        rows = []
        for i, (side, closing) in enumerate([("Bid", bid_close), ("Ask", ask_close)], start=1):
            rows.append({
                "No.": i,
                "Bid/Ask": side,
                "Bank": self.bank_name,
                "Quoting date": quoting_date,
                "Lowest rate of preceeding week": low,
                "Highest rate of preceeding week": high,
                "Closing rate of Friday (last week)": closing,
            })
        return pd.DataFrame(rows, columns=out_cols)
    
    def _parse_forwards(self, email_text: str) -> pd.DataFrame:
        """Parse Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        root = re.split(r"(?i)forward\s+exchange\s+rates", email_text, maxsplit=1)
        if len(root) < 2:
            return pd.DataFrame(columns=out_cols)
        tail = root[1]
        
        bid_parts = re.split(r"(?i)\bBid\s*Price\s*:", tail, maxsplit=1)
        if len(bid_parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        after_bid = bid_parts[1]
        ask_parts = re.split(r"(?i)\bAsk\s*Price\s*:", after_bid, maxsplit=1)
        
        bid_text = ask_parts[0]
        ask_text = ask_parts[1] if len(ask_parts) > 1 else ""
        
        rows = []
        rows += self._parse_forward_side(bid_text, "Bid")
        rows += self._parse_forward_side(ask_text, "Ask")
        
        if not rows:
            return pd.DataFrame(columns=out_cols)
        
        df = pd.DataFrame(rows)
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df
    
    def _parse_forward_side(self, text: str, side: str) -> list:
        """Parse a single side (Bid or Ask)"""
        ROW_RE = re.compile(
            rf"\s*(?P<trd>{self.DATE_DMY})\s+"
            rf"(?P<val>{self.DATE_DMY})\s+"
            rf"(?:(?P<spot>\d{{5}})\s+)?"
            rf"(?P<termnum>\d+)\s*(?P<termunit>[DMWY])?\s*\(\s*\)\s+"
            rf"(?P<gap>-?\d+(?:\.\d+)?)%\s+"
            rf"(?P<fwd>[\d,]+)\s*",
            flags=re.IGNORECASE
        )
        
        rows = []
        current_spot = None
        
        for m in ROW_RE.finditer(text):
            trd_s = m.group("trd")
            val_s = m.group("val")
            spot_s = m.group("spot")
            termnum = int(m.group("termnum"))
            termunit = (m.group("termunit") or "M").upper()
            gap_pct = float(m.group("gap"))
            fwd = self._to_int(m.group("fwd"))
            
            if spot_s:
                current_spot = self._to_int(spot_s)
            if current_spot is None:
                prev = re.findall(r"\b\d{5}\b", text[:m.start()])
                current_spot = self._to_int(prev[-1]) if prev else None
            
            trd = self._to_date(trd_s)
            val = self._to_date(val_s)
            
            if val < trd:
                trd, val = val, trd
            
            term_days = self._days(trd, val)
            term_lookup = round(self._yearfrac_30360_us(trd, val) * 12)
            
            rows.append({
                "Bid/Ask": side,
                "Bank": self.bank_name,
                "Quoting date": trd,
                "Trading date": trd,
                "Value date": val,
                "Spot Exchange rate": current_spot,
                "Gap(%)": gap_pct,
                "Forward Exchange rate": fwd,
                "Term (days)": term_days,
                "% forward (cal)": None,  # Excel formula
                "Diff.": None,  # Excel formula
                "Term (lookup)": term_lookup
            })
        return rows
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub"""
        out_cols = self.get_standard_columns()['central']
        qd = self._first_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)