# -*- coding: utf-8 -*-
"""
TCB (Techcombank) Email Parser - robust version
- Spot: parse by labels (Lowest/Highest/Closing), extract Bid & Ask separately
- Forward: flexible row regex, accepts numbers with/without commas
"""

import re
from datetime import date
import pandas as pd
from typing import Tuple, List, Dict

from ..base import BaseBankProcessor


class TCBProcessor(BaseBankProcessor):
    """TCB-specific email processor (fixed/robust)"""

    def __init__(self):
        super().__init__("TCB")

    # -------------------------------
    # Public API
    # -------------------------------
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        df_forward = self._parse_forwards(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central

    # -------------------------------
    # Spot parsing (label-anchored)
    # -------------------------------
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """
        Parse Spot Exchange Rates as a 3-row table with Bid/Ask per row:
          - Lowest rate of the preceding week
          - Highest rate of the preceding week
          - Closing rate of Friday (last week)
        We isolate the spot section and read numbers adjacent to each label.
        """
        out_cols = self.get_standard_columns()['spot']

        # 1) Isolate Spot section, exclude Forward section
        parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        spot_sec = parts[1]
        spot_sec = re.split(r"(?i)forward\s+exchange\s+rates", spot_sec, maxsplit=1)[0]

        # Normalize lines (remove extra spaces but keep order)
        lines = [re.sub(r"\s+", " ", ln.strip()) for ln in spot_sec.splitlines() if ln.strip()]

        # Helper: find Bid/Ask numbers (with or without comma) in a small window around a label
        def grab_pair(label_regex: str):
            pat = re.compile(label_regex, flags=re.IGNORECASE)
            for i, ln in enumerate(lines):
                if pat.search(ln):
                    window = " ".join(lines[i:i+2])  # current + next line (handles wrapped cells)
                    # Accept 26,350 or 26350
                    nums = re.findall(r"\b\d{2},?\d{3}\b", window)
                    # Expect Bid then Ask
                    if len(nums) >= 2:
                        bid = self._to_int(nums[0])
                        ask = self._to_int(nums[1])
                        return bid, ask
                    # If only one number appears, map it to both sides (rare)
                    if len(nums) == 1:
                        v = self._to_int(nums[0])
                        return v, v
                    return None, None
            return None, None

        # Accept both 'preceding'/'preceeding'
        low_bid,   low_ask   = grab_pair(r"Lowest\s+rate\s+of\s+the\s+pre(?:c|cc)eding\s+week")
        high_bid,  high_ask  = grab_pair(r"Highest\s+rate\s+of\s+the\s+pre(?:c|cc)eding\s+week")
        close_bid, close_ask = grab_pair(r"Closing\s+rate\s+of\s+Friday\s*\(last\s*week\)")

        quoting_date = self._first_date(email_text) or ""

        rows = [
            {
                "No.": 1,
                "Bid/Ask": "Bid",
                "Bank": self.bank_name,
                "Quoting date": quoting_date,
                "Lowest rate of preceeding week": low_bid,
                "Highest rate of preceeding week": high_bid,
                "Closing rate of Friday (last week)": close_bid,
            },
            {
                "No.": 2,
                "Bid/Ask": "Ask",
                "Bank": self.bank_name,
                "Quoting date": quoting_date,
                "Lowest rate of preceeding week": low_ask,
                "Highest rate of preceeding week": high_ask,
                "Closing rate of Friday (last week)": close_ask,
            },
        ]
        return pd.DataFrame(rows, columns=out_cols)

    # -------------------------------
    # Forward parsing (robust rows)
    # -------------------------------
    def _parse_forwards(self, email_text: str) -> pd.DataFrame:
        """
        Forward section has two blocks: 'Bid Price:' and 'Ask Price:'.
        Each data row generally looks like:
          25/08/2025  22/09/2025  26,300  1M ( )  1.20  26,324
        Regex tolerates extra spaces, numbers with/without commas, and integer/decimal Gap.
        """
        out_cols = self.get_standard_columns()['forward']

        root = re.split(r"(?i)forward\s+exchange\s+rates", email_text, maxsplit=1)
        if len(root) < 2:
            return pd.DataFrame(columns=out_cols)

        tail = root[1]
        # Split into Bid/Ask blocks
        bid_split = re.split(r"(?i)\bBid\s*Price\s*:", tail, maxsplit=1)
        if len(bid_split) < 2:
            return pd.DataFrame(columns=out_cols)

        after_bid = bid_split[1]
        ask_split = re.split(r"(?i)\bAsk\s*Price\s*:", after_bid, maxsplit=1)
        bid_text = ask_split[0]
        ask_text = ask_split[1] if len(ask_split) > 1 else ""

        rows: List[Dict] = []
        rows += self._parse_forward_side(bid_text, "Bid")
        rows += self._parse_forward_side(ask_text, "Ask")

        if not rows:
            return pd.DataFrame(columns=out_cols)

        df = pd.DataFrame(rows, columns=[
            "Bid/Ask","Bank","Quoting date","Trading date","Value date",
            "Spot Exchange rate","Gap(%)","Forward Exchange rate",
            "Term (days)","% forward (cal)","Diff.","Term (lookup)"
        ])
        # Sort (Bid first), then by Trading date and tenor
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df

    def _parse_forward_side(self, text: str, side: str) -> List[Dict]:
        """Parse TCB forward side - Handle missing spot rates like VCB"""
        
        # TCB structure: 
        # Trading date (DD/MM/YYYY) -> Value date -> [Spot (only first row)] -> Term -> Gap% -> Forward rate
        
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Skip header lines
        data_start = 0
        for i, line in enumerate(lines):
            if re.match(self.DATE_DMY, line):
                data_start = i
                break
        
        if data_start == 0:
            return []
        
        data_lines = lines[data_start:]
        rows = []
        current_spot = None
        i = 0
        
        while i < len(data_lines):
            # Check if this looks like a trading date (DD/MM/YYYY)
            if not re.match(self.DATE_DMY, data_lines[i]):
                i += 1
                continue
                
            # We need at least 5 more elements: Value date, [Spot], Term, Gap%, Forward
            if i + 4 >= len(data_lines):
                break
                
            trd_s = data_lines[i]
            val_s = data_lines[i + 1]
            
            # Check if next element is a spot rate or a term
            next_elem = data_lines[i + 2]
            
            if re.match(r'\d{2},\d{3}', next_elem):
                # This row has spot rate
                spot_s = next_elem
                current_spot = self._to_int(spot_s)
                term_idx = i + 3
                gap_idx = i + 4
                fwd_idx = i + 5
                next_i = i + 6
            else:
                # This row doesn't have spot rate, use previous spot
                term_idx = i + 2
                gap_idx = i + 3
                fwd_idx = i + 4
                next_i = i + 5
            
            # Check bounds
            if fwd_idx >= len(data_lines):
                break
                
            term_s = data_lines[term_idx]
            gap_s = data_lines[gap_idx]
            fwd_s = data_lines[fwd_idx]
            
            # Extract term number and unit (TCB uses parentheses like ACB)
            term_match = re.match(r'(\d+)\s*([DMWY])\s*\(\s*\)', term_s)
            if not term_match:
                i = next_i
                continue
                
            termnum = int(term_match.group(1))
            termunit = term_match.group(2).upper()
            
            # Validate gap format
            if not re.match(r'\d+\.\d+', gap_s):
                i = next_i
                continue
                
            gap_pct = float(gap_s)
            
            # Validate forward rate format
            if not re.match(r'\d{2},\d{3}', fwd_s):
                i = next_i
                continue
                
            fwd = self._to_int(fwd_s)
            
            trd = self._to_date(trd_s)
            val = self._to_date(val_s)
            
            # Ensure correct order
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
            
            i = next_i
        
        return rows

    # -------------------------------
    # Central Bank stub
    # -------------------------------
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        out_cols = self.get_standard_columns()['central']
        qd = self._first_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)
