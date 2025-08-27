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
        """
        Parse one side (Bid/Ask). Each match yields a row.
        - Accept spot/fwd both '26,300' or '26300'
        - Accept gap '1' or '1.20'
        - Term captured but only used for consistency; Term(days) computed by Value-Trading
        """
        # \s+ matches newlines; labels tolerated
        ROW_RE = re.compile(
            rf"(?P<trd>{self.DATE_DMY})\s+"
            rf"(?P<val>{self.DATE_DMY})\s+"
            rf"(?P<spot>\d{{2}},?\d{{3}})\s+"
            rf"(?P<termnum>\d+)\s*(?P<termunit>[DMWY])?\s*\(\s*\)\s+"
            rf"(?P<gap>-?\d+(?:[.,]\d+)?)\s+"
            rf"(?P<fwd>\d{{2}},?\d{{3}})",
            flags=re.IGNORECASE
        )

        rows: List[Dict] = []
        for m in ROW_RE.finditer(text):
            trd_s = m.group("trd")
            val_s = m.group("val")
            spot_s = m.group("spot")
            termnum = int(m.group("termnum"))
            # ter mun it captured but not strictly needed
            gap_s = m.group("gap").replace(",", ".")
            fwd_s = m.group("fwd")

            spot = self._to_int(spot_s)      # '26,300' -> 26300
            fwd = self._to_int(fwd_s)        # '26,324' -> 26324
            gap_pct = float(gap_s)           # '1.20' -> 1.20

            trd = self._to_date(trd_s)       # date obj
            val = self._to_date(val_s)

            # Ensure non-negative days
            if val < trd:
                trd, val = val, trd

            term_days = self._days(trd, val)
            term_lookup = round(self._yearfrac_30360_us(trd, val) * 12)

            rows.append({
                "Bid/Ask": side,
                "Bank": self.bank_name,
                "Quoting date": trd,     # keep as date (Excel YEARFRAC ok)
                "Trading date": trd,
                "Value date": val,
                "Spot Exchange rate": spot,
                "Gap(%)": gap_pct,       # keep 1.20 (not 0.012)
                "Forward Exchange rate": fwd,
                "Term (days)": term_days,
                "% forward (cal)": None, # to be filled by Excel formula
                "Diff.": None,           # to be filled by Excel formula
                "Term (lookup)": term_lookup
            })
        return rows

    # -------------------------------
    # Central Bank stub
    # -------------------------------
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        out_cols = self.get_standard_columns()['central']
        qd = self._first_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)
