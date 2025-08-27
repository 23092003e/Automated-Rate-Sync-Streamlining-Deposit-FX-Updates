# -*- coding: utf-8 -*-
"""
BIDV (Bank for Investment and Development of Vietnam) Email Parser - Fixed version
Handles BIDV-specific email format with comma separators and decimal points
"""

import re
from datetime import datetime, date
import pandas as pd
from typing import Tuple

from ..base import BaseBankProcessor


class BIDVProcessor(BaseBankProcessor):
    """BIDV-specific email processor - Fixed"""
    
    def __init__(self):
        super().__init__("BIDV")
    
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Parse BIDV email format"""
        df_forward = self._parse_forwards(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_bidv_int(self, s) -> int:
        """Convert BIDV rate format to int (handles commas and decimal points)"""
        if s is None or str(s).strip() == '':
            return None
        # BIDV uses format like 26,284.00 -> 26284
        return int(float(str(s).replace(',', '')))
    
    def _first_bidv_date(self, text: str) -> str:
        """Extract first date from BIDV email (MM/DD/YYYY format)"""
        dates = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', text)
        if dates:
            # Convert MM/DD/YYYY to DD/MM/YYYY
            mdy = dates[0]
            parts = mdy.split('/')
            return f"{parts[1]}/{parts[0]}/{parts[2]}"
        return ""
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse BIDV Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        forward_parts = re.split(r"(?i)forward\s+exchange\s+rates", spot_section, maxsplit=1)
        spot_only = forward_parts[0]
        
        # BIDV format: numbers with commas and decimal points like 26,284.00
        rate_pattern = r'\b\d{2},\d{3}\.\d{2}\b'
        all_rates = re.findall(rate_pattern, spot_only)
        
        # Extract rates in order based on BIDV structure
        bid_close = ask_close = None
        low_bid = low_ask = high_bid = high_ask = None
        
        if len(all_rates) >= 6:
            low_bid = self._to_bidv_int(all_rates[0])   # 26,284.00
            low_ask = self._to_bidv_int(all_rates[1])   # 26,285.00  
            high_bid = self._to_bidv_int(all_rates[2])  # 26,430.00
            high_ask = self._to_bidv_int(all_rates[3])  # 26,431.00
            bid_close = self._to_bidv_int(all_rates[4]) # 26,428.00
            ask_close = self._to_bidv_int(all_rates[5]) # 26,429.00
        
        quoting_date = self._first_bidv_date(email_text) or ""
        
        rows = []
        for i, (side, closing, low, high) in enumerate([
            ("Bid", bid_close, low_bid, high_bid), 
            ("Ask", ask_close, low_ask, high_ask)
        ], start=1):
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
        """Parse BIDV Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        root = re.split(r"(?i)forward\s+exchange\s+rates", email_text, maxsplit=1)
        if len(root) < 2:
            return pd.DataFrame(columns=out_cols)
        tail = root[1]
        
        # Split by Bid Price and Ask Price sections
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
        """Parse BIDV forward side - Handle missing spot rates like VCB"""
        # BIDV uses MM/DD/YYYY date format
        DATE_MDY = r"(?:0[1-9]|1[0-2])/(?:0[1-9]|[12]\d|3[01])/(?:19|20)\d\d"
        
        # BIDV structure: 
        # Trading date (MM/DD/YYYY) -> Value date -> [Spot (only first row)] -> Term -> Gap% -> Forward rate
        
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Skip header lines
        data_start = 0
        for i, line in enumerate(lines):
            if re.match(DATE_MDY, line):
                data_start = i
                break
        
        if data_start == 0:
            return []
        
        data_lines = lines[data_start:]
        rows = []
        current_spot = None
        i = 0
        
        while i < len(data_lines):
            # Check if this looks like a trading date (MM/DD/YYYY)
            if not re.match(DATE_MDY, data_lines[i]):
                i += 1
                continue
                
            # We need at least 5 more elements: Value date, [Spot], Term, Gap%, Forward
            if i + 4 >= len(data_lines):
                break
                
            trd_s = data_lines[i]
            val_s = data_lines[i + 1]
            
            # Check if next element is a spot rate or a term
            next_elem = data_lines[i + 2]
            
            if re.match(r'\d{2},\d{3}\.\d{2}', next_elem):
                # This row has spot rate
                spot_s = next_elem
                current_spot = self._to_bidv_int(spot_s)
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
            
            # Extract term number and unit (BIDV doesn't use parentheses)
            term_match = re.match(r'(\d+)\s*([DMWY])', term_s)
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
            
            # Convert MM/DD/YYYY to DD/MM/YYYY for parsing
            trd_parts = trd_s.split('/')
            val_parts = val_s.split('/')
            trd_dmy = f"{trd_parts[1]}/{trd_parts[0]}/{trd_parts[2]}"
            val_dmy = f"{val_parts[1]}/{val_parts[0]}/{val_parts[2]}"
            
            trd = self._to_date(trd_dmy)
            val = self._to_date(val_dmy)
            
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
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for BIDV"""
        out_cols = self.get_standard_columns()['central']
        qd = self._first_bidv_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)