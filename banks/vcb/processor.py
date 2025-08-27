# -*- coding: utf-8 -*-
"""
VCB (Vietcombank) Email Parser - Fixed version for missing spot rates
VCB format: Only first row has Spot rate, subsequent rows missing it
"""

import re
from datetime import datetime, date
import pandas as pd
from typing import Tuple

from ..base import BaseBankProcessor


class VCBProcessor(BaseBankProcessor):
    """VCB-specific email processor - Fixed for missing spots"""
    
    def __init__(self):
        super().__init__("VCB")
    
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Parse VCB email format"""
        df_forward = self._parse_forwards(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_vcb_int(self, s) -> int:
        """Convert VCB rate format to int (handles decimal points)"""
        if s is None or str(s).strip() == '':
            return None
        # VCB uses format like 26.090 -> 26090
        return int(str(s).replace('.', '').replace(',', ''))
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse VCB Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        forward_parts = re.split(r"(?i)forward\s+exchange\s+rates", spot_section, maxsplit=1)
        spot_only = forward_parts[0]
        
        # VCB format: numbers with decimal points like 26.090, 26.450
        rate_pattern = r'\b\d{2}\.\d{3}\b'  # Matches 26.090 format
        all_rates = re.findall(rate_pattern, spot_only)
        
        # Extract rates in order: low_bid, low_ask, high_bid, high_ask, close_bid, close_ask
        bid_close = ask_close = None
        low_bid = low_ask = high_bid = high_ask = None
        
        if len(all_rates) >= 6:
            low_bid = self._to_vcb_int(all_rates[0])    # 26.090
            low_ask = self._to_vcb_int(all_rates[1])    # 26.450
            high_bid = self._to_vcb_int(all_rates[2])   # 26.242
            high_ask = self._to_vcb_int(all_rates[3])   # 26.562
            bid_close = self._to_vcb_int(all_rates[4])  # 26.242
            ask_close = self._to_vcb_int(all_rates[5])  # 26.562
        
        quoting_date = self._first_date(email_text) or ""
        
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
        """Parse VCB Forward Exchange Rates"""
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
        """Parse VCB forward side - Handle missing spot rates"""
        
        # VCB structure: 
        # Trading date -> Value date -> [Spot (only first row)] -> Term -> Forward rate
        
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
            # Check if this looks like a trading date
            if not re.match(self.DATE_DMY, data_lines[i]):
                i += 1
                continue
                
            # We need at least 4 more elements: Value date, [Spot], Term, Forward
            if i + 3 >= len(data_lines):
                break
                
            trd_s = data_lines[i]
            val_s = data_lines[i + 1]
            
            # Check if next element is a spot rate or a term
            next_elem = data_lines[i + 2]
            
            if re.match(r'\d{2}\.\d{3}', next_elem):
                # This row has spot rate
                spot_s = next_elem
                current_spot = self._to_vcb_int(spot_s)
                term_idx = i + 3
                fwd_idx = i + 4
                next_i = i + 5
            else:
                # This row doesn't have spot rate, use previous spot
                term_idx = i + 2
                fwd_idx = i + 3
                next_i = i + 4
            
            # Check bounds
            if fwd_idx >= len(data_lines):
                break
                
            term_s = data_lines[term_idx]
            fwd_s = data_lines[fwd_idx]
            
            # Extract term number and unit
            term_match = re.match(r'(\d+)\s*([DMWY])\s*\(\s*\)', term_s)
            if not term_match:
                i = next_i
                continue
                
            termnum = int(term_match.group(1))
            termunit = term_match.group(2).upper()
            
            # Validate forward rate format
            if not re.match(r'\d{2}\.\d{3}', fwd_s):
                i = next_i
                continue
                
            fwd = self._to_vcb_int(fwd_s)
            
            trd = self._to_date(trd_s)
            val = self._to_date(val_s)
            
            # Ensure correct order
            if val < trd:
                trd, val = val, trd
            
            term_days = self._days(trd, val)
            term_lookup = round(self._yearfrac_30360_us(trd, val) * 12)
            
            # Calculate gap % from forward and spot (VCB doesn't provide it)
            gap_pct = ((fwd - current_spot) / current_spot * 100) if current_spot else 0
            
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
        """Build Central Bank rate stub for VCB"""
        out_cols = self.get_standard_columns()['central']
        qd = self._first_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)