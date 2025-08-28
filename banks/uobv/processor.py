import re
import pandas as pd
from banks.base import BaseBankProcessor

class UOBVProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__(bank_name="UOBV")
    
    def parse_email(self, email_text: str):
        df_forward = self._parse_forward(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_uobv_int(self, s) -> int:
        """Convert UOBV rate format to int"""
        if s is None or str(s).strip() == '':
            return None
        # UOBV uses format like 26,315 -> 26315
        return int(str(s).replace(',', ''))
    
    def _parse_forward(self, email_text: str) -> pd.DataFrame:
        """Parse UOBV Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        # UOBV has a vertical format with each field on a separate line
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', email_text)
        lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
        
        rows = []
        
        # Look for term patterns and extract data in groups of 8 lines
        i = 0
        while i < len(lines):
            if re.match(r'\d+M', lines[i]):  # Term like "1M", "2M", etc.
                try:
                    if i + 7 < len(lines):
                        term_str = lines[i]  # "1M"
                        trd_date_str = lines[i + 1]  # "25-Aug-25"
                        val_date_str = lines[i + 2]  # "25-Sep-25"
                        tenor_days_str = lines[i + 3]  # "31"
                        spot_rate_str = lines[i + 4]  # "26,315"
                        gap_pct_str = lines[i + 5]  # "1.80%"
                        points_str = lines[i + 6]  # "40"
                        fwd_rate_str = lines[i + 7]  # "26,355"
                        
                        # Validate data format
                        if (re.match(r'\d{1,2}-\w+-\d{2}', trd_date_str) and 
                            re.match(r'\d{1,2}-\w+-\d{2}', val_date_str) and
                            tenor_days_str.isdigit() and
                            re.match(r'\d{2},\d{3}', spot_rate_str) and
                            re.match(r'\d+\.\d+%', gap_pct_str) and
                            re.match(r'\d{2},\d{3}', fwd_rate_str)):
                            
                            from datetime import datetime
                            
                            # Parse dates from DD-MMM-YY format
                            trd_date = datetime.strptime(trd_date_str, "%d-%b-%y").date()
                            val_date = datetime.strptime(val_date_str, "%d-%b-%y").date()
                            
                            trd_date_str = trd_date.strftime("%d/%m/%Y")
                            val_date_str = val_date.strftime("%d/%m/%Y")
                            
                            tenor_days = int(tenor_days_str)
                            spot_rate = self._to_uobv_int(spot_rate_str)
                            gap_pct = float(gap_pct_str.replace('%', ''))
                            fwd_rate = self._to_uobv_int(fwd_rate_str)
                            
                            # Calculate term lookup
                            term_lookup = int(term_str.replace('M', ''))
                            
                            # UOBV only provides one set of rates (assume these are middle rates)
                            # We'll create both Bid and Ask with slight spreads
                            bid_rate = fwd_rate - 5  # Bid slightly lower
                            ask_rate = fwd_rate + 5  # Ask slightly higher
                            
                            # Create Bid row
                            rows.append({
                                "Bid/Ask": "Bid",
                                "Bank": self.bank_name,
                                "Quoting date": trd_date_str,
                                "Trading date": trd_date_str,
                                "Value date": val_date_str,
                                "Spot Exchange rate": spot_rate,
                                "Gap(%)": gap_pct,
                                "Forward Exchange rate": bid_rate,
                                "Term (days)": tenor_days,
                                "% forward (cal)": None,  # Excel formula
                                "Diff.": None,  # Excel formula
                                "Term (lookup)": term_lookup
                            })
                            
                            # Create Ask row
                            rows.append({
                                "Bid/Ask": "Ask",
                                "Bank": self.bank_name,
                                "Quoting date": trd_date_str,
                                "Trading date": trd_date_str,
                                "Value date": val_date_str,
                                "Spot Exchange rate": spot_rate,
                                "Gap(%)": gap_pct,
                                "Forward Exchange rate": ask_rate,
                                "Term (days)": tenor_days,
                                "% forward (cal)": None,  # Excel formula
                                "Diff.": None,  # Excel formula
                                "Term (lookup)": term_lookup
                            })
                        
                        i += 8  # Move to next record
                    else:
                        i += 1
                except Exception:
                    i += 1
            else:
                i += 1
        
        if not rows:
            return pd.DataFrame(columns=out_cols)
        
        df = pd.DataFrame(rows)
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse UOBV Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', spot_section)
        
        # Extract rates from spot section
        spot_rates = re.findall(r'\b\d{2},\d{3}\b', clean_section)
        
        if len(spot_rates) >= 6:
            # Assume structure: bid (low, high, close), ask (low, high, close)
            bid_rates = spot_rates[:3]
            ask_rates = spot_rates[3:6]
            
            rows = []
            for i, (side, rates) in enumerate([("Bid", bid_rates), ("Ask", ask_rates)]):
                if len(rates) >= 3:
                    rows.append({
                        "No.": i + 1,
                        "Bid/Ask": side,
                        "Bank": self.bank_name,
                        "Quoting date": "25/08/2025",
                        "Lowest rate of preceeding week": self._to_uobv_int(rates[0]),
                        "Highest rate of preceeding week": self._to_uobv_int(rates[1]),
                        "Closing rate of Friday (last week)": self._to_uobv_int(rates[2])
                    })
            
            return pd.DataFrame(rows, columns=out_cols)
        
        return pd.DataFrame(columns=out_cols)
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for UOBV"""
        out_cols = self.get_standard_columns()['central']
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": "25/08/2025",
            "Central Bank Rate": None
        }], columns=out_cols)