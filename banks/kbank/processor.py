import re
import pandas as pd
from banks.base import BaseBankProcessor

class KBankProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__(bank_name="KBANK")
    
    def parse_email(self, email_text: str):
        df_forward = self._parse_forward(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_kbank_int(self, s) -> int:
        """Convert KBank rate format to int (handles commas)"""
        if s is None or str(s).strip() == '':
            return None
        # KBank uses format like 26,280 or 26,295.00 -> 26280 or 26295
        return int(float(str(s).replace(',', '')))
    
    def _parse_forward(self, email_text: str) -> pd.DataFrame:
        """Parse KBank Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        # Clean unicode first
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', email_text)
        
        # Find forward section
        parts = re.split(r"(?i)forward\s+exchange\s+rates", clean_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        tail = parts[1]
        
        # Split into Bid and Ask sections
        bid_parts = re.split(r"(?i)KBank\s*s\s*Bid\s*Price", tail, maxsplit=1)
        if len(bid_parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        after_bid = bid_parts[1]
        ask_parts = re.split(r"(?i)KBank\s*s\s*Ask\s*Price", after_bid, maxsplit=1)
        
        bid_text = ask_parts[0]
        ask_text = ask_parts[1] if len(ask_parts) > 1 else ""
        
        rows = []
        rows += self._parse_kbank_forward_side(bid_text, "Bid")
        rows += self._parse_kbank_forward_side(ask_text, "Ask")
        
        if not rows:
            return pd.DataFrame(columns=out_cols)
        
        df = pd.DataFrame(rows)
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df
    
    def _parse_kbank_forward_side(self, text: str, side: str) -> list:
        """Parse KBank forward side (Bid or Ask) - vertical table format"""
        rows = []
        
        # Clean unicode and split into lines
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', text)
        lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
        
        # Skip header lines and find data starting with row numbers
        data_start = -1
        for i, line in enumerate(lines):
            if line == '1':  # First data row number
                data_start = i
                break
        
        if data_start == -1:
            return rows
        
        # Process data lines in groups of 9 (No, Product, Side, Value date, Tenor, Spot rate, Gap%, Point, FWD Rate)
        i = data_start
        while i < len(lines):
            if i + 8 < len(lines) and lines[i].isdigit():
                try:
                    row_no = int(lines[i])
                    product = lines[i + 1]  # "Forward"
                    trade_side = lines[i + 2]  # "Sell" or "Buy"
                    value_date_str = lines[i + 3]  # "24 September 2025"
                    tenor_days_str = lines[i + 4]  # "30"
                    spot_rate_str = lines[i + 5]  # "26,295.00"
                    gap_pct_str = lines[i + 6]  # "1.20%"
                    point_str = lines[i + 7]  # "26.00"
                    fwd_rate_str = lines[i + 8]  # "26,321.00"
                    
                    if product == "Forward" and re.match(r'\d{1,2}\s+\w+\s+\d{4}', value_date_str):
                        # Convert value date
                        from datetime import datetime
                        val_date = datetime.strptime(value_date_str, "%d %B %Y").date()
                        val_date_str = val_date.strftime("%d/%m/%Y")
                        
                        # Use same trading date as value date for KBank
                        trd_date_str = val_date_str
                        
                        tenor_days = int(tenor_days_str)
                        spot_rate = self._to_kbank_int(spot_rate_str)
                        gap_pct = float(gap_pct_str.replace('%', ''))
                        fwd_rate = self._to_kbank_int(fwd_rate_str)
                        
                        # Calculate term lookup
                        term_lookup = round(tenor_days / 30)
                        
                        rows.append({
                            "Bid/Ask": side,
                            "Bank": self.bank_name,
                            "Quoting date": trd_date_str,
                            "Trading date": trd_date_str,
                            "Value date": val_date_str,
                            "Spot Exchange rate": spot_rate,
                            "Gap(%)": gap_pct,
                            "Forward Exchange rate": fwd_rate,
                            "Term (days)": tenor_days,
                            "% forward (cal)": None,  # Excel formula
                            "Diff.": None,  # Excel formula
                            "Term (lookup)": term_lookup
                        })
                    
                    i += 9  # Move to next record
                except Exception:
                    i += 1
            else:
                i += 1
        
        return rows
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse KBank Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        # Clean unicode
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', spot_section)
        
        # Extract all rates from spot section
        rate_matches = re.findall(r'\b\d{2},\d{3}(?:\.\d{2})?\b', clean_section)
        
        # Assume standard structure: bid rates then ask rates
        if len(rate_matches) >= 6:
            # First 3 rates are bid (low, high, close), next 3 are ask
            bid_rates = rate_matches[:3] 
            ask_rates = rate_matches[3:6]
            
            rows = []
            for i, (side, rates) in enumerate([("Bid", bid_rates), ("Ask", ask_rates)]):
                if len(rates) >= 3:
                    rows.append({
                        "No.": i + 1,
                        "Bid/Ask": side,
                        "Bank": self.bank_name,
                        "Quoting date": "25/08/2025",
                        "Lowest rate of preceeding week": self._to_kbank_int(rates[0]),
                        "Highest rate of preceeding week": self._to_kbank_int(rates[1]),
                        "Closing rate of Friday (last week)": self._to_kbank_int(rates[2])
                    })
            
            return pd.DataFrame(rows, columns=out_cols)
        
        return pd.DataFrame(columns=out_cols)
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for KBank"""
        out_cols = self.get_standard_columns()['central']
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": "25/08/2025",
            "Central Bank Rate": None
        }], columns=out_cols)