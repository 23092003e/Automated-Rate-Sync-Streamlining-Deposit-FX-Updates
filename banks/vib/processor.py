import re
import pandas as pd
from banks.base import BaseBankProcessor

class VIBProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__(bank_name="VIB")
    
    def parse_email(self, email_text: str):
        df_forward = self._parse_forward(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_vib_int(self, s) -> int:
        """Convert VIB rate format to int (handles commas and decimals)"""
        if s is None or str(s).strip() == '':
            return None
        # VIB uses format like 26,310.00 -> 26310 or 26285 -> 26285
        return int(float(str(s).replace(',', '')))
    
    def _parse_forward(self, email_text: str) -> pd.DataFrame:
        """Parse VIB Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        # Find forward section
        parts = re.split(r"(?i)forward\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        forward_section = parts[1]
        spot_parts = re.split(r"(?i)spot\s+exchange\s+rates?", forward_section, maxsplit=1)
        forward_only = spot_parts[0]
        
        # Clean unicode
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', forward_only)
        
        # Extract dates for terms
        dates = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', clean_section)
        
        # Extract all rates (both 5-digit and comma formats)
        rates_comma = re.findall(r'\b\d{2},\d{3}(?:\.\d{2})?\b', clean_section)
        rates_plain = re.findall(r'\b\d{5}\b', clean_section)
        
        all_rates = rates_comma + rates_plain
        
        rows = []
        terms = ['1M', '3M', '6M', '9M', '12M']  # Standard terms
        
        # Process rates in pairs (bid, ask) for each term
        rate_index = 0
        for i, term in enumerate(terms):
            if rate_index + 1 < len(all_rates):
                bid_rate = self._to_vib_int(all_rates[rate_index])
                ask_rate = self._to_vib_int(all_rates[rate_index + 1])
                
                # Use corresponding date if available
                trading_date = dates[i] if i < len(dates) else "25/08/2025"
                
                if bid_rate and ask_rate:
                    rows.append({
                        "No.": len(rows) + 1,
                        "Bid/Ask": "Bid",
                        "Bank": self.bank_name,
                        "Terms": term,
                        "Trading date": trading_date,
                        "Forward Rate": bid_rate
                    })
                    rows.append({
                        "No.": len(rows) + 1,
                        "Bid/Ask": "Ask",
                        "Bank": self.bank_name,
                        "Terms": term,
                        "Trading date": trading_date,
                        "Forward Rate": ask_rate
                    })
                
                rate_index += 2
        
        return pd.DataFrame(rows, columns=out_cols)
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse VIB Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', spot_section)
        
        # Extract rates from spot section (prefer comma format first)
        rates_comma = re.findall(r'\b\d{2},\d{3}(?:\.\d{2})?\b', clean_section)
        rates_plain = re.findall(r'\b\d{5}\b', clean_section)
        
        spot_rates = rates_comma + rates_plain
        
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
                        "Quoting date": self._first_date(email_text) or "25/08/2025",
                        "Lowest rate of preceeding week": self._to_vib_int(rates[0]),
                        "Highest rate of preceeding week": self._to_vib_int(rates[1]),
                        "Closing rate of Friday (last week)": self._to_vib_int(rates[2])
                    })
            
            return pd.DataFrame(rows, columns=out_cols)
        
        return pd.DataFrame(columns=out_cols)
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for VIB"""
        out_cols = self.get_standard_columns()['central']
        qd = self._first_date(email_text) or "25/08/2025"
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)