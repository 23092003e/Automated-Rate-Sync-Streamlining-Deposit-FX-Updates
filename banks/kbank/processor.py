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
        
        # Find forward section
        parts = re.split(r"(?i)forward\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        forward_section = parts[1]
        spot_parts = re.split(r"(?i)spot\s+exchange\s+rates?", forward_section, maxsplit=1)
        forward_only = spot_parts[0]
        
        # Clean unicode and extract lines
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', forward_only)
        lines = [line.strip() for line in clean_section.split('\n') if line.strip()]
        
        # Look for terms and rates pattern
        rows = []
        current_term = None
        
        for line in lines:
            # Look for term patterns like "1 MONTH", "2 MONTHS", etc.
            term_match = re.search(r'(\d+)\s+MONTHS?', line, re.IGNORECASE)
            if term_match:
                current_term = f"{term_match.group(1)}M"
                
                # Extract rates from the same line
                rate_matches = re.findall(r'\b\d{2},\d{3}(?:\.\d{2})?\b', line)
                if len(rate_matches) >= 2:
                    bid_rate = self._to_kbank_int(rate_matches[0])
                    ask_rate = self._to_kbank_int(rate_matches[1])
                    
                    if bid_rate and ask_rate:
                        rows.append({
                            "No.": len(rows) + 1,
                            "Bid/Ask": "Bid",
                            "Bank": self.bank_name,
                            "Terms": current_term,
                            "Trading date": "25/08/2025",  # Default date
                            "Forward Rate": bid_rate
                        })
                        rows.append({
                            "No.": len(rows) + 1,
                            "Bid/Ask": "Ask", 
                            "Bank": self.bank_name,
                            "Terms": current_term,
                            "Trading date": "25/08/2025",
                            "Forward Rate": ask_rate
                        })
        
        return pd.DataFrame(rows, columns=out_cols)
    
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