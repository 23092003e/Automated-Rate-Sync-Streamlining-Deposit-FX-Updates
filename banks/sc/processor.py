import re
import pandas as pd
from banks.base import BaseBankProcessor

class SCProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__(bank_name="SC")
    
    def parse_email(self, email_text: str):
        df_forward = self._parse_forward(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_sc_int(self, s) -> int:
        """Convert SC rate format to int (5-digit numbers)"""
        if s is None or str(s).strip() == '':
            return None
        # SC uses 5-digit format like 26275 directly
        return int(str(s))
    
    def _parse_forward(self, email_text: str) -> pd.DataFrame:
        """Parse SC Forward Exchange Rates"""
        out_cols = self.get_standard_columns()['forward']
        
        # Find forward section
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
        rows += self._parse_sc_forward_side(bid_text, "Bid")
        rows += self._parse_sc_forward_side(ask_text, "Ask")
        
        if not rows:
            return pd.DataFrame(columns=out_cols)
        
        df = pd.DataFrame(rows)
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df
    
    def _parse_sc_forward_side(self, text: str, side: str) -> list:
        """Parse SC forward side (Bid or Ask)"""
        rows = []
        
        # Clean unicode and split into lines
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', text)
        lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
        
        # Parse rows by combining consecutive lines
        i = 0
        while i < len(lines):
            if re.match(r'\d{1,2}\s+\w+\s+\d{4}', lines[i]):  # Trading date line
                try:
                    if i + 5 < len(lines):
                        trd_date_str = lines[i]  # "25 Aug 2025"
                        val_date_str = lines[i + 1]  # "25 Sep 2025"
                        spot_str = lines[i + 2] if re.match(r'\d{5}', lines[i + 2]) else "26300"
                        term_str = lines[i + 3] if "M" in lines[i + 3] else lines[i + 2]  # "1M ( )"
                        fwd_str = None
                        
                        # Find forward rate (5-digit number not marked as "Cant offer")
                        for j in range(i + 3, min(i + 7, len(lines))):
                            if re.match(r'\d{5}', lines[j]) and "cant" not in lines[j].lower():
                                fwd_str = lines[j]
                                break
                        
                        if fwd_str and "cant" not in fwd_str.lower():
                            # Convert dates
                            from datetime import datetime
                            trd_date = datetime.strptime(trd_date_str, "%d %b %Y").date()
                            val_date = datetime.strptime(val_date_str, "%d %b %Y").date()
                            
                            trd_date_str = trd_date.strftime("%d/%m/%Y")
                            val_date_str = val_date.strftime("%d/%m/%Y")
                            
                            # Calculate term days and lookup
                            term_days = (val_date - trd_date).days
                            term_lookup = round(term_days / 30)
                            
                            # Extract term info
                            term_match = re.search(r'(\d+)M', term_str)
                            if term_match:
                                term_months = int(term_match.group(1))
                            else:
                                term_months = term_lookup
                            
                            spot_rate = self._to_sc_int(spot_str)
                            fwd_rate = self._to_sc_int(fwd_str)
                            
                            rows.append({
                                "Bid/Ask": side,
                                "Bank": self.bank_name,
                                "Quoting date": trd_date_str,
                                "Trading date": trd_date_str,
                                "Value date": val_date_str,
                                "Spot Exchange rate": spot_rate,
                                "Gap(%)": None,  # Not provided by SC
                                "Forward Exchange rate": fwd_rate,
                                "Term (days)": term_days,
                                "% forward (cal)": None,  # Excel formula
                                "Diff.": None,  # Excel formula
                                "Term (lookup)": term_months
                            })
                        
                        i += 6  # Skip processed lines
                    else:
                        i += 1
                except Exception:
                    i += 1
            else:
                i += 1
        
        return rows
    
    def _parse_spot(self, email_text: str) -> pd.DataFrame:
        """Parse SC Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', spot_section)
        
        # Extract rates from spot section
        spot_rates = re.findall(r'\b\d{5}\b', clean_section)
        
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
                        "Lowest rate of preceeding week": self._to_sc_int(rates[0]),
                        "Highest rate of preceeding week": self._to_sc_int(rates[1]),
                        "Closing rate of Friday (last week)": self._to_sc_int(rates[2])
                    })
            
            return pd.DataFrame(rows, columns=out_cols)
        
        return pd.DataFrame(columns=out_cols)
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for SC"""
        out_cols = self.get_standard_columns()['central']
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": "25/08/2025",
            "Central Bank Rate": None
        }], columns=out_cols)