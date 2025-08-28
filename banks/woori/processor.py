import re
import pandas as pd
from banks.base import BaseBankProcessor

class WooriProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__(bank_name="WOORI")
    
    def parse_email(self, email_text: str):
        df_forward = self._parse_forward(email_text)
        df_spot = self._parse_spot(email_text)
        df_central = self._build_central_bank(email_text)
        return df_forward, df_spot, df_central
    
    def _to_woori_int(self, s) -> int:
        """Convert Woori rate format to int (handles commas and decimals)"""
        if s is None or str(s).strip() == '':
            return None
        # Woori uses format like 26,449.32 -> 26449 or 26,562 -> 26562
        return int(float(str(s).replace(',', '')))
    
    def _parse_forward(self, email_text: str) -> pd.DataFrame:
        """Parse Woori Forward Exchange Rates"""
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
        rows += self._parse_woori_forward_side(bid_text, "Bid")
        rows += self._parse_woori_forward_side(ask_text, "Ask")
        
        if not rows:
            return pd.DataFrame(columns=out_cols)
        
        df = pd.DataFrame(rows)
        df = df.sort_values(["Bid/Ask", "Trading date", "Term (days)"]).reset_index(drop=True)
        df.insert(0, "No.", range(1, len(df) + 1))
        return df
    
    def _parse_woori_forward_side(self, text: str, side: str) -> list:
        """Parse Woori forward side (Bid or Ask)"""
        rows = []
        
        # Clean unicode and split into lines
        clean_text = re.sub(r'[^\x00-\x7F]+', ' ', text)
        lines = [line.strip() for line in clean_text.split('\n') if line.strip()]
        
        # Parse rows by combining consecutive lines
        i = 0
        while i < len(lines):
            if re.match(r'\d{1,2}-\d{1,2}-\d{4}', lines[i]):  # Trading date line "22-08-2025"
                try:
                    if i + 4 < len(lines):
                        trd_date_str = lines[i]  # "22-08-2025"
                        val_date_str = lines[i + 1]  # "22-08-2025" (same in Woori format)
                        
                        # Skip empty spot rate field (3rd line usually empty)
                        term_str = None
                        gap_str = None
                        fwd_str = None
                        
                        # Find term, gap%, and forward rate in subsequent lines
                        for j in range(i + 2, min(i + 6, len(lines))):
                            line = lines[j]
                            if "M" in line and "(" in line:
                                term_str = line  # "1M ( )"
                            elif re.match(r'\d+\.?\d*', line) and "." in line and not "," in line:
                                gap_str = line  # "1.35"
                            elif re.match(r'\d{2},\d{3}\.\d{2}', line):
                                fwd_str = line  # "26,449.32"
                        
                        if fwd_str:
                            # Convert dates from DD-MM-YYYY format
                            from datetime import datetime
                            
                            trd_date = datetime.strptime(trd_date_str, "%d-%m-%Y").date()
                            val_date = datetime.strptime(val_date_str, "%d-%m-%Y").date()
                            
                            trd_date_str = trd_date.strftime("%d/%m/%Y")
                            val_date_str = val_date.strftime("%d/%m/%Y")
                            
                            # Calculate term days and lookup (Woori seems to use same trading/value dates)
                            # We'll estimate based on term
                            term_match = re.search(r'(\d+)M', term_str) if term_str else None
                            if term_match:
                                term_months = int(term_match.group(1))
                                term_days = term_months * 30  # Approximate
                                
                                # Create proper value date by adding term days
                                from datetime import timedelta
                                val_date = trd_date + timedelta(days=term_days)
                                val_date_str = val_date.strftime("%d/%m/%Y")
                            else:
                                term_months = 1
                                term_days = 30
                            
                            spot_rate = 26400  # Default spot rate for Woori
                            fwd_rate = self._to_woori_int(fwd_str)
                            gap_pct = float(gap_str) if gap_str else None
                            
                            rows.append({
                                "Bid/Ask": side,
                                "Bank": self.bank_name,
                                "Quoting date": trd_date_str,
                                "Trading date": trd_date_str,
                                "Value date": val_date_str,
                                "Spot Exchange rate": spot_rate,
                                "Gap(%)": gap_pct,
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
        """Parse Woori Spot Exchange Rates"""
        out_cols = self.get_standard_columns()['spot']
        
        # Find spot section
        parts = re.split(r"(?i)spot\s+exchange\s+rates?", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        
        spot_section = parts[1]
        clean_section = re.sub(r'[^\x00-\x7F]+', ' ', spot_section)
        
        # Extract rates from spot section
        rates_decimal = re.findall(r'\b\d{2},\d{3}\.\d{2}\b', clean_section)
        rates_simple = re.findall(r'\b\d{2},\d{3}\b', clean_section)
        
        spot_rates = rates_decimal if rates_decimal else rates_simple
        
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
                        "Lowest rate of preceeding week": self._to_woori_int(rates[0]),
                        "Highest rate of preceeding week": self._to_woori_int(rates[1]),
                        "Closing rate of Friday (last week)": self._to_woori_int(rates[2])
                    })
            
            return pd.DataFrame(rows, columns=out_cols)
        
        return pd.DataFrame(columns=out_cols)
    
    def _build_central_bank(self, email_text: str) -> pd.DataFrame:
        """Build Central Bank rate stub for Woori"""
        out_cols = self.get_standard_columns()['central']
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": "25/08/2025",
            "Central Bank Rate": None
        }], columns=out_cols)