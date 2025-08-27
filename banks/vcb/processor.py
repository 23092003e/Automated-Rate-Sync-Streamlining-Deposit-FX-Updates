# -*- coding: utf-8 -*-
"""
VCB (Vietcombank) Email Parser - Fixed version
- Forward: handle missing spot on subsequent rows; robust Bid/Ask splitting
- Spot: parse by labels; accept 26.090 or 26090 styles
- Quoting date (Spot & CentralBank) = min(Trading date) from Forward
"""

import re
from datetime import datetime, date
import pandas as pd
from typing import Tuple

from ..base import BaseBankProcessor


class VCBProcessor(BaseBankProcessor):
    """VCB-specific email processor - Fixed for missing spots & quoting date"""

    def __init__(self):
        super().__init__("VCB")

    # -------------------------------
    # Public API
    # -------------------------------
    def parse_email(self, email_text: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Parse VCB email format"""
        df_forward = self._parse_forwards(email_text)

        # quoting date = min Trading date in Forward (fallback: first date in email)
        if not df_forward.empty and "Trading date" in df_forward:
            qd = df_forward["Trading date"].min()
        else:
            qd_str = self._first_date(email_text) or ""
            qd = self._to_date(qd_str) if qd_str else ""

        df_spot = self._parse_spot(email_text, quoting_date_override=qd)
        df_central = self._build_central_bank(quoting_date_override=qd)
        return df_forward, df_spot, df_central

    # -------------------------------
    # Utils
    # -------------------------------
    def _to_vcb_int(self, s) -> int:
        """Convert VCB rate format to int (handles 26.090 → 26090, 26,090 → 26090)"""
        if s is None or str(s).strip() == '':
            return None
        return int(str(s).replace('.', '').replace(',', ''))

    # -------------------------------
    # Spot parsing (label-anchored)
    # -------------------------------
    def _parse_spot(self, email_text: str, quoting_date_override=None) -> pd.DataFrame:
        """
        VCB Spot Exchange Rates (robust):
        - Đọc theo nhãn: Lowest / Highest / Closing
        - Gom số trong block [nhãn .. nhãn kế tiếp] (tối đa 6 dòng),
            hỗ trợ 26.090 / 26,090 / 26090
        - Trả về Bid/Ask riêng; nếu chỉ bắt được 1 số -> Bid = số đó, Ask = None
            (KHÔNG copy Bid sang Ask)
        """
        out_cols = self.get_standard_columns()['spot']

        # 1) Cắt riêng phần Spot (loại Forward)
        parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
        if len(parts) < 2:
            return pd.DataFrame(columns=out_cols)
        spot_section = parts[1]
        spot_only = re.split(r"(?i)forward\s+exchange\s+rates", spot_section, maxsplit=1)[0]

        # 2) Chuẩn hoá dòng
        lines = [re.sub(r"\s+", " ", ln.strip()) for ln in spot_only.splitlines() if ln.strip()]

        # 3) Tìm index của 3 nhãn
        def find_idx(regex: str) -> int:
            pat = re.compile(regex, flags=re.IGNORECASE)
            for i, ln in enumerate(lines):
                if pat.search(ln):
                    return i
            return -1

        idx_low   = find_idx(r"Lowest\s+rate\s+of\s+the\s+pre(?:c|cc)eding\s+week")
        idx_high  = find_idx(r"Highest\s+rate\s+of\s+the\s+pre(?:c|cc)eding\s+week")
        idx_close = find_idx(r"Closing\s+rate\s+of\s+Friday\s*\(last\s*week\)")

        # 4) Hàm trích 2 số (Bid, Ask) trong block từ start..end
        rate_re = re.compile(r"\b\d{2}[.,]?\d{3}\b")  # 26.090 / 26,090 / 26090

        def extract_pair(start_idx: int, end_idx: int) -> tuple:
            if start_idx == -1:
                return (None, None)
            # Giới hạn block tối đa 6 dòng để tránh ăn sang phần khác
            j_end = end_idx if end_idx != -1 else min(len(lines), start_idx + 6)
            block = " ".join(lines[start_idx:j_end])
            nums = rate_re.findall(block)

            if len(nums) >= 2:
                return self._to_vcb_int(nums[0]), self._to_vcb_int(nums[1])
            elif len(nums) == 1:
                # Chỉ có 1 số gần nhãn -> coi là Bid; Ask để None (không copy)
                v = self._to_vcb_int(nums[0])
                return v, None
            else:
                return (None, None)

        # 5) Xác định ranh giới block theo thứ tự nhãn
        #    low ... high ... close ...
        end_low  = idx_high  if idx_high  != -1 else (idx_close if idx_close != -1 else -1)
        end_high = idx_close if idx_close != -1 else -1

        low_bid,   low_ask   = extract_pair(idx_low,   end_low)
        high_bid,  high_ask  = extract_pair(idx_high,  end_high)
        close_bid, close_ask = extract_pair(idx_close, -1)

        # 6) Quoting date đồng bộ (ưu tiên từ Forward) - convert to string format
        if quoting_date_override:
            if hasattr(quoting_date_override, 'strftime'):
                quoting_date = quoting_date_override.strftime('%d/%m/%Y')
            else:
                quoting_date = str(quoting_date_override)
        else:
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
    # Forward parsing (robust, missing spot rows)
    # -------------------------------
    def _parse_forwards(self, email_text: str) -> pd.DataFrame:
        """Parse VCB Forward Exchange Rates (handles missing spot on later rows)"""
        out_cols = self.get_standard_columns()['forward']

        root = re.split(r"(?i)forward\s+exchange\s+rates", email_text, maxsplit=1)
        if len(root) < 2:
            return pd.DataFrame(columns=out_cols)
        tail = root[1]

        # Split Bid/Ask with optional colon (':' or '：' or none)
        bid_m = re.search(r"(?i)\bBid\s*Price\b[:：]?", tail)
        ask_m = re.search(r"(?i)\bAsk\s*Price\b[:：]?", tail)

        if not bid_m:
            return pd.DataFrame(columns=out_cols)

        if ask_m:
            bid_text = tail[bid_m.end():ask_m.start()]
            ask_text = tail[ask_m.end():]
        else:
            bid_text = tail[bid_m.end():]
            ask_text = ""

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
        """
        Parse VCB forward side where only the first row has Spot.
        We tokenize by lines and carry the last seen spot to subsequent rows.
        """
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

        # Find where data starts (first dd/mm/yyyy)
        data_start = -1
        for i, line in enumerate(lines):
            if re.match(self.DATE_DMY, line):
                data_start = i
                break
        if data_start == -1:
            return []

        data_lines = lines[data_start:]
        rows = []
        current_spot = None
        i = 0

        while i < len(data_lines):
            if not re.match(self.DATE_DMY, data_lines[i]):
                i += 1
                continue

            # Need at least Value date and then [Spot?]/Term/Fwd
            if i + 2 >= len(data_lines):
                break

            trd_s = data_lines[i]
            val_s = data_lines[i + 1]

            # Next token can be Spot (26.090) or Term ("1M ()")
            next_tok = data_lines[i + 2]

            # patterns
            spot_re = re.compile(r"\b\d{2}\.\d{3}\b")
            term_re = re.compile(r"(\d+)\s*([DMWY])\s*\(\s*\)")

            if spot_re.match(next_tok):
                # row has spot
                current_spot = self._to_vcb_int(next_tok)
                term_idx = i + 3
                fwd_idx  = i + 4
                next_i   = i + 5
            else:
                # missing spot -> reuse last
                term_idx = i + 2
                fwd_idx  = i + 3
                next_i   = i + 4

            # bounds
            if fwd_idx >= len(data_lines):
                break

            term_s = data_lines[term_idx]
            fwd_s  = data_lines[fwd_idx]

            if not term_re.match(term_s):
                i = next_i
                continue
            mterm = term_re.match(term_s)
            termnum = int(mterm.group(1))
            termunit = mterm.group(2).upper()

            # forward must be 26.242 format
            if not spot_re.match(fwd_s):
                i = next_i
                continue
            fwd = self._to_vcb_int(fwd_s)

            trd = self._to_date(trd_s)
            val = self._to_date(val_s)
            if val < trd:
                trd, val = val, trd

            term_days = self._days(trd, val)
            term_lookup = round(self._yearfrac_30360_us(trd, val) * 12)

            # VCB thường không cung cấp Gap(%): ước tính non-annualized để bạn tham khảo
            gap_pct = ((fwd - current_spot) / current_spot * 100) if current_spot else 0
            gap_pct = round(gap_pct, 2)
            rows.append({
                "Bid/Ask": side,
                "Bank": self.bank_name,
                "Quoting date": trd,   # date type
                "Trading date": trd,
                "Value date": val,
                "Spot Exchange rate": current_spot,
                "Gap(%)": gap_pct,     # ví dụ 0.45 (không phải 0.0045)
                "Forward Exchange rate": fwd,
                "Term (days)": term_days,
                "% forward (cal)": None,
                "Diff.": None,
                "Term (lookup)": term_lookup
            })

            i = next_i

        return rows

    # -------------------------------
    # Central Bank stub (with quoting-date override)
    # -------------------------------
    def _build_central_bank(self, email_text: str = "", quoting_date_override=None) -> pd.DataFrame:
        out_cols = self.get_standard_columns()['central']
        if quoting_date_override:
            if hasattr(quoting_date_override, 'strftime'):
                qd = quoting_date_override.strftime('%d/%m/%Y')
            else:
                qd = str(quoting_date_override)
        else:
            qd = self._first_date(email_text) or ""
        return pd.DataFrame([{
            "No.": 1,
            "Bank": self.bank_name,
            "Quoting date": qd,
            "Central Bank Rate": None
        }], columns=out_cols)
