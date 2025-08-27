# -*- coding: utf-8 -*-
"""
Parse ACB FX email -> structured tables (Forward / Spot / Central Bank)
and export to Excel with your requested calculations.

What you get:
- Sheet "Forward": columns
  [No., Bid/Ask, Bank, Quoting date, Trading date, Value date,
   Spot Exchange rate, Gap(%), Forward Exchange rate, Term (days),
   % forward (cal), Diff., Term (lookup)]
  * % forward (cal) = ((Fwd - Spot) * 365) / (Spot * Term(days))  (Excel formula)
  * Diff. = % forward (cal) - Gap(%)                              (Excel formula)
  * Term (lookup) = ROUND(YEARFRAC(Trading, Value)*12,0)          (Excel formula)

- Sheet "Spot": columns
  [No., Bid/Ask, Bank, Quoting date,
   Lowest rate of preceeding week, Highest rate of preceeding week,
   Closing rate of Friday (last week)]

- Sheet "CentralBank": stub (email didn’t include SBV rate)
  [No., Quoting date, Central Bank Rate]

Dependencies:
    pip install pandas openpyxl python-dateutil
"""

import re
from datetime import datetime, date
import pandas as pd

# ==============================
# PUT YOUR EMAIL CONTENT HERE
# ==============================
EMAIL_TEXT = r"""
"[This email is from an EXTERNAL source. Please use caution when opening attachments, clicking links, or responding]      Hi Thuan,    Please find updated rates as follow:    1.	Update Spot and Forward Exchange Rates:    Spot Exchange Rates:     USD/VND    Bid Price    Ask price    Lowest rate of the preceding week              Highest rate of the preceding week              Closing rate of Friday (last week)    26370    26380         Forward Exchange Rates:     Bid Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate    25/08/2025    24/09/2025    26303    1M ( )    1.11%    26,327    25/08/2025    26/11/2025    3M ( )    1.25%    26,387    25/08/2025    24/02/2026    6M ( )    1.37%    26,484    25/08/2025    22/05/2026    9M ( )    1.48%    26,591    25/08/2025    20/08/2026    12M ( )    1.49%    26,689         Ask Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate    25/08/2025    24/09/2025    26307    1M ( )    1.66%    26,343    25/08/2025    26/11/2025    3M ( )    1.66%    26,418    25/08/2025    24/02/2026    6M ( )    1.73%    26,535    25/08/2025    22/05/2026    9M ( )    1.87%    26,606    25/08/2025    20/08/2026    12M ( )    1.87%    26,622         Should  you need any further assistance, please dont hesitate to contact me.    Have a great week!    Cheers,         (Bella) Duong Bui (Ms.)     Financial Markets Direct Sales Manager    Financial Markets Division         ASIA COMMERCIAL BANK (ACB)    2nd floor, ACB Tower C, 2/20 Cao Thang, Dist 3, Ho Chi Minh City, Vietnam     Mobile: + 84 (0) 938 921 389     Email: duongbt@acb.com.vn <mailto:duongbt@acb.com.vn>      Website: www.acb.com.vn <http://www.acb.com.vn/>           From: Thuan, Tran Duc (Capital Market) <thuan.td@bwidjsc.com>   Sent: Monday, August 25, 2025 7:37 AM  To: Bui Thuy Duong <duongbt@acb.com.vn>  Subject: BWID - Request for Deposit Rate and Exchange Rate              CHÚ Ý : Email này được gửi từ bên ngoài ACB. Vui lòng không truy cập vào các liên kết (URL) hoặc mở các tập tin đính kèm trừ khi bạn nhận biết người gửi và nội dung thư là an toàn./CAUTION : This email is originated from outside of ACB. Do not click links (URL) or open attachments unless you recognize the sender and know content is safe.    ________________________________    Dear Ms. Thi and Ms. Duong,         On behalf of the CM team at BWID, I would kindly request the deposit rates and the exchange rates applied for the BWID Group as follows:         1.	VND Deposit Rate:    Please kindly update the VND deposit rate for BWID Group as of today.     Note: On flexible withdrawal condition.    Tenor    Current Rate    1M    3M    6M    9M    12M    Interest (%/year)                                       Note: Please kindly announce any changes in deposit rates and update weekly, specifically on Mondays. If deposit rates are unchanged, please announce in advance.         2.	Update Spot and Forward Exchange Rates:    Spot Exchange Rates:     USD/VND    Bid Price    Ask price    Lowest rate of the preceding week              Highest rate of the preceding week              Closing rate of Friday (last week)                   Forward Exchange Rates:     Bid Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate                   1M ( )                        3M ( )                        6M ( )                        9M ( )                        12M ( )                   Ask Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate                   1M ( )                        3M ( )                        6M ( )                        9M ( )                        12M ( )                   Note: Please kindly fill in all the blanks above. For any inquiries regarding the information required, please feel free to contact the sender.         Best regards,    Disclaimer:  The information contained in this communication and attachment is confidential and is for the use of the intended recipient only. Any disclosure, copying or distribution of this communication without the sender's consent is strictly prohibited. If you are not the intended recipient, please notify the sender and delete this communication entirely without using, retaining, or disclosing any of its contents. This communication is for information purposes only and shall not be construed as an offer or solicitation of an offer or an acceptance or a confirmation of any contract or transaction. All data or other information contained herein are not warranted to be complete and accurate and are subject to change without notice. Any comments or statements made herein do not necessarily reflect those of BW Industrial Development JSC or any of its affiliates. Internet communications cannot be guaranteed to be virus-free. The recipient is responsible for ensuring that this communication is virus free and the sender accepts no liability for any damages caused by virus transmitted by this communication.           <https://mail.acb.com.vn/owa/auth/acb/footer-emai.png>     Email chứa thông tin dành riêng cho người nhận, có thể là thông tin mật. Nếu bạn không phải là người nhận mong muốn, bạn không nên xem, sử dụng, phổ biến, phân phối, sao chép email hoặc tập tin đính kèm theo email này. Bạn nên xóa nội dung này và thông báo cho người gửi. ACB chịu trách nhiệm về các chuẩn mực chất lượng dịch vụ, thông tin sản phẩm dịch vụ và quy tắc ứng xử được công bố. Ngoài các nội dung nêu trên, email này có thể chứa các quan điểm cá nhân và ý kiến của người gửi hoặc tác giả, không phải là quan điểm và ý kiến của ACB hoặc các công ty trực thuộc. Do đó, ACB không chịu trách nhiệm đối với bất kỳ khiếu nại, tổn thất hoặc thiệt hại nào phát sinh liên quan đến email này. Để biết thêm thông tin chi tiết, tham khảo tại website ACB.    "
"""

# -------------------------------
# Helpers
# -------------------------------
DATE_DMY = r"(?:0[1-9]|[12]\d|3[01])/(?:0[1-9]|1[0-2])/(?:19|20)\d\d"

def _to_int(s):
    return int(str(s).replace(',', '')) if s is not None and str(s).strip() else None

def _to_date(dmy: str) -> date:
    return datetime.strptime(dmy, "%d/%m/%Y").date()

def _days(d1: date, d2: date) -> int:
    return (d2 - d1).days

def _yearfrac_30360_us(d1: date, d2: date) -> float:
    """Excel YEARFRAC basis=0 approximation for Term(lookup)."""
    d1d = 30 if d1.day == 31 else d1.day
    d2d = 30 if (d2.day == 31 and d1d in (30, 31)) else d2.day
    months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
    days = d2d - d1d
    return (months * 30 + days) / 360.0

def _first_date(text: str) -> str | None:
    m = re.search(DATE_DMY, text)
    return m.group(0) if m else None

# -------------------------------
# Spot section
# -------------------------------
def parse_spot(email_text: str) -> pd.DataFrame:
    """
    Heuristics:
      - Find the 'Spot Exchange Rates' block.
      - Take the first two 5-digit numbers after that as
        [Lowest, Highest]; set Closing = Highest (matches your example).
      - Create one Bid and one Ask row.
    """
    out_cols = [
        "No.", "Bid/Ask", "Bank", "Quoting date",
        "Lowest rate of preceeding week",
        "Highest rate of preceeding week",
        "Closing rate of Friday (last week)"
    ]

    parts = re.split(r"(?i)spot\s+exchange\s+rates", email_text, maxsplit=1)
    if len(parts) < 2:
        return pd.DataFrame(columns=out_cols)

    tail = parts[1]
    nums = re.findall(r"\b\d{5}\b", tail)
    nums = list(dict.fromkeys(nums))  # dedupe keep-order

    low = _to_int(nums[0]) if len(nums) >= 1 else None
    high = _to_int(nums[1]) if len(nums) >= 2 else low
    close = high

    quoting_date = _first_date(email_text) or ""

    rows = []
    for i, side in enumerate(["Bid", "Ask"], start=1):
        rows.append({
            "No.": 10 + i,  # continue numbering after forwards (1..10)
            "Bid/Ask": side,
            "Bank": "ACB",
            "Quoting date": quoting_date,
            "Lowest rate of preceeding week": low,
            "Highest rate of preceeding week": high,
            "Closing rate of Friday (last week)": close,
        })
    return pd.DataFrame(rows, columns=out_cols)

# -------------------------------
# Forward section (Bid / Ask)
# -------------------------------
ROW_RE = re.compile(
    rf"\s*(?P<trd>{DATE_DMY})\s+"
    rf"(?P<val>{DATE_DMY})\s+"
    rf"(?:(?P<spot>\d{{5}})\s+)?"
    rf"(?P<termnum>\d+)\s*(?P<termunit>[DMWY])?\s*\(\s*\)\s+"
    rf"(?P<gap>-?\d+(?:\.\d+)?)%\s+"
    rf"(?P<fwd>[\d,]+)\s*",
    flags=re.IGNORECASE
)

def _parse_forward_side(text: str, side: str) -> list[dict]:
    """
    Parse a single side (Bid or Ask). Carries forward the latest spot within that side.
    Ensures Trading date <= Value date (swaps if needed).
    """
    rows = []
    current_spot = None

    for m in ROW_RE.finditer(text):
        trd_s = m.group("trd")
        val_s = m.group("val")
        spot_s = m.group("spot")
        termnum = int(m.group("termnum"))
        termunit = (m.group("termunit") or "M").upper()  # default months
        gap_pct = float(m.group("gap"))                 # keep as 1.11 (not 0.0111)
        fwd = _to_int(m.group("fwd"))

        # update / carry side-specific spot
        if spot_s:
            current_spot = _to_int(spot_s)
        if current_spot is None:
            prev = re.findall(r"\b\d{5}\b", text[:m.start()])
            current_spot = _to_int(prev[-1]) if prev else None

        trd = _to_date(trd_s)
        val = _to_date(val_s)

        # Guarantee non-negative days: if out-of-order, swap
        if val < trd:
            trd, val = val, trd
            trd_s, val_s = trd.strftime("%d/%m/%Y"), val.strftime("%d/%m/%Y")

        term_days = _days(trd, val)
        term_lookup = round(_yearfrac_30360_us(trd, val) * 12)

        rows.append({
            "Bid/Ask": side,
            "Bank": "ACB",
            "Quoting date": trd,         # write as date type
            "Trading date": trd,
            "Value date": val,
            "Spot Exchange rate": current_spot,
            "Gap(%)": gap_pct,           # e.g., 1.11
            "Forward Exchange rate": fwd,
            "Term (days)": term_days,
            "% forward (cal)": None,     # will be Excel formula
            "Diff.": None,               # will be Excel formula
            "Term (lookup)": term_lookup # numeric (could also be an Excel formula)
        })
    return rows

def parse_forwards(email_text: str) -> pd.DataFrame:
    out_cols = [
        "No.","Bid/Ask","Bank","Quoting date","Trading date","Value date",
        "Spot Exchange rate","Gap(%)","Forward Exchange rate",
        "Term (days)","% forward (cal)","Diff.","Term (lookup)"
    ]

    root = re.split(r"(?i)forward\s+exchange\s+rates", email_text, maxsplit=1)
    if len(root) < 2:
        return pd.DataFrame(columns=out_cols)
    tail = root[1]

    # First, split by "Bid Price:" to get everything after it
    bid_parts = re.split(r"(?i)\bBid\s*Price\s*:", tail, maxsplit=1)
    if len(bid_parts) < 2:
        return pd.DataFrame(columns=out_cols)
    
    # Now split the remaining text by "Ask Price:" to separate Bid and Ask sections
    after_bid = bid_parts[1]
    ask_parts = re.split(r"(?i)\bAsk\s*Price\s*:", after_bid, maxsplit=1)
    
    bid_text = ask_parts[0]  # Everything between "Bid Price:" and "Ask Price:"
    ask_text = ask_parts[1] if len(ask_parts) > 1 else ""

    rows = []
    rows += _parse_forward_side(bid_text, "Bid")
    rows += _parse_forward_side(ask_text, "Ask")

    df = pd.DataFrame(rows, columns=[
        "Bid/Ask","Bank","Quoting date","Trading date","Value date",
        "Spot Exchange rate","Gap(%)","Forward Exchange rate",
        "Term (days)","% forward (cal)","Diff.","Term (lookup)"
    ])
    if df.empty:
        df.insert(0, "No.", [])
        return df

    # order: Bid first, then by Trading date, then by Term(days)
    df = df.sort_values(["Bid/Ask","Trading date","Term (days)"]).reset_index(drop=True)
    df.insert(0, "No.", range(1, len(df) + 1))
    return df

# -------------------------------
# Central Bank rate (stub)
# -------------------------------
def build_central_bank(email_text: str) -> pd.DataFrame:
    qd = _first_date(email_text) or ""
    return pd.DataFrame([{
        "No.": 1,
        "Quoting date": qd,
        "Central Bank Rate": None
    }])

# -------------------------------
# Export to Excel (with formulas)
# -------------------------------
def to_excel_with_formulas(df_forward, df_spot, df_central, out_path="ACB_FX_Parsed.xlsx"):
    """
    Writes 3 sheets and injects Excel formulas:
      - % forward (cal) = ((Fwd - Spot) * 365) / (Spot * Term(days))
      - Diff.           = % forward (cal) - Gap(%)/100
      - Term (lookup)   = ROUND(YEARFRAC(Trading, Value)*12,0)   (kept numeric too)
    Ensures date cells are real dates (no #VALUE! in YEARFRAC).
    """
    # ensure date dtype
    for c in ["Quoting date","Trading date","Value date"]:
        if c in df_forward.columns:
            df_forward[c] = pd.to_datetime(df_forward[c], dayfirst=True, errors="coerce").dt.date

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_forward.to_excel(writer, sheet_name="Forward", index=False)
        df_spot.to_excel(writer, sheet_name="Spot", index=False)
        df_central.to_excel(writer, sheet_name="CentralBank", index=False)

        ws = writer.book["Forward"]
        headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
        col = {name: i+1 for i, name in enumerate(headers)}

        # Format date columns
        for r in range(2, ws.max_row + 1):
            for name in ["Quoting date","Trading date","Value date"]:
                ws.cell(row=r, column=col[name]).number_format = "dd/mm/yyyy"

        # Insert formulas
        for r in range(2, ws.max_row + 1):
            c_spot = ws.cell(row=r, column=col["Spot Exchange rate"]).coordinate
            c_fwd  = ws.cell(row=r, column=col["Forward Exchange rate"]).coordinate
            c_term = ws.cell(row=r, column=col["Term (days)"]).coordinate
            c_gap  = ws.cell(row=r, column=col["Gap(%)"]).coordinate
            c_trd  = ws.cell(row=r, column=col["Trading date"]).coordinate
            c_val  = ws.cell(row=r, column=col["Value date"]).coordinate

            # % forward (cal)
            ws.cell(row=r, column=col["% forward (cal)"]).value = (
                f"=IFERROR(({c_fwd}-{c_spot})*365/({c_spot}*{c_term}),0)"
            )
            # Diff. (Gap is in percent units like 1.11, so divide by 100)
            pct_cell = ws.cell(row=r, column=col["% forward (cal)"]).coordinate
            ws.cell(row=r, column=col["Diff."]).value = f"=IFERROR({pct_cell}-{c_gap}/100,0)"
            # Term (lookup)
            ws.cell(row=r, column=col["Term (lookup)"]).value = f"=ROUND(YEARFRAC({c_trd},{c_val})*12,0)"

        # Number formats
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=col["% forward (cal)"]).number_format = "0.000%"
            ws.cell(row=r, column=col["Diff."]).number_format = "0.000%"

# -------------------------------
# Main
# -------------------------------
def parse_acb_email(email_text: str):
    df_forward = parse_forwards(email_text)
    df_spot = parse_spot(email_text)
    df_central = build_central_bank(email_text)
    return df_forward, df_spot, df_central

if __name__ == "__main__":
    forward_df, spot_df, central_df = parse_acb_email(EMAIL_TEXT)

    # Show in console
    pd.set_option("display.max_columns", None)
    print("\nFORWARD:")
    print(forward_df)
    print("\nSPOT:")
    print(spot_df)
    print("\nCENTRAL BANK:")
    print(central_df)

    # Export
    to_excel_with_formulas(forward_df, spot_df, central_df, out_path="ACB_FX_Parsed.xlsx")
    print("\nExported -> ACB_FX_Parsed.xlsx")
