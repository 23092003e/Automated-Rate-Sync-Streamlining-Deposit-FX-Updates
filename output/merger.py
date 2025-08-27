# -*- coding: utf-8 -*-
"""
Output Merger - Combines results from all banks into single Excel output
Maintains consistent 3-sheet structure (Forward/Spot/CentralBank)
"""

import pandas as pd
from typing import List, Tuple, Dict, Any
from datetime import datetime


class OutputMerger:
    """Merges bank results into single consolidated output"""
    
    def __init__(self):
        self.forward_data = []
        self.spot_data = []
        self.central_data = []
    
    def add_bank_results(self, bank_name: str, forward_df: pd.DataFrame, 
                        spot_df: pd.DataFrame, central_df: pd.DataFrame) -> None:
        """Add results from a single bank"""
        if not forward_df.empty:
            self.forward_data.append(forward_df)
        if not spot_df.empty:
            self.spot_data.append(spot_df)
        if not central_df.empty:
            self.central_data.append(central_df)
    
    def merge_all_results(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Merge all bank results into consolidated DataFrames"""
        
        # Merge Forward data
        if self.forward_data:
            merged_forward = pd.concat(self.forward_data, ignore_index=True)
            # Re-number sequentially
            merged_forward['No.'] = range(1, len(merged_forward) + 1)
        else:
            merged_forward = pd.DataFrame(columns=[
                "No.", "Bid/Ask", "Bank", "Quoting date", "Trading date", "Value date",
                "Spot Exchange rate", "Gap(%)", "Forward Exchange rate", 
                "Term (days)", "% forward (cal)", "Diff.", "Term (lookup)"
            ])
        
        # Merge Spot data
        if self.spot_data:
            merged_spot = pd.concat(self.spot_data, ignore_index=True)
            merged_spot['No.'] = range(1, len(merged_spot) + 1)
        else:
            merged_spot = pd.DataFrame(columns=[
                "No.", "Bid/Ask", "Bank", "Quoting date",
                "Lowest rate of preceeding week", "Highest rate of preceeding week",
                "Closing rate of Friday (last week)"
            ])
        
        # Merge Central Bank data
        if self.central_data:
            merged_central = pd.concat(self.central_data, ignore_index=True)
            merged_central['No.'] = range(1, len(merged_central) + 1)
        else:
            merged_central = pd.DataFrame(columns=[
                "No.", "Quoting date", "Central Bank Rate"
            ])
        
        return merged_forward, merged_spot, merged_central
    
    def export_to_excel(self, output_path: str = "All_Banks_FX_Parsed.xlsx") -> None:
        """Export merged results to Excel with formulas"""
        forward_df, spot_df, central_df = self.merge_all_results()
        
        # Ensure date columns are properly formatted
        for c in ["Quoting date", "Trading date", "Value date"]:
            if c in forward_df.columns:
                forward_df[c] = pd.to_datetime(forward_df[c], dayfirst=True, errors="coerce").dt.date
        
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            forward_df.to_excel(writer, sheet_name="Forward", index=False)
            spot_df.to_excel(writer, sheet_name="Spot", index=False)
            central_df.to_excel(writer, sheet_name="CentralBank", index=False)
            
            # Add Excel formulas to Forward sheet
            if not forward_df.empty:
                self._add_excel_formulas(writer.book["Forward"], forward_df)
        
        print(f"SUCCESS: Merged results exported to: {output_path}")
        print(f"Summary:")
        print(f"  - Forward: {len(forward_df)} rows from {forward_df['Bank'].nunique() if 'Bank' in forward_df.columns and not forward_df.empty else 0} banks")
        print(f"  - Spot: {len(spot_df)} rows from {spot_df['Bank'].nunique() if 'Bank' in spot_df.columns and not spot_df.empty else 0} banks")
        print(f"  - Central: {len(central_df)} rows")
    
    def _add_excel_formulas(self, worksheet, forward_df: pd.DataFrame) -> None:
        """Add Excel formulas to Forward sheet"""
        headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
        col = {name: i+1 for i, name in enumerate(headers)}
        
        # Format date columns
        for r in range(2, worksheet.max_row + 1):
            for name in ["Quoting date", "Trading date", "Value date"]:
                if name in col:
                    worksheet.cell(row=r, column=col[name]).number_format = "dd/mm/yyyy"
        
        # Insert formulas
        for r in range(2, worksheet.max_row + 1):
            if all(name in col for name in ["Spot Exchange rate", "Forward Exchange rate", "Term (days)", "Gap(%)"]):
                c_spot = worksheet.cell(row=r, column=col["Spot Exchange rate"]).coordinate
                c_fwd = worksheet.cell(row=r, column=col["Forward Exchange rate"]).coordinate
                c_term = worksheet.cell(row=r, column=col["Term (days)"]).coordinate
                c_gap = worksheet.cell(row=r, column=col["Gap(%)"]).coordinate
                
                # % forward (cal) = ((Fwd - Spot) * 365) / (Spot * Term(days))
                worksheet.cell(row=r, column=col["% forward (cal)"]).value = (
                    f"=IFERROR(({c_fwd}-{c_spot})*365/({c_spot}*{c_term}),0)"
                )
                
                # Diff. = % forward (cal) - Gap(%)/100
                pct_cell = worksheet.cell(row=r, column=col["% forward (cal)"]).coordinate
                worksheet.cell(row=r, column=col["Diff."]).value = f"=IFERROR({pct_cell}-{c_gap}/100,0)"
            
            if all(name in col for name in ["Trading date", "Value date"]):
                c_trd = worksheet.cell(row=r, column=col["Trading date"]).coordinate
                c_val = worksheet.cell(row=r, column=col["Value date"]).coordinate
                
                # Term (lookup) = ROUND(YEARFRAC(Trading, Value)*12,0)
                worksheet.cell(row=r, column=col["Term (lookup)"]).value = (
                    f"=ROUND(YEARFRAC({c_trd},{c_val})*12,0)"
                )
        
        # Number formats
        for r in range(2, worksheet.max_row + 1):
            if "% forward (cal)" in col:
                worksheet.cell(row=r, column=col["% forward (cal)"]).number_format = "0.000%"
            if "Diff." in col:
                worksheet.cell(row=r, column=col["Diff."]).number_format = "0.000%"