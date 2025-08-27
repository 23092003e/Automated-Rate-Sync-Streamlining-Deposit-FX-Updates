# -*- coding: utf-8 -*-
"""
Demo script - Test the multi-bank system with ACB
"""

from banks.acb.processor import ACBProcessor
from output.merger import OutputMerger

# Sample ACB email content (từ ACB.py gốc)
ACB_EMAIL_SAMPLE = """
[This email is from an EXTERNAL source. Please use caution when opening attachments, clicking links, or responding]      Hi Thuan,    Please find updated rates as follow:    1.	Update Spot and Forward Exchange Rates:    Spot Exchange Rates:     USD/VND    Bid Price    Ask price    Lowest rate of the preceding week              Highest rate of the preceding week              Closing rate of Friday (last week)    26370    26380         Forward Exchange Rates:     Bid Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate    25/08/2025    24/09/2025    26303    1M ( )    1.11%    26,327    25/08/2025    26/11/2025    3M ( )    1.25%    26,387    25/08/2025    24/02/2026    6M ( )    1.37%    26,484    25/08/2025    22/05/2026    9M ( )    1.48%    26,591    25/08/2025    20/08/2026    12M ( )    1.49%    26,689         Ask Price:    Trading date    Value date    Spot Exchange rate    Term (days)    Gap(%)    Forward Exchange rate    25/08/2025    24/09/2025    26307    1M ( )    1.66%    26,343    25/08/2025    26/11/2025    3M ( )    1.66%    26,418    25/08/2025    24/02/2026    6M ( )    1.73%    26,535    25/08/2025    22/05/2026    9M ( )    1.87%    26,606    25/08/2025    20/08/2026    12M ( )    1.87%    26,622         Should  you need any further assistance, please dont hesitate to contact me.    Have a great week!    Cheers,         (Bella) Duong Bui (Ms.)
"""

def demo_acb_processing():
    """Demo ACB processing"""
    print("🚀 Demo: ACB Multi-Bank Processing")
    print("-" * 50)
    
    # 1. Initialize ACB processor
    acb_processor = ACBProcessor()
    
    # 2. Process ACB email
    print("📧 Processing ACB email...")
    forward_df, spot_df, central_df = acb_processor.parse_email(ACB_EMAIL_SAMPLE)
    
    print(f"✅ ACB Results:")
    print(f"  - Forward: {len(forward_df)} rows")
    print(f"  - Spot: {len(spot_df)} rows")
    print(f"  - Central: {len(central_df)} rows")
    
    # 3. Use output merger
    print("\n📊 Merging results...")
    merger = OutputMerger()
    merger.add_bank_results("ACB", forward_df, spot_df, central_df)
    
    # 4. Export to Excel
    output_file = "Demo_ACB_Output.xlsx"
    merger.export_to_excel(output_file)
    
    print(f"\n🎉 Demo completed! Check: {output_file}")
    
    # 5. Show sample data
    print("\n📋 Sample Forward Data:")
    if not forward_df.empty:
        print(forward_df[['Bank', 'Bid/Ask', 'Trading date', 'Value date', 'Forward Exchange rate']].head())
    
    print("\n📋 Sample Spot Data:")
    if not spot_df.empty:
        print(spot_df[['Bank', 'Bid/Ask', 'Closing rate of Friday (last week)']].head())

if __name__ == "__main__":
    demo_acb_processing()