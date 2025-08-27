# -*- coding: utf-8 -*-
"""
Multi-Bank FX Rate Processor
Main entry point that processes all bank emails and generates consolidated output

Usage:
    python main.py

This will:
1. Extract emails from all MSG files in W4_Aug_25 folder
2. Process each bank using their specific parser
3. Merge all results into single Excel file with 3 sheets:
   - Forward: Forward exchange rates from all banks
   - Spot: Spot exchange rates from all banks  
   - CentralBank: Central bank rates from all banks
"""

import sys
from pathlib import Path
from typing import Dict, List

# Add project root to path
sys.path.append(str(Path(__file__).parent))

from utils.email_extractor import EmailExtractor
from banks.acb.processor import ACBProcessor
from banks.vcb.processor import VCBProcessor
from banks.bidv.processor import BIDVProcessor
from banks.tcb.processor import TCBProcessor
from output.merger import OutputMerger


class MultiBankProcessor:
    """Main processor that coordinates all bank processing"""
    
    def __init__(self):
        self.email_extractor = EmailExtractor()
        self.output_merger = OutputMerger()
        
        # Initialize bank processors
        self.processors = {
            'ACB': ACBProcessor(),
            'VCB': VCBProcessor(),
            'BIDV': BIDVProcessor(),
            'TCB': TCBProcessor(),
        }
        
        # Banks that will use generic/fallback processing for now
        self.pending_banks = ['KBANK', 'SC', 'UOB', 'UOBV', 'VIB', 'VTB', 'WOORI']
    
    def process_all_banks(self, msg_folder: str, output_file: str = "All_Banks_FX_Parsed.xlsx") -> None:
        """Process all bank MSG files and generate consolidated output"""
        
        print("Starting Multi-Bank FX Rate Processing...")
        print(f"Source folder: {msg_folder}")
        print(f"Output file: {output_file}")
        print("-" * 60)
        
        # Extract emails from MSG files
        emails = self.email_extractor.extract_all_from_folder(msg_folder)
        
        if not emails:
            print("❌ No emails extracted. Please check the MSG folder path.")
            return
        
        print(f"Found emails from {len(emails)} banks: {', '.join(emails.keys())}")
        print("-" * 60)
        
        # Process each bank
        processed_count = 0
        
        for bank_name, email_content in emails.items():
            print(f"Processing {bank_name}...")
            
            try:
                if bank_name in self.processors:
                    # Use specific processor
                    processor = self.processors[bank_name]
                    forward_df, spot_df, central_df = processor.parse_email(email_content)
                    
                    # Add results to merger
                    self.output_merger.add_bank_results(bank_name, forward_df, spot_df, central_df)
                    
                    print(f"  SUCCESS {bank_name}: Forward={len(forward_df)}, Spot={len(spot_df)}, Central={len(central_df)}")
                    processed_count += 1
                    
                else:
                    print(f"  PENDING {bank_name}: No specific processor implemented yet.")
                    
            except Exception as e:
                print(f"  ERROR {bank_name}: Error processing - {str(e)}")
        
        print("-" * 60)
        
        if processed_count > 0:
            # Export consolidated results
            print("Merging results from all banks...")
            self.output_merger.export_to_excel(output_file)
            print("-" * 60)
            print("Multi-Bank processing completed successfully!")
        else:
            print("ERROR: No banks were processed successfully.")
    
    def add_bank_processor(self, bank_name: str, processor):
        """Add a new bank processor (for future expansion)"""
        self.processors[bank_name.upper()] = processor
        if bank_name.upper() in self.pending_banks:
            self.pending_banks.remove(bank_name.upper())
    
    def list_supported_banks(self) -> None:
        """List currently supported banks"""
        print("Bank Processing Status:")
        print(f"Implemented: {', '.join(self.processors.keys())}")
        print(f"Pending: {', '.join(self.pending_banks)}")


def main():
    """Main entry point"""
    
    # Configuration
    MSG_FOLDER = r"C:\Users\ADMIN\Downloads\W4_Aug_25"
    OUTPUT_FILE = "All_Banks_FX_Parsed.xlsx"
    
    # Create and run processor
    processor = MultiBankProcessor()
    
    # Show current status
    processor.list_supported_banks()
    print()
    
    # Process all banks
    processor.process_all_banks(MSG_FOLDER, OUTPUT_FILE)


if __name__ == "__main__":
    main()