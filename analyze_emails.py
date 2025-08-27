# -*- coding: utf-8 -*-
"""
Analyze email formats from different banks
"""

from utils.email_extractor import EmailExtractor
import sys

def analyze_bank_emails():
    extractor = EmailExtractor()
    emails = extractor.extract_all_from_folder('C:/Users/ADMIN/Downloads/W4_Aug_25')
    
    # Analyze each bank's email format
    banks_to_analyze = ['VCB', 'BIDV', 'TCB', 'UOB', 'KBANK']
    
    for bank in banks_to_analyze:
        if bank in emails:
            print(f"\n=== {bank} EMAIL STRUCTURE ===")
            content = emails[bank]
            
            # Look for key sections
            sections = []
            if 'spot' in content.lower() or 'Spot' in content:
                sections.append("Has Spot section")
            if 'forward' in content.lower() or 'Forward' in content:
                sections.append("Has Forward section")
            if 'central' in content.lower() or 'Central' in content:
                sections.append("Has Central Bank section")
            
            # Check for rate patterns
            import re
            rates = re.findall(r'\b\d{5}\b', content)  # 5-digit rates like 26370
            dates = re.findall(r'\d{1,2}/\d{1,2}/\d{4}', content)  # Date patterns
            
            print(f"- Sections found: {', '.join(sections) if sections else 'None detected'}")
            print(f"- Rate numbers found: {len(rates)} (sample: {rates[:3] if rates else 'None'})")
            print(f"- Dates found: {len(dates)} (sample: {dates[:3] if dates else 'None'})")
            print(f"- Content length: {len(content)} characters")
            
            # Show first few lines (safely)
            lines = content.split('\n')[:10]
            print("- First few lines:")
            for i, line in enumerate(lines):
                try:
                    print(f"  {i+1}: {line.strip()[:80]}...")
                except UnicodeEncodeError:
                    print(f"  {i+1}: [Contains non-ASCII characters - {len(line)} chars]")

if __name__ == "__main__":
    analyze_bank_emails()