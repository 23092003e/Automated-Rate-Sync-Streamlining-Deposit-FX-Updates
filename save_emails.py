# -*- coding: utf-8 -*-
"""
Save email contents to individual files for analysis
"""

from utils.email_extractor import EmailExtractor
import os

def save_emails_to_files():
    extractor = EmailExtractor()
    emails = extractor.extract_all_from_folder('C:/Users/ADMIN/Downloads/W4_Aug_25')
    
    # Create output directory
    output_dir = 'email_samples'
    os.makedirs(output_dir, exist_ok=True)
    
    # Save each email to a file
    for bank, content in emails.items():
        filename = f"{output_dir}/{bank}_email.txt"
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Saved {bank} email to: {filename}")

if __name__ == "__main__":
    save_emails_to_files()