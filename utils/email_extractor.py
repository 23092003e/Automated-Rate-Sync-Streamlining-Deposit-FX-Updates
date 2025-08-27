# -*- coding: utf-8 -*-
"""
Email extraction utilities for MSG files
"""

import extract_msg
from pathlib import Path
from typing import Dict, List, Optional


class EmailExtractor:
    """Extract email content from MSG files"""
    
    def __init__(self):
        pass
    
    def extract_from_msg(self, msg_file_path: str) -> Optional[str]:
        """Extract email body text from MSG file"""
        try:
            msg_path = Path(msg_file_path)
            if not msg_path.exists() or msg_path.suffix.lower() != '.msg':
                print(f"ERROR: Invalid MSG file: {msg_file_path}")
                return None
            
            msg = extract_msg.Message(msg_file_path)
            email_body = msg.body or ""
            
            print(f"SUCCESS: Extracted email from: {msg_path.name}")
            return email_body
            
        except Exception as e:
            print(f"ERROR extracting {msg_file_path}: {str(e)}")
            return None
    
    def extract_all_from_folder(self, folder_path: str) -> Dict[str, str]:
        """Extract email content from all MSG files in folder"""
        folder = Path(folder_path)
        if not folder.exists() or not folder.is_dir():
            print(f"ERROR: Folder not found: {folder_path}")
            return {}
        
        results = {}
        msg_files = list(folder.glob("*.msg"))
        
        print(f"Found {len(msg_files)} MSG files in {folder_path}")
        
        for msg_file in msg_files:
            bank_name = msg_file.stem.upper()  # ACB.msg -> ACB
            email_content = self.extract_from_msg(str(msg_file))
            
            if email_content:
                results[bank_name] = email_content
        
        print(f"Successfully extracted {len(results)} emails")
        return results