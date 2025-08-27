# Multi-Bank FX Rate Processing System

## 🏗️ Cấu trúc hệ thống

```
project/
├── main.py                     # Entry point chính
├── banks/                      # Bank processors
│   ├── base.py                # Base class cho tất cả banks
│   ├── acb/
│   │   ├── __init__.py
│   │   └── processor.py       # ACB-specific logic
│   ├── bidv/                  # (sẽ thêm sau)
│   ├── vcb/                   # (sẽ thêm sau)
│   └── ...                    # Các banks khác
├── output/
│   ├── __init__.py
│   └── merger.py              # Gộp kết quả từ tất cả banks
├── utils/
│   ├── __init__.py
│   └── email_extractor.py     # Extract emails từ MSG files
└── requirements.txt
```

## 🚀 Cách sử dụng

### 1. Cài đặt dependencies
```bash
pip install pandas openpyxl extract-msg python-dateutil
```

### 2. Chạy hệ thống
```bash
python main.py
```

### 3. Kết quả
- File output: `All_Banks_FX_Parsed.xlsx`
- 3 sheets: **Forward**, **Spot**, **CentralBank**
- Tất cả banks được gộp trong 1 file duy nhất

## 📊 Output Format

### Sheet "Forward"
| No. | Bid/Ask | Bank | Quoting date | Trading date | Value date | Spot Exchange rate | Gap(%) | Forward Exchange rate | Term (days) | % forward (cal) | Diff. | Term (lookup) |

### Sheet "Spot"
| No. | Bid/Ask | Bank | Quoting date | Lowest rate of preceeding week | Highest rate of preceeding week | Closing rate of Friday (last week) |

### Sheet "CentralBank"
| No. | Quoting date | Central Bank Rate |

## 🏦 Banks Status

### ✅ Implemented
- **ACB**: Hoàn thành

### ⏳ Pending (sẽ implement sau)
- BIDV, KBank, SC, TCB, UOB, UOBV, VCB, VIB, VTB, Woori

## 🔧 Thêm Bank mới

### 1. Tạo folder cho bank
```bash
mkdir banks/vcb
touch banks/vcb/__init__.py
```

### 2. Tạo processor
```python
# banks/vcb/processor.py
from ..base import BaseBankProcessor

class VCBProcessor(BaseBankProcessor):
    def __init__(self):
        super().__init__("VCB")
    
    def parse_email(self, email_text: str):
        # VCB-specific parsing logic
        forward_df = self._parse_forwards(email_text)
        spot_df = self._parse_spot(email_text)
        central_df = self._build_central_bank(email_text)
        return forward_df, spot_df, central_df
```

### 3. Đăng ký trong main.py
```python
from banks.vcb.processor import VCBProcessor

# Thêm vào __init__
self.processors = {
    'ACB': ACBProcessor(),
    'VCB': VCBProcessor(),  # <-- thêm dòng này
}
```

## 🎯 Lợi ích

### ✅ Scalable
- Dễ dàng thêm bank mới
- Mỗi bank có logic riêng
- Code không bị duplicate

### ✅ Maintainable  
- Cấu trúc rõ ràng
- Separation of concerns
- Base class cung cấp utilities chung

### ✅ Unified Output
- 1 file Excel duy nhất
- Cùng format 3 sheets
- Excel formulas tự động

### ✅ Production Ready
- Error handling
- Logging
- Progress tracking

## 🔄 Workflow

```
MSG Files (W4_Aug_25/) 
    ↓
Email Extraction
    ↓
Bank-specific Processing (ACB, VCB, BIDV, ...)
    ↓
Result Merging
    ↓
Excel Export (All_Banks_FX_Parsed.xlsx)
```

## 📝 Notes

- Hiện tại chỉ có ACB được implement đầy đủ
- Các banks khác sẽ được thêm dần theo template
- Base class cung cấp các utilities chung
- Output format đồng nhất cho tất cả banks