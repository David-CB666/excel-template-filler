---
name: excel-template-filler
description: "Excel 模板填充與批量生成系統（融合版）。支持雙引擎架構：openpyxl（無圖片模板）+ ZIP 引擎（完美保留圖片/打印設置）。觸發詞：批量生成報批表、模板填充、佔位符替換、Excel 模板、材料報批、EL 報批表、BQ 頁合併、含圖片模板。核心能力：(1) 自動引擎選擇 (2) 圖片完美保留 (3) 模板填充 (4) BQ 頁合併 (5) 自動鏈接 (6) 數據驗證。"
version: 2.0.0
icon: 📊
license: MIT
metadata:
  clawdbot:
    requires:
      bins:
        - python
        - uv
    commands:
      fill: uv run python {baseDir}/scripts/template_filler.py
      merge: uv run python {baseDir}/scripts/exporters/bq_merger.py
---

# Excel Template Filler - Excel 模板填充與批量生成系統

> 版本：v2.0 融合版（2026-06-01）
> 來源：從 VBA 工程總表宏系統內化 + dem-material-approval 融合

---

## 📚 快速索引

**执行前先读取对应模块文档：**

- **引擎选择与使用** → `references/engines.md`
  - openpyxl 引擎（无图片模板）
  - ZIP 引擎（含图片模板）★核心
  - 自动引擎选择机制
- **BQ 页合并器** → `references/bq-merger.md`
  - 报批表 PDF + BQ 页合并流程
  - 匹配逻辑与参数说明
  - 常见问题与解决方案
- **API 文档** → 见下方详细说明
- **示例代码** → `examples/` 目录

---

## 🎯 v2.0 融合版新特性

### 雙引擎架構

本技能融合了 `excel-template-filler` 和 `dem-material-approval` 兩個技能，實現了**雙引擎架構**：

| 引擎 | 適用場景 | 核心優勢 |
|:---|:---|:---|
| **openpyxl 引擎** | 無圖片模板 | 代碼簡潔、易於維護、純 Python |
| **ZIP 引擎** | 含圖片模板 | 完美保留圖片、打印設置、二進制資源 |

**自動選擇機制：** 根據模板特徵自動檢測並選擇最優引擎，用戶無需手動指定。

### 新增功能

| 功能 | 來源 | 說明 |
|:---|:---|:---|
| **ZIP 引擎** | dem-material-approval | 完美保留圖片、打印設置 |
| **BQ 頁合併** | dem-material-approval | 報批表 PDF + BQ 頁自動合併 |
| **自動引擎選擇** | 新增 | 檢測模板類型，自動選擇引擎 |
| **統一 API** | 新增 | 一個入口覆蓋所有場景 |

---

## 📋 概述

這是一個從 VBA 系統內化並融合 dem-material-approval 的 Python 技能，專為工程材料報批流程設計。核心功能包括：

1. **自動引擎選擇：** 根據模板特徵自動選擇 openpyxl 或 ZIP 引擎
2. **圖片完美保留：** ZIP 引擎完美保留圖片、打印設置、二進制資源
3. **模板填充：** 掃描佔位符 `{{字段名}}` 或 `{字段名}` 並批量替換
4. **BQ 頁合併：** 報批表 PDF 與 BQ 頁自動合併
5. **自動鏈接：** 自動匹配文件並創建超鏈接
6. **數據驗證：** 檢查數據源完整性

---

## 🚀 快速開始

### 1. 統一入口（推薦）

```python
from scripts.template_filler import TemplateFiller

# 初始化（自動檢測引擎）
filler = TemplateFiller(
    data_source="材料審批總表.xlsx",
    template="報批表模板.xlsx"  # 含圖片，自動選擇 ZIP 引擎
)

# 掃描佔位符
placeholders = filler.scan_placeholders()
print(f"發現佔位符: {placeholders}")

# 檢測引擎類型
engine_name = "ZIP" if filler.has_images() else "openpyxl"
print(f"使用引擎: {engine_name}")

# 填充並導出
output_files = filler.fill_and_export(
    start_row=2,
    end_row=11,
    export_format="pdf",
    output_dir="./output",
    field_map={
        "{編號}": "編號",
        "{材料}": "材料",
        "{品牌}": "品牌"
    }
)
```

### 2. BQ 頁合併

```python
from scripts.exporters.bq_merger import BQMerger

# 初始化
merger = BQMerger()

# 載入 BQ PDF
merger.load_bq_pdf("BQ標書.pdf")

# 載入總表
merger.load_zongbiao("材料審批總表.xlsx")

# 匹配 BQ 頁碼
success, failed = merger.match_bq_pages()

# 合併 PDF
merger.merge_pdfs("./報批表PDF", "./最終輸出")
```

### 3. 自動鏈接文件

```python
from scripts.auto_linker import AutoLinker

# 初始化
linker = AutoLinker(
    workbook="工程總表.xlsx",
    target_column=1,  # A 列
    search_root="./files"
)

# 執行自動鏈接
linker.convert_to_hyperlinks(
    extensions=[".pdf", ".docx", ".xlsx"],
    include_subfolders=True
)
```

---

## 📁 文件結構

```
excel-template-filler/
├── SKILL.md                    # 本文件
├── README.md                   # GitHub 展示頁
├── scripts/
│   ├── template_filler.py      # 統一入口（自動選擇引擎）
│   ├── engines/
│   │   ├── base_engine.py      # 引擎基類
│   │   ├── openpyxl_engine.py  # openpyxl 引擎（無圖片模板）
│   │   └── zip_engine.py       # ZIP 引擎（含圖片模板）★核心
│   ├── exporters/
│   │   └── bq_merger.py        # BQ 頁合併器
│   ├── auto_linker.py          # 自動鏈接器
│   ├── file_grabber.py         # 文件名抓取器
│   └── utils.py                # 工具函數
├── examples/
│   ├── example_basic.py        # 基礎示例
│   ├── example_batch_pdf.py    # 批量 PDF 示例
│   └── example_auto_link.py    # 自動鏈接示例
└── tests/
    ├── test_openpyxl_engine.py # openpyxl 引擎測試
    ├── test_zip_engine.py      # ZIP 引擎測試
    └── test_bq_merger.py       # BQ 合併測試
```

---

## 🔧 核心功能詳解

### 1. 統一入口（TemplateFiller）

**自動引擎選擇邏輯：**

```python
def detect_template_type(template_path: str) -> str:
    """
    檢測模板類型
    
    Returns:
        "zip" - 含圖片/打印設置/繪圖
        "openpyxl" - 無圖片
    """
    import zipfile
    
    with zipfile.ZipFile(template_path) as zf:
        names = zf.namelist()
        
        # 檢查是否包含圖片
        has_images = any(name.startswith('xl/media/') for name in names)
        
        # 檢查是否包含打印設置
        has_printer = any('printerSettings' in name for name in names)
        
        # 檢查是否包含 DrawingML
        has_drawings = any('drawings' in name for name in names)
        
        return "zip" if (has_images or has_printer or has_drawings) else "openpyxl"
```

**API：**

```python
class TemplateFiller:
    def __init__(
        self,
        data_source: str = None,
        template: str = None,
        engine_type: str = "auto"  # "auto" / "openpyxl" / "zip"
    ):
        """
        初始化模板填充器
        
        Args:
            data_source: 數據源 Excel 路徑
            template: 模板 Excel 路徑
            engine_type: 引擎類型
                "auto" - 自動檢測（推薦）
                "openpyxl" - 強制使用 openpyxl
                "zip" - 強制使用 ZIP
        """
    
    def has_images(self) -> bool:
        """檢測模板是否含圖片"""
    
    def scan_placeholders(self) -> List[str]:
        """掃描模板中的所有佔位符"""
    
    def validate_data(self) -> Dict[str, Any]:
        """驗證數據源完整性"""
    
    def fill_and_export(
        self,
        start_row: int = None,
        end_row: int = None,
        export_format: str = "excel",  # "excel" / "pdf" / "both"
        output_dir: str = "./output",
        field_map: Dict[str, str] = None,
        **kwargs
    ) -> List[str]:
        """
        填充模板並導出
        
        Returns:
            導出的文件路徑列表
        """
```

### 2. ZIP 引擎（zip_engine.py）★核心

**為什麼需要 ZIP 引擎？**

openpyxl 的 `copy_worksheet()` 和 `save()` 會：
- ❌ 丟失 `printerSettings.bin`（打印設置）
- ❌ 丟失 DrawingML 關聯（圖片）
- ❌ 破壞 workbook.xml.rels 的 rId 映射

**ZIP 引擎解決方案：**

```
模板.xlsx（ZIP）
  → 讀取全部文件到內存
  → 修改 worksheet XML（t="s" → inlineStr，保留 s 屬性）
  → 生成新的 sheet2~sheetN
  → 更新 workbook.xml / rels / Content_Types
  → 直接寫入新 ZIP（保留所有原始資源）
```

**核心算法：共享字符串 → inlineStr 轉換**

```python
def build_inline_sheet(sheet1_xml, ss_entries, data, field_map):
    """
    1. 從 sharedStrings 提取所有佔位符的實際值
    2. 將 t="s" 單元格轉為 inlineStr，同時保留 s="N" 樣式索引
    """
    # 佔位符 → 實際值
    ph_to_val = {ph: esc(data.get(fk, '')) for ph, fk in field_map.items()}
    
    # 構建替換後的 sharedStrings 純文本列表
    new_ss = []
    for content in ss_entries:
        result = content
        for ph, val in ph_to_val.items():
            result = result.replace(ph, val)
        texts = re.findall(r'<t[^>]*>([^<]*)</t>', result)
        new_ss.append(''.join(texts).strip())
    
    # 轉換 t="s" → t="inlineStr"，關鍵：保留原單元格 s 屬性
    def cell_replacer(m):
        cell_xml = m.group(0)
        ref = re.search(r'r="([^"]+)"', cell_xml).group(1)
        s_attr = f' s="{re.search(r'\bs="(\d+)"', cell_xml).group(1)}"'
        idx = int(re.search(r'<v>(\d+)</v>', cell_xml).group(1))
        val = new_ss[idx]
        return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{val}</t></is></c>'
    
    return re.sub(r'<c [^>]*t="s"[^>]*>.*?</c>', cell_replacer, sheet1_xml, flags=re.DOTALL)
```

**優勢：**
- ✅ 保留所有二進制資源（圖片、打印設置）
- ✅ 保留單元格樣式（居中、字體、填充）
- ✅ 無 openpyxl 依賴，純標準庫

### 3. BQ 頁合併器（bq_merger.py）

**功能：**
1. 讀取總表獲取 EL 編號 + 材料名 + BQ 編號
2. 在 BQ PDF 中搜索每個編號所在的頁碼（正則匹配）
3. 將報批表 PDF 與對應 BQ 整頁合併
4. 輸出文件名格式：`EL-XXX 材料名.pdf`

**使用示例：**

```python
from scripts.exporters.bq_merger import BQMerger

merger = BQMerger()

# 載入 BQ PDF
merger.load_bq_pdf("BQ標書.pdf")

# 載入總表
merger.load_zongbiao(
    zongbiao_path="材料審批總表.xlsx",
    sheet_index=0,
    start_row=7,
    col_bq=1,    # A 列：BQ 編號
    col_el=2,    # B 列：EL 編號
    col_name=3   # C 列：材料名
)

# 匹配 BQ 頁碼
success, failed = merger.match_bq_pages()

# 合併 PDF
merger.merge_pdfs("./報批表PDF", "./最終輸出")

merger.close()
```

### 4. 自動鏈接（auto_linker.py）

**功能：** 根據單元格內容自動匹配文件並創建超鏈接

**API：**

```python
class AutoLinker:
    def __init__(
        self,
        workbook: str,
        target_column: int,
        search_root: str
    ):
        """
        初始化自動鏈接器
        
        Args:
            workbook: 工作簿路徑
            target_column: 目標列（1 = A 列）
            search_root: 搜索根目錄
        """
        
    def convert_to_hyperlinks(
        self,
        extensions: List[str] = [".pdf", ".docx", ".xlsx"],
        include_subfolders: bool = True,
        use_wildcard: bool = True
    ) -> Dict[str, int]:
        """
        轉換為超鏈接
        
        Returns:
            {"success": 成功數, "skipped": 跳過數}
        """
```

---

## 📊 使用場景

### 場景 1：含圖片的報批表模板（推薦 ZIP 引擎）

```python
from scripts.template_filler import TemplateFiller

# 自動檢測 → ZIP 引擎
filler = TemplateFiller(
    data_source="材料審批總表.xlsx",
    template="報批表模板.xlsx"  # 含圖片
)

# 驗證
print(f"引擎: {'ZIP' if filler.has_images() else 'openpyxl'}")

# 生成
output_files = filler.fill_and_export(
    field_map={
        "{編號}": "編號",
        "{材料}": "材料",
        "{品牌}": "品牌",
        "{數量}": "數量"
    },
    output_dir="./報批表"
)
```

### 場景 2：完整流程（報批表 + BQ 合併）

```python
from scripts.template_filler import TemplateFiller
from scripts.exporters.bq_merger import BQMerger

# Step 1: 生成報批表 PDF
filler = TemplateFiller(
    data_source="材料審批總表.xlsx",
    template="報批表模板.xlsx"
)

filler.fill_and_export(
    export_format="pdf",
    output_dir="./報批表PDF",
    field_map={"{編號}": "編號", "{材料}": "材料"}
)

# Step 2: BQ 頁合併
merger = BQMerger()
merger.load_bq_pdf("BQ標書.pdf")
merger.load_zongbiao("材料審批總表.xlsx")
merger.match_bq_pages()
merger.merge_pdfs("./報批表PDF", "./最終輸出")
merger.close()
```

---

## ⚠️ 限制與注意事項

### ZIP 引擎注意事項

1. **佔位符必須在共享字符串中**（單元格顯示為文本）
2. **不支持單元格內混合佔位符和普通文字**
3. **每次生成後必須在 Excel 中打開驗證**

### PDF 導出限制

| 方案 | 優點 | 缺點 |
|:---|:---|:---|
| **win32com** | 完美保真 | 需要安裝 Excel（Windows only） |
| **純 Python** | 跨平台 | 需要手動排版 |

### BQ 合併限制

- BQ PDF **必須有文字層**（從 Word/Excel 轉換來的 PDF）
- 如果 BQ PDF 是純掃描圖（無文字），請改用 OCR 或手動提供頁碼映射

---

## 🧪 測試

```bash
# 運行所有測試
python -m pytest tests/

# 測試 openpyxl 引擎
python -m pytest tests/test_openpyxl_engine.py -v

# 測試 ZIP 引擎
python -m pytest tests/test_zip_engine.py -v

# 測試 BQ 合併器
python -m pytest tests/test_bq_merger.py -v
```

---

## 📝 更新日誌

### v2.0（2026-06-01）融合版
- ✅ 融合 `excel-template-filler` 和 `dem-material-approval`
- ✅ 實現雙引擎架構（openpyxl + ZIP）
- ✅ 自動引擎選擇
- ✅ 移植 BQ 頁合併功能
- ✅ 統一 API 接口

### v1.0（2026-06-01）
- 從 VBA 系統內化
- 實現核心功能：模板填充、批量導出、自動鏈接
- 支持 PDF/Excel 輸出

---

## 🔗 相關資源

- **原始 VBA 系統：** `~/工程總表宏/`
- **dem-material-approval：** `~/.qclaw/skills/dem-material-approval/`
- **融合分析文檔：** `memory/2026-06-01_skill_merge_analysis.md`
- **VBA 源碼分析：** `memory/2026-06-01.md`

---

_最後更新：2026-06-01 12:55（v2.0 融合版）_
