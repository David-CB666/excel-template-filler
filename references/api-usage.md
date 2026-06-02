# API 用法詳解

## 基本用法

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from template_filler import TemplateFiller

# 初始化（自動選擇引擎）
filler = TemplateFiller(
    data_source="data/sample_data.xlsx",
    template="templates/sample_template.xlsx"
)

# 檢查引擎類型
print(f"Engine: {'ZIP' if filler.has_images() else 'openpyxl'}")

# 加載數據
filler.load_data()

# 掃描佔位符
placeholders = filler.scan_placeholders()

# 填充並導出
output_files = filler.fill_and_export(
    field_map={
        "{ID}": "ID",
        "{Name}": "Name",
        "{Brand}": "Brand",
        "{Qty}": "Qty"
    },
    output_dir="./output"
)
```

---

## 雙引擎架構

| 引擎 | 適用場景 | 優勢 |
|:---|:---|:---|
| **openpyxl** | 無圖片模板 | API 乾淨，易維護 |
| **ZIP** | 含圖片模板 | 完美保留圖片、打印設置、二進制資源 |

`TemplateFiller` 根據模板內容自動選擇引擎，無需手動配置。

---

## 完整流程（模板填充 + BQ 合併）

```python
# Step 1: 生成填充 PDF
filler = TemplateFiller("data.xlsx", "template.xlsx")
filler.load_data()
filler.fill_and_export(
    export_format="pdf", output_dir="./pdfs"
)

# Step 2: BQ 頁合併
from exporters.bq_merger import BQMerger
merger = BQMerger()
merger.load_bq_pdf("bq_reference.pdf")
merger.load_zongbiao("data.xlsx")
merger.match_bq_pages()
merger.merge_pdfs("./pdfs", "./final_output")
```

---

## 自動鏈接

```python
from auto_linker import AutoLinker

linker = AutoLinker("summary.xlsx", 1, "./files")
linker.link_all()
```

---

## 限制

1. 佔位符必須在 shared strings 中（普通單元格文本）
2. 單元格中佔位符+文字混合不支持
3. BQ 合併需要帶文本層的 PDF（不支持掃描圖片）
