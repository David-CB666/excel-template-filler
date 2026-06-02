# 目錄結構與模塊說明

```
excel-template-filler/
├── src/
│   ├── template_filler.py       # 統一入口（自動引擎選擇）
│   ├── engines/
│   │   ├── base_engine.py       # 引擎接口基類
│   │   ├── openpyxl_engine.py   # openpyxl 引擎（無圖片）
│   │   └── zip_engine.py        # ZIP 引擎（圖片保留）
│   ├── exporters/
│   │   └── bq_merger.py         # BQ 頁合併器
│   ├── auto_linker.py           # 自動超鏈接創建器
│   ├── file_grabber.py          # 文件名抓取器
│   └── utils.py                 # 工具函數庫
├── examples/
│   ├── data/                    # 示例數據文件
│   ├── templates/               # 示例模板
│   ├── example_basic.py         # 基本填充示例
│   ├── example_batch_pdf.py    # 批量 PDF 示例
│   └── example_auto_link.py    # 自動鏈接示例
├── references/                  # 模塊文檔
├── pyproject.toml
├── requirements.txt
└── README.md
```

## 依賴

| 依賴 | 版本 | 用途 |
|:---|:---|:---|
| openpyxl | >= 3.1.0 | 無圖片模板引擎 |
| pywin32 | 可選 | PDF 導出（Windows） |
