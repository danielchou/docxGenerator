# DocxGenerator - Word 文檔自動生成工具

這是一個自動化的 Word 文檔生成工具，可以根據 YAML 配置檔案自動填充 Word 模板，生成各種申請單據。

## 功能特色

- 📄 **自動化文檔生成**：根據 YAML 配置自動填充 Word 模板
- 🔍 **模板診斷工具**：檢查模板中的佔位符是否正確配置
- 📁 **批量處理**：支援同時處理多個申請單
- 🛡️ **錯誤處理**：完整的錯誤處理和檔案檢查機制

## 目錄結構

```
DocsGenerator/
├── main.py                    # 主程式（推薦使用）
├── main.ipynb                 # Jupyter Notebook 版本
├── diagnose_template.py       # 模板診斷工具
├── diagonose_template.ipynb   # 診斷工具 Notebook 版本
├── histData/                  # YAML 配置檔案目錄
│   └── APR-*.yaml            # 申請單配置檔案
├── template/                  # Word 模板目錄
│   ├── 設計變更申請單.docx
│   ├── 系統過版申請單.docx
│   ├── 系統測試表.docx
│   └── 聯邦銀行廠商進館上線函.docx
└── output/                    # 生成的文檔輸出目錄
```

## 使用方法

### 1. 環境準備

安裝必要的 Python 套件：
```bash
pip install python-docx PyYAML
```

### 2. 配置申請單資料

在 `histData/` 目錄中創建 YAML 配置檔案，檔案格式參考 `APR-20240628-01-sample.yaml`：

```yaml
申請單編號: APR-20240628-01
申請單位: 信用卡行企
replacements: 
    "{{申請單位}}": 信用卡行企
    "{{申請日期}}": 2024-06-29
    "{{申請單編號}}": APR-20240628-01
    # ... 更多替換變數
```

### 3. 執行程式

**方法一：使用 Python 腳本（推薦）**
```bash
python main.py
```

**方法二：使用 Jupyter Notebook**
```bash
jupyter notebook main.ipynb
```

### 4. 診斷模板

如果需要檢查模板中的佔位符配置是否正確：
```bash
python diagnose_template.py
```

## 輸出結果

生成的文檔會保存在 `output/` 目錄下，按申請單編號分類：
```
output/
└── APR-20240628-01/
    ├── 信用卡行企-設計變更申請單-APR-20240628-01.docx
    ├── 信用卡行企-系統過版申請單-APR-20240628-01.docx
    ├── 信用卡行企-系統測試表-APR-20240628-01.docx
    └── 信用卡行企-聯邦銀行廠商進館上線函-APR-20240628-01.docx
```

## 注意事項

- 確保 `histData/` 和 `template/` 目錄存在
- YAML 檔案必須使用 UTF-8 編碼
- Word 模板中的佔位符格式為 `{{變數名稱}}`
- 建議先使用診斷工具檢查模板配置



