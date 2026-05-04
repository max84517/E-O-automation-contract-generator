# E&O Settlement Contract Generator

自動將 Excel 資料套入 Word 範本，批次產出 E&O 結算協議書（.docx）。

## 目錄結構

```
project_root/
├── data/
│   └── E&O summary table.xlsx        ← 輸入 Excel（請自行放入）
├── template/
│   ├── Settlement Form_Draft.docx                 ← 一般範本
│   └── Settlement Form_Draft_Elan & SYNA.docx     ← Elan/Synaptics 專用
├── output/                                        ← 產出的 Word 結算書
├── src/
│   ├── main.py
│   ├── core/
│   │   ├── loader.py      — InfoLoader
│   │   ├── processor.py   — DataProcessor
│   │   └── writer.py      — WordWriter
│   ├── ui/
│   │   └── app.py         — CustomTkinter GUI (dark mode)
│   └── config/
│       └── paths.py       — 統一路徑管理
├── pyproject.toml
└── requirements.txt
```

## 安裝與執行

### 使用 Poetry（建議）

```bash
poetry install
poetry run eo-generator
```

### 使用 pip

```bash
pip install -r requirements.txt
python src/main.py
```

## Excel 格式需求

工作簿需包含兩個工作表：

### Summary 表（必要欄位）

| 欄位 | 說明 |
|------|------|
| GTK Supplier | 供應商代碼（merge key） |
| Platform | 平台代號 |
| GBU | 產品類別（如 cNB、bDT），用於 Signer 對應 |
| Actual Payment | 結算金額 |
| Sub-Category | 零件類別 |

### Info 表（必要欄位）

| 欄位 | 說明 |
|------|------|
| Supplier | 供應商代碼（merge key） |
| GBU | NB / DT（供應商有多個 Signer 時區分用） |
| Supplier name | 供應商全名 |
| Master Agreement | 合約編號 |
| Effective Date | 合約生效日 |
| Signer | 簽署人姓名 |
| Signer title | 簽署人職稱 |

> **Signer 對應邏輯**：若同一 Supplier 在 Info 表有多列（如 Chicony NB / DT），
> 系統會取 Summary.GBU 最後兩碼（cNB→NB、bDT→DT）來比對 Info.GBU，
> 自動選出對應的 Signer。

## Word 佔位符對應

| 佔位符 | 來源 |
|--------|------|
| `<Master Agreement>` | Info.Master Agreement |
| `<Supplier Name>` | Info.Supplier name |
| `<Effective Date>` | Info.Effective Date（格式：Month DD, YYYY） |
| `<Sub-Category>` | Summary.Sub-Category |
| `<Platform>` | Summary.Platform |
| `<Actual Payment>` | Summary.Actual Payment（格式：$X,XXX.XX） |
| `<Capital Money Letter>` | Summary.Actual Payment（英文大寫） |
| `<Signer>` | Info.Signer |
| `<Title>` | Info.Signer title |

> 以上 9 個欄位**全部必填**，任何一個空值都會禁止生成，並在 Log 顯示缺少哪些欄位。

## GUI 功能說明

- **表格分三區段**：Pending（待生成）→ Missing Data（資料不全）→ Generated（已生成）
- **Pending** 項目預設勾選；**Missing Data** 項目 checkbox 停用
- **Select All / Deselect All**：批次勾選可生成的項目
- **Refresh**：重新掃描 output/ 資料夾，更新狀態
- **Generate Selected**：背景執行緒批次產出，進度條 + Log 即時顯示結果
- Log 失敗訊息範例：`❌ LiteOn Holmes: Missing fields — Effective Date, Master Agreement`

## 範本選擇邏輯

- GTK Supplier 含有 **Elan** 或 **Synaptics**（不區分大小寫）→ `Settlement Form_Draft_Elan & SYNA.docx`
- 其他廠商 → `Settlement Form_Draft.docx`

## 輸出檔名格式

```
Settlement Form_<GTK Supplier> <Platform>.docx
```

## 打包成 exe（選用）

```bash
poetry run pyinstaller --noconfirm --onedir --windowed \
    --name "EO_Contract_Generator" \
    src/main.py
```

產出於 `dist/EO_Contract_Generator/`，將整個資料夾連同 `data/`、`template/`、`output/` 一起部署。

