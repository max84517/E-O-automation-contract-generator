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
│   │   └── app.py         — CustomTkinter GUI
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

| 工作表 | 必要欄位 |
|--------|---------|
| **Summary** | GTK Supplier, Platform, Actual Payment |
| **Info** | GTK Supplier, Platform, Master Agreement, Supplier name, Effective Date, Sub-Category, Signer, Signer title |

> `GTK Supplier + Platform` 作為兩張表的複合索引進行合併。

## Word 佔位符對應

| 佔位符 | 來源 |
|--------|------|
| `<Master Agreement>` | Info.Master Agreement |
| `<Supplier Name>` | Info.Supplier name |
| `<Effective Date>` | Info.Effective Date（格式：Month DD, YYYY） |
| `<Sub-Category>` | Info.Sub-Category |
| `<Platform>` | Info/Summary.Platform |
| `<Actual Payment>` | Summary.Actual Payment（格式：$X,XXX.XX） |
| `<Capital Money Letter>` | Summary.Actual Payment（英文大寫） |
| `<Signer>` | Info.Signer |
| `<Title>` | Info.Signer title |

## 範本選擇邏輯

- GTK Supplier 含有 **Elan** 或 **Synaptics**（不區分大小寫） → 使用 `Settlement Form_Draft_Elan & SYNA.docx`
- 其他廠商 → 使用 `Settlement Form_Draft.docx`

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
