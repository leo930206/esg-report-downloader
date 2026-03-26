# ESG 報告書自動化工具

台灣上市公司永續報告書（ESG）下載與圖表萃取工具，供學術研究使用。

---

## 專案結構

```
esg-report/
├── report-downloader/         ← 爬蟲：自動下載永續報告書 PDF
├── pdf-cuter/                 ← 萃取：切割 PDF 圖表並輸出文字
├── dashboard/                 ← 主控台：查看下載與萃取進度
├── data/                      ← 所有輸出資料（gitignore，不進版控）
│   ├── 2015/
│   │   ├── ESG_Download_Progress_2015.xlsx
│   │   ├── ESG_Extract_Results_2015.xlsx
│   │   └── 2015_1101_台泥/
│   │       ├── 2015_1101_台泥.pdf
│   │       ├── images/        ← 萃取出的圖表 JPEG
│   │       └── texts/         ← 每頁文字 TXT
│   └── 2016/ ... 2022/
└── requirements.txt
```

---

## 環境安裝

需要 **Python 3.12 以上**（macOS 16 Tahoe 以上請勿使用 Python 3.9，Tk 不相容）。

### macOS / Linux

```bash
git clone <repo-url>
cd esg-report
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows

```bat
git clone <repo-url>
cd esg-report
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

安裝完成後，VS Code 選擇 interpreter 為根目錄的 `.venv`（`Cmd+Shift+P` → `Python: Select Interpreter`），之後直接按 **Run** 執行任一腳本。

---

## 使用流程

### Step 1：下載 PDF

```bash
python report-downloader/esg_downloader.py
```

- 選擇年度（可多選），點擊「▶ 開始下載」
- PDF 自動儲存到 `data/{year}/{公司代號_名稱}/`
- 進度記錄在 `data/{year}/ESG_Download_Progress_{year}.xlsx`

### Step 2：萃取圖表

```bash
python pdf-cuter/esg_pdf_cuter.py
```

- 選擇年度，點擊「▶ 開始萃取」
- 圖表 JPEG → `data/{year}/{公司}/images/`
- 頁面文字 TXT → `data/{year}/{公司}/texts/`
- 統計 Excel → `data/{year}/ESG_Extract_Results_{year}.xlsx`

### Step 3：查看主控台

```bash
python dashboard/esg-dashboard.py
```

或在下載／萃取視窗點擊「🖥 查看主控台」。

---

## report-downloader 操作說明

### 補抓模式

點「▶ 開始下載」後會詢問補抓模式：

| 選擇 | 說明 |
|------|------|
| **是** | 重新嘗試「失敗/錯誤/未找到」的公司，跳過「成功/確認無報告」 |
| **否** | 只跑從來沒有任何紀錄的新公司 |

### 下載狀態

| 狀態 | 說明 | 下次執行（選「是」） |
|------|------|------|
| ✅ 成功 | PDF 已下載 | 跳過 |
| ⚠️ 未找到中文版報告 | 當次找不到 | 重試 |
| 🔒 已確認無報告 | 跨 3 次都找不到 | 跳過 |
| ❌ 下載失敗 | 找到但下載失敗 | 重試 |

### 自動防護機制

- 連續 5 家失敗 → 自動等待解封（最長 2 小時）
- 每 50 家重啟一次 Chrome，降低被偵測機率
- 網路斷線 → 自動暫停等待，最多 30 分鐘
- 下載卡住 → 偵測停滯 60 秒後自動放棄繼續下一家

---

## pdf-cuter 操作說明

### 萃取原理

1. **點陣圖偵測**：`page.get_images()` 抓嵌入圖片
2. **向量圖聚類**：`page.get_drawings()` 路徑用 Union-Find 分群，每群獨立裁切
3. **標籤擴張**：偵測框周圍 50pt 內的文字標籤一起納入裁切範圍
4. **JPEG 輸出**：q85 壓縮，2× 渲染倍率（144 DPI）

### 主要參數（esg_pdf_cuter.py 頂部）

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `RENDER_SCALE` | 2 | 渲染倍率 |
| `CLUSTER_GAP_PT` | 40 | 向量路徑聚類距離（pt） |
| `EXPAND_PT` | 50 | 偵測框擴張距離 |
| `SAVE_TXT` | True | 是否輸出頁面文字 TXT |

---

## 注意事項

- `tw_listed.xlsx`（上市公司清單）需放在 `report-downloader/` 資料夾內
- `data/` 內的 PDF、images/、texts/ 已加入 `.gitignore`，不進版控
- 進度 Excel（`ESG_Download_Progress_*.xlsx`、`ESG_Extract_Results_*.xlsx`）會進版控
- 執行日誌儲存於各程式的 `logs/` 資料夾
