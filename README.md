# ESG 報告書自動化工具

台灣上市公司永續報告書（ESG）下載與圖表萃取工具，供學術研究使用。

---

## 專案結構

```
esg-report/
├── tools/
│   ├── report-downloader/     ← 爬蟲：自動下載永續報告書 PDF
│   ├── pdf-cuter/             ← 萃取：切割 PDF 圖表並輸出文字
│   ├── chart-counter/         ← 計數：CLIP zero-shot 統計各公司圖表數量
│   └── dashboard/             ← 主控台：查看下載與萃取進度
├── data/                      ← 所有輸出資料（gitignore，不進版控）
│   ├── 2015/
│   │   ├── ESG_Download_Progress_2015.xlsx
│   │   ├── ESG_Extract_Results_2015.xlsx
│   │   └── 2015_1101_台泥/
│   │       ├── 2015_1101_台泥.pdf
│   │       ├── images/        ← 萃取出的圖表 JPEG
│   │       ├── charts/        ← 判定為圖表的圖片（chart-counter 輸出）
│   │       └── texts/         ← 每頁文字 TXT
│   ├── 2016/ ... 2024/
│   └── chart_statistics.xlsx  ← 圖表計數結果（chart-counter 輸出）
├── ESG.png
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
python tools/report-downloader/esg_downloader.py
```

- 選擇年度（可多選），點擊「▶ 開始下載」
- PDF 自動儲存到 `data/{year}/{公司代號_名稱}/`
- 進度記錄在 `data/{year}/ESG_Download_Progress_{year}.xlsx`

### Step 2：萃取圖表

```bash
python tools/pdf-cuter/esg_pdf_cuter.py
```

- 選擇年度，點擊「▶ 開始萃取」
- 圖表 JPEG → `data/{year}/{公司}/images/`
- 頁面文字 TXT → `data/{year}/{公司}/texts/`
- 統計 Excel → `data/{year}/ESG_Extract_Results_{year}.xlsx`

### Step 3：統計圖表數量

```bash
python tools/chart-counter/chart_counter.py
```

- 選擇年度，調整 CLIP 機率門檻（預設 0.55）
- 程式讀取 `data/{year}/*/images/*.jpg`，以 CLIP zero-shot 分類每張圖是否為「統計圖表或表格」
- 結果輸出至 `data/chart_statistics.xlsx`（11 個 sheet：總覽 + 2015–2024 各年）
- **首次執行**：自動下載 PyTorch（~300 MB）及 CLIP 模型（~600 MB），之後快取於 `~/.cache/huggingface/`

### Step 4：查看主控台

```bash
python tools/dashboard/esg-dashboard.py
```

或在下載／萃取視窗點擊「🖥 查看主控台」。

---

## chart-counter 操作說明

### 判斷標準

CLIP 比較兩個 prompt 的相似度：
- **圖表**：含統計數字的圖示（長條、折線、圓餅、散點等）與表格
- **非圖表**：logo、照片、裝飾圖、地圖、純文字

### 主要參數（chart_counter.py 頂部）

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `CHART_THRESHOLD` | 0.55 | CLIP chart 機率門檻（可在 GUI 滑桿調整） |
| `BATCH_SIZE` | 16 | 每批處理張數 |
| `CLIP_MODEL_ID` | `openai/clip-vit-base-patch32` | 使用的 CLIP 模型 |

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

- `tw_listed.xlsx`（上市公司清單）需放在 `tools/report-downloader/` 資料夾內
- `data/` 內的 PDF、images/、texts/ 已加入 `.gitignore`，不進版控
- 進度 Excel（`ESG_Download_Progress_*.xlsx`、`ESG_Extract_Results_*.xlsx`）會進版控
- 執行日誌儲存於各程式的 `logs/` 資料夾

## 讀取與下載路徑

esg_downloader.py（下載器）
讀取：

tools/report-downloader/tw_listed.xlsx — 上市公司清單
data/<year>/ESG_Download_Progress_<year>.xlsx — 讀取舊進度（斷點續傳）
寫入：

data/<year>/<公司>/2015_1101_台泥.pdf — 下載的 PDF
data/<year>/ESG_Download_Progress_<year>.xlsx — 更新下載進度
tools/report-downloader/logs/ESG_Log_*.txt — 執行日誌
esg_pdf_cuter.py（圖表萃取）
讀取：

data/<year>/<公司>/*.pdf — 原始 PDF
寫入：

data/<year>/<公司>/images/*.jpg — 萃取的圖片
data/<year>/<公司>/texts/*.txt — 每頁文字
data/<year>/<公司>/garbled_pages.txt — 無法讀取的頁面記錄
data/<year>/ESG_Extract_Results_<year>.xlsx — 萃取統計
chart_counter.py（圖表計數）
讀取：

data/<year>/<公司>/images/*.jpg — 萃取的圖片
寫入：

data/<year>/<公司>/charts/*.jpg — 判定為圖表的圖片（複製）
data/chart_statistics.xlsx — 各公司圖表數量統計
esg-dashboard.py（主控台）
讀取（只讀，不寫入）：

data/<year>/ESG_Download_Progress_<year>.xlsx
data/<year>/ESG_Extract_Results_<year>.xlsx
data/<year>/<公司>/images/*.jpg（只計數）
data/<year>/<公司>/garbled_pages.txt
