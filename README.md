# ESG 報告書自動化工具

台灣上市公司永續報告書（ESG）下載與圖表萃取工具，供學術研究使用。

---

## 專案結構

```
esg-report/
├── esg-report-downloader/     ← 爬蟲：自動下載永續報告書 PDF
├── esg-pdf-cuter/             ← 萃取：切割 PDF 圖表並輸出文字
└── data/                      ← 所有輸出資料（gitignore，不進版控）
    ├── 2015/
    │   ├── ESG_Download_Progress_2015.xlsx
    │   ├── ESG_Extract_Results_2015.xlsx
    │   └── 2015_1101_台泥/
    │       ├── 2015_1101_台泥.pdf
    │       ├── images/          ← 萃取出的圖表 PNG
    │       └── texts/           ← 每頁文字 TXT
    └── 2016/ ... 2022/
```

---

## 使用流程

### Step 1：下載 PDF（esg-report-downloader）

```bash
cd esg-report-downloader
python esg_downloader.py
```

- 選擇年度（可多選），點擊「▶ 開始下載」
- PDF 自動儲存到 `data/{year}/{公司代號_名稱}/`
- 進度記錄在 `data/{year}/ESG_Download_Progress_{year}.xlsx`

### Step 2：萃取圖表（esg-pdf-cuter）

```bash
cd esg-pdf-cuter
python esg_pdf_cuter.py
```

- 選擇年度，點擊「▶ 開始萃取」
- 圖表 PNG → `data/{year}/{公司}/images/`
- 頁面文字 TXT → `data/{year}/{公司}/texts/`
- 統計 Excel → `data/{year}/ESG_Extract_Results_{year}.xlsx`

---

## 安裝依賴

每個工具各自有獨立的 venv：

```bash
# 下載工具
cd esg-report-downloader
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt

# 萃取工具
cd esg-pdf-cuter
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## esg-report-downloader 操作說明

### 啟動設定視窗

- 勾選要下載的年份（可多選）
- 點「▶ 開始下載」後會詢問**補抓模式**：

| 選擇 | 說明 |
|------|------|
| **是** | 重新嘗試「失敗/錯誤/未找到」的公司，跳過「成功/確認無報告」 |
| **否** | 只跑從來沒有任何紀錄的新公司 |

### 下載狀態說明

| 狀態 | 說明 | 下次執行（選「是」） |
|------|------|------|
| ✅ 成功 | PDF 已下載 | 跳過 |
| ⚠️ 未找到中文版報告 | 當次找不到（可能被封鎖誤判） | 重試 |
| 🔒 已確認無報告 | 跨 3 次執行都找不到 | 跳過 |
| ❌ 下載失敗 / 處理錯誤 | 找到但下載失敗 | 重試 |

### 自動防護機制

- **封鎖偵測**：連續 5 家失敗 → 自動等待解封（最長 2 小時），解封後補抓
- **主動重啟**：每 50 家重啟一次 Chrome，降低被偵測機率
- **網路斷線**：自動暫停等待，最多 30 分鐘
- **下載卡住**：偵測檔案大小停滯 60 秒 → 自動放棄繼續下一家
- **進度防損毀**：先寫暫存檔再替換，確保 Excel 不損毀

### 暫停與關閉

1. 點擊「⏸ 暫停」→ 確認後等當前公司處理完才暫停
2. 右上角狀態變為「⏸ 已暫停」或「■ 已完成」後才可安全關閉

---

## esg-pdf-cuter 操作說明

### 萃取原理

1. **點陣圖偵測**：`page.get_images()` 抓嵌入圖片
2. **向量圖聚類**：`page.get_drawings()` 路徑用 Union-Find 分群，每群獨立裁切
3. **標籤擴張**：偵測框周圍 30pt 內的文字標籤一起納入裁切範圍
4. **無關文字塗白**：裁切範圍內與圖無關的文字區塊塗白
5. **216 DPI 輸出**：3× 渲染倍率，適合 CNN 分析

### 重新處理

```bash
cd esg-pdf-cuter
python clean_output.py   # 互動式清除所有 images/ 與 texts/ 資料夾
```

或手動刪除單一公司的 `data/{year}/{公司名}/images/` 即可觸發重新萃取（不影響其他公司）。

### 調整參數（esg_pdf_cuter.py 頂部）

| 參數 | 預設值 | 說明 |
|------|--------|------|
| `CLUSTER_GAP_PT` | 80 | 向量路徑聚類距離（pt） |
| `EXPAND_PT` | 20 | 偵測框基礎擴張距離 |
| `TEXT_LINK_GAP_PT` | 30 | 相鄰標籤的最大距離（pt） |
| `MASK_UNRELATED` | True | 是否塗白無關文字 |
| `SAVE_TXT` | True | 是否輸出頁面文字 TXT |
| `QR_MAX_AREA_PCT` | 4.0 | 正方形 Raster 圖 < 此面積（%）視為 QR code，跳過 |
| `RASTER_MAX_AREA_PCT` | 70 | Raster 圖最大面積佔比（%），過濾封面照片 |
| `DECO_ZONE_PCT` | 0.12 | 頁面頂/底各此比例視為裝飾區 |
| `DECO_MAX_HT_PT` | 40 | Vector cluster 高度 < 此值且橫跨頁面 → 視為裝飾線，跳過 |

---

## 注意事項

- `tw_listed.xlsx`（上市公司清單）需放在 `esg-report-downloader/` 資料夾內
- `data/` 內的 PDF、images/、texts/ 已加入 `.gitignore`，不進版控；進度 Excel（`ESG_Download_Progress_*.xlsx`、`ESG_Extract_Results_*.xlsx`）會進版控
- 執行日誌儲存於各程式的 `logs/` 資料夾
