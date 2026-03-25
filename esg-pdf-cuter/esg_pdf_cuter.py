import os
import subprocess
import threading
import queue
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox, scrolledtext
from pathlib import Path
from datetime import datetime

import fitz           # PyMuPDF
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ============================================================
# 路徑設定
# ============================================================
BASE_DIR = Path(__file__).parent.absolute()
DATA_DIR = BASE_DIR.parent / "data"          # 統一輸出根目錄
DATA_DIR.mkdir(exist_ok=True)

def year_dir(year: str) -> Path:
    return DATA_DIR / str(year)

def year_excel(year: str) -> Path:
    """每個年份各自的萃取統計 Excel 路徑"""
    return year_dir(year) / f"ESG_Extract_Results_{year}.xlsx"

def _year_range():
    return range(2015, 2025)

def available_years():
    if not DATA_DIR.is_dir():
        return [str(y) for y in _year_range()]
    dirs = [d for d in os.listdir(DATA_DIR)
            if (DATA_DIR / d).is_dir() and d.isdigit()]
    return sorted(dirs) or [str(y) for y in _year_range()]

# ============================================================
# App Icon（Dock / 視窗）
# ============================================================
def set_app_icon(root: tk.Tk, emoji: str = "🌱") -> None:
    """把 emoji 渲染成圖片並設為視窗與 Dock 圖示（macOS 透過 AppKit）。"""
    try:
        import base64
        from io import BytesIO

        size = 256
        img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        font = None
        for fp in ("/System/Library/Fonts/Apple Color Emoji.ttc",
                   "/System/Library/Fonts/AppleColorEmoji.ttf"):
            try:
                font = ImageFont.truetype(fp, size - 20)
                break
            except Exception:
                pass

        if font:
            draw.text((10, 10), emoji, font=font, embedded_color=True)

        buf = BytesIO()
        img.save(buf, format="PNG")
        png_bytes = buf.getvalue()

        # macOS：透過 AppKit 設定 Dock 圖示
        try:
            from AppKit import NSApplication, NSImage
            from Foundation import NSData
            data     = NSData.dataWithBytes_length_(png_bytes, len(png_bytes))
            ns_image = NSImage.alloc().initWithData_(data)
            NSApplication.sharedApplication().setApplicationIconImage_(ns_image)
        except Exception:
            pass

        # 所有平台：設定 tkinter 視窗圖示
        photo = tk.PhotoImage(data=base64.b64encode(png_bytes).decode())
        root.iconphoto(True, photo)
        root._icon_ref = photo  # 防止被 GC 回收

    except Exception:
        pass


# ============================================================
# Apple 風格配色
# ============================================================
APPLE_BG     = '#f5f5f7'
APPLE_CARD   = '#ffffff'
APPLE_BLUE   = '#0071e3'
APPLE_TEXT   = '#1d1d1f'
APPLE_GREY   = '#6e6e73'
APPLE_BORDER = '#d2d2d7'

FONT_TITLE = ('Helvetica Neue', 13, 'bold')
FONT_MAIN  = ('Helvetica Neue', 10)
FONT_LABEL = ('Helvetica Neue', 9)
FONT_STAT  = ('Helvetica Neue', 20, 'bold')
FONT_LOG   = ('Menlo', 9)

# ============================================================
# 執行緒狀態
# ============================================================
log_queue    = queue.Queue()
program_done = threading.Event()
pause_event  = threading.Event()
paused_event = threading.Event()   # 執行緒真正停下來才 set

ui_stats = {
    'total': 0, 'done': 0, 'images': 0, 'skipped': 0, 'error': 0,
}

# ============================================================
# 萃取參數
# ============================================================
RENDER_SCALE     = 3      # 渲染倍率（3x = 216 DPI）
CLUSTER_GAP_PT   = 80    # 向量路徑聚類距離（PDF 點座標，80pt ≈ 1.1 吋）
EXPAND_PT        = 20    # 偵測框基礎擴張距離
MIN_AREA_PCT     = 0.5   # 最小面積佔比（%），過濾微小雜訊
MAX_AREA_PCT     = 90    # 最大面積佔比（%），Vector 過濾整頁背景
MIN_DIM_PT       = 100   # 最小寬/高（PDF 點），過濾細長雜訊
MAX_PAGE_RATIO   = 0.95  # 超過此比例視為整頁背景（width 判斷）
MIN_PATHS        = 10    # 頁面向量路徑數 >= 此值才做聚類

# [A] QR code 過濾：長寬比接近 1:1 且面積極小 → 跳過
QR_ASPECT_MIN    = 0.8   # 長寬比下限（width/height）
QR_ASPECT_MAX    = 1.25  # 長寬比上限
QR_MAX_AREA_PCT  = 9.0   # Raster 面積 < 此值（%）且為正方形 → 視為 QR code（提高以涵蓋較大 QR）

# [B] Raster 全頁圖過濾：章節封面照片
RASTER_MAX_AREA_PCT = 60  # Raster 專屬最大面積佔比（%），比 Vector 嚴格（降低以抓更多封面照）

# [C] Vector 裝飾線過濾：扁平 cluster 且位於頁首/頁尾區域 → 跳過
DECO_ZONE_PCT        = 0.12  # 頁面頂/底各 12% 為「裝飾區」
DECO_MAX_HT_PT       = 40    # cluster 高度 < 此值（pt）才可能是裝飾線
DECO_MIN_WIDTH_RATIO = 0.65  # cluster 寬度 > 頁面此比例才算「橫跨型裝飾」

# [D] Vector cluster 路徑數門檻：過少代表是裝飾圓形/單一 icon → 跳過
MIN_CLUSTER_PATHS = 5    # cluster 內原始路徑數 < 此值 → 跳過（裝飾形狀通常只有 1~3 條路徑）

TEXT_LINK_GAP_PT = 30    # 距圖表 <= 此距離的文字視為「相鄰標籤」，納入裁切並保留
MASK_UNRELATED   = True  # 塗白裁切範圍內與圖無關的文字區塊
SAVE_TXT         = True  # 每張圖另存同名 .txt（含圖表周邊文字）

# ============================================================
# 核心函式
# ============================================================
def _cluster_drawing_rects(rects: list, gap: float) -> list[tuple]:
    """
    把向量路徑的 fitz.Rect 用 Union-Find 聚類。
    兩個 Rect 若互相擴張 gap/2 後重疊，就歸同一群。
    回傳每群 (合併後的 fitz.Rect, 路徑數量) 列表。
    """
    if not rects:
        return []

    n = len(rects)
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    half = gap / 2
    for i in range(n):
        r1_exp = rects[i] + (-half, -half, half, half)
        for j in range(i + 1, n):
            r2_exp = rects[j] + (-half, -half, half, half)
            if (r1_exp & r2_exp).is_valid:
                union(i, j)

    groups: dict[int, fitz.Rect] = {}
    counts: dict[int, int] = {}
    for i in range(n):
        root = find(i)
        if root in groups:
            groups[root] |= rects[i]
            counts[root] += 1
        else:
            groups[root] = rects[i]
            counts[root] = 1

    return [(groups[k], counts[k]) for k in groups]


def _detect_chart_regions(page) -> list[tuple]:
    """
    偵測頁面上的圖表區域，回傳 (fitz.Rect, type_str) 列表。
    type_str 為 'Raster' 或 'Vector'。
    """
    page_rect = page.rect
    candidates = []

    page_area = page_rect.width * page_rect.height

    # 方法一：點陣圖片
    for img_info in page.get_images(full=True):
        xref = img_info[0]
        for r in page.get_image_rects(xref):
            if r.width <= 50 or r.height <= 50:
                continue
            area_pct = r.width * r.height / page_area * 100

            # [A] QR code 過濾：正方形且極小
            aspect = r.width / r.height if r.height > 0 else 0
            if QR_ASPECT_MIN <= aspect <= QR_ASPECT_MAX and area_pct < QR_MAX_AREA_PCT:
                continue

            # [B] Raster 全頁圖過濾：章節封面/背景照片
            if area_pct > RASTER_MAX_AREA_PCT:
                continue

            candidates.append((r, 'Raster'))

    # 方法二：向量圖形聚類
    paths = page.get_drawings()
    drawing_rects = [
        p["rect"] for p in paths
        if p["rect"].width > 5 and p["rect"].height > 5
    ]

    if len(drawing_rects) >= MIN_PATHS:
        clusters = _cluster_drawing_rects(drawing_rects, CLUSTER_GAP_PT)
        for cluster_rect, path_count in clusters:
            # [D] 路徑數門檻：單一裝飾形狀通常只有 1~3 條路徑
            if path_count < MIN_CLUSTER_PATHS:
                continue

            expanded = cluster_rect + (-EXPAND_PT, -EXPAND_PT, EXPAND_PT, EXPAND_PT)
            expanded &= page_rect          # 不超出頁面邊界
            if not (expanded.width > MIN_DIM_PT and expanded.height > MIN_DIM_PT
                    and expanded.width < page_rect.width * MAX_PAGE_RATIO):
                continue

            # [C] 裝飾線過濾：扁平 cluster 且位於頁首/頁尾區域
            in_header = expanded.y0 < page_rect.height * DECO_ZONE_PCT
            in_footer = expanded.y1 > page_rect.height * (1 - DECO_ZONE_PCT)
            is_flat   = (expanded.height < DECO_MAX_HT_PT
                         and expanded.width > page_rect.width * DECO_MIN_WIDTH_RATIO)
            if (in_header or in_footer) and is_flat:
                continue

            candidates.append((expanded, 'Vector'))

    return candidates


def _find_related_text_rects(page, chart_rect) -> list:
    """
    回傳距 chart_rect 在指定距離以內的文字 block 矩形列表。
    上方使用較大擴張（TEXT_LINK_GAP_TOP）以涵蓋圖表標題，
    其餘方向使用 TEXT_LINK_GAP_OTHERS。
    """
    probe = chart_rect + (
        -TEXT_LINK_GAP_PT, -TEXT_LINK_GAP_PT,
         TEXT_LINK_GAP_PT,  TEXT_LINK_GAP_PT
    )
    probe &= page.rect
    related = []
    for b in page.get_text("blocks", clip=probe):
        if b[6] != 0:           # 跳過圖片 block
            continue
        b_rect = fitz.Rect(b[0], b[1], b[2], b[3])
        if b[4].strip():        # 只要有實際文字
            related.append(b_rect)
    return related


def _mask_unrelated_text(pix, crop_rect, related_rects, page):
    """
    把 pix 轉成 PIL Image，找出 crop_rect 內「不相鄰」的文字 block 並塗白。
    回傳處理後的 PIL Image。
    """
    img  = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    draw = ImageDraw.Draw(img)
    s    = RENDER_SCALE

    for b in page.get_text("blocks", clip=crop_rect):
        if b[6] != 0:
            continue
        b_rect = fitz.Rect(b[0], b[1], b[2], b[3])
        # 若此 block 與任何 related_rect 有重疊則保留
        if any((b_rect & rr).is_valid for rr in related_rects):
            continue
        # 否則塗白（轉換成像素座標）
        x0 = max(0,         (b_rect.x0 - crop_rect.x0) * s)
        y0 = max(0,         (b_rect.y0 - crop_rect.y0) * s)
        x1 = min(pix.width, (b_rect.x1 - crop_rect.x0) * s)
        y1 = min(pix.height,(b_rect.y1 - crop_rect.y0) * s)
        draw.rectangle([x0, y0, x1, y1], fill=(255, 255, 255))
    return img


def process_pdf(pdf_path: str, year: str) -> list[dict]:
    """
    偵測單一 PDF 每頁的圖表區域（點陣圖 + 向量圖聚類），
    高解析度裁切後存成 PNG；每頁另存全文 .txt。
    """
    doc       = fitz.open(pdf_path)
    file_stem = Path(pdf_path).stem
    base_dir  = DATA_DIR / str(year) / file_stem
    img_dir   = base_dir / "images"
    txt_dir   = base_dir / "texts"
    img_dir.mkdir(parents=True, exist_ok=True)
    if SAVE_TXT:
        txt_dir.mkdir(parents=True, exist_ok=True)

    results: list[dict] = []
    garbled_pages = 0

    for page_index, page in enumerate(doc):
        page_num  = page_index + 1
        try:
            page_area = page.rect.width * page.rect.height

            # ── 全頁文字存檔（每頁一份）──
            if SAVE_TXT:
                page_text = page.get_text("text").strip()
                if page_text:
                    cjk = sum(1 for c in page_text if '\u4e00' <= c <= '\u9fff'
                              or '\u3400' <= c <= '\u4dbf')
                    is_garbled = (cjk / len(page_text) < 0.05) and len(page_text) > 50
                    if is_garbled:
                        garbled_pages += 1
                    else:
                        txt_path = txt_dir / f"{file_stem}_p{page_num}.txt"
                        txt_path.write_text(page_text, encoding="utf-8")

            candidates = _detect_chart_regions(page)

            asset_idx = 0
            for r, rtype in candidates:
                area_pct = round(r.width * r.height / page_area * 100, 4)
                if area_pct < MIN_AREA_PCT or area_pct > MAX_AREA_PCT:
                    continue

                related_rects = _find_related_text_rects(page, r)
                crop_rect = r
                for tr in related_rects:
                    crop_rect |= tr
                crop_rect &= page.rect

                asset_idx += 1
                type_code = "RA" if rtype == 'Raster' else "VC"
                img_name  = f"{file_stem}_p{page_num}_{asset_idx}_{type_code}.png"
                save_path = img_dir / img_name

                # 嘗試渲染，失敗則降解析度重試
                pix = None
                for scale in (RENDER_SCALE, 2, 1):
                    try:
                        pix = page.get_pixmap(
                            matrix=fitz.Matrix(scale, scale),
                            clip=crop_rect, alpha=False)
                        break
                    except Exception:
                        pix = None
                if pix is None:
                    log_queue.put(('warning',
                        f'  無法渲染 {file_stem} p{page_num} 區塊 {asset_idx}，跳過'))
                    asset_idx -= 1
                    continue

                if MASK_UNRELATED and related_rects:
                    img = _mask_unrelated_text(pix, crop_rect, related_rects, page)
                    img.save(str(save_path))
                else:
                    pix.save(str(save_path))
                pix = None

                results.append({
                    "年份":           year,
                    "PDF檔名":        file_stem,
                    "PDF總頁數":      len(doc),
                    "頁碼":           page_num,
                    "圖片編號":       asset_idx,
                    "圖片面積佔比(%)": area_pct,
                    "類型":           rtype,
                    "圖片檔名":       img_name,
                    "存檔路徑":       str(save_path),
                    "亂碼頁數":       garbled_pages,
                })

        except Exception as e:
            log_queue.put(('warning', f'  跳過 {file_stem} 第 {page_num} 頁：{e}'))

    doc.close()
    return results

# ============================================================
# 萃取執行緒
# ============================================================
def _is_already_processed(pdf_path: str, year: str) -> bool:
    """
    判斷此 PDF 是否已處理過：
    data/<year>/<file_stem>/images/ 資料夾存在且至少有一個 PNG 檔案。
    只要刪除該資料夾即可觸發重新處理。
    """
    file_stem = Path(pdf_path).stem
    img_dir   = DATA_DIR / str(year) / file_stem / "images"
    if not img_dir.is_dir():
        return False
    return any(img_dir.glob("*.png"))


def run_extraction(years):
    # 收集待處理 PDF
    tasks = []
    for year in years:
        pdf_folder = DATA_DIR / year
        if not pdf_folder.is_dir():
            log_queue.put(('warning', f'找不到資料夾：{pdf_folder}'))
            continue
        for pdf_file in sorted(pdf_folder.rglob("*.pdf")):
            tasks.append((str(pdf_file), year))

    total   = len(tasks)
    pending = [(p, y) for p, y in tasks if not _is_already_processed(p, y)]
    skipped = total - len(pending)

    ui_stats.update({'total': total, 'done': skipped, 'images': 0,
                     'skipped': skipped, 'error': 0})
    log_queue.put(('info', f'共 {total} 個 PDF，已有輸出跳過 {skipped} 個，待處理 {len(pending)} 個（刪除 graph/ 子資料夾可重新處理）'))

    if not pending:
        log_queue.put(('info', '所有檔案皆已處理完成'))
        program_done.set()
        return

    # 每個年份各自維護一份 data list（對應各自的 Excel）
    year_data: dict[str, list] = {}
    for y in set(yr for _, yr in pending):
        xls = year_excel(y)
        if xls.exists():
            try:
                year_data[y] = pd.read_excel(xls).to_dict('records')
            except Exception:
                year_data[y] = []
        else:
            year_data[y] = []

    for i, (pdf_path, year) in enumerate(pending):
        # 暫停檢查點：處理每個 PDF 之前先看有沒有暫停請求
        if pause_event.is_set():
            log_queue.put(('warning', '⏸ 已暫停，進度已儲存，可安全關閉視窗'))
            paused_event.set()
            while pause_event.is_set():
                if program_done.is_set():
                    return
                time.sleep(0.2)
            paused_event.clear()
            log_queue.put(('info', '▶ 繼續執行'))

        fname = os.path.basename(pdf_path)
        log_queue.put(('info', f'[{i+1}/{len(pending)}] 處理 {fname}'))

        try:
            results = process_pdf(pdf_path, year)
            year_data[year].extend(results)
            ui_stats['images'] += len(results)
            ui_stats['done']   += 1

            xls = year_excel(year)
            xls.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(year_data[year]).to_excel(xls, index=False)
            log_queue.put(('success', f'  完成：切割 {len(results)} 個區塊'))

        except Exception as e:
            ui_stats['error'] += 1
            log_queue.put(('error', f'  錯誤：{fname} — {e}'))

    log_queue.put(('success',
                   f'全部完成！共切割 {ui_stats["images"]} 個區塊，錯誤 {ui_stats["error"]} 個'))
    program_done.set()

# ============================================================
# 啟動設定視窗
# ============================================================
def create_startup_window():
    selected_years = []

    root = tk.Tk()
    root.title("🌱 ESG 圖表萃取系統")
    root.geometry("480x380")
    root.configure(bg=APPLE_BG)
    root.resizable(False, False)
    set_app_icon(root)

    header = tk.Frame(root, bg=APPLE_BLUE, pady=14)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 圖表萃取系統", font=FONT_TITLE,
             fg='white', bg=APPLE_BLUE).pack()
    tk.Label(header, text="永續報告書圖表自動擷取 · CNN 前處理", font=FONT_LABEL,
             fg='#a8d4ff', bg=APPLE_BLUE).pack(pady=(2, 0))

    if not DATA_DIR.is_dir():
        messagebox.showerror(
            "找不到資料來源",
            f"找不到以下資料夾：\n{DATA_DIR}\n\n"
            "請確認 data/ 資料夾存在於專案根目錄。"
        )
        root.destroy()
        return selected_years

    content = tk.Frame(root, bg=APPLE_BG, padx=25, pady=15)
    content.pack(fill=tk.BOTH, expand=True)

    tk.Label(content, text="請選擇要處理的年份（可多選）",
             font=('Helvetica Neue', 11, 'bold'),
             fg=APPLE_TEXT, bg=APPLE_BG).pack(anchor='w', pady=(0, 10))

    grid = tk.Frame(content, bg=APPLE_BG)
    grid.pack(fill=tk.X)

    all_years = available_years()
    year_vars = {}
    for i, y in enumerate(all_years):
        var = tk.BooleanVar(value=False)
        cb  = tk.Checkbutton(grid, text=str(y), variable=var,
                             font=FONT_MAIN, bg=APPLE_BG, fg=APPLE_TEXT,
                             activebackground=APPLE_BG, selectcolor=APPLE_CARD,
                             cursor='hand2')
        cb.grid(row=i // 5, column=i % 5, sticky='w', padx=10, pady=4)
        year_vars[y] = var

    def on_start():
        years = sorted(y for y, v in year_vars.items() if v.get())
        if not years:
            messagebox.showwarning("未選擇年份", "請至少選擇一個年份")
            return
        selected_years.extend(years)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())

    btn_frame = tk.Frame(root, bg=APPLE_BG, pady=15)
    btn_frame.pack()
    tk.Button(btn_frame, text="▶  開始萃取",
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=22, pady=9, cursor='hand2',
              command=on_start).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_frame, text="📁  開啟輸出資料夾",
              font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
              activebackground=APPLE_BORDER, relief='flat', padx=22, pady=9,
              cursor='hand2',
              command=lambda: subprocess.Popen(['open', str(DATA_DIR)])).pack(side=tk.LEFT, padx=8)

    root.mainloop()
    return selected_years

# ============================================================
# 進度視窗
# ============================================================
def create_progress_window(years):
    year_label = '、'.join(str(y) for y in years)
    root = tk.Tk()
    root.title(f"🌱 ESG 圖表萃取系統 | {year_label} 年")
    root.geometry("1000x700")
    root.configure(bg=APPLE_BG)
    root.resizable(True, True)
    set_app_icon(root)

    def on_close():
        if program_done.is_set():
            root.destroy()
        elif paused_event.is_set():
            # 暫停中：進度已儲存，可以安全關閉
            program_done.set()
            pause_event.clear()   # 喚醒執行緒讓它讀到 program_done 後自行結束
            root.destroy()
        else:
            messagebox.showinfo(
                "程式仍在執行",
                "程式正在處理中，請稍候。\n\n"
                "點「⏸ 暫停」等目前這份 PDF 處理完後再關閉。"
            )

    root.protocol("WM_DELETE_WINDOW", on_close)

    # --- Header ---
    header = tk.Frame(root, bg=APPLE_BLUE, pady=12)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 圖表萃取系統",
             font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)
    status_dot = tk.Label(header, text='● 初始化', font=FONT_MAIN,
                          fg='#ffdd57', bg=APPLE_BLUE)
    status_dot.pack(side=tk.RIGHT, padx=20)

    # --- 統計卡片 ---
    cards_frame = tk.Frame(root, bg=APPLE_BG, pady=10)
    cards_frame.pack(fill=tk.X, padx=15)

    def make_stat_card(parent, label):
        card = tk.Frame(parent, bg=APPLE_CARD, bd=0,
                        highlightthickness=1, highlightbackground=APPLE_BORDER)
        card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)
        tk.Label(card, text=label, font=FONT_LABEL,
                 fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(8, 0))
        val_var = tk.StringVar(value='—')
        tk.Label(card, textvariable=val_var, font=FONT_STAT,
                 fg=APPLE_TEXT, bg=APPLE_CARD).pack(pady=(0, 8))
        return val_var

    stat_processed = make_stat_card(cards_frame, '已處理')
    stat_images    = make_stat_card(cards_frame, '圖表張數')
    stat_skipped   = make_stat_card(cards_frame, '已跳過')
    stat_error     = make_stat_card(cards_frame, '錯誤')

    # --- 進度條 ---
    prog_frame = tk.Frame(root, bg=APPLE_BG)
    prog_frame.pack(fill=tk.X, padx=20, pady=(0, 8))
    progress_bar = ttk.Progressbar(prog_frame, mode='determinate', length=960)
    progress_bar.pack(fill=tk.X)

    status_frame = tk.Frame(root, bg=APPLE_BG)
    status_frame.pack(fill=tk.X, padx=20)
    last_status_var = tk.StringVar(value='等待開始...')
    tk.Label(status_frame, textvariable=last_status_var,
             font=FONT_LABEL, fg=APPLE_GREY, bg=APPLE_BG, anchor='w').pack(fill=tk.X)

    tk.Frame(root, bg=APPLE_BORDER, height=1).pack(fill=tk.X, padx=15, pady=6)

    # --- Log 區 ---
    log_frame = tk.Frame(root, bg=APPLE_CARD,
                         highlightthickness=1, highlightbackground=APPLE_BORDER)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 8))
    log_text = scrolledtext.ScrolledText(
        log_frame, state='disabled', wrap=tk.WORD,
        font=FONT_LOG, bg=APPLE_CARD, fg=APPLE_TEXT,
        relief='flat', borderwidth=0, padx=8, pady=6)
    log_text.pack(fill=tk.BOTH, expand=True)
    log_text.tag_configure('success', foreground='#1a7f37')
    log_text.tag_configure('error',   foreground='#cf222e')
    log_text.tag_configure('warning', foreground='#9a6700')
    log_text.tag_configure('info',    foreground=APPLE_BLUE)
    log_text.tag_configure('skip',    foreground=APPLE_GREY)

    # --- 底部列 ---
    bottom = tk.Frame(root, bg=APPLE_BG, pady=8)
    bottom.pack(fill=tk.X, padx=15)

    pause_btn_text = tk.StringVar(value='⏸  暫停（目前 PDF 完成後生效）')

    def toggle_pause():
        if pause_event.is_set():
            pause_event.clear()
            pause_btn_text.set('⏸  暫停（目前 PDF 完成後生效）')
        else:
            if messagebox.askyesno("確認暫停", "確定要暫停嗎？\n目前這份 PDF 處理完後會暫停，進度自動儲存。"):
                pause_event.set()
                pause_btn_text.set('▶  繼續執行')

    tk.Button(bottom, textvariable=pause_btn_text,
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=16, pady=7, cursor='hand2', bd=0,
              command=toggle_pause).pack(side=tk.LEFT)
    tk.Button(bottom, text="📁  開啟輸出資料夾",
              font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
              activebackground=APPLE_BORDER, relief='flat', padx=16, pady=7,
              cursor='hand2', bd=0,
              command=lambda: subprocess.Popen(['open', str(DATA_DIR)])).pack(side=tk.LEFT, padx=8)
    time_label = tk.Label(bottom, text='', font=FONT_LABEL,
                          fg=APPLE_GREY, bg=APPLE_BG)
    time_label.pack(side=tk.RIGHT)

    # --- UI 更新 ---
    def update_ui():
        while not log_queue.empty():
            tag, msg = log_queue.get()
            log_text.configure(state='normal')
            ts = datetime.now().strftime('%H:%M:%S')
            log_text.insert(tk.END, f'[{ts}] ', 'skip')
            log_text.insert(tk.END, msg + '\n', tag)
            log_text.see(tk.END)
            log_text.configure(state='disabled')
            if tag in ('success', 'error', 'info', 'warning'):
                last_status_var.set(msg.strip()[:120])

        tot  = ui_stats['total']
        done = ui_stats['done']
        stat_processed.set(f'{done}/{tot}' if tot else '—')
        stat_images.set(str(ui_stats['images']))
        stat_skipped.set(str(ui_stats['skipped']))
        stat_error.set(str(ui_stats['error']) if ui_stats['error'] else '—')

        if tot > 0:
            progress_bar['value'] = done / tot * 100

        if program_done.is_set():
            status_dot.config(text='■ 已完成', fg='#8e8e93')
        elif paused_event.is_set():
            status_dot.config(text='⏸ 已暫停', fg='#ff9f0a')
        elif done > ui_stats['skipped']:
            status_dot.config(text='● 執行中', fg='#34c759')

        time_label.config(text=f'更新時間 {datetime.now().strftime("%H:%M:%S")}')
        root.after(500, update_ui)

    # --- 啟動執行緒 ---
    threading.Thread(target=run_extraction, args=(years,), daemon=True).start()
    update_ui()
    root.mainloop()

# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    years = create_startup_window()
    if years:
        create_progress_window(years)
