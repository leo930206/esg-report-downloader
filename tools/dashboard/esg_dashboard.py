"""
dashboard/esg_dashboard.py
ESG 研究主控台 — 一覽下載進度、圖表萃取、CLIP 分類與 ResNet-50 訓練狀態。
可與所有工具同時執行，自動偵測資料更新。
"""
import os
import sys
import threading
import subprocess
import tkinter as tk
import tkinter.ttk as ttk
from pathlib import Path
from datetime import datetime

import pandas as pd

# ============================================================
# 路徑設定
# ============================================================
BASE_DIR   = Path(__file__).parent.parent.parent.absolute()
DATA_DIR   = BASE_DIR / "data"
CHARTS_DIR = DATA_DIR / "charts"
MODELS_DIR = BASE_DIR / "models"

DOWNLOAD_STATUSES  = ['成功', '未找到中文版報告', '已確認無報告', '下載失敗']
CHART_CATEGORIES   = ['bar', 'line', 'pie', 'map', 'non_chart']
AUTO_REFRESH_MS    = 30_000   # 每 30 秒偵測一次
TOTAL_COMPANIES    = 1078     # 固定分母：台灣上市公司總數

# ============================================================
# Apple 風格配色
# ============================================================
APPLE_BG     = '#f5f5f7'
APPLE_CARD   = '#ffffff'
APPLE_BLUE   = '#0071e3'
APPLE_GREEN  = '#34c759'
APPLE_RED    = '#ff3b30'
APPLE_ORANGE = '#ff9f0a'
APPLE_PURPLE = '#bf5af2'
APPLE_TEXT   = '#1d1d1f'
APPLE_GREY   = '#6e6e73'
APPLE_BORDER = '#d2d2d7'

FONT_TITLE  = ('Helvetica Neue', 13, 'bold')
FONT_MAIN   = ('Helvetica Neue', 10)
FONT_LABEL  = ('Helvetica Neue', 9)
FONT_NUM    = ('Helvetica Neue', 11, 'bold')

# ============================================================
# App Icon
# ============================================================
def set_app_icon(root: tk.Tk) -> None:
    icon_path = Path(__file__).parent.parent.parent / "ESG.png"
    if not icon_path.exists():
        return
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ESG.Report')
        except Exception:
            pass
    else:
        try:
            from AppKit import NSApplication, NSImage
            ns_img = NSImage.alloc().initWithContentsOfFile_(str(icon_path))
            if ns_img:
                NSApplication.sharedApplication().setApplicationIconImage_(ns_img)
        except Exception:
            pass
    try:
        from PIL import Image as PILImage, ImageTk
        photo = ImageTk.PhotoImage(PILImage.open(str(icon_path)))
        root.iconphoto(True, photo)
        root._icon_ref = photo
    except Exception:
        try:
            photo = tk.PhotoImage(file=str(icon_path))
            root.iconphoto(True, photo)
            root._icon_ref = photo
        except Exception:
            pass

# ============================================================
# 資料讀取
# ============================================================
def _file_fingerprint() -> str:
    parts = []
    if not DATA_DIR.is_dir():
        return ''
    for p in sorted(DATA_DIR.glob("*/ESG_Download_Progress_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    for p in sorted(DATA_DIR.glob("*/ESG_Extract_Results_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    for p in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        if p.is_dir():
            parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    # CLIP 分類結果
    xl = CHARTS_DIR / "clip_labeling_results.xlsx"
    if xl.exists():
        parts.append(f"{xl}:{xl.stat().st_mtime:.0f}")
    # ResNet 模型
    pth = MODELS_DIR / "resnet50_chart_best.pth"
    if pth.exists():
        parts.append(f"{pth}:{pth.stat().st_mtime:.0f}")
    return '|'.join(parts)


def load_download_stats() -> dict[str, dict]:
    stats = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        xls  = year_dir / f"ESG_Download_Progress_{year}.xlsx"
        if not xls.exists():
            stats[year] = {'_missing': True}
            continue
        try:
            try:
                df = pd.read_excel(xls, sheet_name='詳細記錄', engine='openpyxl')
            except Exception:
                df = pd.read_excel(xls, engine='openpyxl')
            counts = df['status'].value_counts().to_dict()
            stats[year] = {s: counts.get(s, 0) for s in DOWNLOAD_STATUSES}
            stats[year]['_total'] = len(df)
            stats[year]['_df']    = df
        except Exception as e:
            stats[year] = {'_error': str(e)}
    return stats


def load_cutter_stats() -> dict[str, dict]:
    """用 os.scandir 取代 rglob，速度快 10x 以上。"""
    stats = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        processed_dirs = []
        total_images   = 0
        garbled_files  = []
        try:
            with os.scandir(year_dir) as entries:
                for entry in entries:
                    if not entry.is_dir():
                        continue
                    images_path = os.path.join(entry.path, "images")
                    if os.path.isdir(images_path):
                        try:
                            img_count = sum(
                                1 for f in os.scandir(images_path)
                                if f.name.endswith('.jpg')
                            )
                        except OSError:
                            img_count = 0
                        if img_count > 0:
                            processed_dirs.append(Path(entry.path))
                            total_images += img_count
                    garbled = os.path.join(entry.path, "garbled_pages.txt")
                    if os.path.exists(garbled):
                        garbled_files.append(Path(garbled))
        except OSError:
            pass

        all_pdf_stems   = {p.stem for p in year_dir.glob("*.pdf")}
        processed_stems = {d.name for d in processed_dirs}
        pending = len(all_pdf_stems - processed_stems)

        stats[year] = {
            'processed':      len(processed_dirs),
            'pending':        pending,
            'images':         total_images,
            'garbled':        len(garbled_files),
            'garbled_files':  garbled_files,
            'processed_dirs': processed_dirs,
        }
    return stats


def load_classifier_stats() -> dict:
    """
    讀取 CLIP 分類結果。
    優先讀 clip_labeling_results.xlsx（Sheet1：年度統計），
    fallback：掃描 data/charts/{category}/ 目錄，從檔名解析年份。

    回傳：
      {
        'by_year': { year: {cat: int, 'total': int}, ... },
        'by_cat':  { cat: int, ... },
        'total':   int,
        'source':  'excel' | 'scan' | 'empty',
      }
    """
    empty = {
        'by_year': {},
        'by_cat':  {c: 0 for c in CHART_CATEGORIES},
        'total':   0,
        'source':  'empty',
    }

    if not CHARTS_DIR.is_dir():
        return empty

    # ── 優先讀 Excel ──
    xl = CHARTS_DIR / "clip_labeling_results.xlsx"
    if xl.exists():
        try:
            df = pd.read_excel(xl, sheet_name=0, engine='openpyxl')
            # Sheet1 格式：年份 | bar | line | pie | map | 合計
            by_year: dict[str, dict] = {}
            by_cat:  dict[str, int]  = {c: 0 for c in CHART_CATEGORIES}
            for _, row in df.iterrows():
                year = str(row.iloc[0]).strip()
                if not year.isdigit():
                    continue
                yr_dict: dict[str, int] = {}
                for cat in CHART_CATEGORIES:
                    # 欄位名稱可能是中英文，按位置取
                    val = 0
                    if cat in df.columns:
                        val = int(row[cat]) if pd.notna(row[cat]) else 0
                    yr_dict[cat] = val
                    by_cat[cat] += val
                yr_dict['total'] = sum(yr_dict.values())
                by_year[year] = yr_dict
            total = sum(by_cat.values())
            return {'by_year': by_year, 'by_cat': by_cat, 'total': total, 'source': 'excel'}
        except Exception:
            pass  # fallback to scan

    # ── Fallback：掃描目錄 ──
    by_year: dict[str, dict] = {}
    by_cat:  dict[str, int]  = {c: 0 for c in CHART_CATEGORIES}
    for cat in CHART_CATEGORIES:
        cat_dir = CHARTS_DIR / cat
        if not cat_dir.is_dir():
            continue
        try:
            for f in os.scandir(cat_dir):
                if not f.name.endswith('.jpg'):
                    continue
                # 檔名格式：{year}_{company}_{filename}.jpg
                parts = f.name.split('_', 1)
                year  = parts[0] if parts[0].isdigit() and len(parts[0]) == 4 else 'unknown'
                by_year.setdefault(year, {c: 0 for c in CHART_CATEGORIES})
                by_year.setdefault(year, {})['total'] = by_year.get(year, {}).get('total', 0)
                by_year[year][cat] = by_year[year].get(cat, 0) + 1
                by_cat[cat] += 1
        except OSError:
            pass

    # 補 total
    for yr in by_year:
        by_year[yr]['total'] = sum(by_year[yr].get(c, 0) for c in CHART_CATEGORIES)

    total = sum(by_cat.values())
    source = 'scan' if total > 0 else 'empty'
    return {'by_year': by_year, 'by_cat': by_cat, 'total': total, 'source': source}


def load_trainer_stats() -> dict:
    """
    讀取 ResNet-50 訓練狀態。
    回傳：
      {
        'model_exists':  bool,
        'best_val_acc':  float | None,
        'best_epoch':    int | None,
        'epochs_done':   int,
        'data_counts':   { cat: int },   # data/charts/{cat}/ 下的圖片數
        'data_total':    int,
        'data_no_nonchart': int,          # 不含 non_chart 的訓練圖片數
      }
    """
    model_exists = (MODELS_DIR / "resnet50_chart_best.pth").exists()

    best_val_acc = None
    best_epoch   = None
    epochs_done  = 0
    log_path = MODELS_DIR / "training_log.csv"
    if log_path.exists():
        try:
            import csv
            with open(log_path, newline='', encoding='utf-8') as f:
                rows = list(csv.DictReader(f))
            if rows:
                epochs_done = len(rows)
                best_row = max(rows, key=lambda r: float(r.get('val_acc', 0)))
                best_val_acc = float(best_row['val_acc'])
                best_epoch   = int(best_row['epoch'])
        except Exception:
            pass

    # 計算 data/charts/ 各類圖片數
    data_counts: dict[str, int] = {c: 0 for c in CHART_CATEGORIES}
    if CHARTS_DIR.is_dir():
        for cat in CHART_CATEGORIES:
            cat_dir = CHARTS_DIR / cat
            if cat_dir.is_dir():
                try:
                    data_counts[cat] = sum(
                        1 for f in os.scandir(cat_dir) if f.name.endswith('.jpg')
                    )
                except OSError:
                    pass

    data_total = sum(data_counts.values())
    data_no_nonchart = data_total - data_counts.get('non_chart', 0)

    return {
        'model_exists':      model_exists,
        'best_val_acc':      best_val_acc,
        'best_epoch':        best_epoch,
        'epochs_done':       epochs_done,
        'data_counts':       data_counts,
        'data_total':        data_total,
        'data_no_nonchart':  data_no_nonchart,
    }

# ============================================================
# 進度條工具
# ============================================================
def _bar(pct: int, n: int = 10) -> str:
    filled = round(pct / 100 * n)
    return f"{'█' * filled}{'░' * (n - filled)}  {pct:3d}%"

# ============================================================
# 細節視窗
# ============================================================
class DetailWindow:
    _instances: dict[str, 'DetailWindow'] = {}

    @classmethod
    def open(cls, year: str, dl_row: dict, ct_row: dict):
        if year in cls._instances:
            try:
                cls._instances[year].win.lift()
                return
            except tk.TclError:
                pass
        inst = cls(year, dl_row, ct_row)
        cls._instances[year] = inst

    def __init__(self, year: str, dl_row: dict, ct_row: dict):
        self.year = year
        self.win  = tk.Toplevel()
        self.win.title(f"📋 {year} 年度明細")
        self.win.geometry("860x560")
        self.win.configure(bg=APPLE_BG)
        self.win.protocol("WM_DELETE_WINDOW",
                          lambda: (DetailWindow._instances.pop(year, None),
                                   self.win.destroy()))
        self._build(dl_row, ct_row)

    def _build(self, dl_row: dict, ct_row: dict):
        hdr = tk.Frame(self.win, bg=APPLE_BLUE, pady=8)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"{self.year} 年度明細",
                 font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=16)

        sf = tk.Frame(self.win, bg=APPLE_BG, pady=8)
        sf.pack(fill=tk.X, padx=16)
        tk.Label(sf, text="搜尋：", font=FONT_MAIN, bg=APPLE_BG).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        tk.Entry(sf, textvariable=self.search_var, font=FONT_MAIN, width=28,
                 relief='solid', bd=1).pack(side=tk.LEFT, padx=(4, 12))
        tk.Label(sf, text="狀態：", font=FONT_MAIN, bg=APPLE_BG).pack(side=tk.LEFT)
        self.filter_var = tk.StringVar(value='全部')
        ttk.Combobox(sf, textvariable=self.filter_var,
                     values=['全部'] + DOWNLOAD_STATUSES,
                     width=16, state='readonly').pack(side=tk.LEFT)
        self.search_var.trace_add('write', lambda *_: self._apply_filter())
        self.filter_var.trace_add('write', lambda *_: self._apply_filter())

        cols = ('公司代碼', '公司名稱', '下載狀態', '已萃取', '圖片數', '亂碼頁')
        frame = tk.Frame(self.win, bg=APPLE_BG)
        frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))

        vsb = ttk.Scrollbar(frame, orient='vertical')
        hsb = ttk.Scrollbar(frame, orient='horizontal')
        self.tree = ttk.Treeview(
            frame, columns=cols, show='headings',
            yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=20
        )
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        for col, w in zip(cols, [80, 180, 130, 70, 70, 70]):
            self.tree.heading(col, text=col, command=lambda c=col: self._sort(c))
            self.tree.column(col, width=w, anchor='center')

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind('<MouseWheel>',
                       lambda e: (self.tree.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])

        self._build_rows(dl_row, ct_row)
        self._apply_filter()
        tk.Label(self.win, text=f"共 {len(self.all_rows)} 筆",
                 font=FONT_LABEL, fg=APPLE_GREY, bg=APPLE_BG).pack(pady=(0, 8))

    def _build_rows(self, dl_row: dict, ct_row: dict):
        processed_stems = {d.name for d in ct_row.get('processed_dirs', [])}
        img_counts = {
            d.name: len(list((d / "images").glob("*.jpg")))
            for d in ct_row.get('processed_dirs', [])
        }
        garbled_map: dict[str, str] = {}
        for gf in ct_row.get('garbled_files', []):
            try:
                garbled_map[gf.parent.name] = gf.read_text(encoding='utf-8').strip()
            except Exception:
                garbled_map[gf.parent.name] = '?'

        self.all_rows = []
        df = dl_row.get('_df')
        if df is not None:
            for _, r in df.iterrows():
                raw    = str(r.get('file_name', r.get('filename', r.get('檔名', ''))))
                stem   = Path(raw).stem if raw else ''
                code   = str(r.get('stock_id', r.get('stock_code', r.get('代碼', ''))))
                name   = str(r.get('company_name', r.get('公司名稱', '')))
                status = str(r.get('status', r.get('狀態', '')))
                extracted = '✅ 已萃取' if stem in processed_stems else (
                    '—' if status != '成功' else '⏳ 待萃取')
                imgs = img_counts.get(stem, 0)
                self.all_rows.append((
                    code, name, status, extracted,
                    str(imgs) if imgs else '—',
                    garbled_map.get(stem, '—'),
                ))

    def _apply_filter(self):
        kw     = self.search_var.get().strip().lower()
        status = self.filter_var.get()
        self.tree.delete(*self.tree.get_children())
        for row in self.all_rows:
            if status != '全部' and row[2] != status:
                continue
            if kw and not any(kw in str(c).lower() for c in row):
                continue
            tag = 'ok' if row[2] == '成功' else (
                  'fail' if row[2] == '下載失敗' else 'other')
            self.tree.insert('', 'end', values=row, tags=(tag,))
        self.tree.tag_configure('ok',    foreground=APPLE_GREEN)
        self.tree.tag_configure('fail',  foreground=APPLE_RED)
        self.tree.tag_configure('other', foreground=APPLE_GREY)

    _sort_reverse: dict[str, bool] = {}
    def _sort(self, col: str):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        rev = self._sort_reverse.get(col, False)
        items.sort(reverse=rev)
        for i, (_, k) in enumerate(items):
            self.tree.move(k, '', i)
        self._sort_reverse[col] = not rev
        self.tree.heading(col, text=col + (' ↑' if not rev else ' ↓'))

# ============================================================
# 主控台
# ============================================================
class Dashboard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ESG 研究主控台")
        self.root.geometry("980x820")
        self.root.configure(bg=APPLE_BG)
        self.root.resizable(True, True)
        set_app_icon(self.root)

        self._last_fingerprint   = ''
        self._dl_stats:  dict    = {}
        self._ct_stats:  dict    = {}
        self._clf_stats: dict    = {}
        self._trn_stats: dict    = {}

        self._setup_treeview_style()
        self._build_ui()
        self.refresh(force=True)
        self._schedule_auto_refresh()

    def _setup_treeview_style(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure('ESG.Treeview',
                        background=APPLE_CARD,
                        foreground=APPLE_TEXT,
                        fieldbackground=APPLE_CARD,
                        rowheight=28,
                        font=('Helvetica Neue', 10))
        style.configure('ESG.Treeview.Heading',
                        background='#e8e8ed',
                        foreground=APPLE_GREY,
                        font=('Helvetica Neue', 9, 'bold'),
                        relief='flat')
        style.map('ESG.Treeview',
                  background=[('selected', '#d0e8ff')],
                  foreground=[('selected', APPLE_TEXT)])
        style.map('ESG.Treeview.Heading',
                  background=[('active', '#d2d2d7')])

    # ── Header ──────────────────────────────────────────────
    def _build_ui(self):
        header = tk.Frame(self.root, bg=APPLE_BLUE, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="📊  ESG 研究主控台",
                 font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)

        self.status_dot = tk.Label(header, text='●', font=FONT_MAIN,
                                   fg='#a8d4ff', bg=APPLE_BLUE)
        self.status_dot.pack(side=tk.RIGHT, padx=(0, 8))
        self.last_updated = tk.Label(header, text='', font=FONT_LABEL,
                                     fg='#a8d4ff', bg=APPLE_BLUE)
        self.last_updated.pack(side=tk.RIGHT, padx=(0, 4))

        for sym, label, cmd in [
            ('↺', '重新整理',    lambda: self.refresh(force=True)),
            ('📁', '開啟輸出資料夾', lambda: subprocess.Popen(['open', str(DATA_DIR)])),
        ]:
            btn = tk.Frame(header, bg=APPLE_CARD, cursor='hand2',
                           highlightthickness=1, highlightbackground='#ccc')
            sl  = tk.Label(btn, text=sym,   font=FONT_MAIN, fg=APPLE_BLUE,
                           bg=APPLE_CARD, padx=8, pady=3)
            tl  = tk.Label(btn, text=label, font=FONT_MAIN, fg=APPLE_TEXT,
                           bg=APPLE_CARD, padx=4, pady=3)
            sl.pack(side=tk.LEFT)
            tl.pack(side=tk.LEFT, padx=(0, 8))
            def _enter(_, b=btn, s=sl, t=tl):
                for w in (b, s, t): w.config(bg='#e8e8ed')
            def _leave(_, b=btn, s=sl, t=tl):
                for w in (b, s, t): w.config(bg=APPLE_CARD)
            c = cmd
            for w in (btn, sl, tl):
                w.bind('<Enter>',    _enter)
                w.bind('<Leave>',    _leave)
                w.bind('<Button-1>', lambda _, c=c: c())
            btn.pack(side=tk.RIGHT, padx=4)

        # ── Scrollable body ──
        body_outer = tk.Frame(self.root, bg=APPLE_BG)
        body_outer.pack(fill=tk.BOTH, expand=True)

        canvas    = tk.Canvas(body_outer, bg=APPLE_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.body = tk.Frame(canvas, bg=APPLE_BG)
        self.body_window = canvas.create_window((0, 0), window=self.body, anchor='nw')

        self.body.bind('<Configure>',
                       lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(self.body_window, width=e.width))
        self._canvas = canvas
        self.root.bind_all('<MouseWheel>', self._on_canvas_scroll)

    def _on_canvas_scroll(self, e):
        if e.widget.winfo_toplevel() is self.root:
            self._canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

    # ── 自動重整 ────────────────────────────────────────────
    def _schedule_auto_refresh(self):
        self.refresh(force=False)
        self.root.after(AUTO_REFRESH_MS, self._schedule_auto_refresh)

    def refresh(self, force: bool = True):
        self.status_dot.config(fg='#ffcc00')

        def _load():
            fp = _file_fingerprint()
            if not force and fp == self._last_fingerprint:
                self.root.after(0, lambda: self.status_dot.config(fg='#a8d4ff'))
                return
            dl  = load_download_stats()
            ct  = load_cutter_stats()
            clf = load_classifier_stats()
            trn = load_trainer_stats()
            self.root.after(0, lambda: self._render(dl, ct, clf, trn, fp))

        threading.Thread(target=_load, daemon=True).start()

    def _render(self, dl_stats: dict, ct_stats: dict,
                clf_stats: dict, trn_stats: dict, fingerprint: str):
        self._last_fingerprint = fingerprint

        for y in range(2015, 2025):
            yr = str(y)
            if yr not in dl_stats:
                dl_stats[yr] = {'_missing': True}
            if yr not in ct_stats:
                ct_stats[yr] = {
                    'processed': 0, 'pending': 0, 'images': 0,
                    'garbled': 0, 'garbled_files': [], 'processed_dirs': [],
                }

        self._dl_stats  = dl_stats
        self._ct_stats  = ct_stats
        self._clf_stats = clf_stats
        self._trn_stats = trn_stats

        for w in self.body.winfo_children():
            w.destroy()

        self._build_download_section(dl_stats)
        self._build_cutter_section(ct_stats)
        self._build_classifier_section(clf_stats, ct_stats)
        self._build_trainer_section(trn_stats)
        tk.Frame(self.body, bg=APPLE_BG, height=20).pack()

        self.last_updated.config(text=f"更新 {datetime.now().strftime('%H:%M:%S')}")
        self.status_dot.config(fg='#5cff7a')

    # ── 區塊標題 ────────────────────────────────────────────
    def _section_title(self, text, subtext=''):
        f = tk.Frame(self.body, bg=APPLE_BG)
        f.pack(fill=tk.X, padx=20, pady=(24, 8))
        tk.Label(f, text=text, font=('Helvetica Neue', 12, 'bold'),
                 fg=APPLE_TEXT, bg=APPLE_BG).pack(side=tk.LEFT)
        if subtext:
            tk.Label(f, text=subtext, font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_BG).pack(side=tk.LEFT, padx=(8, 0))
        tk.Frame(self.body, bg=APPLE_BORDER, height=1).pack(fill=tk.X, padx=20, pady=(0, 10))

    # ── 下載狀態表 ───────────────────────────────────────────
    def _build_download_section(self, stats: dict):
        self._section_title("Step 1  下載狀態",
                            "來源：ESG_Download_Progress_YYYY.xlsx  ·  點選列查看明細")
        frame = tk.Frame(self.body, bg=APPLE_BG)
        frame.pack(fill=tk.X, padx=20, pady=(0, 4))

        cols = ('年度', '✅ 成功', '⚠️ 未找到', '🔒 已確認無', '❌ 失敗', '總共爬蟲數', '進度')
        col_widths = [60, 70, 80, 90, 60, 90, 300]
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            height=len(stats) + 1, style='ESG.Treeview',
                            selectmode='browse')
        for col, w in zip(cols, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w,
                        anchor='w' if col == '進度' else 'center',
                        stretch=(col == '進度'))

        tot = {'成功': 0, '未找到中文版報告': 0, '已確認無報告': 0, '下載失敗': 0, '_total': 0}
        for year, s in sorted(stats.items()):
            if s.get('_missing'):
                tree.insert('', 'end', values=(year, '—', '—', '—', '—', '—', '尚無資料'),
                            tags=('missing',))
                continue
            if s.get('_error'):
                tree.insert('', 'end', values=(year, '錯誤', '', '', '', '', s['_error']),
                            tags=('error',))
                continue
            success  = s.get('成功', 0)
            crawled  = s.get('_total', 0)
            pct      = int(crawled / TOTAL_COMPANIES * 100)
            bar      = _bar(pct)
            tag = 'done' if crawled >= TOTAL_COMPANIES else ('partial' if pct > 0 else 'none')
            tree.insert('', 'end', values=(
                year, success,
                s.get('未找到中文版報告', 0),
                s.get('已確認無報告', 0),
                s.get('下載失敗', 0) or '—',
                crawled, bar,
            ), tags=(tag,))
            for k in tot:
                tot[k] += s.get(k, 0)

        tot_pct_dl = int(tot['_total'] / 10780 * 100)
        tree.insert('', 'end', values=(
            '總計', tot['成功'], tot['未找到中文版報告'],
            tot['已確認無報告'], tot['下載失敗'] or '—',
            tot['_total'], _bar(tot_pct_dl),
        ), tags=('total',))

        tree.tag_configure('done',    foreground=APPLE_GREEN)
        tree.tag_configure('partial', foreground=APPLE_BLUE)
        tree.tag_configure('none',    foreground=APPLE_GREY)
        tree.tag_configure('missing', foreground=APPLE_GREY)
        tree.tag_configure('error',   foreground=APPLE_RED)
        tree.tag_configure('total',   foreground=APPLE_TEXT, font=('Helvetica Neue', 10, 'bold'))

        def _on_click(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            vals = tree.item(iid, 'values')
            yr   = vals[0]
            if yr == '總計':
                return
            DetailWindow.open(yr, stats.get(yr, {}), self._ct_stats.get(yr, {}))
        tree.bind('<Button-1>', _on_click)
        tree.bind('<MouseWheel>',
                  lambda e: (self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])
        tree.pack(fill=tk.X)
        tk.Label(frame, text='年度進度 = 爬蟲數 ÷ 1078　　總計進度 = 爬蟲數總和 ÷ 10780',
                 font=('Helvetica Neue', 9), fg=APPLE_GREY, bg=APPLE_BG).pack(anchor='e', pady=(2, 0))

    # ── 圖表萃取狀態表 ───────────────────────────────────────
    def _build_cutter_section(self, stats: dict):
        self._section_title("Step 2  圖表萃取狀態",
                            "來源：掃描 data/{year}/*/images/*.jpg  ·  點選列查看明細")
        frame = tk.Frame(self.body, bg=APPLE_BG)
        frame.pack(fill=tk.X, padx=20, pady=(0, 4))

        cols = ('年度', '✅ 已萃取', '🖼 圖片數', '⚠️ 亂碼公司', '進度')
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            height=len(stats) + 1, style='ESG.Treeview',
                            selectmode='browse')
        for col, w in zip(cols, [60, 130, 130, 130, 300]):
            tree.heading(col, text=col)
            tree.column(col, width=w,
                        anchor='w' if col == '進度' else 'center',
                        stretch=(col == '進度'))

        tot_processed = 0
        tot_images    = 0
        tot_garbled   = 0
        tot_dl        = 0

        for year, s in sorted(stats.items()):
            processed  = s['processed']
            dl_success = self._dl_stats.get(year, {}).get('成功', 0)
            pct        = int(processed / dl_success * 100) if dl_success else 0

            tot_processed += processed
            tot_images    += s['images']
            tot_garbled   += s['garbled']
            tot_dl        += dl_success

            if dl_success == 0:
                bar = '尚無資料'
                tag = 'missing'
            elif pct >= 100:
                bar = _bar(100)
                tag = 'done'
            elif pct == 0:
                bar = '尚未開始'
                tag = 'none'
            else:
                bar = _bar(pct)
                tag = 'partial'

            tree.insert('', 'end', values=(
                year,
                processed or '—',
                f"{s['images']:,}" if s['images'] else '—',
                s['garbled'] or '—',
                bar,
            ), tags=(tag,))

        tot_pct = int(tot_processed / tot_dl * 100) if tot_dl else 0
        tree.insert('', 'end', values=(
            '總計',
            tot_processed or '—',
            f"{tot_images:,}" if tot_images else '—',
            tot_garbled or '—',
            _bar(tot_pct) if tot_dl else '—',
        ), tags=('total',))

        tree.tag_configure('done',    foreground=APPLE_GREEN)
        tree.tag_configure('partial', foreground=APPLE_BLUE)
        tree.tag_configure('none',    foreground=APPLE_GREY)
        tree.tag_configure('missing', foreground=APPLE_GREY)
        tree.tag_configure('total',   foreground=APPLE_TEXT, font=('Helvetica Neue', 10, 'bold'))

        def _on_click(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            vals = tree.item(iid, 'values')
            year = vals[0]
            if year == '總計':
                return
            DetailWindow.open(year, self._dl_stats.get(year, {}), stats.get(year, {}))
        tree.bind('<Button-1>', _on_click)
        tree.bind('<MouseWheel>',
                  lambda e: (self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])
        tree.pack(fill=tk.X)
        tk.Label(frame, text='年度進度 = 已萃取 ÷ 該年下載成功數　　總計進度 = 已萃取總和 ÷ 下載成功總和',
                 font=('Helvetica Neue', 9), fg=APPLE_GREY, bg=APPLE_BG).pack(anchor='e', pady=(2, 0))

    # ── CLIP 分類狀態表 ──────────────────────────────────────
    def _build_classifier_section(self, clf_stats: dict, ct_stats: dict):
        source = clf_stats.get('source', 'empty')
        source_note = {
            'excel': '來源：data/charts/clip_labeling_results.xlsx',
            'scan':  '來源：掃描 data/charts/{category}/ 目錄（Excel 尚未產生）',
            'empty': '尚未執行 clip_classifier.py',
        }.get(source, '')
        self._section_title("Step 3  CLIP 分類狀態", source_note)

        frame = tk.Frame(self.body, bg=APPLE_BG)
        frame.pack(fill=tk.X, padx=20, pady=(0, 4))

        if source == 'empty':
            tk.Label(frame,
                     text="尚無分類資料。請先完成 Step 2 萃取，再執行：\npython tools/chart-classifier/clip_classifier.py",
                     font=FONT_MAIN, fg=APPLE_GREY, bg=APPLE_BG,
                     justify='left').pack(anchor='w', pady=8)
            return

        # 總計橫條（各類別）
        by_cat = clf_stats.get('by_cat', {})
        total  = clf_stats.get('total', 0)
        if total > 0:
            cat_frame = tk.Frame(frame, bg=APPLE_BG)
            cat_frame.pack(fill=tk.X, pady=(0, 8))
            cat_labels = {
                'bar': '長條圖', 'line': '折線圖', 'pie': '圓餅圖',
                'map': '地圖', 'non_chart': '非圖表',
            }
            cat_colors = {
                'bar': APPLE_BLUE, 'line': APPLE_GREEN, 'pie': APPLE_ORANGE,
                'map': APPLE_PURPLE, 'non_chart': APPLE_GREY,
            }
            for cat in CHART_CATEGORIES:
                cnt = by_cat.get(cat, 0)
                pct = int(cnt / total * 100) if total else 0
                card = tk.Frame(cat_frame, bg=APPLE_CARD,
                                highlightthickness=1, highlightbackground=APPLE_BORDER)
                card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=3)
                tk.Label(card, text=cat_labels[cat], font=FONT_LABEL,
                         fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(6, 0))
                tk.Label(card, text=f"{cnt:,}", font=('Helvetica Neue', 14, 'bold'),
                         fg=cat_colors[cat], bg=APPLE_CARD).pack()
                tk.Label(card, text=f"{pct}%", font=FONT_LABEL,
                         fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(0, 6))

        # 年份明細表
        by_year = clf_stats.get('by_year', {})
        if not by_year:
            return

        # 計算進度分母：各年度的萃取圖片總數
        ct_images_by_year = {yr: ct_stats.get(yr, {}).get('images', 0)
                              for yr in range(2015, 2025)}

        cols = ('年度', '長條圖', '折線圖', '圓餅圖', '地圖', '非圖表', '合計', '進度')
        all_years = sorted(set(list(by_year.keys()) + [str(y) for y in range(2015, 2025)]))
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            height=len(all_years) + 1, style='ESG.Treeview',
                            selectmode='none')
        col_widths = [60, 70, 70, 70, 70, 70, 70, 270]
        for col, w in zip(cols, col_widths):
            tree.heading(col, text=col)
            tree.column(col, width=w,
                        anchor='w' if col == '進度' else 'center',
                        stretch=(col == '進度'))

        tot = {c: 0 for c in CHART_CATEGORIES}
        tot_classified = 0
        tot_images_all = 0

        for year in all_years:
            yr_data   = by_year.get(year, {})
            classified = yr_data.get('total', 0)
            total_imgs = ct_images_by_year.get(int(year) if year.isdigit() else 0, 0)

            if classified == 0 and total_imgs == 0:
                tree.insert('', 'end',
                            values=(year, '—', '—', '—', '—', '—', '—', '尚無資料'),
                            tags=('missing',))
                continue

            pct = int(classified / total_imgs * 100) if total_imgs else 0
            if pct == 0:
                bar = '尚未分類'
                tag = 'none'
            elif pct >= 100:
                bar = _bar(100)
                tag = 'done'
            else:
                bar = _bar(pct)
                tag = 'partial'

            row_vals = (
                year,
                yr_data.get('bar', 0) or '—',
                yr_data.get('line', 0) or '—',
                yr_data.get('pie', 0) or '—',
                yr_data.get('map', 0) or '—',
                yr_data.get('non_chart', 0) or '—',
                classified or '—',
                bar,
            )
            tree.insert('', 'end', values=row_vals, tags=(tag,))

            for c in CHART_CATEGORIES:
                tot[c] += yr_data.get(c, 0)
            tot_classified += classified
            tot_images_all += total_imgs

        tot_pct = int(tot_classified / tot_images_all * 100) if tot_images_all else 0
        tree.insert('', 'end', values=(
            '總計',
            tot['bar'] or '—', tot['line'] or '—',
            tot['pie'] or '—', tot['map'] or '—',
            tot['non_chart'] or '—',
            tot_classified or '—',
            _bar(tot_pct) if tot_images_all else '—',
        ), tags=('total',))

        tree.tag_configure('done',    foreground=APPLE_GREEN)
        tree.tag_configure('partial', foreground=APPLE_BLUE)
        tree.tag_configure('none',    foreground=APPLE_GREY)
        tree.tag_configure('missing', foreground=APPLE_GREY)
        tree.tag_configure('total',   foreground=APPLE_TEXT, font=('Helvetica Neue', 10, 'bold'))
        tree.bind('<MouseWheel>',
                  lambda e: (self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])
        tree.pack(fill=tk.X)
        tk.Label(frame, text='年度進度 = 已分類圖片 ÷ 該年萃取圖片數',
                 font=('Helvetica Neue', 9), fg=APPLE_GREY, bg=APPLE_BG).pack(anchor='e', pady=(2, 0))

    # ── ResNet-50 訓練狀態 ───────────────────────────────────
    def _build_trainer_section(self, trn_stats: dict):
        self._section_title("Step 4  ResNet-50 訓練狀態",
                            "來源：models/resnet50_chart_best.pth + models/training_log.csv")
        frame = tk.Frame(self.body, bg=APPLE_BG, padx=20)
        frame.pack(fill=tk.X, pady=(0, 8))

        # ── 訓練資料卡片 ──
        data_frame = tk.Frame(frame, bg=APPLE_BG)
        data_frame.pack(fill=tk.X, pady=(0, 10))

        data_counts     = trn_stats.get('data_counts', {})
        data_total      = trn_stats.get('data_total', 0)
        data_no_nc      = trn_stats.get('data_no_nonchart', 0)
        model_exists    = trn_stats.get('model_exists', False)
        best_val_acc    = trn_stats.get('best_val_acc')
        best_epoch      = trn_stats.get('best_epoch')
        epochs_done     = trn_stats.get('epochs_done', 0)

        cat_labels = {
            'bar': '長條圖', 'line': '折線圖', 'pie': '圓餅圖',
            'map': '地圖', 'non_chart': '非圖表',
        }
        cat_colors = {
            'bar': APPLE_BLUE, 'line': APPLE_GREEN, 'pie': APPLE_ORANGE,
            'map': APPLE_PURPLE, 'non_chart': APPLE_GREY,
        }

        # 各類別圖片數卡片
        cards_frame = tk.Frame(data_frame, bg=APPLE_BG)
        cards_frame.pack(fill=tk.X)
        for cat in CHART_CATEGORIES:
            cnt = data_counts.get(cat, 0)
            card = tk.Frame(cards_frame, bg=APPLE_CARD,
                            highlightthickness=1, highlightbackground=APPLE_BORDER)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=3)
            tk.Label(card, text=cat_labels[cat], font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(6, 0))
            tk.Label(card, text=f"{cnt:,}" if cnt else '—',
                     font=('Helvetica Neue', 14, 'bold'),
                     fg=cat_colors[cat] if cnt else APPLE_GREY,
                     bg=APPLE_CARD).pack()
            tk.Label(card, text='張', font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(0, 6))

        # 訓練狀態資訊列
        info_frame = tk.Frame(frame, bg=APPLE_CARD,
                              highlightthickness=1, highlightbackground=APPLE_BORDER)
        info_frame.pack(fill=tk.X, pady=(8, 0))

        # 左：訓練資料統計
        left = tk.Frame(info_frame, bg=APPLE_CARD, padx=16, pady=12)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(left, text="訓練資料", font=FONT_LABEL,
                 fg=APPLE_GREY, bg=APPLE_CARD).pack(anchor='w')
        if data_total == 0:
            tk.Label(left, text="尚無資料  →  請先執行 clip_classifier.py",
                     font=FONT_MAIN, fg=APPLE_GREY, bg=APPLE_CARD).pack(anchor='w')
        else:
            ready = data_no_nc >= 50  # 每類至少各10張才算可訓練
            status_text = f"共 {data_total:,} 張（訓練用：{data_no_nc:,} 張）"
            hint = "  ✅ 可執行訓練" if ready else "  ⚠️ 建議各類至少各 50 張再訓練"
            tk.Label(left, text=status_text + hint,
                     font=FONT_MAIN,
                     fg=APPLE_GREEN if ready else APPLE_ORANGE,
                     bg=APPLE_CARD).pack(anchor='w')

        # 右：模型狀態
        right = tk.Frame(info_frame, bg=APPLE_CARD, padx=16, pady=12)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(right, text="模型狀態", font=FONT_LABEL,
                 fg=APPLE_GREY, bg=APPLE_CARD).pack(anchor='w')

        if not model_exists and epochs_done == 0:
            tk.Label(right, text="未訓練  →  執行 resnet_trainer.py 開始訓練",
                     font=FONT_MAIN, fg=APPLE_GREY, bg=APPLE_CARD).pack(anchor='w')
        elif model_exists:
            acc_text = f"最佳 val acc：{best_val_acc:.1f}%（第 {best_epoch} / {epochs_done} epoch）" \
                       if best_val_acc is not None else f"已訓練 {epochs_done} epochs"
            tk.Label(right,
                     text=f"✅ 已有模型  ·  {acc_text}",
                     font=FONT_MAIN, fg=APPLE_GREEN, bg=APPLE_CARD).pack(anchor='w')
        else:
            tk.Label(right, text=f"訓練中？（log 顯示 {epochs_done} epoch，但尚無 .pth）",
                     font=FONT_MAIN, fg=APPLE_ORANGE, bg=APPLE_CARD).pack(anchor='w')

    def run(self):
        self.root.mainloop()


# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    Dashboard().run()
