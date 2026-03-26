"""
dashboard/esg-dashboard.py
ESG 研究主控台 — 一覽下載進度與圖表萃取狀態。
可與 esg_downloader.py / esg_pdf_cuter.py 同時執行，自動偵測資料更新。
"""
import os
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
BASE_DIR = Path(__file__).parent.parent.absolute()
DATA_DIR = BASE_DIR / "data"

DOWNLOAD_STATUSES  = ['成功', '未找到中文版報告', '已確認無報告', '下載失敗']
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
def set_app_icon(root: tk.Tk, emoji: str = "📊") -> None:
    try:
        from AppKit import NSApplication, NSImage, NSAttributedString, NSFont
        from Foundation import NSMakeSize
        size   = 256
        ns_img = NSImage.alloc().initWithSize_(NSMakeSize(size, size))
        ns_img.lockFocus()
        attrs  = {"NSFont": NSFont.systemFontOfSize_(200)}
        s      = NSAttributedString.alloc().initWithString_attributes_(emoji, attrs)
        s.drawAtPoint_((20, 20))
        ns_img.unlockFocus()
        NSApplication.sharedApplication().setApplicationIconImage_(ns_img)
        import base64
        from io import BytesIO
        from PIL import Image as PILImage
        tiff  = ns_img.TIFFRepresentation()
        pil   = PILImage.open(BytesIO(bytes(tiff)))
        buf   = BytesIO()
        pil.save(buf, format="PNG")
        photo = tk.PhotoImage(data=base64.b64encode(buf.getvalue()).decode())
        root.iconphoto(True, photo)
        root._icon_ref = photo
    except Exception:
        pass

# ============================================================
# 資料讀取（效能優化：os.scandir + 淺層 glob）
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

        all_pdf_stems  = {p.stem for p in year_dir.glob("*.pdf")}
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
        self.root.title("📊 ESG 研究主控台")
        self.root.geometry("980x720")
        self.root.configure(bg=APPLE_BG)
        self.root.resizable(True, True)
        set_app_icon(self.root)

        self._last_fingerprint = ''
        self._dl_stats: dict = {}
        self._ct_stats: dict = {}

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
        # 只在事件來自主視窗時捲動（避免影響 DetailWindow）
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
            dl = load_download_stats()
            ct = load_cutter_stats()
            self.root.after(0, lambda: self._render(dl, ct, fp))

        threading.Thread(target=_load, daemon=True).start()

    def _render(self, dl_stats: dict, ct_stats: dict, fingerprint: str):
        self._last_fingerprint = fingerprint
        self._dl_stats = dl_stats
        self._ct_stats = ct_stats

        for w in self.body.winfo_children():
            w.destroy()

        self._build_summary(dl_stats, ct_stats)
        self._build_download_section(dl_stats)
        self._build_cutter_section(ct_stats)
        tk.Frame(self.body, bg=APPLE_BG, height=20).pack()

        self.last_updated.config(text=f"更新 {datetime.now().strftime('%H:%M:%S')}")
        self.status_dot.config(fg='#5cff7a')

    # ── 摘要卡 ──────────────────────────────────────────────
    def _build_summary(self, dl_stats: dict, ct_stats: dict):
        frame = tk.Frame(self.body, bg=APPLE_BG, padx=20, pady=10)
        frame.pack(fill=tk.X)

        total_success   = sum(s.get('成功', 0) for s in dl_stats.values()
                              if not s.get('_missing') and not s.get('_error'))
        total_processed = sum(s['processed'] for s in ct_stats.values())
        total_images    = sum(s['images']    for s in ct_stats.values())
        total_garbled   = sum(s['garbled']   for s in ct_stats.values())

        for label, val, color in [
            ("已下載 PDF",  f"{total_success:,}",   APPLE_GREEN),
            ("已萃取公司",  f"{total_processed:,}", APPLE_BLUE),
            ("圖片總數",    f"{total_images:,}",    APPLE_TEXT),
            ("亂碼公司",    f"{total_garbled:,}",
             APPLE_ORANGE if total_garbled else APPLE_GREY),
        ]:
            card = tk.Frame(frame, bg=APPLE_CARD,
                            highlightthickness=1, highlightbackground=APPLE_BORDER)
            card.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=5)
            tk.Label(card, text=label, font=FONT_LABEL,
                     fg=APPLE_GREY, bg=APPLE_CARD).pack(pady=(8, 0))
            tk.Label(card, text=val, font=('Helvetica Neue', 18, 'bold'),
                     fg=color, bg=APPLE_CARD).pack(pady=(2, 8))

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

    # ── 下載狀態表（Treeview）───────────────────────────────
    def _build_download_section(self, stats: dict):
        self._section_title("下載狀態",
                            "來源：ESG_Download_Progress_YYYY.xlsx  ·  點選列查看明細")
        frame = tk.Frame(self.body, bg=APPLE_BG)
        frame.pack(fill=tk.X, padx=20, pady=(0, 4))

        cols = ('年度', '✅ 成功', '⚠️ 未找到', '🔒 已確認無', '❌ 失敗', '總共爬蟲數', '進度')
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            height=len(stats), style='ESG.Treeview',
                            selectmode='browse')
        for col, w in zip(cols, [60, 70, 80, 90, 60, 90, 300]):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor='center', stretch=(col == '進度'))

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
            crawled  = s.get('_total', 0)   # 實際爬過的公司數
            pct      = int(crawled / TOTAL_COMPANIES * 100)
            bar      = _bar(pct)
            ct = self._ct_stats.get(year, {})
            tag = 'done' if crawled >= TOTAL_COMPANIES else ('partial' if pct > 0 else 'none')
            tree.insert('', 'end', values=(
                year,
                success,
                s.get('未找到中文版報告', 0),
                s.get('已確認無報告', 0),
                s.get('下載失敗', 0) or '—',
                crawled,
                bar,
            ), tags=(tag,))

        tree.tag_configure('done',    foreground=APPLE_GREEN)
        tree.tag_configure('partial', foreground=APPLE_BLUE)
        tree.tag_configure('none',    foreground=APPLE_GREY)
        tree.tag_configure('missing', foreground=APPLE_GREY)
        tree.tag_configure('error',   foreground=APPLE_RED)

        def _on_click(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            vals = tree.item(iid, 'values')
            year = vals[0]
            ds   = stats.get(year, {})
            cs   = self._ct_stats.get(year, {})
            DetailWindow.open(year, ds, cs)
        tree.bind('<Button-1>', _on_click)
        tree.bind('<MouseWheel>',
                  lambda e: (self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])
        tree.pack(fill=tk.X)

    # ── 圖表萃取狀態表（Treeview）───────────────────────────
    def _build_cutter_section(self, stats: dict):
        self._section_title("圖表萃取狀態",
                            "來源：掃描 data/{year}/*/images/*.jpg  ·  點選列查看明細")
        frame = tk.Frame(self.body, bg=APPLE_BG)
        frame.pack(fill=tk.X, padx=20, pady=(0, 4))

        cols = ('年度', '✅ 已萃取', '⏳ 待處理', '🖼 圖片數', '⚠️ 亂碼公司', '進度')
        tree = ttk.Treeview(frame, columns=cols, show='headings',
                            height=len(stats), style='ESG.Treeview',
                            selectmode='browse')
        for col, w in zip(cols, [60, 80, 80, 80, 90, 300]):
            tree.heading(col, text=col)
            tree.column(col, width=w, anchor='center', stretch=(col == '進度'))

        for year, s in sorted(stats.items()):
            processed = s['processed']
            pending   = s['pending']
            total     = processed + pending or 1
            pct       = int(processed / total * 100)

            if processed == 0 and pending == 0:
                bar = '尚無資料'
                tag = 'missing'
            elif pending == 0:
                bar = _bar(100)
                tag = 'done'
            else:
                bar = _bar(pct)
                tag = 'partial' if pct > 0 else 'none'

            tree.insert('', 'end', values=(
                year,
                processed or '—',
                pending   or '—',
                f"{s['images']:,}" if s['images'] else '—',
                s['garbled'] or '—',
                bar,
            ), tags=(tag,))

        tree.tag_configure('done',    foreground=APPLE_GREEN)
        tree.tag_configure('partial', foreground=APPLE_BLUE)
        tree.tag_configure('none',    foreground=APPLE_GREY)
        tree.tag_configure('missing', foreground=APPLE_GREY)

        def _on_click(event):
            iid = tree.identify_row(event.y)
            if not iid:
                return
            vals = tree.item(iid, 'values')
            year = vals[0]
            ds   = self._dl_stats.get(year, {})
            cs   = stats.get(year, {})
            DetailWindow.open(year, ds, cs)
        tree.bind('<Button-1>', _on_click)
        tree.bind('<MouseWheel>',
                  lambda e: (self._canvas.yview_scroll(int(-1*(e.delta/120)), 'units'), 'break')[1])
        tree.pack(fill=tk.X)

    def run(self):
        self.root.mainloop()


# ============================================================
# 主程式
# ============================================================
if __name__ == '__main__':
    Dashboard().run()
