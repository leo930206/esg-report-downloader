import pandas as pd
import os
import time
import random
import threading
import queue
import socket
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import scrolledtext, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import re
import tempfile
from pathlib import Path

# ============================================================
# 設定區
# ============================================================
HEADLESS         = True
STANDARD_PATTERN = re.compile(r'^\d{4}_\d+_.+\.pdf$')
_BASE_DIR        = Path(__file__).parent.absolute()
_DATA_DIR        = _BASE_DIR.parent / "data"    # 統一輸出根目錄（與 esg_pdf_cuter 共用）
logs_folder      = str(_BASE_DIR / "logs")
os.makedirs(logs_folder, exist_ok=True)
os.makedirs(str(_DATA_DIR), exist_ok=True)

def year_pdf_folder(year):
    return str(_DATA_DIR / str(year))

def year_progress_file(year):
    return str(_DATA_DIR / str(year) / f"ESG_Download_Progress_{year}.xlsx")

def _year_range():
    """產生有效年份範圍（2015 ～ 2024）"""
    return range(2015, 2025)

# ============================================================
# 執行緒控制
# ============================================================
log_queue              = queue.Queue()
ui_cmd_queue           = queue.Queue()
pause_event            = threading.Event()
truly_paused_event     = threading.Event()  # 程式真正進入暫停等待時才 set
window_closed          = threading.Event()
program_done           = threading.Event()
stop_event             = threading.Event()
network_down_event     = threading.Event()
download_started_event = threading.Event()  # 下載執行緒已啟動，可開進度視窗
log_history            = []
selected_years         = []
retry_failures         = True   # 是否重新補抓之前失敗的公司
_progress_win_opened   = False  # 是否已開過進度視窗（用來判斷是否為重開）

ui_stats = {
    'total': 0, 'processed': 0, 'success': 0,
    'failed': 0, 'start_time': None,
}

# ============================================================
# App Icon（Dock / 視窗）
# ============================================================
def set_app_icon(root: tk.Tk, emoji: str = "🌱") -> None:
    """把 emoji 渲染成圖片並設為視窗與 Dock 圖示（macOS 透過 AppKit）。"""
    try:
        import base64
        from io import BytesIO
        from PIL import Image, ImageDraw, ImageFont

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

        try:
            from AppKit import NSApplication, NSImage
            from Foundation import NSData
            data     = NSData.dataWithBytes_length_(png_bytes, len(png_bytes))
            ns_image = NSImage.alloc().initWithData_(data)
            NSApplication.sharedApplication().setApplicationIconImage_(ns_image)
        except Exception:
            pass

        photo = tk.PhotoImage(data=base64.b64encode(png_bytes).decode())
        root.iconphoto(True, photo)
        root._icon_ref = photo

    except Exception:
        pass


# ============================================================
# macOS 風格顏色
# ============================================================
APPLE_BG     = '#f5f5f7'
APPLE_CARD   = '#ffffff'
APPLE_BLUE   = '#0071e3'
APPLE_TEXT   = '#1d1d1f'
APPLE_GREY   = '#6e6e73'
APPLE_BORDER = '#d2d2d7'
FONT_MAIN    = ('Helvetica Neue', 10)
FONT_TITLE   = ('Helvetica Neue', 13, 'bold')
FONT_STAT    = ('Helvetica Neue', 20, 'bold')
FONT_LABEL   = ('Helvetica Neue', 9)
FONT_LOG     = ('Menlo', 9)

# ============================================================
# 查看進度視窗
# ============================================================
def classify_status(stat):
    """將狀態分成三類，回傳 (顯示文字, tag)"""
    if stat == '成功':
        return '成功', 'success'
    elif stat in ('未找到中文版報告', '已確認無報告'):
        return stat, 'no_report'
    else:
        return stat, 'failed'


def create_detail_window(parent, year, df):
    """點擊年度卡片後的詳細清單視窗"""
    win = tk.Toplevel(parent)
    win.title(f"{year} 年度詳細記錄")
    win.geometry("900x650")
    win.configure(bg=APPLE_BG)

    success   = len(df[df['status'] == '成功'])
    not_found = len(df[df['status'] == '未找到中文版報告'])
    confirmed = len(df[df['status'] == '已確認無報告'])
    dl_fail   = len(df) - success - not_found - confirmed

    # 標題
    header = tk.Frame(win, bg=APPLE_BLUE, pady=12)
    header.pack(fill=tk.X)
    tk.Label(header, text=f"{year} 年度詳細記錄", font=FONT_TITLE,
             fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)
    tk.Label(header,
             text=f"✅ 成功 {success}　⚠️ 未找到報告 {not_found}　🔒 已確認無報告 {confirmed}　❌ 下載失敗 {dl_fail}　共 {len(df)} 筆",
             font=FONT_MAIN, fg='#a8d4ff', bg=APPLE_BLUE).pack(side=tk.RIGHT, padx=20)

    # 篩選 + 搜尋
    filter_frame = tk.Frame(win, bg=APPLE_BG, padx=15, pady=6)
    filter_frame.pack(fill=tk.X)
    filter_var = tk.StringVar(value='全部')
    search_var = tk.StringVar()

    for label, val in [('全部', '全部'), ('✅ 成功', '成功'),
                       ('⚠️ 未找到報告', '未找到中文版報告'),
                       ('🔒 已確認無報告', '已確認無報告'), ('❌ 下載失敗', '下載失敗')]:
        tk.Radiobutton(filter_frame, text=label, variable=filter_var, value=val,
                       font=FONT_LABEL, bg=APPLE_BG, fg=APPLE_TEXT,
                       activebackground=APPLE_BG, cursor='hand2',
                       command=lambda: refresh(search_var.get().strip(), filter_var.get())
                       ).pack(side=tk.LEFT, padx=6)

    tk.Label(filter_frame, text="  搜尋：", font=FONT_MAIN,
             fg=APPLE_TEXT, bg=APPLE_BG).pack(side=tk.LEFT)
    tk.Entry(filter_frame, textvariable=search_var, font=FONT_MAIN, width=18,
             relief='flat', highlightthickness=1,
             highlightbackground=APPLE_BORDER).pack(side=tk.LEFT, padx=4)

    # Treeview
    style = ttk.Style(win)
    style.configure('Detail.Treeview', rowheight=24, font=FONT_MAIN,
                     background=APPLE_CARD, fieldbackground=APPLE_CARD)
    style.configure('Detail.Treeview.Heading',
                     font=('Helvetica Neue', 10, 'bold'), background=APPLE_BG)

    container = tk.Frame(win)
    container.pack(fill=tk.BOTH, expand=True, padx=15, pady=4)

    cols      = ('stock_id', 'company_name', 'status', 'filename')
    col_names = ('股票代號', '公司名稱', '狀態', '檔名')
    col_widths = (80, 150, 160, 420)

    tree = ttk.Treeview(container, columns=cols, show='headings',
                        style='Detail.Treeview')
    for col, name, width in zip(cols, col_names, col_widths):
        tree.heading(col, text=name)
        tree.column(col, width=width, minwidth=60, anchor='w')

    tree.tag_configure('success',   foreground='#1a7f37')
    tree.tag_configure('no_report', foreground='#9a6700')
    tree.tag_configure('failed',    foreground='#cf222e')

    vsb = ttk.Scrollbar(container, orient='vertical', command=tree.yview)
    hsb = ttk.Scrollbar(container, orient='horizontal', command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    container.grid_rowconfigure(0, weight=1)
    container.grid_columnconfigure(0, weight=1)

    all_rows = df.to_dict('records')

    def refresh(keyword='', filter_status='全部'):
        for item in tree.get_children():
            tree.delete(item)
        for r in all_rows:
            sid  = str(r.get('stock_id', ''))
            name = str(r.get('company_name', ''))
            stat = str(r.get('status', ''))
            fn   = str(r.get('filename', ''))
            if filter_status != '全部':
                if filter_status == '下載失敗':
                    if stat in ('成功', '未找到中文版報告', '已確認無報告'):
                        continue
                elif stat != filter_status:
                    continue
            if keyword and keyword not in sid and keyword not in name and keyword not in stat:
                continue
            display_stat, tag = classify_status(stat)
            tree.insert('', 'end', values=(sid, name, display_stat, fn), tags=(tag,))

    refresh()
    search_var.trace_add('write',
        lambda *_: refresh(search_var.get().strip(), filter_var.get()))

    tk.Button(win, text="關閉", font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              relief='flat', padx=16, pady=6, cursor='hand2',
              command=win.destroy).pack(pady=6)


def create_view_window(parent):
    win = tk.Toplevel(parent)
    win.title("下載進度總覽")
    win.geometry("620x500")
    win.configure(bg=APPLE_BG)

    header = tk.Frame(win, bg=APPLE_BLUE, pady=12)
    header.pack(fill=tk.X)
    tk.Label(header, text="📊 下載進度總覽", font=FONT_TITLE,
             fg='white', bg=APPLE_BLUE).pack()
    tk.Label(header, text="點擊年度卡片查看詳細記錄", font=FONT_LABEL,
             fg='#a8d4ff', bg=APPLE_BLUE).pack(pady=(2, 0))

    content = tk.Frame(win, bg=APPLE_BG, padx=20, pady=15)
    content.pack(fill=tk.BOTH, expand=True)

    found_any = False
    for year in _year_range():
        pf = year_progress_file(year)
        if not os.path.exists(pf):
            continue
        found_any = True
        try:
            df = pd.read_excel(pf, sheet_name='詳細記錄', engine='openpyxl')
            df = df.dropna(subset=['stock_id'])
            df['stock_id'] = df['stock_id'].apply(lambda x: str(int(float(x))))
            total      = len(df)
            success    = len(df[df['status'] == '成功'])
            not_found  = len(df[df['status'] == '未找到中文版報告'])
            confirmed  = len(df[df['status'] == '已確認無報告'])
            dl_fail    = total - success - not_found - confirmed

            card = tk.Frame(content, bg=APPLE_CARD, highlightthickness=1,
                            highlightbackground=APPLE_BORDER, cursor='hand2')
            card.pack(fill=tk.X, pady=4)

            tk.Label(card, text=f"{year} 年度", font=('Helvetica Neue', 11, 'bold'),
                     fg=APPLE_TEXT, bg=APPLE_CARD, cursor='hand2').pack(side=tk.LEFT, padx=12, pady=10)
            tk.Label(card,
                     text=f"✅ 成功 {success}　⚠️ 未找到報告 {not_found}　🔒 已確認無報告 {confirmed}　❌ 下載失敗 {dl_fail}　共 {total} 筆  ›",
                     font=FONT_MAIN, fg=APPLE_GREY, bg=APPLE_CARD, cursor='hand2').pack(side=tk.RIGHT, padx=12)

            df_copy = df.copy()
            for widget in (card,) + tuple(card.winfo_children()):
                widget.bind('<Button-1>', lambda e, y=year, d=df_copy: create_detail_window(win, y, d))
                widget.bind('<Enter>', lambda e, c=card: c.configure(bg='#f0f0f5'))
                widget.bind('<Leave>', lambda e, c=card: c.configure(bg=APPLE_CARD))

        except Exception as e:
            tk.Label(content, text=f"{year} 年度：讀取失敗 ({e})",
                     font=FONT_MAIN, fg='#cf222e', bg=APPLE_BG).pack(anchor='w')

    if not found_any:
        tk.Label(content, text="尚無任何下載記錄", font=FONT_MAIN,
                 fg=APPLE_GREY, bg=APPLE_BG).pack(pady=30)

    tk.Button(win, text="關閉", font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              relief='flat', padx=16, pady=6, cursor='hand2',
              command=win.destroy).pack(pady=10)

# ============================================================
# 啟動設定視窗
# ============================================================
def create_startup_window():
    root = tk.Tk()
    root.title("🌱 ESG 報告下載系統")
    root.geometry("480x380")
    root.configure(bg=APPLE_BG)
    root.resizable(False, False)
    set_app_icon(root)

    header = tk.Frame(root, bg=APPLE_BLUE, pady=14)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 報告下載系統", font=FONT_TITLE,
             fg='white', bg=APPLE_BLUE).pack()
    tk.Label(header, text="台灣上市公司永續報告書", font=FONT_LABEL,
             fg='#a8d4ff', bg=APPLE_BLUE).pack(pady=(2, 0))

    content = tk.Frame(root, bg=APPLE_BG, padx=25, pady=15)
    content.pack(fill=tk.BOTH, expand=True)

    tk.Label(content, text="請選擇下載年度（可多選）",
             font=('Helvetica Neue', 11, 'bold'),
             fg=APPLE_TEXT, bg=APPLE_BG).pack(anchor='w', pady=(0, 10))

    grid = tk.Frame(content, bg=APPLE_BG)
    grid.pack(fill=tk.X)

    all_years = list(_year_range())
    year_vars = {}
    for i, y in enumerate(all_years):
        var = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(grid, text=str(y), variable=var,
                            font=FONT_MAIN, bg=APPLE_BG, fg=APPLE_TEXT,
                            activebackground=APPLE_BG, selectcolor=APPLE_CARD,
                            cursor='hand2')
        cb.grid(row=i // 5, column=i % 5, sticky='w', padx=10, pady=4)
        year_vars[y] = var

    def on_start():
        global retry_failures
        years = sorted(y for y, v in year_vars.items() if v.get())
        if not years:
            messagebox.showwarning("未選擇年度", "請至少選擇一個年度")
            return
        retry_failures = messagebox.askyesno(
            "是否要補抓失敗的公司？",
            "【是】：Excel 裡「成功」和「確認無報告」的跳過，其餘全部重新抓\n\n"
            "【否】：Excel 裡有任何紀錄的都跳過，只抓從來沒處理過的"
        )
        selected_years.extend(years)
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", lambda: root.destroy())

    btn_frame = tk.Frame(root, bg=APPLE_BG, pady=15)
    btn_frame.pack()
    tk.Button(btn_frame, text="▶  開始下載",
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=22, pady=9, cursor='hand2',
              command=on_start).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_frame, text="📊  查看進度",
              font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
              activebackground=APPLE_BORDER, relief='flat', padx=22, pady=9,
              cursor='hand2',
              command=lambda: create_view_window(root)).pack(side=tk.LEFT, padx=8)

    root.mainloop()

# ============================================================
# 進度視窗
# ============================================================
def create_progress_window():
    global _progress_win_opened
    is_reopen = _progress_win_opened
    _progress_win_opened = True

    year_label = '、'.join(str(y) for y in selected_years)
    root = tk.Tk()
    root.title(f"🌱 ESG 報告下載系統 | {year_label} 年 台灣上市公司")
    root.geometry("1000x700")
    root.configure(bg=APPLE_BG)
    root.resizable(True, True)
    set_app_icon(root)

    def on_close():
        if program_done.is_set():
            try:
                fn = os.path.join(logs_folder,
                    f"ESG_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                with open(fn, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(ts + msg for ts, _, msg in log_history))
            except Exception:
                pass
            window_closed.set()
            root.destroy()
        elif pause_event.is_set() and truly_paused_event.is_set():
            try:
                fn = os.path.join(logs_folder,
                    f"ESG_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                with open(fn, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(ts + msg for ts, _, msg in log_history))
            except Exception:
                pass
            stop_event.set()
            pause_event.clear()
            program_done.set()
            window_closed.set()
            root.destroy()
        else:
            # 下載仍在進行中，提示等待
            messagebox.showinfo(
                "程式仍在執行",
                "程式仍在執行，請稍後再關閉。\n\n"
                "右上角狀態為「⏸ 已暫停」或「■ 已完成」才可關閉。"
            )

    root.protocol("WM_DELETE_WINDOW", on_close)

    header = tk.Frame(root, bg=APPLE_BLUE, pady=12)
    header.pack(fill=tk.X)
    tk.Label(header, text="ESG 報告下載系統",
             font=FONT_TITLE, fg='white', bg=APPLE_BLUE).pack(side=tk.LEFT, padx=20)
    if program_done.is_set():
        _dot_text, _dot_color = '■ 已完成', '#8e8e93'
    elif is_reopen and pause_event.is_set() and truly_paused_event.is_set():
        _dot_text, _dot_color = '⏸ 已暫停', '#ff9f0a'
    elif is_reopen:
        _dot_text, _dot_color = '● 執行中', '#34c759'
    else:
        _dot_text, _dot_color = '● 初始化', '#ffdd57'
    status_dot = tk.Label(header, text=_dot_text, font=FONT_MAIN,
                          fg=_dot_color, bg=APPLE_BLUE)
    status_dot.pack(side=tk.RIGHT, padx=20)

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
    stat_success   = make_stat_card(cards_frame, '成功')
    stat_failed    = make_stat_card(cards_frame, '失敗')

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
    log_text.tag_configure('sep',     foreground=APPLE_BORDER)

    # 重開視窗時還原歷史紀錄（第一次開啟不需要，queue 裡已有）
    if is_reopen and log_history:
        log_text.configure(state='normal')
        for ts_str, entry_tag, entry_msg in log_history:
            log_text.insert(tk.END, ts_str, 'skip')
            log_text.insert(tk.END, entry_msg + '\n', entry_tag)
        log_text.see(tk.END)
        log_text.configure(state='disabled')

    bottom = tk.Frame(root, bg=APPLE_BG, pady=8)
    bottom.pack(fill=tk.X, padx=15)
    pause_btn_text = tk.StringVar(value='⏸  暫停（下載完成後生效）')
    tk.Button(bottom, textvariable=pause_btn_text,
              font=FONT_MAIN, bg=APPLE_BLUE, fg='white',
              activebackground='#0051a8', activeforeground='white',
              relief='flat', padx=16, pady=7, cursor='hand2', bd=0,
              command=lambda: toggle_pause(pause_btn_text)).pack(side=tk.LEFT)

    view_btn = tk.Button(bottom, text='📊  查看進度',
                         font=FONT_MAIN, bg=APPLE_CARD, fg=APPLE_TEXT,
                         activebackground=APPLE_BORDER,
                         relief='flat', padx=16, pady=7, cursor='hand2', bd=0,
                         state='disabled',
                         command=lambda: create_view_window(root))
    view_btn.pack(side=tk.LEFT, padx=8)

    time_label = tk.Label(bottom, text='', font=FONT_LABEL,
                          fg=APPLE_GREY, bg=APPLE_BG)
    time_label.pack(side=tk.RIGHT)

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

        while not ui_cmd_queue.empty():
            cmd, val = ui_cmd_queue.get()
            if cmd == 'pause_btn':
                pause_btn_text.set(val)
            elif cmd == 'status_dot':
                status_dot.config(text=val[0], fg=val[1])

        # 暫停中或完成後才開放查看進度按鈕
        if pause_event.is_set() or program_done.is_set():
            view_btn.config(state='normal')
        else:
            view_btn.config(state='disabled')

        p   = ui_stats['processed']
        tot = ui_stats['total']
        stat_processed.set(f'{p}/{tot}' if tot else '—')
        stat_success.set(str(ui_stats['success']))
        stat_failed.set(str(ui_stats['failed']))
        if tot > 0:
            progress_bar['value'] = p / tot * 100
        time_label.config(text=f'更新時間 {datetime.now().strftime("%H:%M:%S")}')
        root.after(500, update_ui)

    update_ui()
    root.mainloop()


def toggle_pause(btn_var):
    if pause_event.is_set():
        pause_event.clear()
        btn_var.set('⏸  暫停（下載完成後生效）')
        ui_cmd_queue.put(('status_dot', ('● 執行中', '#34c759')))
    else:
        if messagebox.askyesno("確認暫停", "確定要暫停嗎？\n程式將在當前公司處理完成後暫停。"):
            pause_event.set()
            btn_var.set('▶  繼續執行')
            # 狀態點不在此更新，等 check_pause_point 真正暫停後才顯示


# ============================================================
# 視窗管理器（主執行緒）
# ============================================================
def window_manager():
    create_startup_window()
    if not selected_years:
        return

    threading.Thread(target=run_download, daemon=True).start()
    download_started_event.wait(timeout=5)  # 等執行緒真正啟動，比固定 sleep 可靠

    while not program_done.is_set():
        create_progress_window()
        if not program_done.is_set():
            time.sleep(0.3)

    if not window_closed.is_set():
        create_progress_window()


# ============================================================
# 日誌 / 暫停
# ============================================================
def log(msg, tag=None):
    if tag is None:
        if '✅' in msg:   tag = 'success'
        elif '❌' in msg: tag = 'error'
        elif '⚠️' in msg: tag = 'warning'
        elif any(c in msg for c in ('📊','🗓️','🎉','📂','🔍','▶','⏸')): tag = 'info'
        elif '⏭️' in msg: tag = 'skip'
        else:             tag = 'normal'
    log_history.append((f"[{datetime.now().strftime('%H:%M:%S')}] ", tag, msg))
    log_queue.put((tag, msg))


def check_pause_point():
    # 網路斷線等待（最多 30 分鐘，超過就停止）
    if network_down_event.is_set():
        deadline = time.time() + 1800
        while network_down_event.is_set():
            if time.time() > deadline:
                log('❌ 網路斷線超過 30 分鐘，停止程式', 'error')
                stop_event.set()
                return False
            time.sleep(1)
        if stop_event.is_set():
            return False
    # 手動暫停等待
    if pause_event.is_set():
        log('⏸ 已暫停，等待繼續執行...', 'info')
        ui_cmd_queue.put(('status_dot', ('⏸ 已暫停', '#ff9f0a')))
        truly_paused_event.set()
        while pause_event.is_set():
            time.sleep(0.3)
        truly_paused_event.clear()
        if stop_event.is_set():
            return False
        log('▶ 繼續執行', 'info')
        ui_cmd_queue.put(('status_dot', ('● 執行中', '#34c759')))
    return True


def is_network_available():
    """修正版：用 context manager 確保 socket 一定關閉，不汙染全域 timeout"""
    try:
        with socket.create_connection(('8.8.8.8', 53), timeout=3):
            return True
    except Exception:
        return False


def network_monitor():
    """背景執行緒：監測網路，斷線時設定 event 並更新狀態點"""
    was_down = False
    while not program_done.is_set():
        time.sleep(5)
        if not is_network_available():
            if not was_down:
                was_down = True
                network_down_event.set()
                log('📡 網路斷線！等待網路恢復...', 'warning')
                ui_cmd_queue.put(('status_dot', ('📡 網路中斷', '#ff3b30')))
        else:
            if was_down:
                was_down = False
                network_down_event.clear()
                log('📡 網路已恢復，繼續執行', 'info')
                ui_cmd_queue.put(('status_dot', ('● 執行中', '#34c759')))


# ============================================================
# 進度 Excel 管理（每年獨立）
# ============================================================
progress_records = []
completed_keys   = set()   # (year, stock_id)


def load_progress():
    global progress_records, completed_keys
    progress_records = []
    completed_keys   = set()
    for year in _year_range():
        pf = year_progress_file(year)
        if not os.path.exists(pf):
            log(f"📂 {year} 年：無進度檔")
            continue
        try:
            df_p = pd.read_excel(pf, sheet_name='詳細記錄', engine='openpyxl')
            df_p = df_p.dropna(subset=['stock_id'])
            df_p['stock_id'] = df_p['stock_id'].apply(lambda x: str(int(float(x))))
            df_p['year']     = df_p['year'].apply(lambda x: int(float(x)))
            records = df_p.to_dict('records')
            progress_records.extend(records)
            done_ok = 0
            done_confirmed = 0
            for r in records:
                if r.get('status') == '成功':
                    completed_keys.add((r['year'], r['stock_id']))
                    done_ok += 1
                elif r.get('status') == '已確認無報告':
                    completed_keys.add((r['year'], r['stock_id']))
                    done_confirmed += 1
            log(f"📂 {year} 年：{len(records)} 筆，成功抓取：{done_ok} 家（已確認無報告：{done_confirmed} 家）")
        except Exception as e:
            log(f"⚠️  {year} 年進度載入失敗（{pf}）：{e}")
    total_ok = sum(1 for r in progress_records if r.get('status') == '成功')
    total_confirmed = sum(1 for r in progress_records if r.get('status') == '已確認無報告')
    log(f"📂 合計載入：{len(progress_records)} 筆，成功抓取：{total_ok} 家（已確認無報告：{total_confirmed} 家）")


def save_to_excel(year):
    try:
        year_records = [r for r in progress_records if r.get('year') == year]
        if not year_records:
            return
        year_folder = os.path.abspath(str(year))
        os.makedirs(year_folder, exist_ok=True)

        df_p = pd.DataFrame(year_records)
        df_p['stock_id'] = df_p['stock_id'].astype(str)
        df_p = df_p.sort_values('stock_id').reset_index(drop=True)

        success_count = len(df_p[df_p['status'] == '成功'])
        fail_df       = df_p[df_p['status'] != '成功']
        summary_rows  = [{'項目': '✅ 成功', '數量': success_count}]
        for status, count in fail_df['status'].value_counts().items():
            if str(status).startswith('處理錯誤'):
                label = '❌ 處理錯誤'
            elif str(status) == '已確認無報告':
                label = '🔒 已確認無報告'
            elif str(status) == '未找到中文版報告':
                label = '⚠️ 未找到中文版報告'
            else:
                label = f'❌ {status}'
            summary_rows.append({'項目': label, '數量': int(count)})
        summary_rows.append({'項目': '📊 總計', '數量': len(df_p)})

        pf      = year_progress_file(year)
        pf_dir  = os.path.dirname(os.path.abspath(pf))
        tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=pf_dir)
        os.close(tmp_fd)
        try:
            with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
                df_p.to_excel(writer, sheet_name='詳細記錄', index=False)
                pd.DataFrame(summary_rows).to_excel(writer, sheet_name='統計', index=False)
            os.replace(tmp_path, pf)   # 寫完才替換，中途中斷不影響原檔
        except Exception:
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            raise
    except Exception as e:
        log(f"⚠️  寫入 {year} 進度檔案失敗: {e}")


def save_progress(record):
    year = record.get('year')
    key  = (year, str(record.get('stock_id', '')))
    for i, r in enumerate(progress_records):
        if (r.get('year'), str(r.get('stock_id', ''))) == key:
            progress_records[i] = record
            break
    else:
        progress_records.append(record)
    completed_keys.add(key)
    save_to_excel(year)


# ============================================================
# 啟動清理
# ============================================================
def startup_cleanup(download_folder, year):
    changed = False
    if os.path.exists(download_folder):
        strange = [f for f in os.listdir(download_folder)
                   if f.endswith('.pdf') and not STANDARD_PATTERN.match(f)]
        if strange:
            log(f"⚠️  發現 {len(strange)} 個原始檔名，刪除中...")
            for f in strange:
                try:
                    os.remove(os.path.join(download_folder, f))
                    log(f"   🗑️  已刪除: {f}")
                except Exception as e:
                    log(f"   ⚠️  無法刪除 {f}: {e}")
            changed = True

    pdf_files = set(os.listdir(download_folder)) if os.path.exists(download_folder) else set()
    to_remove = [r for r in progress_records
                 if r.get('year') == year and r.get('status') == '成功'
                 and r.get('filename', '') not in pdf_files]
    for r in to_remove:
        progress_records.remove(r)
        completed_keys.discard((year, str(r.get('stock_id', ''))))
        log(f"   ❌ 移除無對應檔案記錄: {r.get('stock_id')} {r.get('company_name')}")
        changed = True

    if changed:
        save_to_excel(year)
        log("✅ 清理完成")
    else:
        log("✅ 檢查完成，無需清理")


# ============================================================
# Selenium 工具函數
# ============================================================
driver = None
wait   = None


def wait_and_click(xpath, timeout=15):
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", el)
        time.sleep(0.3)
        driver.execute_script("arguments[0].click();", el)
        time.sleep(0.5)
        return True
    except Exception:
        return False


def select_dropdown_option(input_placeholder, option_text):
    try:
        if not wait_and_click(f"//input[@placeholder='{input_placeholder}']"):
            return False
        # 等 dropdown 選項出現，比固定 sleep 更快也更可靠
        try:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((By.XPATH,
                    f"//li[contains(text(),'{option_text}')]")))
        except Exception:
            time.sleep(0.8)  # fallback：dropdown 較慢時補足等待
        for xpath in [
            f"//li[normalize-space()='{option_text}']",
            f"//li[contains(text(),'{option_text}')]",
            f"//div[contains(@class,'option') and normalize-space()='{option_text}']",
            f"//*[normalize-space()='{option_text}']",
        ]:
            for opt in driver.find_elements(By.XPATH, xpath):
                if opt.is_displayed():
                    driver.execute_script("arguments[0].click();", opt)
                    time.sleep(0.2)
                    return True
        return False
    except Exception:
        return False


def find_pdf_download_button():
    try:
        title_xpath = "//div[contains(@class,'section-subtitle') and contains(.,'中文版報告書')]"
        try:
            title_el = driver.find_element(By.XPATH, title_xpath)
        except NoSuchElementException:
            return None
        for table_xpath in [
            "./following-sibling::*//table[contains(@class,'inquiry-table')]",
            "./parent::*/following-sibling::*//table[contains(@class,'inquiry-table')]",
            "./following-sibling::*//table",
            "./parent::*/following-sibling::*//table",
        ]:
            try:
                table = title_el.find_element(By.XPATH, table_xpath)
                break
            except NoSuchElementException:
                continue
        else:
            return None
        # 掃所有 tr，每行的第二格都找看看，避免報告在非第一行
        rows = table.find_elements(By.XPATH, ".//tbody/tr")
        for row in rows:
            tds = row.find_elements(By.XPATH, ".//td")
            if len(tds) < 2:
                continue
            second_td = tds[1]
            for xpath in [
                ".//div[contains(@class,'_link-icon-button-box_') and contains(@class,'img-left')]",
                ".//*[contains(text(),'下載PDF')]",
                ".//*[contains(text(),'下載') and contains(text(),'PDF')]",
            ]:
                els = second_td.find_elements(By.XPATH, xpath)
                if els:
                    return els[0]
        return None
    except Exception:
        return None


def wait_for_download(year, stock_id, company_name, download_folder):
    initial          = set(os.listdir(download_folder)) if os.path.exists(download_folder) else set()
    last_size        = 0
    download_started = False

    def rename_and_return(filename):
        old_path = os.path.join(download_folder, filename)
        clean    = re.sub(r'[<>:"/\\|?*]', '', company_name).strip() or stock_id
        desired  = f"{year}_{stock_id}_{clean}.pdf"
        if filename != desired:
            new_path = os.path.join(download_folder, desired)
            if os.path.exists(new_path):
                os.remove(new_path)
            try:
                os.rename(old_path, new_path)
                log(f"✅ 下載完成並重新命名: {desired}")
                return desired
            except Exception as e:
                log(f"✅ 下載完成: {filename} (改名失敗: {e})")
                return filename
        log(f"✅ 下載完成: {filename}")
        return filename

    i             = 0
    stall_seconds = 0   # 大小連續未增長的秒數
    STALL_LIMIT   = 60  # 大小停止增長超過 60 秒 → 確認卡住
    while True:
        time.sleep(1)
        i += 1
        if not os.path.exists(download_folder):
            continue
        current     = set(os.listdir(download_folder))
        new_files   = current - initial
        new_pdfs    = [f for f in new_files if f.endswith('.pdf')]
        crdownloads = [f for f in new_files if f.endswith('.crdownload')]

        if new_pdfs:
            return rename_and_return(new_pdfs[0])
        elif crdownloads:
            download_started = True
            cr_path = os.path.join(download_folder, crdownloads[0])
            try:
                size_kb = os.path.getsize(cr_path) // 1024
                if size_kb != last_size:
                    stall_seconds = 0   # 有進展，重置卡住計數
                    total_mb = sum(
                        os.path.getsize(os.path.join(download_folder, f))
                        for f in os.listdir(download_folder) if f.endswith('.pdf')
                    ) / 1024 / 1024
                    log(f"⏳ 下載中... 已等待 {i} 秒，已下載 {size_kb} KB（檔案總計 {total_mb:.1f} MB）")
                    last_size = size_kb
                else:
                    stall_seconds += 1
                    if stall_seconds >= STALL_LIMIT:
                        log(f"❌ 下載停止增長超過 {STALL_LIMIT} 秒，視為卡住失敗")
                        return None
                    if stall_seconds % 10 == 0:
                        log(f"⏳ 下載中... 已等待 {i} 秒，大小停滯 {stall_seconds} 秒（{size_kb} KB）")
            except Exception:
                stall_seconds += 1
                if i % 5 == 0:
                    log(f"⏳ 下載中... 已等待 {i} 秒")
        else:
            if not download_started and i >= 30:
                log("❌ 等待 30 秒後下載未開始，視為失敗")
                return None


def handle_download_click(element, year, stock_id, company_name, download_folder):
    try:
        driver.execute_script("arguments[0].click();", element)
        filename = wait_for_download(year, stock_id, company_name, download_folder)
        return filename is not None, filename
    except Exception as e:
        log(f"❌ 點擊下載時發生錯誤: {e}")
        return False, None


# ============================================================
# 共用查詢 + 下載邏輯（主流程和重試共用，消除重複程式碼）
# ============================================================
def _query_and_download(year, stock_id, company_name, download_folder):
    """
    查詢指定公司並嘗試下載 PDF。
    回傳: (status_str, filename_or_None, updated_company_name)
      status_str: '成功' | '下載失敗' | '未找到中文版報告' | '處理錯誤: ...'
    """
    if not select_dropdown_option("市場別*", "上市"):
        return '處理錯誤: 市場別失敗', None, company_name
    if not select_dropdown_option("報告年度*", str(year)):
        return '處理錯誤: 年度失敗', None, company_name
    if not select_dropdown_option("產業別", "全選"):
        return '處理錯誤: 產業別失敗', None, company_name

    try:
        ci = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//input[@placeholder='公司代號']")))
        driver.execute_script("arguments[0].scrollIntoView(true);", ci)
        time.sleep(0.3)
        ci.click(); ci.clear(); time.sleep(0.3)
        ci.send_keys(stock_id)
        # 等自動完成出現
        try:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located((By.XPATH,
                    f"//li[contains(text(),'{stock_id}')]")))
        except Exception:
            time.sleep(1.0)  # fallback
        found_suggestion = False
        for xpath in [f"//li[contains(text(),'{stock_id}')]",
                      f"//div[contains(text(),'{stock_id}')]"]:
            if found_suggestion:
                break
            for sug in driver.find_elements(By.XPATH, xpath):
                if sug.is_displayed() and stock_id in sug.text:
                    extracted = re.sub(r'^[-\s]+|[-\s]+$', '',
                                       sug.text.strip().replace(stock_id, ''))
                    if extracted and len(extracted) < 50:
                        company_name = extracted
                    driver.execute_script("arguments[0].click();", sug)
                    time.sleep(0.3)
                    found_suggestion = True
                    break
    except Exception:
        return '處理錯誤: 輸入代號失敗', None, company_name

    clicked = False
    for xpath in ["//button[contains(text(),'查詢')]",
                  "//span[text()='查詢']/parent::button"]:
        if clicked:
            break
        for btn in driver.find_elements(By.XPATH, xpath):
            if btn.is_displayed() and btn.is_enabled():
                driver.execute_script("arguments[0].click();", btn)
                clicked = True
                break
    if not clicked:
        try:
            driver.find_element(By.XPATH, "//input[@placeholder='公司代號']")\
                  .send_keys(Keys.RETURN)
            clicked = True
        except Exception:
            pass
    if not clicked:
        return '處理錯誤: 查詢失敗', None, company_name

    try:
        WebDriverWait(driver, 20).until(
            lambda d: (
                d.find_elements(By.XPATH, "//table[contains(@class,'inquiry-table')]") or
                d.find_elements(By.XPATH, "//*[contains(text(),'查無') or contains(text(),'無資料')]")
            ))
    except Exception:
        time.sleep(5)

    btn_el = find_pdf_download_button()
    if btn_el:
        success, filename = handle_download_click(
            btn_el, year, stock_id, company_name, download_folder)
        if success:
            return '成功', filename, company_name
        else:
            return '下載失敗', None, company_name
    else:
        return '未找到中文版報告', None, company_name


# ============================================================
# 主下載流程（單一年度）
# ============================================================
def process_year(year, df_companies, download_folder, opts, year_offset=0):
    log(f"\n{'─'*60}", 'sep')
    log(f"🗓️  開始處理 {year} 年度報告", 'info')
    log(f"{'─'*60}", 'sep')

    year_success = 0
    year_failed  = []

    total_all   = len(df_companies)

    # 「否」模式：預先建立這一年有紀錄的代號集合，迴圈裡直接查
    year_recorded_ids = {
        str(r['stock_id']) for r in progress_records
        if int(float(r['year'])) == year
    } if not retry_failures else set()

    # 已處理初始值：從既有進度開始，不從 0 開始
    if not retry_failures:
        initial_processed = len(year_recorded_ids)
    else:
        initial_processed = sum(
            1 for r in progress_records
            if int(float(r['year'])) == year
            and r.get('status') in ('成功', '已確認無報告')
        )

    # 跨年度累積：ui_stats['total'] 由呼叫端設好，只更新 processed
    ui_stats['processed']  = year_offset + initial_processed
    ui_stats['success']    = 0
    ui_stats['failed']     = 0
    ui_stats['start_time'] = datetime.now()
    ui_cmd_queue.put(('status_dot', ('● 執行中', '#34c759')))

    processed_this_run    = 0
    consecutive_not_found = 0
    consecutive_errors    = 0    # 連續 exception 錯誤數，與 not_found 合計判斷封鎖
    RATE_LIMIT_THRESHOLD  = 5    # 連續幾家找不到就視為被封鎖（正常最多4連失敗）
    recent_failures       = []   # 記錄最近失敗的 (stock_id, company_name)
    last_restart_at       = 0    # 記錄上次重啟時是第幾家，避免雙重重啟並動態縮短等待

    def proactive_restart():
        nonlocal last_restart_at
        global driver, wait
        log('🔄 每50家主動重啟瀏覽器，重置 session...', 'info')
        try:
            driver.quit()
        except Exception:
            pass
        time.sleep(3)
        driver = webdriver.Chrome(options=opts)
        wait   = WebDriverWait(driver, 20)
        driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
        time.sleep(4)
        last_restart_at = processed_this_run

    def restart_browser_after_block():
        nonlocal last_restart_at
        global driver, wait
        log('🔄 偵測到連續未找到，可能被網站封鎖，重啟瀏覽器...', 'warning')
        try:
            driver.quit()
        except Exception:
            pass

        # 自動偵測是否解封：每 60 秒重試一次，最多等 2 小時
        wait_minutes = 0
        MAX_BLOCK_WAIT = 120
        while True:
            wait_minutes += 1
            if wait_minutes > MAX_BLOCK_WAIT:
                log(f'⚠️  封鎖超過 {MAX_BLOCK_WAIT} 分鐘，強制繼續（可能仍被封）', 'warning')
                try:
                    driver = webdriver.Chrome(options=opts)
                    wait   = WebDriverWait(driver, 20)
                    driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                    time.sleep(4)
                except Exception:
                    pass
                break
            log(f'⏳ 等待解封中... 第 {wait_minutes} 分鐘', 'warning')
            ui_cmd_queue.put(('status_dot', (f'🚫 封鎖中 ({wait_minutes}分)', '#ff3b30')))
            time.sleep(60)
            if stop_event.is_set():
                return
            try:
                driver = webdriver.Chrome(options=opts)
                wait   = WebDriverWait(driver, 20)
                driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                time.sleep(4)
                # 固定用 2022 年 + 1101 台泥 做解封確認
                # 原因：不管目前在處理哪一年，2022 的 1101 確定有報告，
                # 能保證解封後看到結果表格，而非依賴「查無資料」這種模糊判斷
                if not select_dropdown_option("市場別*", "上市"):
                    raise Exception("市場別無法選取")
                if not select_dropdown_option("報告年度*", "2022"):
                    raise Exception("年度無法選取")
                select_dropdown_option("產業別", "全選")
                ci = driver.find_element(By.XPATH, "//input[@placeholder='公司代號']")
                ci.click(); ci.clear(); ci.send_keys("1101"); time.sleep(2.5)

                # 選自動完成建議（改用 flag 取代 raise StopIteration）
                clicked_suggestion = False
                for xp in ["//li[contains(text(),'1101')]", "//div[contains(text(),'1101')]"]:
                    if clicked_suggestion:
                        break
                    for sug in driver.find_elements(By.XPATH, xp):
                        if sug.is_displayed():
                            driver.execute_script("arguments[0].click();", sug)
                            time.sleep(0.5)
                            clicked_suggestion = True
                            break

                # 點查詢按鈕（改用 flag 取代 raise StopIteration）
                clicked_query = False
                for xp in ["//button[contains(text(),'查詢')]",
                            "//span[text()='查詢']/parent::button"]:
                    if clicked_query:
                        break
                    for btn in driver.find_elements(By.XPATH, xp):
                        if btn.is_displayed() and btn.is_enabled():
                            driver.execute_script("arguments[0].click();", btn)
                            clicked_query = True
                            break

                # 只要頁面回傳「結果表格」或「查無資料」任一個就算解封成功
                WebDriverWait(driver, 15).until(
                    lambda d: (
                        d.find_elements(By.XPATH, "//table[contains(@class,'inquiry-table')]") or
                        d.find_elements(By.XPATH, "//*[contains(text(),'查無') or contains(text(),'無資料')]")
                    ))
                if wait_minutes >= 10:
                    log(f'⚠️  封鎖持續了 {wait_minutes} 分鐘，建議考慮切換網路或重新撥號', 'warning')
                log(f'✅ 解封確認：頁面正常回應（等了 {wait_minutes} 分鐘）', 'info')
                ui_cmd_queue.put(('status_dot', ('● 執行中', '#34c759')))
                driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                time.sleep(3)
                last_restart_at = processed_this_run
                break
            except Exception:
                log(f'🚫 仍被封鎖，繼續等待...', 'warning')
                try:
                    driver.quit()
                except Exception:
                    pass

    try:
        driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
        time.sleep(4)

        for abs_pos, (_, row) in enumerate(df_companies.iterrows(), start=1):
            if stop_event.is_set():
                break

            stock_id     = str(row['公司代號'])
            company_name = str(row['公司簡稱'])

            if (year, stock_id) in completed_keys:
                continue
            if stock_id in year_recorded_ids:   # 「否」模式：有紀錄的也跳過
                continue

            processed_this_run += 1
            ui_stats['processed'] = year_offset + initial_processed + processed_this_run

            log(f"\n{'─'*50}", 'sep')
            log(f"[{abs_pos}/{total_all}] {year}年 {stock_id} {company_name}", 'info')

            try:
                time.sleep(0.3)

                def fail_record(reason, count_retry=True):
                    if stock_id not in year_failed:
                        year_failed.append(stock_id)
                    retry_count = 0
                    # 只有「未找到中文版報告」才追蹤次數，累積到第3次才標成確認
                    if reason == '未找到中文版報告' and count_retry:
                        existing = next((r for r in progress_records
                                         if r.get('year') == year
                                         and str(r.get('stock_id')) == stock_id), None)
                        raw = existing.get('retry_count', 0) if existing else 0
                        retry_count = int(raw if raw == raw else 0) + 1  # 處理 NaN（NaN != NaN）
                        if retry_count >= 3:
                            reason = '已確認無報告'
                            log(f"🔒 {stock_id} 第3次仍未找到，標記為已確認無報告", 'warning')
                    r = {'year': year, 'stock_id': stock_id, 'company_name': company_name,
                         'status': reason, 'filename': '', 'retry_count': retry_count}
                    save_progress(r)
                    ui_stats['failed'] = len(year_failed)

                status, filename, company_name = _query_and_download(
                    year, stock_id, company_name, download_folder)

                if status == '成功':
                    year_success += 1
                    consecutive_not_found = 0
                    consecutive_errors    = 0
                    recent_failures.clear()
                    ui_stats['success'] = year_success
                    save_progress({'year': year, 'stock_id': stock_id,
                                   'company_name': company_name,
                                   'status': '成功', 'filename': filename})
                    log(f"✅ [{abs_pos}/{total_all}] 成功: {stock_id} {company_name}")
                elif status == '未找到中文版報告':
                    consecutive_not_found += 1
                    recent_failures.append((stock_id, company_name))
                    log(f"⚠️  {stock_id} 未找到中文版報告")
                    fail_record('未找到中文版報告')
                elif status == '下載失敗':
                    consecutive_not_found += 1
                    recent_failures.append((stock_id, company_name))
                    log(f"❌ {stock_id} 下載失敗")
                    fail_record('下載失敗')
                else:  # 處理錯誤
                    log(f"❌ {stock_id} {status}")
                    fail_record(status)

                if consecutive_not_found + consecutive_errors >= RATE_LIMIT_THRESHOLD:
                    retry_list = recent_failures[-RATE_LIMIT_THRESHOLD:]
                    consecutive_not_found = 0
                    consecutive_errors    = 0
                    recent_failures.clear()
                    restart_browser_after_block()

                    # 解封後：清除那幾家的失敗記錄，重新爬取
                    log(f'🔄 解封後重試前 {len(retry_list)} 家公司...', 'info')
                    for retry_sid, retry_cname in retry_list:
                        completed_keys.discard((year, retry_sid))
                        progress_records[:] = [
                            r for r in progress_records
                            if not (r.get('year') == year and str(r.get('stock_id')) == retry_sid)
                        ]
                        year_failed[:] = [f for f in year_failed if f != retry_sid]
                    save_to_excel(year)

                    for retry_sid, retry_cname in retry_list:
                        if stop_event.is_set():
                            break
                        log(f"\n{'─'*50}", 'sep')
                        log(f"[重試] {year}年 {retry_sid} {retry_cname}", 'info')
                        try:
                            time.sleep(0.3)
                            r_status, r_filename, retry_cname = _query_and_download(
                                year, retry_sid, retry_cname, download_folder)

                            if r_status == '成功':
                                year_success += 1
                                ui_stats['success'] = year_success
                                save_progress({'year': year, 'stock_id': retry_sid,
                                               'company_name': retry_cname,
                                               'status': '成功', 'filename': r_filename})
                                log(f"✅ [重試成功] {retry_sid} {retry_cname}")
                            elif r_status == '未找到中文版報告':
                                if retry_sid not in year_failed:
                                    year_failed.append(retry_sid)
                                # 重試期間的失敗不累加 retry_count（避免被封鎖時冤枉標成確認）
                                existing = next((r for r in progress_records
                                                 if r.get('year') == year
                                                 and str(r.get('stock_id')) == retry_sid), None)
                                raw = existing.get('retry_count', 0) if existing else 0
                                kept_count = int(raw if raw == raw else 0)
                                save_progress({'year': year, 'stock_id': retry_sid,
                                               'company_name': retry_cname,
                                               'status': '未找到中文版報告', 'filename': '',
                                               'retry_count': kept_count})
                                log(f"⚠️  [重試] {retry_sid} 仍未找到中文版報告（重試次數保持 {kept_count}）")
                            else:
                                if retry_sid not in year_failed:
                                    year_failed.append(retry_sid)
                                save_progress({'year': year, 'stock_id': retry_sid,
                                               'company_name': retry_cname,
                                               'status': r_status, 'filename': ''})
                                log(f"❌ [重試失敗] {retry_sid} {r_status}")

                            driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                            time.sleep(random.uniform(2, 3.5))
                        except Exception as e:
                            log(f"❌ [重試] {retry_sid} 錯誤: {str(e).split(chr(10))[0][:60]}")

                    # 重試結束後歸零，避免重試失敗再次觸發無限輪迴
                    consecutive_not_found = 0
                    consecutive_errors    = 0
                    recent_failures.clear()

                if not check_pause_point():
                    log('⏹ 使用者停止程式', 'warning')
                    return

                driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                # 剛重啟後短暫等待即可；否則用正常較長等待降低被偵測機率
                since_restart = processed_this_run - last_restart_at
                delay = random.uniform(2, 3.5) if since_restart <= 5 else random.uniform(5, 10)
                time.sleep(delay)

                # 每 50 家主動重啟，但若距上次重啟不到 10 家則跳過（避免剛重啟完又重啟）
                if processed_this_run % 50 == 0 and since_restart >= 10:
                    proactive_restart()

            except Exception as e:
                short_err = str(e).split('\n')[0][:80]
                log(f"❌ {stock_id} 處理錯誤: {short_err}")
                fail_record(f'處理錯誤: {short_err}')
                consecutive_errors += 1
                recent_failures.append((stock_id, company_name))
                try:
                    driver.get("https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW")
                    time.sleep(3)
                except Exception:
                    pass

    except Exception as e:
        log(f"❌ 嚴重錯誤: {e}")

    log(f"\n{'─'*60}", 'sep')
    log(f"🎉 {year} 年度完畢！成功: {year_success} | 失敗: {len(year_failed)}", 'info')


# ============================================================
# 下載執行緒
# ============================================================
def run_download():
    global driver, wait

    threading.Thread(target=network_monitor, daemon=True).start()
    download_started_event.set()  # 通知 window_manager 執行緒已啟動，可開進度視窗

    try:
        df_companies = pd.read_excel("tw_listed.xlsx", engine='openpyxl')
    except FileNotFoundError:
        log("❌ 找不到 tw_listed.xlsx，請確認檔案與程式在同一資料夾", 'error')
        program_done.set()
        return
    except Exception as e:
        log(f"❌ 讀取 tw_listed.xlsx 失敗: {e}", 'error')
        program_done.set()
        return

    df_companies.columns = ['公司代號', '公司簡稱']
    df_companies['公司代號'] = df_companies['公司代號'].astype(str)
    log(f"讀取到 {len(df_companies)} 家公司")

    load_progress()

    if not retry_failures:
        log(f":track_next: 略過補抓模式：Excel 裡有任何紀錄的都跳過，只抓從來沒處理過的")
    else:
        log(f":arrows_counterclockwise: 補抓模式：Excel 裡「成功」和「確認無報告」的跳過，其餘全部重新抓")

    # 設定跨年度總進度（進度條不會在換年時歸零）
    ui_stats['total']     = len(df_companies) * len(selected_years)
    ui_stats['processed'] = 0
    ui_stats['success']   = 0
    ui_stats['failed']    = 0

    try:
        for year_idx, year in enumerate(selected_years):
            if stop_event.is_set():
                break

            download_folder = year_pdf_folder(year)
            os.makedirs(download_folder, exist_ok=True)

            log(f"🔍 {year} 年度啟動清理中（檢查 PDF 資料夾，刪除不完整下載、補齊遺失紀錄）...", 'info')
            startup_cleanup(download_folder, year)

            prefs = {
                "download.default_directory": download_folder,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": False,
                "safebrowsing.disable_download_protection": True,
                "profile.default_content_settings.popups": 0,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
                "plugins.always_open_pdf_externally": True,
                "plugins.plugins_disabled": ["Chrome PDF Viewer"],
                "download.open_pdf_in_system_reader": False,
                "download.extensions_to_open": "",
            }
            opts = Options()
            if HEADLESS:
                opts.add_argument("--headless=new")
            opts.add_argument("--no-sandbox")
            opts.add_argument("--disable-dev-shm-usage")
            opts.add_argument("--disable-gpu")
            opts.add_argument("--window-size=1920,1080")
            opts.add_argument("--disable-blink-features=AutomationControlled")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option('useAutomationExtension', False)
            opts.add_argument("--disable-popup-blocking")
            opts.add_argument("--disable-extensions")
            opts.add_experimental_option("prefs", prefs)

            if driver is not None:
                try:
                    driver.quit()
                except Exception:
                    pass
            driver = webdriver.Chrome(options=opts)
            wait   = WebDriverWait(driver, 20)

            year_offset = year_idx * len(df_companies)
            process_year(year, df_companies, download_folder, opts, year_offset)

            if stop_event.is_set():
                break

    except KeyboardInterrupt:
        log("\n⚠️  程式被用戶中斷", 'warning')
    except Exception as e:
        log(f"❌ 程式執行時發生嚴重錯誤: {e}")
    finally:
        try:
            if driver:
                driver.quit()
                log("✅ 瀏覽器已關閉")
        except Exception:
            pass

        for year in selected_years:
            dl = year_pdf_folder(year)
            if os.path.exists(dl):
                cnt = len([f for f in os.listdir(dl) if f.endswith('.pdf')])
                log(f"📊 {year} 年度共 {cnt} 個 PDF", 'info')

        ui_cmd_queue.put(('status_dot', ('■ 已完成', '#6e6e73')))
        log("\n程式結束，請關閉視窗", 'info')
        program_done.set()


# ============================================================
# 進入點（tkinter 在主執行緒）
# ============================================================
window_manager()
