"""
esg_dashboard.py — ESG 研究主控台 v2.0 (customtkinter)

四個步驟全覽：下載 → 萃取 → CLIP 分類 → ResNet-50 訓練
"""
from __future__ import annotations

import os
import sys
import threading
import subprocess
import tkinter.font as tkfont
from pathlib import Path
from datetime import datetime
from typing import Optional

import customtkinter as ctk
import pandas as pd

# ── 外觀 ─────────────────────────────────────────────────────
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ── 路徑 ─────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent.parent.parent.absolute()
DATA_DIR   = BASE_DIR / "data"
CHARTS_DIR = DATA_DIR / "charts"
MODELS_DIR = BASE_DIR / "models"

DOWNLOAD_STATUSES = ["成功", "未找到中文版報告", "已確認無報告", "下載失敗"]
CHART_CATEGORIES  = ["bar", "line", "pie", "map", "non_chart"]
AUTO_REFRESH_MS   = 30_000
TOTAL_COMPANIES   = 1078

# ── 顏色 ─────────────────────────────────────────────────────
C_BG      = "#F5F5F7"
C_CARD    = "#FFFFFF"
C_HDR_BAR = "#1D1D1F"      # 頂欄背景（幾乎黑色）
C_BLUE    = "#0071E3"
C_GREEN   = "#34C759"
C_RED     = "#FF3B30"
C_ORANGE  = "#FF9F0A"
C_PURPLE  = "#AF52DE"
C_TEXT    = "#1D1D1F"
C_SUB     = "#6E6E73"
C_BORDER  = "#E5E5EA"
C_ROW_ALT = "#F9F9FB"
C_HOVER   = "#EEF5FF"
C_TOTAL   = "#F0F0F5"
C_SEP     = "#E5E5EA"

# ── 字體偵測 ─────────────────────────────────────────────────
def _pick_font() -> str:
    available = set(tkfont.families())
    for f in ["Inter", "Segoe UI", "SF Pro Display", "Helvetica Neue", "Arial"]:
        if f in available:
            return f
    return "Arial"

_FONT_FAMILY: str = ""   # lazy init after Tk root exists

def F(size: int = 13, weight: str = "normal") -> ctk.CTkFont:
    global _FONT_FAMILY
    if not _FONT_FAMILY:
        _FONT_FAMILY = _pick_font()
    return ctk.CTkFont(family=_FONT_FAMILY, size=size, weight=weight)


# ── ProgressCell 標記 ────────────────────────────────────────
class ProgressCell:
    """在 Table.add_row 中標記進度欄位。"""
    def __init__(self, pct: int, done: bool = False, empty: bool = False, label: str = ""):
        self.pct   = max(0, min(100, pct))
        self.done  = done
        self.empty = empty
        self.label = label   # 覆寫顯示文字（如「尚無資料」）


# ── 資料讀取（與原版相同）────────────────────────────────────
def _file_fingerprint() -> str:
    parts = []
    if not DATA_DIR.is_dir():
        return ""
    for p in sorted(DATA_DIR.glob("*/ESG_Download_Progress_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    for p in sorted(DATA_DIR.glob("*/ESG_Extract_Results_*.xlsx")):
        parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    for p in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        if p.is_dir():
            parts.append(f"{p}:{p.stat().st_mtime:.0f}")
    xl = CHARTS_DIR / "clip_labeling_results.xlsx"
    if xl.exists():
        parts.append(f"{xl}:{xl.stat().st_mtime:.0f}")
    pth = MODELS_DIR / "resnet50_chart_best.pth"
    if pth.exists():
        parts.append(f"{pth}:{pth.stat().st_mtime:.0f}")
    return "|".join(parts)


def load_download_stats() -> dict[str, dict]:
    stats: dict[str, dict] = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        xls  = year_dir / f"ESG_Download_Progress_{year}.xlsx"
        if not xls.exists():
            stats[year] = {"_missing": True}
            continue
        try:
            try:
                df = pd.read_excel(xls, sheet_name="詳細記錄", engine="openpyxl")
            except Exception:
                df = pd.read_excel(xls, engine="openpyxl")
            counts = df["status"].value_counts().to_dict()
            stats[year] = {s: counts.get(s, 0) for s in DOWNLOAD_STATUSES}
            stats[year]["_total"] = len(df)
            stats[year]["_df"]    = df
        except Exception as e:
            stats[year] = {"_error": str(e)}
    return stats


def load_cutter_stats() -> dict[str, dict]:
    stats: dict[str, dict] = {}
    if not DATA_DIR.is_dir():
        return stats
    for year_dir in sorted(DATA_DIR.glob("[0-9][0-9][0-9][0-9]")):
        year = year_dir.name
        processed_dirs: list[Path] = []
        total_images   = 0
        garbled_files: list[Path] = []
        try:
            with os.scandir(year_dir) as entries:
                for entry in entries:
                    if not entry.is_dir():
                        continue
                    images_path = os.path.join(entry.path, "images")
                    if os.path.isdir(images_path):
                        try:
                            cnt = sum(1 for f in os.scandir(images_path) if f.name.endswith(".jpg"))
                        except OSError:
                            cnt = 0
                        if cnt > 0:
                            processed_dirs.append(Path(entry.path))
                            total_images += cnt
                    garbled = os.path.join(entry.path, "garbled_pages.txt")
                    if os.path.exists(garbled):
                        garbled_files.append(Path(garbled))
        except OSError:
            pass
        all_pdf_stems   = {p.stem for p in year_dir.glob("*.pdf")}
        processed_stems = {d.name for d in processed_dirs}
        stats[year] = {
            "processed":      len(processed_dirs),
            "pending":        len(all_pdf_stems - processed_stems),
            "images":         total_images,
            "garbled":        len(garbled_files),
            "garbled_files":  garbled_files,
            "processed_dirs": processed_dirs,
        }
    return stats


def load_classifier_stats() -> dict:
    empty = {"by_year": {}, "by_cat": {c: 0 for c in CHART_CATEGORIES}, "total": 0, "source": "empty"}
    if not CHARTS_DIR.is_dir():
        return empty
    xl = CHARTS_DIR / "clip_labeling_results.xlsx"
    if xl.exists():
        try:
            df = pd.read_excel(xl, sheet_name=0, engine="openpyxl")
            by_year: dict[str, dict] = {}
            by_cat:  dict[str, int]  = {c: 0 for c in CHART_CATEGORIES}
            for _, row in df.iterrows():
                year = str(row.iloc[0]).strip()
                if not year.isdigit():
                    continue
                yr: dict[str, int] = {}
                for cat in CHART_CATEGORIES:
                    val = int(row[cat]) if cat in df.columns and pd.notna(row[cat]) else 0
                    yr[cat] = val
                    by_cat[cat] += val
                yr["total"] = sum(yr.values())
                by_year[year] = yr
            return {"by_year": by_year, "by_cat": by_cat, "total": sum(by_cat.values()), "source": "excel"}
        except Exception:
            pass
    by_year = {}
    by_cat  = {c: 0 for c in CHART_CATEGORIES}
    for cat in CHART_CATEGORIES:
        cat_dir = CHARTS_DIR / cat
        if not cat_dir.is_dir():
            continue
        try:
            for f in os.scandir(cat_dir):
                if not f.name.endswith(".jpg"):
                    continue
                parts = f.name.split("_", 1)
                year  = parts[0] if len(parts[0]) == 4 and parts[0].isdigit() else "unknown"
                by_year.setdefault(year, {c: 0 for c in CHART_CATEGORIES})
                by_year[year][cat] = by_year[year].get(cat, 0) + 1
                by_cat[cat] += 1
        except OSError:
            pass
    for yr in by_year:
        by_year[yr]["total"] = sum(by_year[yr].get(c, 0) for c in CHART_CATEGORIES)
    total = sum(by_cat.values())
    return {"by_year": by_year, "by_cat": by_cat, "total": total, "source": "scan" if total else "empty"}


def load_trainer_stats() -> dict:
    model_exists = (MODELS_DIR / "resnet50_chart_best.pth").exists()
    best_val_acc = best_epoch = None
    epochs_done  = 0
    log_path = MODELS_DIR / "training_log.csv"
    if log_path.exists():
        try:
            import csv
            with open(log_path, newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
            if rows:
                epochs_done = len(rows)
                best_row = max(rows, key=lambda r: float(r.get("val_acc", 0)))
                best_val_acc = float(best_row["val_acc"])
                best_epoch   = int(best_row["epoch"])
        except Exception:
            pass
    data_counts: dict[str, int] = {c: 0 for c in CHART_CATEGORIES}
    if CHARTS_DIR.is_dir():
        for cat in CHART_CATEGORIES:
            cat_dir = CHARTS_DIR / cat
            if cat_dir.is_dir():
                try:
                    data_counts[cat] = sum(1 for f in os.scandir(cat_dir) if f.name.endswith(".jpg"))
                except OSError:
                    pass
    data_total = sum(data_counts.values())
    return {
        "model_exists":     model_exists,
        "best_val_acc":     best_val_acc,
        "best_epoch":       best_epoch,
        "epochs_done":      epochs_done,
        "data_counts":      data_counts,
        "data_total":       data_total,
        "data_no_nonchart": data_total - data_counts.get("non_chart", 0),
    }


# ════════════════════════════════════════════════════════════
# Table 元件
# ════════════════════════════════════════════════════════════
class Table(ctk.CTkFrame):
    """
    簡潔表格元件。
    - headers: 欄位標題清單
    - col_widths: 各欄 px 寬度（最後一欄可留較寬給進度條）
    - on_click: 點選 row 的 callback(row_data)
    """

    def __init__(self, master, headers: list[str], col_widths: list[int],
                 on_click=None, **kw):
        super().__init__(master, fg_color=C_CARD, corner_radius=12,
                         border_width=1, border_color=C_BORDER, **kw)
        self._headers    = headers
        self._widths     = col_widths
        self._on_click   = on_click
        self._row_widgets: list[ctk.CTkFrame] = []
        self._draw_header()

    def _draw_header(self):
        hdr = ctk.CTkFrame(self, fg_color=C_BG, corner_radius=0)
        hdr.pack(fill="x")
        for i, (h, w) in enumerate(zip(self._headers, self._widths)):
            ctk.CTkLabel(hdr, text=h, width=w, anchor="center",
                         font=F(11, "bold"), text_color=C_SUB,
                         fg_color="transparent"
                         ).grid(row=0, column=i, padx=2, pady=10, sticky="ew")
        hdr.grid_columnconfigure(len(self._headers) - 1, weight=1)
        # separator line
        ctk.CTkFrame(self, fg_color=C_SEP, height=1, corner_radius=0).pack(fill="x")

    def clear(self):
        for w in self._row_widgets:
            w.destroy()
        self._row_widgets.clear()

    def add_row(self, values: list, colors: list[str] | None = None,
                is_total: bool = False, data=None):
        """
        values: 每欄的值。ProgressCell 物件會渲染成進度條。
        colors: 對應各欄的文字顏色（None = 預設）。
        is_total: True → 粗體 + 較深背景。
        data: 點選時回傳給 on_click 的 payload。
        """
        idx = len(self._row_widgets)
        bg  = C_TOTAL if is_total else (C_CARD if idx % 2 == 0 else C_ROW_ALT)
        colors = colors or [C_TEXT] * len(values)

        # 分隔線
        ctk.CTkFrame(self, fg_color=C_SEP, height=1, corner_radius=0).pack(fill="x")

        row = ctk.CTkFrame(self, fg_color=bg, corner_radius=0)
        row.pack(fill="x")

        for i, (val, color, w) in enumerate(zip(values, colors, self._widths)):
            if isinstance(val, ProgressCell):
                cell = self._make_progress_cell(row, val, w)
                cell.grid(row=0, column=i, padx=6, pady=8, sticky="ew")
            else:
                text = str(val) if val != 0 else "—"
                lbl  = ctk.CTkLabel(row, text=text, width=w, anchor="center",
                                    font=F(12, "bold" if is_total else "normal"),
                                    text_color=color, fg_color="transparent")
                lbl.grid(row=0, column=i, padx=2, pady=8, sticky="ew")

        row.grid_columnconfigure(len(self._headers) - 1, weight=1)

        if self._on_click and not is_total and data is not None:
            self._bind_hover_click(row, bg, data)

        self._row_widgets.append(row)

    def _make_progress_cell(self, parent, pc: ProgressCell, width: int) -> ctk.CTkFrame:
        cell = ctk.CTkFrame(parent, fg_color="transparent", width=width)
        cell.grid_propagate(False)

        if pc.label:
            ctk.CTkLabel(cell, text=pc.label, font=F(11),
                         text_color=C_SUB, fg_color="transparent").pack(pady=10)
            return cell

        bar_w   = max(60, width - 56)
        fill_pct = pc.pct / 100
        bar = ctk.CTkProgressBar(cell, width=bar_w, height=8, corner_radius=4,
                                 fg_color=C_BORDER,
                                 progress_color=C_GREEN if pc.done else C_BLUE)
        bar.set(fill_pct)
        bar.pack(side="left", padx=(6, 6), pady=14)

        ctk.CTkLabel(cell, text=f"{pc.pct}%", font=F(11),
                     text_color=C_GREEN if pc.done else C_TEXT,
                     fg_color="transparent", width=36).pack(side="left")
        return cell

    def _bind_hover_click(self, row: ctk.CTkFrame, bg_orig: str, data):
        # 只改 row 本身的 fg_color（子 widget 皆 transparent，自動跟隨）
        # on_leave 加 bounding-box 檢查：游標移到子 widget 時 row 仍視為 hover，不閃爍

        def _in_row(e) -> bool:
            try:
                return (row.winfo_rootx() <= e.x_root <= row.winfo_rootx() + row.winfo_width()
                        and row.winfo_rooty() <= e.y_root <= row.winfo_rooty() + row.winfo_height())
            except Exception:
                return False

        def on_enter(_):
            row.configure(fg_color=C_HOVER)

        def on_leave(e):
            if not _in_row(e):
                row.configure(fg_color=bg_orig)

        def on_click(_):
            if self._on_click:
                self._on_click(data)

        def _all(w):
            yield w
            for c in w.winfo_children():
                yield from _all(c)

        for w in _all(row):
            w.bind("<Enter>",    on_enter)
            w.bind("<Leave>",    on_leave)
            w.bind("<Button-1>", on_click)
        row.configure(cursor="hand2")


# ════════════════════════════════════════════════════════════
# 細節彈窗
# ════════════════════════════════════════════════════════════
class DetailWindow:
    _instances: dict[str, "DetailWindow"] = {}

    @classmethod
    def open(cls, year: str, dl_row: dict, ct_row: dict):
        if year in cls._instances:
            try:
                cls._instances[year].win.lift()
                return
            except Exception:
                pass
        cls._instances[year] = cls(year, dl_row, ct_row)

    def __init__(self, year: str, dl_row: dict, ct_row: dict):
        self.year = year
        self.win  = ctk.CTkToplevel()
        self.win.title(f"{year} 年度明細")
        self.win.geometry("860x580")
        self.win.configure(fg_color=C_BG)
        self.win.protocol("WM_DELETE_WINDOW",
                          lambda: (DetailWindow._instances.pop(year, None), self.win.destroy()))
        self._build(dl_row, ct_row)

    def _build(self, dl_row: dict, ct_row: dict):
        # header — Mac 風格白色頂欄
        hdr = ctk.CTkFrame(self.win, fg_color=C_CARD, corner_radius=0, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text=f"  {self.year} 年度明細",
                     font=F(15, "bold"), text_color=C_TEXT,
                     fg_color="transparent").pack(side="left", padx=12)
        # 底部分隔線
        ctk.CTkFrame(self.win, fg_color=C_BORDER, height=1, corner_radius=0).pack(fill="x")

        # search bar
        sf = ctk.CTkFrame(self.win, fg_color=C_BG, corner_radius=0)
        sf.pack(fill="x", padx=16, pady=8)
        ctk.CTkLabel(sf, text="搜尋", font=F(12), text_color=C_TEXT).pack(side="left")
        self.search_var = ctk.StringVar()
        ctk.CTkEntry(sf, textvariable=self.search_var, width=240,
                     font=F(12), placeholder_text="公司代碼 / 名稱 / 狀態…"
                     ).pack(side="left", padx=(8, 16))
        ctk.CTkLabel(sf, text="狀態", font=F(12), text_color=C_TEXT).pack(side="left")
        self.filter_var = ctk.StringVar(value="全部")
        ctk.CTkOptionMenu(sf, variable=self.filter_var, width=160, font=F(12),
                          values=["全部"] + DOWNLOAD_STATUSES,
                          command=lambda _: self._apply()
                          ).pack(side="left", padx=8)
        self.search_var.trace_add("write", lambda *_: self._apply())

        # 可捲動區域
        self._scroll_body = ctk.CTkScrollableFrame(self.win, fg_color=C_BG, corner_radius=0)
        self._scroll_body.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        # table 放進可捲動區域
        headers    = ["代碼", "公司名稱", "下載狀態", "已萃取", "圖片數", "亂碼頁"]
        col_widths = [70, 190, 130, 90, 80, 80]
        self.table = Table(self._scroll_body, headers, col_widths)
        self.table.pack(fill="x")

        self._build_rows(dl_row, ct_row)
        self._apply()

    def _bind_scroll(self, widget):
        """遞迴把 MouseWheel 綁到 DetailWindow 的 scroll_body。"""
        def _scroll(e):
            try:
                units = int(-e.delta / 3) if sys.platform == "darwin" else int(-e.delta / 120)
                self._scroll_body._parent_canvas.yview_scroll(units or (-1 if e.delta > 0 else 1), "units")
            except Exception:
                pass
        widget.bind("<MouseWheel>", _scroll)
        for child in widget.winfo_children():
            self._bind_scroll(child)

    def _build_rows(self, dl_row: dict, ct_row: dict):
        processed_stems = {d.name for d in ct_row.get("processed_dirs", [])}
        img_counts = {
            d.name: len(list((d / "images").glob("*.jpg")))
            for d in ct_row.get("processed_dirs", [])
        }
        garbled_map: dict[str, str] = {}
        for gf in ct_row.get("garbled_files", []):
            try:
                garbled_map[gf.parent.name] = gf.read_text(encoding="utf-8").strip()
            except Exception:
                garbled_map[gf.parent.name] = "?"

        self.all_rows: list[tuple] = []
        df = dl_row.get("_df")
        if df is not None:
            for _, r in df.iterrows():
                raw    = str(r.get("file_name", r.get("filename", r.get("檔名", ""))))
                stem   = Path(raw).stem if raw else ""
                code   = str(r.get("stock_id", r.get("stock_code", r.get("代碼", ""))))
                name   = str(r.get("company_name", r.get("公司名稱", "")))
                status = str(r.get("status", r.get("狀態", "")))
                extracted = "✅ 已萃取" if stem in processed_stems else (
                    "—" if status != "成功" else "⏳ 待萃取")
                imgs = img_counts.get(stem, 0)
                self.all_rows.append((code, name, status, extracted,
                                      str(imgs) if imgs else "—",
                                      garbled_map.get(stem, "—")))

    def _apply(self):
        kw     = self.search_var.get().strip().lower()
        status = self.filter_var.get()
        self.table.clear()
        STATUS_COLORS = {
            "成功":           C_GREEN,
            "下載失敗":        C_RED,
            "未找到中文版報告":  C_ORANGE,
            "已確認無報告":     C_SUB,
        }
        for row in self.all_rows:
            if status != "全部" and row[2] != status:
                continue
            if kw and not any(kw in str(c).lower() for c in row):
                continue
            sc = STATUS_COLORS.get(row[2], C_TEXT)
            colors = [C_SUB, C_TEXT, sc, C_TEXT, C_TEXT, C_TEXT]
            self.table.add_row(list(row), colors=colors)
        # 重新綁定捲動（新增的 row widget 也要綁）
        self._bind_scroll(self.win)


# ════════════════════════════════════════════════════════════
# Dashboard 主視窗
# ════════════════════════════════════════════════════════════
class Dashboard(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ESG 研究主控台")
        self.geometry("1080x820")
        self.configure(fg_color=C_BG)
        self.resizable(True, True)
        self._set_icon()

        self._last_fp   = ""
        self._dl_stats: dict = {}
        self._ct_stats: dict = {}
        self._clf_stats: dict = {}
        self._trn_stats: dict = {}

        self._build_header()
        self._build_body()
        self.refresh(force=True)
        self._schedule_refresh()

    def _set_icon(self):
        icon = BASE_DIR / "ESG.png"
        if not icon.exists():
            return
        try:
            if sys.platform != "win32":
                from AppKit import NSApplication, NSImage
                img = NSImage.alloc().initWithContentsOfFile_(str(icon))
                if img:
                    NSApplication.sharedApplication().setApplicationIconImage_(img)
        except Exception:
            pass
        try:
            from PIL import Image as PILImage, ImageTk
            photo = ImageTk.PhotoImage(PILImage.open(str(icon)))
            self.iconphoto(True, photo)
            self._icon_ref = photo
        except Exception:
            pass

    # ── Header ───────────────────────────────────────────────
    def _build_header(self):
        # Mac 風格：白色背景 + 底部分隔線
        bar = ctk.CTkFrame(self, fg_color=C_CARD, corner_radius=0, height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        ctk.CTkLabel(bar, text="  ESG 研究主控台",
                     font=F(16, "bold"), text_color=C_TEXT,
                     fg_color="transparent").pack(side="left", padx=12)

        # right side
        right = ctk.CTkFrame(bar, fg_color="transparent")
        right.pack(side="right", padx=12)

        self._status_lbl = ctk.CTkLabel(right, text="", font=F(11), text_color=C_SUB)
        self._status_lbl.pack(side="right", padx=(12, 0))

        for label, cmd in [("↺  重新整理", lambda: self.refresh(force=True)),
                           ("📁  開啟資料夾", lambda: subprocess.Popen(["open", str(DATA_DIR)]))]:
            btn = ctk.CTkButton(right, text=label, font=F(12), width=110, height=30,
                                corner_radius=8, fg_color=C_BG, border_width=1,
                                border_color=C_BORDER, hover_color="#E8E8ED",
                                text_color=C_TEXT, command=cmd)
            btn.pack(side="right", padx=4)

        # 底部分隔線
        ctk.CTkFrame(self, fg_color=C_BORDER, height=1, corner_radius=0).pack(fill="x")

    # ── Body ─────────────────────────────────────────────────
    def _build_body(self):
        self._body = ctk.CTkScrollableFrame(self, fg_color=C_BG, corner_radius=0)
        self._body.pack(fill="both", expand=True)

    # ── 自動刷新 ─────────────────────────────────────────────
    def _schedule_refresh(self):
        self.refresh(force=False)
        self.after(AUTO_REFRESH_MS, self._schedule_refresh)

    def refresh(self, force: bool = True):
        self._status_lbl.configure(text="⟳  更新中…")

        def _load():
            fp = _file_fingerprint()
            if not force and fp == self._last_fp:
                self.after(0, lambda: self._status_lbl.configure(text=f"更新 {datetime.now().strftime('%H:%M:%S')}"))
                return
            dl  = load_download_stats()
            ct  = load_cutter_stats()
            clf = load_classifier_stats()
            trn = load_trainer_stats()
            self.after(0, lambda: self._render(dl, ct, clf, trn, fp))

        threading.Thread(target=_load, daemon=True).start()

    def _render(self, dl: dict, ct: dict, clf: dict, trn: dict, fp: str):
        self._last_fp = fp
        for y in range(2015, 2025):
            yr = str(y)
            if yr not in dl:
                dl[yr] = {"_missing": True}
            if yr not in ct:
                ct[yr] = {"processed": 0, "pending": 0, "images": 0,
                           "garbled": 0, "garbled_files": [], "processed_dirs": []}
        self._dl_stats, self._ct_stats = dl, ct
        self._clf_stats, self._trn_stats = clf, trn

        for w in self._body.winfo_children():
            w.destroy()

        self._section("Step 1 — 下載狀態",      "ESG_Download_Progress_YYYY.xlsx · 點選年份查看明細", C_BLUE)
        self._build_download(dl)
        self._section("Step 2 — 圖表萃取狀態",  "掃描 data/{year}/*/images/*.jpg · 點選年份查看明細", C_BLUE)
        self._build_cutter(ct)
        self._section("Step 3 — CLIP 分類狀態", self._clf_source_note(clf), C_PURPLE)
        self._build_classifier(clf, ct)
        self._section("Step 4 — ResNet-50 訓練狀態", "models/resnet50_chart_best.pth + training_log.csv", C_ORANGE)
        self._build_trainer(trn)
        ctk.CTkFrame(self._body, fg_color="transparent", height=24).pack()

        # 全視窗任何位置都能捲動（遞迴把 MouseWheel 綁到所有子 widget）
        self._bind_mousewheel_recursive(self._body)
        self._status_lbl.configure(text=f"更新 {datetime.now().strftime('%H:%M:%S')}")

    # ── 全域捲動 ─────────────────────────────────────────────
    def _bind_mousewheel_recursive(self, widget):
        """讓整個視窗任意位置都能觸發主 ScrollableFrame 的捲動。"""
        def _scroll(e):
            try:
                units = int(-e.delta / 3) if sys.platform == "darwin" else int(-e.delta / 120)
                self._body._parent_canvas.yview_scroll(units or (-1 if e.delta > 0 else 1), "units")
            except Exception:
                pass
        widget.bind("<MouseWheel>", _scroll)
        for child in widget.winfo_children():
            self._bind_mousewheel_recursive(child)

    # ── 區塊標題 ─────────────────────────────────────────────
    def _section(self, title: str, subtitle: str, accent: str):
        f = ctk.CTkFrame(self._body, fg_color="transparent")
        f.pack(fill="x", padx=24, pady=(28, 6))

        dot = ctk.CTkFrame(f, fg_color=accent, width=4, height=22, corner_radius=2)
        dot.pack(side="left", padx=(0, 10))

        texts = ctk.CTkFrame(f, fg_color="transparent")
        texts.pack(side="left")
        ctk.CTkLabel(texts, text=title, font=F(14, "bold"),
                     text_color=C_TEXT, fg_color="transparent").pack(anchor="w")
        ctk.CTkLabel(texts, text=subtitle, font=F(11),
                     text_color=C_SUB, fg_color="transparent").pack(anchor="w")

    def _wrap(self) -> ctk.CTkFrame:
        """建立帶 padding 的卡片容器。"""
        f = ctk.CTkFrame(self._body, fg_color="transparent")
        f.pack(fill="x", padx=24, pady=(0, 4))
        return f

    # ── Step 1：下載狀態 ──────────────────────────────────────
    def _build_download(self, stats: dict):
        frame = self._wrap()
        hdrs  = ["年度", "✅ 成功", "⚠️ 未找到", "🔒 確認無", "❌ 失敗", "爬蟲數", "進度"]
        widths = [60, 72, 80, 80, 72, 80, 260]

        def on_click(data):
            yr = data["year"]
            DetailWindow.open(yr, stats.get(yr, {}), self._ct_stats.get(yr, {}))

        tbl = Table(frame, hdrs, widths, on_click=on_click)
        tbl.pack(fill="x")

        tot = {k: 0 for k in ["成功", "未找到中文版報告", "已確認無報告", "下載失敗", "_total"]}
        for year, s in sorted(stats.items()):
            if s.get("_missing"):
                tbl.add_row([year, "—", "—", "—", "—", "—",
                             ProgressCell(0, label="尚無資料")],
                            colors=[C_TEXT] + [C_SUB] * 6)
                continue
            if s.get("_error"):
                tbl.add_row([year, "錯誤", "", "", "", "", ProgressCell(0, label=s["_error"])],
                            colors=[C_TEXT] + [C_RED] * 6)
                continue
            crawled = s.get("_total", 0)
            pct     = int(crawled / TOTAL_COMPANIES * 100)
            done    = pct >= 100
            clr     = C_GREEN if done else (C_BLUE if pct > 0 else C_SUB)
            tbl.add_row(
                [year, s.get("成功", 0), s.get("未找到中文版報告", 0),
                 s.get("已確認無報告", 0), s.get("下載失敗", 0) or "—",
                 crawled, ProgressCell(pct, done=done)],
                colors=[C_TEXT, C_GREEN, C_ORANGE, C_SUB, C_RED, clr, C_TEXT],
                data={"year": year},
            )
            for k in tot:
                tot[k] += s.get(k, 0)

        tot_pct = int(tot["_total"] / 10780 * 100)
        tbl.add_row(
            ["總計", tot["成功"], tot["未找到中文版報告"],
             tot["已確認無報告"], tot["下載失敗"] or "—",
             tot["_total"], ProgressCell(tot_pct, done=tot_pct >= 100)],
            colors=[C_TEXT, C_GREEN, C_ORANGE, C_SUB, C_RED, C_TEXT, C_TEXT],
            is_total=True,
        )
        ctk.CTkLabel(frame, text="年度進度 = 爬蟲數 ÷ 1078 ｜ 總計 = 爬蟲數 ÷ 10780",
                     font=F(10), text_color=C_SUB, fg_color="transparent"
                     ).pack(anchor="e", pady=(4, 0))

    # ── Step 2：圖表萃取 ──────────────────────────────────────
    def _build_cutter(self, stats: dict):
        frame = self._wrap()
        hdrs   = ["年度", "✅ 已萃取", "🖼 圖片數", "⚠️ 亂碼", "進度"]
        widths = [60, 110, 110, 100, 340]

        def on_click(data):
            yr = data["year"]
            DetailWindow.open(yr, self._dl_stats.get(yr, {}), stats.get(yr, {}))

        tbl = Table(frame, hdrs, widths, on_click=on_click)
        tbl.pack(fill="x")

        tot_proc = tot_imgs = tot_garb = tot_dl = 0
        for year, s in sorted(stats.items()):
            processed  = s["processed"]
            dl_success = self._dl_stats.get(year, {}).get("成功", 0)
            pct        = int(processed / dl_success * 100) if dl_success else 0
            tot_proc += processed;  tot_imgs += s["images"]
            tot_garb += s["garbled"];  tot_dl += dl_success

            if dl_success == 0:
                prog = ProgressCell(0, label="尚無資料")
                clr  = C_SUB
            elif pct >= 100:
                prog = ProgressCell(100, done=True)
                clr  = C_GREEN
            elif pct == 0:
                prog = ProgressCell(0, label="尚未開始")
                clr  = C_SUB
            else:
                prog = ProgressCell(pct)
                clr  = C_BLUE

            tbl.add_row(
                [year, processed or "—",
                 f"{s['images']:,}" if s["images"] else "—",
                 s["garbled"] or "—", prog],
                colors=[C_TEXT, clr, C_TEXT, C_ORANGE if s["garbled"] else C_SUB, C_TEXT],
                data={"year": year},
            )

        tot_pct = int(tot_proc / tot_dl * 100) if tot_dl else 0
        tbl.add_row(
            ["總計", tot_proc or "—", f"{tot_imgs:,}" if tot_imgs else "—",
             tot_garb or "—", ProgressCell(tot_pct, done=tot_pct >= 100)],
            is_total=True,
        )
        ctk.CTkLabel(frame, text="年度進度 = 已萃取 ÷ 下載成功數",
                     font=F(10), text_color=C_SUB, fg_color="transparent"
                     ).pack(anchor="e", pady=(4, 0))

    # ── Step 3：CLIP 分類 ─────────────────────────────────────
    @staticmethod
    def _clf_source_note(clf: dict) -> str:
        return {
            "excel": "來源：data/charts/clip_labeling_results.xlsx",
            "scan":  "來源：掃描 data/charts/ 目錄（Excel 尚未產生）",
            "empty": "尚未執行 clip_classifier.py",
        }.get(clf.get("source", "empty"), "")

    def _build_classifier(self, clf: dict, ct: dict):
        frame = self._wrap()
        if clf.get("source") == "empty":
            ctk.CTkLabel(frame,
                         text="尚無分類資料。請先完成 Step 2，再執行：\npython tools/chart-classifier/clip_classifier.py",
                         font=F(12), text_color=C_SUB, justify="left",
                         fg_color="transparent").pack(anchor="w", pady=8)
            return

        # 類別卡片橫排
        by_cat = clf.get("by_cat", {})
        total  = clf.get("total", 0)
        CAT_LABELS = {"bar": "長條圖", "line": "折線圖", "pie": "圓餅圖", "map": "地圖", "non_chart": "非圖表"}
        CAT_COLORS = {"bar": C_BLUE, "line": C_GREEN, "pie": C_ORANGE, "map": C_PURPLE, "non_chart": C_SUB}
        cards = ctk.CTkFrame(frame, fg_color="transparent")
        cards.pack(fill="x", pady=(0, 10))
        for cat in CHART_CATEGORIES:
            cnt = by_cat.get(cat, 0)
            pct = int(cnt / total * 100) if total else 0
            c = ctk.CTkFrame(cards, fg_color=C_CARD, corner_radius=10,
                             border_width=1, border_color=C_BORDER)
            c.pack(side="left", expand=True, fill="both", padx=4)
            ctk.CTkLabel(c, text=CAT_LABELS[cat], font=F(11), text_color=C_SUB).pack(pady=(10, 0))
            ctk.CTkLabel(c, text=f"{cnt:,}" if cnt else "—",
                         font=F(20, "bold"), text_color=CAT_COLORS[cat]).pack()
            ctk.CTkLabel(c, text=f"{pct}%", font=F(11), text_color=C_SUB).pack(pady=(0, 10))

        # 年份表
        by_year = clf.get("by_year", {})
        if not by_year:
            return
        ct_imgs = {str(y): ct.get(str(y), {}).get("images", 0) for y in range(2015, 2025)}
        hdrs   = ["年度", "長條圖", "折線圖", "圓餅圖", "地圖", "非圖表", "合計", "進度"]
        widths = [60, 72, 72, 72, 72, 72, 72, 230]
        tbl = Table(frame, hdrs, widths)
        tbl.pack(fill="x")

        tot_clf = {c: 0 for c in CHART_CATEGORIES}
        tot_cls = tot_imgs_all = 0
        for year in sorted({str(y) for y in range(2015, 2025)} | set(by_year.keys())):
            yd  = by_year.get(year, {})
            cls = yd.get("total", 0)
            tim = ct_imgs.get(year, 0)
            if cls == 0 and tim == 0:
                tbl.add_row([year, "—", "—", "—", "—", "—", "—", ProgressCell(0, label="尚無資料")],
                            colors=[C_TEXT] + [C_SUB] * 7)
                continue
            pct  = int(cls / tim * 100) if tim else 0
            done = pct >= 100
            tbl.add_row(
                [year,
                 yd.get("bar", 0) or "—", yd.get("line", 0) or "—",
                 yd.get("pie", 0) or "—", yd.get("map", 0) or "—",
                 yd.get("non_chart", 0) or "—",
                 cls or "—",
                 ProgressCell(pct, done=done) if tim else ProgressCell(0, label="—")],
                colors=[C_TEXT, C_BLUE, C_GREEN, C_ORANGE, C_PURPLE, C_SUB, C_TEXT, C_TEXT],
            )
            for c in CHART_CATEGORIES:
                tot_clf[c] += yd.get(c, 0)
            tot_cls += cls;  tot_imgs_all += tim

        tot_pct = int(tot_cls / tot_imgs_all * 100) if tot_imgs_all else 0
        tbl.add_row(
            ["總計",
             tot_clf["bar"] or "—", tot_clf["line"] or "—",
             tot_clf["pie"] or "—", tot_clf["map"] or "—",
             tot_clf["non_chart"] or "—",
             tot_cls or "—",
             ProgressCell(tot_pct, done=tot_pct >= 100) if tot_imgs_all else ProgressCell(0, label="—")],
            is_total=True,
        )

    # ── Step 4：ResNet 訓練 ───────────────────────────────────
    def _build_trainer(self, trn: dict):
        frame = self._wrap()
        dc    = trn.get("data_counts", {})
        CAT_LABELS = {"bar": "長條圖", "line": "折線圖", "pie": "圓餅圖", "map": "地圖", "non_chart": "非圖表"}
        CAT_COLORS = {"bar": C_BLUE, "line": C_GREEN, "pie": C_ORANGE, "map": C_PURPLE, "non_chart": C_SUB}

        # 各類別訓練資料卡片
        cards = ctk.CTkFrame(frame, fg_color="transparent")
        cards.pack(fill="x", pady=(0, 10))
        for cat in CHART_CATEGORIES:
            cnt = dc.get(cat, 0)
            c = ctk.CTkFrame(cards, fg_color=C_CARD, corner_radius=10,
                             border_width=1, border_color=C_BORDER)
            c.pack(side="left", expand=True, fill="both", padx=4)
            ctk.CTkLabel(c, text=CAT_LABELS[cat], font=F(11), text_color=C_SUB).pack(pady=(10, 0))
            ctk.CTkLabel(c, text=f"{cnt:,}" if cnt else "—",
                         font=F(20, "bold"), text_color=CAT_COLORS[cat] if cnt else C_SUB).pack()
            ctk.CTkLabel(c, text="張", font=F(11), text_color=C_SUB).pack(pady=(0, 10))

        # 狀態資訊卡
        info = ctk.CTkFrame(frame, fg_color=C_CARD, corner_radius=10,
                            border_width=1, border_color=C_BORDER)
        info.pack(fill="x")

        left  = ctk.CTkFrame(info, fg_color="transparent")
        left.pack(side="left", fill="both", expand=True, padx=20, pady=14)
        right = ctk.CTkFrame(info, fg_color=C_BORDER, width=1, corner_radius=0)
        right.pack(side="left", fill="y", pady=12)
        right2 = ctk.CTkFrame(info, fg_color="transparent")
        right2.pack(side="left", fill="both", expand=True, padx=20, pady=14)

        # 左：訓練資料狀態
        ctk.CTkLabel(left, text="訓練資料", font=F(11), text_color=C_SUB).pack(anchor="w")
        data_total = trn.get("data_total", 0)
        data_nc    = trn.get("data_no_nonchart", 0)
        if data_total == 0:
            ctk.CTkLabel(left, text="尚無資料 — 請先執行 clip_classifier.py",
                         font=F(13), text_color=C_SUB).pack(anchor="w", pady=(4, 0))
        else:
            ready = data_nc >= 50
            text  = f"共 {data_total:,} 張（含訓練用 {data_nc:,} 張）"
            hint  = "  ✅ 可以開始訓練" if ready else "  ⚠️ 建議各類至少 50 張再訓練"
            ctk.CTkLabel(left, text=text + hint, font=F(13),
                         text_color=C_GREEN if ready else C_ORANGE).pack(anchor="w", pady=(4, 0))

        # 右：模型狀態
        ctk.CTkLabel(right2, text="模型狀態", font=F(11), text_color=C_SUB).pack(anchor="w")
        model_exists = trn.get("model_exists", False)
        best_acc     = trn.get("best_val_acc")
        best_ep      = trn.get("best_epoch")
        epochs_done  = trn.get("epochs_done", 0)
        if not model_exists and epochs_done == 0:
            ctk.CTkLabel(right2, text="未訓練 — 執行 resnet_trainer.py 開始",
                         font=F(13), text_color=C_SUB).pack(anchor="w", pady=(4, 0))
        elif model_exists:
            acc_str = (f"最佳 val acc：{best_acc:.1f}%（第 {best_ep}/{epochs_done} epoch）"
                       if best_acc else f"已訓練 {epochs_done} epochs")
            ctk.CTkLabel(right2, text=f"✅ 已有模型  ·  {acc_str}",
                         font=F(13), text_color=C_GREEN).pack(anchor="w", pady=(4, 0))
        else:
            ctk.CTkLabel(right2, text=f"訓練中？（log 有 {epochs_done} epoch，但尚無 .pth）",
                         font=F(13), text_color=C_ORANGE).pack(anchor="w", pady=(4, 0))

    def run(self):
        self.mainloop()


# ════════════════════════════════════════════════════════════
# 入口
# ════════════════════════════════════════════════════════════
if __name__ == "__main__":
    Dashboard().run()
