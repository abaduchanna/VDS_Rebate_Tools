# -*- coding: utf-8 -*-
"""
Rebate Folder Tools (PySide6/Qt build)
========================================
Same four-step tool as the Tkinter build - Store Rename, Add Suffix,
Delete Year Rows, Convert Legacy Excel. Only the UI toolkit differs;
see VDS_Rebate_Tools.py's module docstring for the full feature
description.
"""
import os
import sys
import time
import json
import threading
import traceback
from pathlib import Path
from datetime import date

# --- 3SVerse runtime tamper guard (t197): frozen EXE main debugger = silent exit ---
import runtime_guard as _rguard3sv
_rguard3sv.enforce()


def _pip_install(pkg_name):
    if getattr(sys, "frozen", False):
        return
    print(f"[WARN] {pkg_name} is missing; related features may be unavailable.")
    return


for _mod, _pip in [
    ("openpyxl", "openpyxl"),
    ("xlrd", "xlrd==1.2.0"),
    ("xlwt", "xlwt"),
    ("xlutils", "xlutils"),
    ("win32com", "pywin32"),
    ("PIL", "pillow"),
]:
    try:
        __import__(_mod)
    except ImportError:
        _pip_install(_pip)

try:
    from openpyxl import load_workbook as _load_xlsx
except ImportError:
    _load_xlsx = None

try:
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
except ImportError:
    xlrd = xlwt = xl_copy = None

try:
    import win32com.client as _win32com_client
except ImportError:
    _win32com_client = None

from theme_manager_qt import ThemeManagerQt


def _script_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(name: str) -> str:
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", _script_dir())
        return os.path.join(base, name)
    return os.path.join(_script_dir(), name)


# ═══════════════════════════════════════════════════════════════════════════
# STORE MANAGEMENT — stores.json (user-editable list of rename pairs)
# ═══════════════════════════════════════════════════════════════════════════
STORES_FILE = _resource_path("stores.json")


def _default_stores():
    return [
        {"old": "Eliff", "new": "E Iliff"},
        {"old": "FM 1960", "new": "FM1960"},
        {"old": "Mckellips", "new": "McKellips"},
        {"old": "N 19th", "new": "N 19"},
        {"old": "N 35th", "new": "N 35"},
        {"old": "N 51st", "new": "N 51"},
        {"old": "New Thomas", "new": "W Thomas"},
        {"old": "Mt View RD", "new": "Mount View"},
        {"old": "4363 W Fuqua St", "new": "Fuqua"},
    ]


def load_stores():
    try:
        if os.path.exists(STORES_FILE):
            with open(STORES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return [(s.get("old", ""), s.get("new", "")) for s in data if s.get("old")]
    except Exception:
        pass
    pairs = _default_stores()
    save_stores(pairs)
    return [(s["old"], s["new"]) for s in pairs]


def save_stores(pairs):
    try:
        data = [{"old": o, "new": n} for o, n in pairs]
        with open(STORES_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def step1_store_rename(folder: Path, log, stores=None):
    log("\n-- STEP 1: Store Rename --------------------------------")
    if stores is None:
        stores = load_stores()
    log(f"  Using {len(stores)} store rename rules.")
    renamed = skipped = 0
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        new_name = f.name
        for old, new in stores:
            if old and old in new_name:
                new_name = new_name.replace(old, new)
        if new_name != f.name:
            dest = f.parent / new_name
            try:
                f.rename(dest)
                log(f"  Renamed: {f.name}  ->  {new_name}")
                renamed += 1
            except Exception as e:
                log(f"  ERROR renaming {f.name}: {e}")
        else:
            skipped += 1
    log(f"  Done — {renamed} renamed, {skipped} unchanged.")
    return renamed


def step2_add_suffix(folder: Path, log, suffix="2026"):
    log(f"\n-- STEP 2: Add ' {suffix}' Suffix -----------------------")
    renamed = skipped = 0
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".xlsx", ".xls"):
            continue
        stem = f.stem
        if f" {suffix}" in stem:
            log(f"  Skipped (already has ' {suffix}'): {f.name}")
            skipped += 1
            continue
        new_name = f"{stem} {suffix}{f.suffix}"
        dest = f.parent / new_name
        try:
            f.rename(dest)
            log(f"  Renamed: {f.name}  ->  {new_name}")
            renamed += 1
        except Exception as e:
            log(f"  ERROR renaming {f.name}: {e}")
    log(f"  Done — {renamed} renamed, {skipped} skipped.")
    return renamed


def _delete_year_xlsx(path: Path, log, year="2025") -> str:
    if _load_xlsx is None:
        return "SKIP (openpyxl not installed)"
    wb = _load_xlsx(str(path))
    ws = wb.active
    rows_to_delete = [
        r for r in range(1, ws.max_row + 1)
        if year in str(ws.cell(r, 1).value or "")
    ]
    if not rows_to_delete:
        return f"SKIP (no {year} rows)"
    for r in reversed(rows_to_delete):
        ws.delete_rows(r)
    wb.save(str(path))
    return f"OK — deleted {len(rows_to_delete)} row(s)"


def _delete_year_xls(path: Path, log, year="2025") -> str:
    if xlrd is None:
        return "SKIP (xlrd not installed)"
    if _load_xlsx is None:
        return "SKIP (openpyxl not installed)"

    rb = xlrd.open_workbook(str(path), formatting_info=False)
    rs = rb.sheet_by_index(0)

    keep_rows = [
        r for r in range(rs.nrows)
        if year not in str(rs.cell_value(r, 0))
    ]
    deleted = rs.nrows - len(keep_rows)
    if deleted == 0:
        return f"SKIP (no {year} rows)"

    import openpyxl
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = rs.name

    for new_r_idx, old_r in enumerate(keep_rows, start=1):
        for c in range(rs.ncols):
            cell_val = rs.cell_value(old_r, c)
            if isinstance(cell_val, float) and cell_val == int(cell_val):
                cell_val = int(cell_val)
            new_ws.cell(row=new_r_idx, column=c + 1, value=cell_val)

    xlsx_path = path.with_suffix(".xlsx")
    new_wb.save(str(xlsx_path))

    try:
        path.unlink()
    except Exception:
        pass

    return f"OK — deleted {deleted} row(s), saved as {xlsx_path.name}"


def step3_delete_year_rows(folder: Path, log, year="2025"):
    log(f"\n-- STEP 3: Delete {year} Rows ----------------------------")
    modified = skipped = errors = 0
    xls_files = [f for f in sorted(folder.iterdir())
                 if f.is_file() and f.suffix.lower() in (".xls", ".xlsx")]
    if not xls_files:
        log("  No Excel files found.")
        return 0
    for f in xls_files:
        log(f"  Processing: {f.name}")
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                ext = f.suffix.lower()
                if ext == ".xlsx":
                    result = _delete_year_xlsx(f, log, year)
                else:
                    result = _delete_year_xls(f, log, year)
                log(f"    {result}")
                if result.startswith("OK"):
                    modified += 1
                else:
                    skipped += 1
                break
            except Exception as e:
                log(f"    Attempt {attempt}/{max_retries} failed: {e}")
                if attempt < max_retries:
                    time.sleep(2)
                else:
                    log("    FINAL FAILURE — skipping file.")
                    errors += 1
    log(f"  Done — {modified} modified, {skipped} skipped, {errors} errors.")
    return modified


XLSX_FILE_FORMAT = 51
LEGACY_EXTS = (".xls", ".xlsm", ".xlt", ".xlsb", ".xlc")
EXCEL_RESTART_EVERY = 60

_CANCEL_CONVERT = threading.Event()


def _find_legacy_files(folder: Path, recurse: bool):
    out = []
    if recurse:
        for root, _, files in os.walk(folder):
            for f in files:
                out.append(Path(root) / f)
    else:
        out = [f for f in folder.iterdir() if f.is_file()]
    res = []
    for p in out:
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() in LEGACY_EXTS:
            res.append(p)
    return sorted(res)


class _ExcelConverter:
    def __init__(self, log):
        self.log = log
        self.xl = None
        self._opened = 0

    def _start_excel(self):
        if _win32com_client is None:
            raise RuntimeError("pywin32 not installed — cannot drive Excel.")
        self.xl = _win32com_client.DispatchEx("Excel.Application")
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        try:
            self.xl.AutomationSecurity = 3
        except Exception:
            pass
        try:
            self.xl.AskToUpdateLinks = False
        except Exception:
            pass

    def _stop_excel(self):
        if self.xl is not None:
            try:
                self.xl.Quit()
            except Exception:
                pass
        self.xl = None

    def _recycle_if_needed(self):
        self._opened += 1
        if self._opened % EXCEL_RESTART_EVERY == 0:
            self._stop_excel()
            time.sleep(1)
            self._start_excel()

    def convert_one(self, path: Path, overwrite: bool, delete_original: bool) -> str:
        out = path.with_suffix(".xlsx")
        if out.exists() and not overwrite:
            return "skip"
        wb = None
        try:
            try:
                wb = self.xl.Workbooks.Open(os.fspath(path.resolve()),
                                             UpdateLinks=0, ReadOnly=True)
            except Exception:
                wb = self.xl.Workbooks.Open(os.fspath(path.resolve()),
                                             UpdateLinks=0, ReadOnly=True,
                                             CorruptLoad=1)
            wb.SaveAs(os.fspath(out.resolve()), FileFormat=XLSX_FILE_FORMAT)
            wb.Close(False)
            wb = None
            self._recycle_if_needed()
            if delete_original:
                try:
                    path.unlink()
                except Exception as e:
                    self.log(f"      (kept original — delete failed: {e})")
            return "ok"
        except Exception as e:
            if wb is not None:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            try:
                self._stop_excel()
                self._start_excel()
            except Exception:
                pass
            return f"error: {e}"


def step4_convert_legacy_excel(folder: Path, log,
                                recurse: bool = False,
                                overwrite: bool = False,
                                delete_original: bool = False):
    log("\n-- STEP 4: Convert Legacy Excel -> .xlsx ---------------")
    if _win32com_client is None:
        log("  SKIP — pywin32 not installed. Install with: pip install pywin32")
        return 0

    _CANCEL_CONVERT.clear()
    files = _find_legacy_files(folder, recurse)
    if not files:
        log("  No legacy Excel files found in folder.")
        return 0

    log(f"  Found {len(files)} legacy Excel file(s). Starting Excel...")
    conv = _ExcelConverter(log)
    try:
        conv._start_excel()
    except Exception as e:
        log(f"  ERROR starting Excel: {e}")
        log("  Make sure Microsoft Excel is installed on this machine.")
        return 0

    ok = skip = err = 0
    try:
        for i, p in enumerate(files, 1):
            if _CANCEL_CONVERT.is_set():
                log("  Cancelled by user.")
                break
            log(f"  [{i}/{len(files)}] {p.name}")
            r = conv.convert_one(p, overwrite, delete_original)
            if r == "ok":
                ok += 1
                log(f"      OK converted -> {p.stem}.xlsx")
            elif r == "skip":
                skip += 1
                log("      SKIP (.xlsx already exists)")
            else:
                err += 1
                log(f"      FAIL {r}")
    finally:
        conv._stop_excel()

    log(f"  Done — {ok} converted, {skip} skipped, {err} errors.")
    return ok


# ==========================================================
# QT UI LAYER
# ==========================================================
from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
    QPushButton, QCheckBox, QProgressBar, QPlainTextEdit, QFileDialog, QMessageBox,
    QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QLabel
)

from header_manager_qt import FixedHeaderManagerQt

EMBEDDED_ICON_B64 = (
    open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip()
    if not getattr(sys, "frozen", False)
    else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()
)
ICON_ICO_NAME = "VDS_icon.ico"
LOGO_PNG_NAME = "VDS_Logo.png"
APP_TITLE = "VDS Desk Solutions - Rebate Tools"


def _extract_embedded_icon(b64, filename):
    try:
        if not b64:
            return None
        import base64 as _b64, tempfile
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None


def _set_window_icon_qt(window):
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        ico_path = os.path.join(meipass, ICON_ICO_NAME)
        if os.path.exists(ico_path):
            window.setWindowIcon(QIcon(ico_path))
            return
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, ICON_ICO_NAME)
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    try:
        ico_path = _extract_embedded_icon(EMBEDDED_ICON_B64, "VDS_app_icon.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
    except Exception:
        pass


class LogBridge(QObject):
    message = Signal(str)
    finished = Signal()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        _set_window_icon_qt(self)
        self._running = False

        self.theme_manager = ThemeManagerQt(default="dark", app_name="VDS-Rebate-Tools")

        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header_mgr = FixedHeaderManagerQt(outer, title="VDS Rebate Folder Tools")
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_PNG_NAME)
        if os.path.exists(logo_path):
            self.header_mgr.set_logo(logo_path=logo_path, text="VDS")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(24, 18, 24, 18)
        outer.addWidget(body, 1)

        # Folder row
        folder_row = QHBoxLayout()
        self.folder_edit = QLineEdit()
        folder_row.addWidget(self.folder_edit, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self._browse)
        folder_row.addWidget(browse_btn)
        body_layout.addLayout(folder_row)

        # Step checkboxes
        opts_row = QHBoxLayout()
        self.step_cb = {}
        for key, label in [
            (1, "1. Store Rename"),
            (2, "2. Add Suffix"),
            (3, "3. Delete Year Rows"),
            (4, "4. Convert Legacy Excel -> .xlsx"),
        ]:
            cb = QCheckBox(label)
            cb.setChecked(True)
            opts_row.addWidget(cb)
            self.step_cb[key] = cb
        opts_row.addStretch(1)
        body_layout.addLayout(opts_row)

        # Suffix / year fields
        yr_row = QHBoxLayout()
        yr_row.addWidget(QLabel("Step 2 — Suffix to add:"))
        self.suffix_edit = QLineEdit(str(date.today().year))
        self.suffix_edit.setFixedWidth(80)
        yr_row.addWidget(self.suffix_edit)
        yr_row.addSpacing(20)
        yr_row.addWidget(QLabel("Step 3 — Year to delete:"))
        self.year_edit = QLineEdit(str(date.today().year - 1))
        self.year_edit.setFixedWidth(80)
        yr_row.addWidget(self.year_edit)
        yr_row.addStretch(1)
        body_layout.addLayout(yr_row)

        # Store rename rules
        store_box = QGroupBox("Store Rename Rules (Step 1)")
        store_layout = QVBoxLayout(store_box)
        self.store_table = QTableWidget(0, 2)
        self.store_table.setHorizontalHeaderLabels(["Find (in filename)", "Replace with"])
        self.store_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.store_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.store_table.setMaximumHeight(160)
        store_layout.addWidget(self.store_table)

        edit_row = QHBoxLayout()
        edit_row.addWidget(QLabel("Find:"))
        self.new_old_edit = QLineEdit()
        edit_row.addWidget(self.new_old_edit)
        edit_row.addWidget(QLabel("Replace:"))
        self.new_new_edit = QLineEdit()
        edit_row.addWidget(self.new_new_edit)
        add_btn = QPushButton("+ Add")
        add_btn.clicked.connect(self._add_store)
        edit_row.addWidget(add_btn)
        edit_btn = QPushButton("Edit")
        edit_btn.clicked.connect(self._edit_store)
        edit_row.addWidget(edit_btn)
        del_btn = QPushButton("Delete")
        del_btn.clicked.connect(self._delete_store)
        edit_row.addWidget(del_btn)
        store_layout.addLayout(edit_row)
        body_layout.addWidget(store_box)
        self._reload_store_table()

        # Step-4 sub-options
        sub_row = QHBoxLayout()
        self.recurse_cb = QCheckBox("Include subfolders (Step 4)")
        self.recurse_cb.setChecked(True)
        self.overwrite_cb = QCheckBox("Overwrite existing .xlsx (Step 4)")
        self.overwrite_cb.setChecked(True)
        self.delete_original_cb = QCheckBox("Delete original after converting (Step 4)")
        self.delete_original_cb.setChecked(True)
        for cb in (self.recurse_cb, self.overwrite_cb, self.delete_original_cb):
            sub_row.addWidget(cb)
        sub_row.addStretch(1)
        body_layout.addLayout(sub_row)

        # Run / Cancel
        act_row = QHBoxLayout()
        self.run_btn = QPushButton("Run Selected Steps")
        self.run_btn.clicked.connect(self._start)
        act_row.addWidget(self.run_btn)
        self.cancel_btn = QPushButton("Cancel Step 4")
        self.cancel_btn.clicked.connect(lambda: _CANCEL_CONVERT.set())
        self.cancel_btn.setEnabled(False)
        act_row.addWidget(self.cancel_btn)
        self.status_label = QLabel("Ready.")
        act_row.addWidget(self.status_label)
        act_row.addStretch(1)
        body_layout.addLayout(act_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        body_layout.addWidget(self.progress)

        body_layout.addWidget(QLabel("Activity Log"))
        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background: #10182e; color: #a8d8ff; font-family: Consolas;")
        body_layout.addWidget(self.log_box, 1)

        self.header_mgr.add_copyright(self.theme_manager, outer)

        self._bridge = LogBridge()
        self._bridge.message.connect(self.log_box.appendPlainText)
        self._bridge.finished.connect(self._done)

        self._apply_theme()
        self.resize(1200, 900)
        QTimer.singleShot(0, self.showMaximized)

    def _apply_theme(self):
        self.theme_manager.apply_theme_to_app(QApplication.instance())
        self.header_mgr.update_for_theme()

    def _browse(self):
        d = QFileDialog.getExistingDirectory(self, "Select Rebate folder")
        if d:
            self.folder_edit.setText(d)

    # ── Store management helpers ──────────────────────────────────────────
    def _reload_store_table(self):
        self.store_table.setRowCount(0)
        for old, new in load_stores():
            row = self.store_table.rowCount()
            self.store_table.insertRow(row)
            self.store_table.setItem(row, 0, QTableWidgetItem(old))
            self.store_table.setItem(row, 1, QTableWidgetItem(new))

    def _add_store(self):
        old = self.new_old_edit.text().strip()
        new = self.new_new_edit.text().strip()
        if not old:
            QMessageBox.warning(self, "Missing", "Please enter text to find in filenames.")
            return
        pairs = load_stores()
        for o, n in pairs:
            if o == old:
                QMessageBox.warning(self, "Duplicate", f"'{old}' already exists in the list.")
                return
        pairs.append((old, new))
        save_stores(pairs)
        self._reload_store_table()
        self.new_old_edit.setText("")
        self.new_new_edit.setText("")

    def _edit_store(self):
        row = self.store_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "No selection", "Select a row to edit first.")
            return
        old = self.new_old_edit.text().strip()
        new = self.new_new_edit.text().strip()
        if not old:
            QMessageBox.warning(self, "Missing", "Please enter text to find in filenames.")
            return
        old_old = self.store_table.item(row, 0).text()
        pairs = load_stores()
        new_pairs = []
        for o, n in pairs:
            if o == old_old:
                new_pairs.append((old, new))
            else:
                new_pairs.append((o, n))
        save_stores(new_pairs)
        self._reload_store_table()
        self.new_old_edit.setText("")
        self.new_new_edit.setText("")

    def _delete_store(self):
        row = self.store_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "No selection", "Select a row to delete first.")
            return
        old_to_delete = self.store_table.item(row, 0).text()
        resp = QMessageBox.question(self, "Confirm", f"Delete rule for '{old_to_delete}'?")
        if resp != QMessageBox.Yes:
            return
        pairs = load_stores()
        new_pairs = [(o, n) for o, n in pairs if o != old_to_delete]
        save_stores(new_pairs)
        self._reload_store_table()

    def _log(self, msg):
        self._bridge.message.emit(msg)

    def _start(self):
        if self._running:
            return
        folder = Path(self.folder_edit.text().strip())
        if not folder.exists():
            QMessageBox.critical(self, "Folder not found", str(folder))
            return
        if self.step_cb[4].isChecked() and _win32com_client is None:
            resp = QMessageBox.question(
                self, "Excel driver missing",
                "Step 4 needs pywin32 + Microsoft Excel, which doesn't appear to be installed.\n\n"
                "Run Step 4 anyway? (It will skip itself if Excel can't start.)")
            if resp != QMessageBox.Yes:
                return
        self._running = True
        self.run_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.log_box.setPlainText("")
        self.progress.setVisible(True)
        threading.Thread(target=self._worker, args=(folder,), daemon=True).start()

    def _worker(self, folder: Path):
        try:
            self.status_label.setText("Running...")
            self._log("=" * 55)
            self._log(f"Folder: {folder}")
            steps = "".join(f"{k} " for k in (1, 2, 3, 4) if self.step_cb[k].isChecked()).rstrip()
            self._log(f"Steps: {steps}")
            suffix = self.suffix_edit.text().strip() or "2026"
            year = self.year_edit.text().strip() or str(date.today().year - 1)
            if self.step_cb[2].isChecked():
                self._log(f"  Step 2 suffix: '{suffix}'")
            if self.step_cb[3].isChecked():
                self._log(f"  Step 3 year to delete: '{year}'")
            self._log("=" * 55)

            if self.step_cb[1].isChecked():
                step1_store_rename(folder, self._log, stores=load_stores())
            if self.step_cb[2].isChecked():
                step2_add_suffix(folder, self._log, suffix=suffix)
            if self.step_cb[3].isChecked():
                step3_delete_year_rows(folder, self._log, year=year)
            if self.step_cb[4].isChecked():
                step4_convert_legacy_excel(
                    folder, self._log,
                    recurse=self.recurse_cb.isChecked(),
                    overwrite=self.overwrite_cb.isChecked(),
                    delete_original=self.delete_original_cb.isChecked(),
                )

            self._log("")
            self._log("COMPLETE — all selected steps finished.")
            self.status_label.setText("Done.")
        except Exception as e:
            self._log(f"\nCRITICAL ERROR: {e}")
            self._log(traceback.format_exc())
            self.status_label.setText("Error — see log.")
        finally:
            self._bridge.finished.emit()

    def _done(self):
        self.progress.setVisible(False)
        self.run_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self._running = False


# ==========================================================
# ENTRY POINT
# ==========================================================
def _enable_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("VDSDesk.App")
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _close_pyi_splash():
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    _close_pyi_splash()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
