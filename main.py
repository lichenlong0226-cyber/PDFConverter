#!/usr/bin/env python3
"""
PDFConverter: Word/Excel -> PDF 转换器（PySide6 GUI）
功能：
- 拖拽 / 添加文件、输出目录选择
- Windows 使用 pywin32 调用 Office COM 导出（高保真）
- 非 Windows 或回退使用 LibreOffice headless 转换
- 多线程并发转换（QThreadPool + QRunnable）
- 转换完成后可合并所有生成的 PDF（pypdf）
- 自动更新：从 GitHub Releases 检查新版本，下载安装包并使用 SHA256 校验后运行安装器
"""
import sys
import os
import platform
import subprocess
import shutil
import tempfile
import traceback
import hashlib
from pathlib import Path
from typing import Optional

from PySide6.QtCore import (Qt, QThreadPool, QRunnable, Signal, QObject, QTimer)
from PySide6.QtGui import QIcon, QAction, QCursor, QFont
from PySide6.QtWidgets import (QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
                               QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QHeaderView, QMessageBox, QCheckBox,
                               QProgressBar, QMenu, QTextEdit, QLineEdit, QFrame,
                               QSizePolicy)

IS_WINDOWS = platform.system() == "Windows"

# Optional Windows COM
if IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client
    except Exception:
        win32com = None
else:
    win32com = None

# PDF merger lib
try:
    from pypdf import PdfWriter
except Exception:
    PdfWriter = None

# Networking for auto-update
try:
    import requests
except Exception:
    requests = None

SUPPORTED_EXT = (".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".xlsb",
                 ".odt", ".ods", ".rtf", ".docm", ".pdf")

# ----------------- CONFIG -----------------
APP_NAME = "PDFConverter"
APP_VERSION = "1.1.6"
GITHUB_OWNER = "lichenlong0226-cyber"
GITHUB_REPO = "pdf"
ASSET_PREFIX = f"{APP_NAME}-setup-"
# ------------------------------------------

MAX_WORKERS = 3  # Office COM 实例并发数上限，太高反而变慢


def log_exc_text(e: Exception) -> str:
    return "".join(traceback.format_exception_only(type(e), e)).strip()

def sha256_of_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()

def parse_sha256_text(content: str) -> str:
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split()
        for p in parts:
            p2 = p.strip()
            if len(p2) == 64 and all(c in "0123456789abcdefABCDEF" for c in p2):
                return p2.lower()
        if len(parts[0]) >= 32 and all(c in "0123456789abcdefABCDEF" for c in parts[0]):
            return parts[0].lower()
    return ""

def convert_with_libreoffice(in_path: str, out_dir: str) -> str:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("LibreOffice (`soffice`) not found on PATH.")
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, in_path]
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out_pdf = Path(out_dir) / (Path(in_path).stem + ".pdf")
    if not out_pdf.exists():
        raise RuntimeError(f"LibreOffice conversion failed: {p.stderr.strip() or p.stdout.strip()}")
    return str(out_pdf)

def convert_word_windows(in_path: str, out_path: str):
    if win32com is None:
        raise RuntimeError("pywin32 is required on Windows for Word conversion.")
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    wdExportFormatPDF = 17
    word = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(in_path, ReadOnly=True)
        doc.ExportAsFixedFormat(out_path, wdExportFormatPDF)
        doc.Close(False)
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def convert_excel_windows(in_path: str, out_path: str):
    if win32com is None:
        raise RuntimeError("pywin32 is required on Windows for Excel conversion.")
    in_path = os.path.abspath(in_path)
    out_path = os.path.abspath(out_path)
    xlTypePDF = 0
    excel = None
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(in_path, ReadOnly=True)
        wb.ExportAsFixedFormat(xlTypePDF, out_path)
        wb.Close(False)
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def ensure_pdf_merger_available():
    if PdfWriter is None:
        raise RuntimeError("pypdf is required for merging PDFs.")


class WorkerSignals(QObject):
    started = Signal(str)
    finished = Signal(str, str)
    log = Signal(str)

class ConvertWorker(QRunnable):
    def __init__(self, in_path: str, out_dir: str):
        super().__init__()
        self.in_path = in_path
        self.out_dir = out_dir
        self.signals = WorkerSignals()

    def run(self):
        self.signals.started.emit(self.in_path)
        try:
            ext = Path(self.in_path).suffix.lower()
            base = Path(self.in_path).stem
            target = Path(self.out_dir) / f"{base}.pdf"
            cnt = 1
            while target.exists():
                target = Path(self.out_dir) / f"{base}({cnt}).pdf"
                cnt += 1
            if ext == ".pdf":
                target = Path(self.out_dir) / Path(self.in_path).name
                cnt = 1
                while target.exists():
                    target = Path(self.out_dir) / f"{Path(self.in_path).stem}({cnt}).pdf"
                    cnt += 1
                shutil.copy2(self.in_path, str(target))
                self.signals.log.emit(f"复制 PDF：{Path(self.in_path).name}")
                self.signals.finished.emit(self.in_path, str(target))
                return
            if IS_WINDOWS and ext in (".doc", ".docx", ".docm", ".rtf"):
                self.signals.log.emit(f"使用 Word COM 转换：{Path(self.in_path).name}")
                convert_word_windows(self.in_path, str(target))
            elif IS_WINDOWS and ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
                self.signals.log.emit(f"使用 Excel COM 导出：{Path(self.in_path).name}")
                convert_excel_windows(self.in_path, str(target))
            else:
                self.signals.log.emit(f"使用 LibreOffice 转换：{Path(self.in_path).name}")
                tmpd = tempfile.mkdtemp()
                try:
                    outpdf = convert_with_libreoffice(self.in_path, tmpd)
                    shutil.move(outpdf, str(target))
                finally:
                    shutil.rmtree(tmpd, ignore_errors=True)
            self.signals.finished.emit(self.in_path, str(target))
        except Exception as e:
            err = f"ERR: {log_exc_text(e)}"
            self.signals.log.emit(f"转换失败：{Path(self.in_path).name} -> {err}")
            self.signals.finished.emit(self.in_path, err)


# ----------------- UI -----------------

class DropTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(0, 3, parent)
        self.setHorizontalHeaderLabels(["文件", "状态", "大小"])
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.verticalHeader().setVisible(False)

    def dragEnterEvent(self, event):
        if event.source() == self:
            event.acceptProposedAction()
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.source() == self:
            event.acceptProposedAction()
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.source() == self:
            super().dropEvent(event)
            self.parent()._update_file_count()
            return
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        event.setDropAction(Qt.CopyAction)
        event.accept()
        urls = event.mimeData().urls()
        added = 0
        for u in urls:
            path = u.toLocalFile()
            if not path:
                continue
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORTED_EXT):
                            if self.add_file(os.path.join(root, f)):
                                added += 1
            else:
                if self.add_file(path):
                    added += 1
        if added > 0:
            self.parent()._update_file_count()

    @staticmethod
    def _file_emoji(path):
        ext = Path(path).suffix.lower()
        if ext in (".doc", ".docx", ".docm", ".rtf"): return "📝 "
        if ext in (".xls", ".xlsx", ".xlsm", ".xlsb"): return "📊 "
        if ext == ".pdf": return "📄 "
        if ext in (".odt", ".ods"): return "📃 "
        return "📁 "

    def add_file(self, path):
        if not os.path.exists(path):
            return False
        ext = Path(path).suffix.lower()
        if ext not in SUPPORTED_EXT:
            return False
        for row in range(self.rowCount()):
            if self.item(row, 0).data(Qt.UserRole) == path:
                return False
        row = self.rowCount()
        self.insertRow(row)
        size_text = f"{Path(path).stat().st_size // 1024} KB"
        item = QTableWidgetItem(self._file_emoji(path) + os.path.basename(path))
        item.setData(Qt.UserRole, path)
        self.setItem(row, 0, item)
        item_status = QTableWidgetItem("待处理")
        item_status.setTextAlignment(Qt.AlignCenter)
        self.setItem(row, 1, item_status)
        item_size = QTableWidgetItem(size_text)
        item_size.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.setItem(row, 2, item_size)
        return True

    def _on_context_menu(self, pos):
        row = self.indexAt(pos).row()
        if row < 0:
            return
        path = self.item(row, 0).data(Qt.UserRole)
        menu = QMenu()
        open_act = QAction("打开文件所在目录", self)
        open_act.triggered.connect(lambda: self._open_folder(path))
        remove_act = QAction("移除", self)
        remove_act.triggered.connect(lambda: self._remove_row(row))
        retry_act = QAction("重试转换", self)
        retry_act.triggered.connect(lambda: self.parent().retry_single(path))
        menu.addAction(open_act)
        menu.addAction(remove_act)
        menu.addAction(retry_act)
        menu.exec(QCursor.pos())

    def _remove_row(self, row):
        self.removeRow(row)
        self.parent()._update_file_count()

    def _open_folder(self, path):
        folder = str(Path(path).parent)
        if IS_WINDOWS:
            subprocess.Popen(f'explorer /select,"{path}"')
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])


class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} - Word/Excel -> PDF")
        self.resize(960, 620)
        self.setMinimumSize(640, 400)
        icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # --- 基础样式 ---
        self._apply_style()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        # ---- 顶栏 ----
        top = QHBoxLayout()
        title = QLabel("📄 Word/Excel → PDF")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #e0e0e0;")
        top.addWidget(title)
        top.addStretch()
        self.btn_check_update = QPushButton(f"检查更新 v{APP_VERSION}")
        self.btn_check_update.setFixedWidth(130)
        self.btn_check_update.clicked.connect(self.manual_check_update)
        top.addWidget(self.btn_check_update)
        self.btn_about = QPushButton("关于")
        self.btn_about.setFixedWidth(60)
        self.btn_about.clicked.connect(self.show_about)
        top.addWidget(self.btn_about)
        main_layout.addLayout(top)

        # ---- 文件计数 ----
        self.lbl_status = QLabel("已添加 0 个文件   |   拖拽或点击「添加文件」")
        self.lbl_status.setStyleSheet("color: #999; font-size: 11px; padding: 2px 0;")
        main_layout.addWidget(self.lbl_status)

        # ---- 文件列表 ----
        file_group = QGroupBox("文件列表（拖拽排序 / 从外部添加）")
        file_group.setStyleSheet("QGroupBox { font-weight: bold; color: #a6adc8; border: 1px solid #313244; border-radius: 6px; margin-top: 10px; padding-top: 16px; } QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }")
        file_layout = QVBoxLayout(file_group)
        file_layout.setContentsMargins(6, 6, 6, 6)
        self.table = DropTable(self)
        file_layout.addWidget(self.table, stretch=1)
        main_layout.addWidget(file_group, stretch=1)

        # ---- 按钮行 ----
        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("+ 添加文件")
        self.btn_remove = QPushButton("− 移除选中")
        self.btn_clear = QPushButton("清空全部")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        btn_row.addStretch()
        self.btn_add.clicked.connect(self.open_add_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.clear_all)
        main_layout.addLayout(btn_row)

        out_group = QGroupBox("输出设置")
        out_group.setStyleSheet("QGroupBox { font-weight: bold; color: #a6adc8; border: 1px solid #313244; border-radius: 6px; margin-top: 10px; padding-top: 16px; } QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }")
        out_layout = QVBoxLayout(out_group)
        out_layout.setContentsMargins(8, 8, 8, 8)
        out_layout.setSpacing(6)

        out_row = QHBoxLayout()
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("输出目录（留空则使用桌面/PDFConverter_output）")
        self.btn_out = QPushButton("选择目录")
        self.btn_out.setFixedWidth(100)
        out_row.addWidget(QLabel("输出:"))
        out_row.addWidget(self.out_edit, stretch=1)
        out_row.addWidget(self.btn_out)
        self.btn_out.clicked.connect(self.choose_out_dir)
        out_layout.addLayout(out_row)

        # ---- 操作行 ----
        ops_row = QHBoxLayout()
        self.chk_merge = QCheckBox("合并为单个 PDF")
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        self.progress.setFixedHeight(20)
        self.progress.setTextVisible(True)
        self.progress.setFormat("就绪")
        self.btn_convert = QPushButton("▶ 开始转换")
        self.btn_convert.setFixedWidth(110)
        self.btn_convert.setStyleSheet("QPushButton { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #89b4fa, stop:1 #74c7ec); color: #1e1e2e; font-weight: bold; border: none; border-radius: 6px; padding: 7px 18px; } QPushButton:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #b4d0fb, stop:1 #89dcf0); } QPushButton:pressed { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #74c7ec, stop:1 #89b4fa); }")
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedWidth(80)
        self.btn_convert.clicked.connect(self.start_conversion)
        self.btn_cancel.clicked.connect(self.cancel_all)
        ops_row.addWidget(self.chk_merge)
        ops_row.addWidget(self.progress, stretch=1)
        ops_row.addWidget(self.btn_convert)
        ops_row.addWidget(self.btn_cancel)
        out_layout.addLayout(ops_row)
        main_layout.addWidget(out_group)

        # ---- 日志面板（默认折叠） ----
        self.log_toggle = QPushButton("▶ 显示日志")
        self.log_toggle.setCheckable(True)
        self.log_toggle.setStyleSheet("QPushButton { text-align: left; padding: 4px 8px; font-size: 11px; color: #888; border: 1px solid #333; border-radius: 3px; } QPushButton:checked { color: #ccc; }")
        self.log_toggle.toggled.connect(self._toggle_log)
        main_layout.addWidget(self.log_toggle)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setVisible(False)
        self.log_edit.setFixedHeight(140)
        self.log_edit.setStyleSheet("background: #1a1a1a; color: #c0c0c0; font-family: Consolas, monospace; font-size: 10px; border: 1px solid #333;")
        main_layout.addWidget(self.log_edit)

        # ---- 状态 ----
        # ---- 状态栏 ----
        self.status_bar = QLabel("就绪")
        self.status_bar.setStyleSheet("color: #6c7086; font-size: 11px; padding: 2px 4px; border-top: 1px solid #313244;")
        main_layout.addWidget(self.status_bar)

        self.pool = QThreadPool.globalInstance()
        self.pool.setMaxThreadCount(MAX_WORKERS)
        self.active_workers = {}

        self.update_timer = QTimer(self)
        self.update_timer.setInterval(1000 * 60 * 60 * 24)
        self.update_timer.timeout.connect(lambda: self.check_for_updates(background=True))
        self.update_timer.start()

        self.output_dir = str(Path.home() / "Desktop" / "PDFConverter_output")
        self.cancel_requested = False

    def _apply_style(self):
        self.setStyleSheet("""
            QWidget { font-family: "Segoe UI", "Microsoft YaHei", Arial, sans-serif; font-size: 12px; background: #1e1e2e; color: #cdd6f4; }
            QPushButton { padding: 6px 16px; border: 1px solid #45475a; border-radius: 5px; background: #313244; color: #cdd6f4; }
            QPushButton:hover { background: #45475a; border-color: #585b70; }
            QPushButton:pressed { background: #585b70; }
            QTableWidget { background: #181825; alternate-background-color: #1e1e2e; border: 1px solid #313244; gridline-color: #313244; border-radius: 5px; font-size: 11px; outline: none; }
            QTableWidget::item { padding: 5px 8px; border-radius: 3px; }
            QTableWidget::item:selected { background: #45475a; color: #cdd6f4; }
            QHeaderView::section { background: #313244; color: #a6adc8; padding: 5px 8px; border: none; border-bottom: 1px solid #45475a; font-size: 11px; font-weight: 600; }
            QProgressBar { background: #181825; border: 1px solid #313244; border-radius: 4px; text-align: center; color: #a6adc8; font-size: 11px; }
            QProgressBar::chunk { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #89b4fa, stop:1 #74c7ec); border-radius: 3px; }
            QCheckBox { spacing: 8px; color: #cdd6f4; }
            QCheckBox::indicator { width: 18px; height: 18px; border-radius: 3px; border: 1px solid #585b70; background: #313244; }
            QCheckBox::indicator:checked { background: #89b4fa; border-color: #89b4fa; }
            QLineEdit { background: #181825; border: 1px solid #313244; border-radius: 5px; padding: 5px 8px; color: #cdd6f4; selection-background-color: #45475a; }
            QLineEdit:focus { border-color: #89b4fa; }
            QTextEdit { background: #11111b; color: #a6adc8; font-family: Consolas, "Cascadia Code", monospace; font-size: 10px; border: 1px solid #313244; border-radius: 4px; padding: 4px; }
            QLabel { color: #cdd6f4; }
            QFrame { color: #313244; }
        """)

    def _update_file_count(self):
        count = self.table.rowCount()
        self.lbl_status.setText(f"已添加 {count} 个文件   |   拖拽或点击「添加文件」")
        self.lbl_status.repaint()
        self.status_bar.setText(f"就绪  |  文件: {count}  |  输出: {self.output_dir}")

    def _toggle_log(self, checked):
        self.log_edit.setVisible(checked)
        self.log_toggle.setText("▼ 隐藏日志" if checked else "▶ 显示日志")

    def open_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", os.getcwd(),
                                                "Word/Excel (*.doc *.docx *.xls *.xlsx *.xlsm *.odt *.ods *.rtf)")
        added = 0
        for f in files:
            if self.table.add_file(f):
                added += 1
        if added:
            self._update_file_count()

    def remove_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)
        self._update_file_count()

    def clear_all(self):
        self.table.setRowCount(0)
        self._update_file_count()

    def choose_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_dir or os.getcwd())
        if d:
            self.output_dir = d
            self.out_edit.setText(d)

    def retry_single(self, path):
        self.table.add_file(path)
        self._update_file_count()
        if not self.active_workers:
            self.start_conversion()

    def append_log(self, text: str):
        self.log_edit.append(text)
        # 自定展开日志
        if not self.log_toggle.isChecked():
            self.log_toggle.setChecked(True)

    def cancel_all(self):
        self.append_log("取消请求：等待正在运行的线程结束...")
        self.cancel_requested = True

    def start_conversion(self):
        n = self.table.rowCount()
        if n == 0:
            QMessageBox.information(self, "提示", "请先添加要转换的文件。")
            return
        if IS_WINDOWS and win32com is None:
            QMessageBox.critical(self, "错误", "pywin32 未正确打包，请确保构建时包含该模块。")
            return
        if self.chk_merge.isChecked() and PdfWriter is None:
            QMessageBox.critical(self, "错误", "pypdf 未正确打包，请确保构建时包含该模块。")
            return
        if requests is None:
            self.append_log("警告：requests 未安装，自动更新不可用。")

        od = self.out_edit.text().strip() or self.output_dir or os.getcwd()
        if not os.path.exists(od):
            try:
                os.makedirs(od, exist_ok=True)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"无法创建输出目录：{e}")
                return
        self.output_dir = od

        self.progress.setMaximum(n)
        self.progress.setValue(0)
        self.progress.setFormat(f"0 / {n}")
        self.cancel_requested = False
        self.append_log(f"开始转换：{n} 个文件 → {self.output_dir}")
        self.status_bar.setText(f"转换中  |  剩余: {n}  |  输出: {self.output_dir}")

        self.pdfs_generated = []
        self.remaining = n

        for i in range(n):
            in_path = self.table.item(i, 0).data(Qt.UserRole)
            self.table.setItem(i, 1, QTableWidgetItem("等待中"))
            worker = ConvertWorker(in_path, self.output_dir)
            worker.signals.started.connect(lambda p: self.on_started(p))
            worker.signals.log.connect(lambda s: self.on_worker_log(s))
            worker.signals.finished.connect(lambda p, out: self.on_finished(p, out))
            self.active_workers[in_path] = worker
            self.pool.start(worker)

    def on_started(self, path):
        self.append_log(f"[启动] {Path(path).name}")
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0).data(Qt.UserRole) == path:
                item = QTableWidgetItem("处理中")
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(r, 1, item)

    def on_worker_log(self, text):
        self.append_log(text)

    def on_finished(self, path, out_or_err):
        self.active_workers.pop(path, None)
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0).data(Qt.UserRole) == path:
                if out_or_err.startswith("ERR:"):
                    item = QTableWidgetItem("❌ 失败")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(r, 1, item)
                    self.append_log(f"[失败] {Path(path).name} -> {out_or_err}")
                else:
                    item = QTableWidgetItem("✅ 已完成")
                    item.setTextAlignment(Qt.AlignCenter)
                    self.table.setItem(r, 1, item)
                    self.append_log(f"[完成] {Path(path).name} -> {Path(out_or_err).name}")
                    self.pdfs_generated.append(out_or_err)
                break
        done = self.progress.value() + 1
        self.progress.setValue(done)
        self.progress.setFormat(f"{done} / {self.remaining + done}")
        self.remaining -= 1
        if self.remaining <= 0 or (self.cancel_requested and not self.active_workers):
            self.progress.setFormat("完成")
            self.append_log("全部任务已结束。")
            if self.chk_merge.isChecked() and self.pdfs_generated:
                self.merge_after_convert()
            else:
                try:
                    subprocess.Popen(["explorer", self.output_dir])
                except Exception:
                    pass
                QMessageBox.information(self, "完成", f"已完成转换，生成 {len(self.pdfs_generated)} 个 PDF\n输出目录：{self.output_dir}")

    def merge_after_convert(self):
        ensure_pdf_merger_available()
        default_name = os.path.join(self.output_dir, "merged.pdf")
        merged_name, _ = QFileDialog.getSaveFileName(self, "保存合并后的 PDF", default_name, "PDF (*.pdf)")
        if not merged_name:
            QMessageBox.information(self, "完成", "已转换完成（未保存合并结果）。")
            return
        merger = PdfWriter()
        try:
            for p in self.pdfs_generated:
                merger.append(p)
            merger.write(merged_name)
            try:
                subprocess.Popen(f"explorer /select,\"{merged_name}\"")
            except Exception:
                pass
            QMessageBox.information(self, "完成", f"已完成转换并合并\n合并文件：{merged_name}")
            self.append_log(f"合并完成：{merged_name}")
        except Exception as e:
            QMessageBox.warning(self, "合并失败", f"合并 PDF 失败：{e}")
            self.append_log(f"合并失败：{e}")
        finally:
            merger.close()

    # ---- 自动更新 ----
    def manual_check_update(self):
        self.check_for_updates(background=False)

    def check_for_updates(self, background: bool = True):
        if requests is None:
            self.append_log("更新检查：requests 未安装。")
            if not background:
                QMessageBox.warning(self, "更新检查", "requests 未安装，无法检查更新。")
            return

        api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
        try:
            self.append_log("检查更新...")
            r = requests.get(api_url, timeout=10)
            if r.status_code != 200:
                self.append_log(f"检查更新失败：HTTP {r.status_code}")
                if not background:
                    QMessageBox.warning(self, "更新检查", f"HTTP {r.status_code}")
                return
            data = r.json()
            tag_name = data.get("tag_name", "")
            latest_version = tag_name.lstrip("v")
            if self._is_newer_version(latest_version, APP_VERSION):
                assets = data.get("assets", [])
                chosen = None
                for a in assets:
                    name = a.get("name", "")
                    if name.startswith(ASSET_PREFIX) and name.endswith(".exe"):
                        chosen = a
                        break
                if not chosen:
                    self.append_log("找到新版 {latest_version}，但无安装包资产。")
                    if not background:
                        QMessageBox.information(self, "更新", f"找到新版 {latest_version}，但无安装包。")
                    return

                checksum_asset = None
                possible_names = [chosen["name"] + ".sha256", chosen["name"] + ".sha256.txt"]
                for a in assets:
                    if a.get("name", "") in possible_names:
                        checksum_asset = a
                        break

                if not background:
                    ask = QMessageBox.question(self, "更新可用",
                                               f"发现新版本 {latest_version}（当前 {APP_VERSION}），是否下载并安装？")
                    if ask != QMessageBox.Yes:
                        return

                download_url = chosen.get("browser_download_url")
                checksum_url = checksum_asset.get("browser_download_url") if checksum_asset else None
                self._download_and_run_installer(download_url, chosen.get("name"), checksum_url)
            else:
                self.append_log("当前已是最新版本。")
                if not background:
                    QMessageBox.information(self, "更新检查", "当前已是最新版本。")
        except Exception as e:
            self.append_log(f"检查更新异常：{e}")
            if not background:
                QMessageBox.warning(self, "更新检查", f"检查更新失败：{e}")

    def _is_newer_version(self, v_new: str, v_current: str) -> bool:
        def parse(v):
            parts = []
            for x in v.split("."):
                try:
                    parts.append(int(x))
                except Exception:
                    parts.append(0)
            return parts
        return parse(v_new) > parse(v_current)

    def _download_and_run_installer(self, url: str, name: str, checksum_url: Optional[str]):
        if requests is None:
            self.append_log("requests 未安装，无法下载更新。")
            return
        tmp_installer = None
        try:
            self.append_log(f"下载更新：{name}")
            with requests.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                fd, tmp_installer = tempfile.mkstemp(suffix=".exe", prefix="installer_")
                os.close(fd)
                with open(tmp_installer, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
            self.append_log(f"下载完成：{tmp_installer}")

            expected_hash = None
            if checksum_url:
                try:
                    r2 = requests.get(checksum_url, timeout=10)
                    r2.raise_for_status()
                    expected_hash = parse_sha256_text(r2.text)
                    if not expected_hash:
                        self.append_log("无法解析校验文件，跳过校验。")
                except Exception as e:
                    self.append_log(f"获取校验文件失败：{e}")

            if expected_hash:
                actual = sha256_of_file(tmp_installer)
                self.append_log(f"SHA256 校验：{actual[:16]}...")
                if actual.lower() != expected_hash.lower():
                    self.append_log("校验失败，取消安装。")
                    try:
                        os.remove(tmp_installer)
                    except Exception:
                        pass
                    QMessageBox.critical(self, "校验失败", "下载的安装包校验失败，取消安装。")
                    return

            if IS_WINDOWS:
                # Copy installer to Desktop for easy user access
                import shutil
                desktop_installer = str(Path.home() / "Desktop" / f"PDFConverter-setup-{APP_VERSION}.exe")
                try:
                    shutil.copy2(tmp_installer, desktop_installer)
                    launch_path = desktop_installer
                except Exception:
                    launch_path = tmp_installer

                launched = False
                # Method 1: PowerShell Start-Process -Verb RunAs
                try:
                    import subprocess
                    subprocess.Popen(
                        ["powershell", "-NoProfile", "-Command",
                         f'Start-Process -FilePath "{launch_path}" -Verb RunAs'],
                        creationflags=0x08000000
                    )
                    launched = True
                except Exception:
                    pass
                # Method 2: ShellExecuteW runas
                if not launched:
                    try:
                        import ctypes
                        ctypes.windll.shell32.ShellExecuteW(None, "runas", launch_path, None, None, 1)
                        launched = True
                    except Exception:
                        pass
                self.append_log(f"安装包已下载到桌面：{desktop_installer}")
                if launched:
                    QMessageBox.information(self, "更新", "安装程序已启动，请在 UAC 对话框中点击「是」以完成安装。")
                else:
                    QMessageBox.information(self, "更新",
                        f"未能自动启动安装程序。\n\n安装包已保存到桌面：\n{desktop_installer}\n\n"
                        "请右键点击该文件 → 「以管理员身份运行」")
            else:
                self.append_log("自动安装仅支持 Windows。")
                QMessageBox.information(self, "更新", "已下载更新，但自动安装仅支持 Windows。")
        except Exception as e:
            self.append_log(f"下载/运行安装器失败：{e}")
            QMessageBox.warning(self, "更新失败", f"下载/运行安装器失败：{e}")


    def show_about(self):
        QMessageBox.about(self, f"关于 {APP_NAME}",
            f"<h3>{APP_NAME}</h3>"
            f"<p>版本 {APP_VERSION}</p>"
            f"<p>Word/Excel → PDF 转换器（PySide6 GUI）</p>"
            f"<hr>"
            f"<p>支持格式：doc/docx/xls/xlsx/pdf/odt/ods/rtf</p>"
            f"<p>功能：批量转换、多线程、PDF合并、自动更新</p>"
            f"<hr>"
            f"<p><a href='https://github.com/{GITHUB_OWNER}/{GITHUB_REPO}'>GitHub</a></p>"
        )

def main():
    app = QApplication(sys.argv)
    w = ConverterApp()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()













