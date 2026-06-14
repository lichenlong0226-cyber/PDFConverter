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

使用前须修改（在构建前）:
- 在下方 CONFIG 区设置 GITHUB_OWNER 与 GITHUB_REPO 为你的仓库（已设置为 lichenlong0226-cyber/pdf）
- 设置 APP_NAME（默认 PDFConverter）和 APP_VERSION（与 release tag 对应）
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
from PySide6.QtGui import QIcon, QAction, QCursor
from PySide6.QtWidgets import (QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
                               QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
                               QAbstractItemView, QHeaderView, QMessageBox, QCheckBox,
                               QProgressBar, QMenu, QTextEdit, QSplitter, QLineEdit)

IS_WINDOWS = platform.system() == "Windows"

# Optional Windows COM
if IS_WINDOWS:
    try:
        import pythoncom  # noqa: F401
        import win32com.client
    except Exception:
        win32com = None
else:
    win32com = None

# PDF merger lib
try:
    from pypdf import PdfMerger
except Exception:
    PdfMerger = None

# Networking for auto-update
try:
    import requests
except Exception:
    requests = None

SUPPORTED_EXT = (".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".xlsb",
                 ".odt", ".ods", ".rtf", ".docm")

# ----------------- CONFIG (请在 build 前调整为你的仓库信息) -----------------
APP_NAME = "PDFConverter"
APP_VERSION = "1.0.0"               # 构建时应与 release tag 对应（例如 v1.0.0 -> "1.0.0")
# 已替换为你提供的仓库
GITHUB_OWNER = "lichenlong0226-cyber"
GITHUB_REPO = "pdf"
ASSET_PREFIX = f"{APP_NAME}-setup-"  # 安装包前缀（workflow 也将生成以此为前缀的文件）
# -------------------------------------------------------------------------


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
    if PdfMerger is None:
        raise RuntimeError("pypdf is required for merging PDFs. Install with `pip install pypdf`.")

# ----------------- Worker & UI classes -----------------
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
            if IS_WINDOWS and ext in (".doc", ".docx", ".docm", ".rtf"):
                self.signals.log.emit(f"使用 Windows Word COM 转换：{self.in_path}")
                convert_word_windows(self.in_path, str(target))
            elif IS_WINDOWS and ext in (".xls", ".xlsx", ".xlsm", ".xlsb"):
                self.signals.log.emit(f"使用 Windows Excel COM 导出（所有 sheets）：{self.in_path}")
                convert_excel_windows(self.in_path, str(target))
            else:
                self.signals.log.emit(f"使用 LibreOffice 转换（后台）：{self.in_path}")
                tmpd = tempfile.mkdtemp()
                try:
                    outpdf = convert_with_libreoffice(self.in_path, tmpd)
                    shutil.move(outpdf, str(target))
                finally:
                    shutil.rmtree(tmpd, ignore_errors=True)
            self.signals.finished.emit(self.in_path, str(target))
        except Exception as e:
            err = f"ERR: {log_exc_text(e)}"
            self.signals.log.emit(f"转换失败：{self.in_path} -> {err}")
            self.signals.finished.emit(self.in_path, err)

class DropTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(0, 3, parent)
        self.setHorizontalHeaderLabels(["文件", "状态", "大小"])
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setAcceptDrops(True)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        for u in urls:
            path = u.toLocalFile()
            if os.path.isdir(path):
                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORTED_EXT):
                            self.add_file(os.path.join(root, f))
            else:
                self.add_file(path)
        event.acceptProposedAction()

    def add_file(self, path):
        if not os.path.exists(path):
            return False
        ext = Path(path).suffix.lower()
        if ext not in SUPPORTED_EXT:
            return False
        for row in range(self.rowCount()):
            if self.item(row, 0).text() == path:
                return False
        row = self.rowCount()
        self.insertRow(row)
        size_text = f"{Path(path).stat().st_size // 1024} KB"
        self.setItem(row, 0, QTableWidgetItem(path))
        self.setItem(row, 1, QTableWidgetItem("待处理"))
        self.setItem(row, 2, QTableWidgetItem(size_text))
        return True

    def _on_context_menu(self, pos):
        row = self.indexAt(pos).row()
        if row < 0:
            return
        path = self.item(row, 0).text()
        menu = QMenu()
        open_act = QAction("打开文件所在目录", self)
        open_act.triggered.connect(lambda: self._open_folder(path))
        remove_act = QAction("移除", self)
        remove_act.triggered.connect(lambda: self.removeRow(row))
        retry_act = QAction("重试转换", self)
        retry_act.triggered.connect(lambda: self.parent().retry_single(path))
        menu.addAction(open_act)
        menu.addAction(remove_act)
        menu.addAction(retry_act)
        menu.exec(QCursor.pos())

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
        self.resize(1000, 600)
        icon_path = os.path.join(os.path.dirname(__file__), "app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.layout = QVBoxLayout(self)

        header = QLabel("<h2>Word/Excel → PDF 转换器</h2>")
        header.setTextFormat(Qt.RichText)
        self.layout.addWidget(header)

        top_row = QHBoxLayout()
        hint = QLabel("拖拽文件到下方列表，或使用“添加文件”。 支持 doc/docx/xls/xlsx/ods 等。")
        top_row.addWidget(hint)
        self.btn_check_update = QPushButton("检查更新")
        self.btn_check_update.clicked.connect(self.manual_check_update)
        top_row.addWidget(self.btn_check_update)
        self.layout.addLayout(top_row)

        splitter = QSplitter(Qt.Horizontal)
        left = QVBoxLayout()
        container_left = QWidget()
        container_left.setLayout(left)
        splitter.addWidget(container_left)

        right = QVBoxLayout()
        container_right = QWidget()
        container_right.setLayout(right)
        splitter.addWidget(container_right)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        self.layout.addWidget(splitter)

        self.table = DropTable(self)
        left.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_add = QPushButton("添加文件...")
        self.btn_remove = QPushButton("移除选中")
        self.btn_clear = QPushButton("清空")
        btn_row.addWidget(self.btn_add)
        btn_row.addWidget(self.btn_remove)
        btn_row.addWidget(self.btn_clear)
        left.addLayout(btn_row)

        out_row = QHBoxLayout()
        self.out_label = QLabel("输出目录:")
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("留空使用当前目录或选择输出目录")
        self.btn_out = QPushButton("选择输出目录")
        out_row.addWidget(self.out_label)
        out_row.addWidget(self.out_edit)
        out_row.addWidget(self.btn_out)
        left.addLayout(out_row)

        ops_row = QHBoxLayout()
        self.chk_merge = QCheckBox("转换后合并为单个 PDF")
        self.btn_convert = QPushButton("开始转换")
        self.btn_cancel = QPushButton("取消所有")
        self.progress = QProgressBar()
        self.progress.setMinimum(0)
        self.progress.setValue(0)
        ops_row.addWidget(self.chk_merge)
        ops_row.addWidget(self.btn_convert)
        ops_row.addWidget(self.btn_cancel)
        ops_row.addWidget(self.progress)
        left.addLayout(ops_row)

        log_label = QLabel("日志")
        right.addWidget(log_label)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        right.addWidget(self.log_edit)

        self.btn_add.clicked.connect(self.open_add_files)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_clear.clicked.connect(self.clear_all)
        self.btn_out.clicked.connect(self.choose_out_dir)
        self.btn_convert.clicked.connect(self.start_conversion)
        self.btn_cancel.clicked.connect(self.cancel_all)

        self.pool = QThreadPool.globalInstance()
        self.pool.setMaxThreadCount(max(1, min(8, os.cpu_count() or 4)))
        self.active_workers = {}

        self.update_timer = QTimer(self)
        self.update_timer.setInterval(1000 * 60 * 60 * 24)  # daily
        self.update_timer.timeout.connect(lambda: self.check_for_updates(background=True))
        self.update_timer.start()

        self.output_dir = os.getcwd()
        self.cancel_requested = False

        self.setStyleSheet("""
            QWidget { font-family: "Segoe UI", Arial, sans-serif; font-size: 11px; }
            QPushButton { padding: 6px 10px; }
            QProgressBar { min-width: 150px; }
            QTextEdit { background: #111; color: #eee; font-family: monospace; }
        """)

    def open_add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", os.getcwd(),
                                                "Word/Excel Files (*.doc *.docx *.xls *.xlsx *.xlsm *.odt *.ods *.rtf)")
        for f in files:
            self.table.add_file(f)

    def remove_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.table.removeRow(r)

    def clear_all(self):
        self.table.setRowCount(0)

    def choose_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.output_dir or os.getcwd())
        if d:
            self.output_dir = d
            self.out_edit.setText(d)

    def retry_single(self, path):
        self.table.add_file(path)
        if not self.active_workers:
            self.start_conversion()

    def append_log(self, text: str):
        self.log_edit.append(text)

    def cancel_all(self):
        self.append_log("取消请求：正在等待正在运行的线程结束（不可立即中止 COM 调用）。")
        self.cancel_requested = True

    def start_conversion(self):
        n = self.table.rowCount()
        if n == 0:
            QMessageBox.information(self, "提示", "请先添加要转换的文件。")
            return
        if IS_WINDOWS and win32com is None:
            QMessageBox.critical(self, "错误", "Windows: 需要安装 pywin32（pip install pywin32）。")
            return
        if self.chk_merge.isChecked() and PdfMerger is None:
            QMessageBox.critical(self, "错误", "合并功能需要 pypdf（pip install pypdf）。")
            return
        if requests is None:
            self.append_log("警告：requests 未安装，自动更新功能不可用（pip install requests）。")

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
        self.cancel_requested = False
        self.append_log(f"开始转换，共 {n} 个文件，输出目录：{self.output_dir}")

        self.pdfs_generated = []
        self.remaining = n

        for i in range(n):
            in_path = self.table.item(i, 0).text()
            self.table.setItem(i, 1, QTableWidgetItem("等待中"))
            worker = ConvertWorker(in_path, self.output_dir)
            worker.signals.started.connect(lambda p: self.on_started(p))
            worker.signals.log.connect(lambda s: self.on_worker_log(s))
            worker.signals.finished.connect(lambda p, out: self.on_finished(p, out))
            self.active_workers[in_path] = worker
            self.pool.start(worker)

    def on_started(self, path):
        self.append_log(f"[启动] {path}")
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0).text() == path:
                self.table.setItem(r, 1, QTableWidgetItem("处理中"))

    def on_worker_log(self, text):
        self.append_log(text)

    def on_finished(self, path, out_or_err):
        self.active_workers.pop(path, None)
        for r in range(self.table.rowCount()):
            if self.table.item(r, 0).text() == path:
                if out_or_err.startswith("ERR:"):
                    self.table.setItem(r, 1, QTableWidgetItem(f"失败"))
                    self.append_log(f"[失败] {path} -> {out_or_err}")
                else:
                    self.table.setItem(r, 1, QTableWidgetItem("已完成"))
                    self.append_log(f"[完成] {path} -> {out_or_err}")
                    self.pdfs_generated.append(out_or_err)
                break
        self.progress.setValue(self.progress.value() + 1)
        self.remaining -= 1
        if self.remaining <= 0 or (self.cancel_requested and not self.active_workers):
            self.append_log("全部任务已结束。")
            if self.chk_merge.isChecked() and self.pdfs_generated:
                self.merge_after_convert()
            else:
                QMessageBox.information(self, "完成", f"已完成转换，生成 {len(self.pdfs_generated)} 个 PDF，输出目录：\n{self.output_dir}")

    def merge_after_convert(self):
        ensure_pdf_merger_available()
        default_name = os.path.join(self.output_dir, "merged.pdf")
        merged_name, _ = QFileDialog.getSaveFileName(self, "保存合并后的 PDF 为", default_name, "PDF Files (*.pdf)")
        if not merged_name:
            QMessageBox.information(self, "完成", "已完成转换（未保存合并结果）。")
            return
        merger = PdfMerger()
        try:
            for p in self.pdfs_generated:
                merger.append(p)
            merger.write(merged_name)
            QMessageBox.information(self, "完成", f"已完成转换并合并，合并文件：\n{merged_name}")
            self.append_log(f"合并完成：{merged_name}")
        except Exception as e:
            QMessageBox.warning(self, "合并失败", f"合并 PDF 失败：{e}")
            self.append_log(f"合并失败：{e}")
        finally:
            merger.close()

    # ----------------- 自动更新（SHA256 校验） -----------------
    def manual_check_update(self):
        self.check_for_updates(background=False)

    def check_for_updates(self, background: bool = True):
        if requests is None:
            self.append_log("自动更新：requests 未安装，无法检查更新。")
            if not background:
                QMessageBox.warning(self, "更新检查", "requests 未安装，无法检查更新（pip install requests）。")
            return

        api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
        try:
            self.append_log("检查更新中...")
            r = requests.get(api_url, timeout=10)
            if r.status_code != 200:
                self.append_log(f"更新检查失败：HTTP {r.status_code}")
                if not background:
                    QMessageBox.warning(self, "更新检查", f"更新检查失败：HTTP {r.status_code}")
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
                    self.append_log("找到新版，但没有匹配的安装包资产（.exe）。")
                    if not background:
                        QMessageBox.information(self, "更新", f"找到新版 {latest_version}，但未找到安装包资产。")
                    return

                checksum_asset = None
                possible_names = [chosen["name"] + ".sha256", chosen["name"] + ".sha256.txt", "checksums.json"]
                for a in assets:
                    if a.get("name", "") in possible_names:
                        checksum_asset = a
                        break

                if not background:
                    ask = QMessageBox.question(self, "更新可用",
                                               f"发现新版本 {latest_version}（当前 {APP_VERSION}），是否下载并运行安装？")
                    if ask != QMessageBox.Yes:
                        return

                download_url = chosen.get("browser_download_url")
                checksum_url = checksum_asset.get("browser_download_url") if checksum_asset else None
                self._download_and_run_installer(download_url, chosen.get("name"), checksum_url)
            else:
                self.append_log("当前为最新版本。")
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
            self.append_log(f"开始下载更新：{url}")
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
                    self.append_log(f"下载校验文件：{checksum_url}")
                    r2 = requests.get(checksum_url, timeout=10)
                    r2.raise_for_status()
                    txt = r2.text
                    expected_hash = parse_sha256_text(txt)
                    if not expected_hash:
                        self.append_log("无法解析校验文件内容，跳过校验。")
                except Exception as e:
                    self.append_log(f"获取校验文件失败：{e}")

            if expected_hash:
                actual = sha256_of_file(tmp_installer)
                self.append_log(f"校验：实际 {actual} 期望 {expected_hash}")
                if actual.lower() != expected_hash.lower():
                    self.append_log("校验失败：下载的安装包哈希与发布页面不匹配，已删除下载文件。")
                    try:
                        os.remove(tmp_installer)
                    except Exception:
                        pass
                    QMessageBox.critical(self, "更新校验失败", "下载的安装包校验失败，取消安装。")
                    return

            if IS_WINDOWS:
                subprocess.Popen([tmp_installer], shell=False)
                self.append_log("已启动安装程序，程序将继续运行；请手动完成安装步骤。")
                QMessageBox.information(self, "更新", "安装程序已启动，完成安装后请重新启动程序。")
            else:
                self.append_log("自动安装仅支持 Windows 可执行安装包（.exe）。")
                QMessageBox.information(self, "更新", "已下载更新，但自动运行仅支持 Windows 可执行安装包。")
        except Exception as e:
            self.append_log(f"下载或运行安装器失败：{e}")
            QMessageBox.warning(self, "更新失败", f"下载或运行安装器失败：{e}")

def main():
    app = QApplication(sys.argv)
    w = ConverterApp()
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
