"""
MD Toolkit — GUI con PyQt6
Herramientas para convertir DOCX → MD, unir archivos y procesar carpetas.

Instalación:
    pip install PyQt6
    brew install pandoc   ← requerido para la conversión

Uso:
    python md_toolkit.py
"""

import sys
import os
import glob
import threading
import subprocess
from pathlib import Path
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QLineEdit, QListWidget,
    QFileDialog, QTextEdit, QComboBox, QCheckBox, QFrame,
    QSplitter, QGroupBox, QAbstractItemView, QStatusBar, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QFont, QColor, QPalette, QIcon


# ════════════════════════════════════════════════════════════════════════════
#  DEBUG GLOBAL
# ════════════════════════════════════════════════════════════════════════════

class DebugConfig:
    """Singleton para controlar el modo debug desde cualquier parte de la app."""
    enabled = False
    _log_fn = None

    @classmethod
    def set_log(cls, fn):
        cls._log_fn = fn

    @classmethod
    def log(cls, msg: str):
        if cls.enabled and cls._log_fn:
            cls._log_fn(f"🔍 [DEBUG] {msg}")


# ════════════════════════════════════════════════════════════════════════════
#  ESTILOS
# ════════════════════════════════════════════════════════════════════════════

STYLESHEET = """
QMainWindow, QWidget {
    background-color: #1e1e2e;
    color: #cdd6f4;
    font-family: 'SF Pro Display', 'Helvetica Neue', Arial, sans-serif;
    font-size: 13px;
}
QTabWidget::pane {
    border: 1px solid #313244;
    border-radius: 8px;
    background-color: #181825;
}
QTabBar::tab {
    background-color: #313244;
    color: #a6adc8;
    padding: 8px 20px;
    margin-right: 2px;
    border-radius: 6px 6px 0 0;
    font-weight: 500;
}
QTabBar::tab:selected {
    background-color: #89b4fa;
    color: #1e1e2e;
    font-weight: 700;
}
QTabBar::tab:hover:!selected {
    background-color: #45475a;
    color: #cdd6f4;
}
QPushButton {
    background-color: #89b4fa;
    color: #1e1e2e;
    border: none;
    padding: 8px 16px;
    border-radius: 6px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #b4d0ff;
}
QPushButton:pressed {
    background-color: #6a9fd8;
}
QPushButton.danger {
    background-color: #f38ba8;
    color: #1e1e2e;
}
QPushButton.danger:hover {
    background-color: #f5a0b5;
}
QPushButton.secondary {
    background-color: #45475a;
    color: #cdd6f4;
}
QPushButton.secondary:hover {
    background-color: #585b70;
}
QPushButton.action {
    background-color: #a6e3a1;
    color: #1e1e2e;
    font-size: 14px;
    padding: 10px 20px;
}
QPushButton.action:hover {
    background-color: #b9f0b4;
}
QListWidget {
    background-color: #11111b;
    border: 1px solid #313244;
    border-radius: 6px;
    padding: 4px;
    color: #cdd6f4;
    font-family: 'SF Mono', 'Menlo', monospace;
    font-size: 12px;
}
QListWidget::item {
    padding: 4px 8px;
    border-radius: 4px;
}
QListWidget::item:selected {
    background-color: #313244;
    color: #89b4fa;
}
QListWidget::item:hover {
    background-color: #1e1e2e;
}
QTextEdit {
    background-color: #11111b;
    border: 1px solid #313244;
    border-radius: 6px;
    padding: 8px;
    color: #a6e3a1;
    font-family: 'SF Mono', 'Menlo', monospace;
    font-size: 12px;
}
QLineEdit {
    background-color: #11111b;
    border: 1px solid #313244;
    border-radius: 6px;
    padding: 6px 10px;
    color: #cdd6f4;
}
QLineEdit:focus {
    border-color: #89b4fa;
}
QComboBox {
    background-color: #313244;
    border: 1px solid #45475a;
    border-radius: 6px;
    padding: 6px 10px;
    color: #cdd6f4;
    min-width: 160px;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #313244;
    border: 1px solid #45475a;
    selection-background-color: #89b4fa;
    selection-color: #1e1e2e;
}
QCheckBox {
    color: #cdd6f4;
    spacing: 6px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border-radius: 4px;
    border: 2px solid #45475a;
    background-color: #11111b;
}
QCheckBox::indicator:checked {
    background-color: #89b4fa;
    border-color: #89b4fa;
}
QGroupBox {
    border: 1px solid #313244;
    border-radius: 8px;
    margin-top: 12px;
    padding: 12px 8px 8px 8px;
    font-weight: 600;
    color: #a6adc8;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #89b4fa;
}
QLabel {
    color: #a6adc8;
}
QLabel.title {
    color: #cdd6f4;
    font-size: 18px;
    font-weight: 700;
}
QLabel.subtitle {
    color: #6c7086;
    font-size: 12px;
}
QStatusBar {
    background-color: #181825;
    color: #6c7086;
    border-top: 1px solid #313244;
    font-size: 11px;
}
QSplitter::handle {
    background-color: #313244;
}
"""


# ════════════════════════════════════════════════════════════════════════════
#  LÓGICA DE CONVERSIÓN  (usa pandoc, igual que tus scripts originales)
# ════════════════════════════════════════════════════════════════════════════

SEPARATOR = '=' * 80


def check_pandoc() -> bool:
    """Verifica que pandoc esté disponible."""
    try:
        subprocess.run(['pandoc', '--version'], capture_output=True, timeout=5)
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def safe_dest_path(path: Path) -> Path:
    """Si el archivo ya existe, agrega _New (o _New2, _New3, etc.) al nombre."""
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    candidate = parent / f"{stem}_New{suffix}"
    counter = 2
    while candidate.exists():
        candidate = parent / f"{stem}_New{counter}{suffix}"
        counter += 1
    return candidate


def docx_to_md_file(docx_path: str, out_dir: str = None) -> tuple[bool, str, str]:
    """
    Convierte un .docx a .md usando pandoc (igual que word2md.sh / word2mdFolder.sh).
    Extrae imágenes a una subcarpeta <nombre>_media/.
    Si el .md destino ya existe lo renombra agregando _New.
    Retorna (éxito, ruta_md, mensaje).
    """
    docx = Path(docx_path)
    dest_dir = Path(out_dir) if out_dir else docx.parent
    dest_dir.mkdir(parents=True, exist_ok=True)

    md_path = safe_dest_path(dest_dir / f"{docx.stem}.md")
    media_folder = dest_dir / f"{docx.stem}_media"

    try:
        result = subprocess.run(
            ['pandoc', str(docx), '-o', str(md_path),
             f'--extract-media={media_folder}'],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0:
            img_count = len(list(media_folder.rglob('*'))) if media_folder.exists() else 0
            msg = f"✓  {docx.name}  →  {md_path.name}"
            if img_count:
                msg += f"  ({img_count} imágenes en {media_folder.name}/)"
            return True, str(md_path), msg
        else:
            return False, "", f"✗  {docx.name}: {result.stderr.strip()}"
    except subprocess.TimeoutExpired:
        return False, "", f"✗  {docx.name}: timeout"
    except FileNotFoundError:
        return False, "", "✗  pandoc no encontrado. Instalalo con: brew install pandoc"


def build_consolidated_md(md_files: list, output_path: str, input_dir: str) -> None:
    """
    Une varios .md en un archivo consolidado con índice, separadores y metadata,
    igual que word2mdFolderConsolidado.py / unir_markdownFolder.py.
    """
    paths = [Path(p) for p in md_files]
    with open(output_path, 'w', encoding='utf-8') as out:
        # Encabezado
        out.write("# Documentación Consolidada\n\n")
        out.write(f"**Fecha de generación**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        out.write(f"**Archivos incluidos**: {len(paths)}\n")
        out.write(f"**Directorio origen**: `{input_dir}`\n\n")

        # Índice
        out.write("## 📑 Índice de Archivos\n\n")
        for idx, p in enumerate(paths, 1):
            anchor = p.stem.lower().replace(' ', '-').replace('_', '-')
            out.write(f"{idx}. [{p.name}](#{anchor})\n")
        out.write(f"\n{SEPARATOR}\n\n")

        # Contenido
        for idx, p in enumerate(paths, 1):
            anchor = p.stem.lower().replace(' ', '-').replace('_', '-')
            out.write(f"\n{SEPARATOR}\n")
            out.write(f"## 📄 Archivo {idx}: {p.name}\n")
            out.write(f"<a name='{anchor}'></a>\n\n")
            out.write(f"**Ruta**: `{p}`\n\n")
            out.write(f"{SEPARATOR}\n\n")
            try:
                out.write(p.read_text(encoding='utf-8'))
            except Exception as e:
                out.write(f"*[Error al leer el archivo: {e}]*")
            out.write("\n\n")

        # Footer
        out.write(f"\n{SEPARATOR}\n")
        out.write(f"## 📊 Resumen\n\n")
        out.write(f"- **Total archivos procesados**: {len(paths)}\n")
        out.write(f"- **Fecha**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        out.write(f"- **Generado por**: MD Toolkit\n")
        out.write(f"\n{SEPARATOR}\n")


# ════════════════════════════════════════════════════════════════════════════
#  WORKER THREAD  (para no bloquear la UI)
# ════════════════════════════════════════════════════════════════════════════

class Worker(QObject):
    log = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self._fn = fn
        self._args = args
        self._kwargs = kwargs

    def run(self):
        self._fn(self.log.emit, *self._args, **self._kwargs)
        self.finished.emit()


# ════════════════════════════════════════════════════════════════════════════
#  WIDGET: LISTA DE ARCHIVOS
# ════════════════════════════════════════════════════════════════════════════

class FileListWidget(QWidget):
    def __init__(self, extensions=None, parent=None):
        super().__init__(parent)
        self.extensions = extensions or []
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # Botones
        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)

        self.btn_files = QPushButton("＋ Archivos")
        self.btn_folder = QPushButton("＋ Carpeta")
        self.btn_remove = QPushButton("✕ Quitar")
        self.btn_clear = QPushButton("Limpiar")
        self.btn_remove.setProperty("class", "danger")
        self.btn_clear.setProperty("class", "secondary")

        for b in [self.btn_files, self.btn_folder, self.btn_remove, self.btn_clear]:
            b.setFixedHeight(32)
            btn_row.addWidget(b)
        btn_row.addStretch()

        self.count_label = QLabel("0 archivos")
        self.count_label.setProperty("class", "subtitle")
        btn_row.addWidget(self.count_label)

        layout.addLayout(btn_row)

        # Lista
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout.addWidget(self.list_widget)

        # Conexiones
        self.btn_files.clicked.connect(self._add_files)
        self.btn_folder.clicked.connect(self._add_folder)
        self.btn_remove.clicked.connect(self._remove_selected)
        self.btn_clear.clicked.connect(self._clear)

        # Drag & drop
        self.list_widget.setAcceptDrops(True)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path):
                self._add_path(path)
            elif os.path.isdir(path):
                self._add_dir(path)

    @property
    def files(self):
        return [self.list_widget.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.list_widget.count())]

    def _filter(self, path):
        """Valida extensión case-insensitive (.md .MD .Md etc)."""
        if not self.extensions:
            return True
        ext = os.path.splitext(path)[1].lstrip('.').lower()
        allowed = [e.lstrip('.').lower() for e in self.extensions]
        return ext in allowed

    def _add_path(self, path):
        # Normalizar path (resuelve aliases y symlinks de macOS)
        path = os.path.realpath(path)
        DebugConfig.log(f"_add_path: path='{path}'")
        pasa = self._filter(path)
        DebugConfig.log(f"_add_path: _filter={pasa} | extensiones={self.extensions}")
        if not pasa:
            DebugConfig.log(f"_add_path: RECHAZADO por filtro")
            return
        existing = self.files
        if path in existing:
            DebugConfig.log(f"_add_path: ya existe, ignorado")
            return
        self.list_widget.addItem(os.path.basename(path))
        it = self.list_widget.item(self.list_widget.count() - 1)
        it.setData(Qt.ItemDataRole.UserRole, path)
        it.setToolTip(path)
        DebugConfig.log(f"_add_path: AGREGADO '{os.path.basename(path)}'")
        self._update_count()

    def _add_dir(self, folder):
        exts = self.extensions or ["*"]
        DebugConfig.log(f"_add_dir: folder='{folder}' | exts={exts}")
        for ext in exts:
            found = glob.glob(os.path.join(folder, f"*.{ext.lstrip('.')}"))
            DebugConfig.log(f"_add_dir: glob *.{ext} → {len(found)} archivos: {found}")
            for p in found:
                self._add_path(p)

    def _add_files(self):
        # Filtro según extensión esperada, con fallback a "todos" por si macOS oculta algo
        if self.extensions:
            exts_lower = [e.lstrip('.').lower() for e in self.extensions]
            filt_parts = " ".join(f"*.{e} *.{e.upper()}" for e in exts_lower)
            label = "/".join(e.upper() for e in exts_lower)
            filetypes = f"{label} ({filt_parts});;Todos los archivos (*)"
        else:
            filetypes = "Todos los archivos (*)"
        paths, _ = QFileDialog.getOpenFileNames(self, "Seleccionar archivos", "", filetypes)
        DebugConfig.log(f"_add_files: extensiones={self.extensions} | paths recibidos={paths}")
        for p in paths:
            self._add_path(p)

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta")
        DebugConfig.log(f"_add_folder: carpeta='{folder}'")
        if folder:
            self._add_dir(folder)

    def _remove_selected(self):
        for item in self.list_widget.selectedItems():
            self.list_widget.takeItem(self.list_widget.row(item))
        self._update_count()

    def _clear(self):
        self.list_widget.clear()
        self._update_count()

    def _update_count(self):
        n = self.list_widget.count()
        self.count_label.setText(f"{n} archivo{'s' if n != 1 else ''}")


# ════════════════════════════════════════════════════════════════════════════
#  PANEL DE LOG
# ════════════════════════════════════════════════════════════════════════════

class LogPanel(QWidget):
    append = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        header = QHBoxLayout()
        lbl = QLabel("  📋 Log de operaciones")
        lbl.setProperty("class", "subtitle")
        header.addWidget(lbl)
        header.addStretch()
        btn_clear = QPushButton("Limpiar")
        btn_clear.setProperty("class", "secondary")
        btn_clear.setFixedSize(70, 24)
        header.addWidget(btn_clear)
        layout.addLayout(header)

        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setFixedHeight(130)
        layout.addWidget(self.text)

        btn_clear.clicked.connect(self.text.clear)
        self.append.connect(self._append_safe)

    def _append_safe(self, msg):
        self.text.append(msg)
        self.text.verticalScrollBar().setValue(
            self.text.verticalScrollBar().maximum())

    def log(self, msg):
        self.append.emit(msg)


# ════════════════════════════════════════════════════════════════════════════
#  TABS
# ════════════════════════════════════════════════════════════════════════════

class BaseTab(QWidget):
    def __init__(self, log: LogPanel, parent=None):
        super().__init__(parent)
        self.log = log
        self._thread = None
        self._worker = None

    def _run_in_thread(self, fn, *args, **kwargs):
        self._thread = QThread()
        self._worker = Worker(fn, *args, **kwargs)
        self._worker.moveToThread(self._thread)
        self._worker.log.connect(self.log.log)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._thread.quit)
        self._thread.start()


class DocxToMdTab(BaseTab):
    def __init__(self, log, parent=None):
        super().__init__(log, parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Opciones de salida
        out_group = QGroupBox("Destino")
        out_layout = QHBoxLayout(out_group)
        self.out_same = QCheckBox("Misma carpeta que el archivo original")
        self.out_same.setChecked(True)
        self.out_same.toggled.connect(self._toggle_out)
        out_layout.addWidget(self.out_same)
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("Carpeta de destino…")
        self.out_edit.setEnabled(False)
        out_layout.addWidget(self.out_edit)
        self.out_btn = QPushButton("📁")
        self.out_btn.setFixedWidth(36)
        self.out_btn.setEnabled(False)
        self.out_btn.clicked.connect(self._pick_out)
        out_layout.addWidget(self.out_btn)
        layout.addWidget(out_group)

        # Lista
        list_group = QGroupBox("Archivos DOCX  (arrastrá o usá los botones)")
        list_layout = QVBoxLayout(list_group)
        self.flist = FileListWidget(extensions=["docx"])
        list_layout.addWidget(self.flist)
        layout.addWidget(list_group)

        # Botón acción
        self.run_btn = QPushButton("⚙   Convertir DOCX → Markdown")
        self.run_btn.setProperty("class", "action")
        self.run_btn.setFixedHeight(42)
        self.run_btn.clicked.connect(self._run)
        layout.addWidget(self.run_btn)

    def _toggle_out(self, checked):
        self.out_edit.setEnabled(not checked)
        self.out_btn.setEnabled(not checked)

    def _pick_out(self):
        folder = QFileDialog.getExistingDirectory(self, "Carpeta destino")
        if folder:
            self.out_edit.setText(folder)

    def _run(self):
        files = self.flist.files
        if not files:
            self.log.log("⚠  Sin archivos. Agregá al menos un .docx.")
            return
        out_dir = None if self.out_same.isChecked() else self.out_edit.text()
        self._run_in_thread(self._convert, files, out_dir)

    @staticmethod
    def _convert(emit, files, out_dir):
        if not check_pandoc():
            emit("✗  pandoc no encontrado. Instalalo con: brew install pandoc")
            return
        emit(f"\n▶  Convirtiendo {len(files)} archivo(s) con pandoc…")
        ok = 0
        for path in files:
            success, _, msg = docx_to_md_file(path, out_dir)
            emit(f"  {msg}")
            if success:
                ok += 1
        emit(f"✔  Listo: {ok}/{len(files)} convertidos.")


class MergeDocxTab(BaseTab):
    def __init__(self, log, parent=None):
        super().__init__(log, parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        sep_group = QGroupBox("Separador entre documentos")
        sep_layout = QHBoxLayout(sep_group)
        sep_layout.addWidget(QLabel("Texto separador:"))
        self.sep_edit = QLineEdit("\\n\\n---\\n\\n")
        sep_layout.addWidget(self.sep_edit)
        layout.addWidget(sep_group)

        list_group = QGroupBox("Archivos DOCX a unir  (el orden importa)")
        list_layout = QVBoxLayout(list_group)
        self.flist = FileListWidget(extensions=["docx"])
        list_layout.addWidget(self.flist)
        layout.addWidget(list_group)

        self.run_btn = QPushButton("⚙   Unir DOCX → un solo Markdown")
        self.run_btn.setProperty("class", "action")
        self.run_btn.setFixedHeight(42)
        self.run_btn.clicked.connect(self._run)
        layout.addWidget(self.run_btn)

    def _run(self):
        files = self.flist.files
        if not files:
            self.log.log("⚠  Sin archivos.")
            return
        # Diálogo en el thread principal (requerido por macOS)
        dest, _ = QFileDialog.getSaveFileName(self, "Guardar como", "",
                                               "Markdown (*.md)")
        if not dest:
            return
        self._run_in_thread(self._merge, files, dest)

    @staticmethod
    def _merge(emit, files, dest):
        if not check_pandoc():
            emit("✗  pandoc no encontrado. Instalalo con: brew install pandoc")
            return
        emit(f"\n▶  Convirtiendo y uniendo {len(files)} archivo(s)…")
        try:
            import tempfile, shutil
            tmp_dir = tempfile.mkdtemp()
            md_files = []
            for path in files:
                emit(f"  →  {os.path.basename(path)}")
                success, md_path, msg = docx_to_md_file(path, tmp_dir)
                emit(f"     {msg}")
                if success:
                    md_files.append(md_path)
            if md_files:
                input_dir = os.path.dirname(files[0])
                build_consolidated_md(md_files, dest, input_dir)
                emit(f"✔  Guardado en: {dest}")
            else:
                emit("✗  No se pudo convertir ningún archivo.")
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception as e:
            emit(f"✗  Error: {e}")


class MergeMdTab(BaseTab):
    def __init__(self, log, parent=None):
        super().__init__(log, parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        opt_group = QGroupBox("Opciones")
        opt_layout = QHBoxLayout(opt_group)
        opt_layout.addWidget(QLabel("Separador:"))
        self.sep_edit = QLineEdit("\\n\\n---\\n\\n")
        self.sep_edit.setFixedWidth(220)
        opt_layout.addWidget(self.sep_edit)
        opt_layout.addSpacing(20)
        opt_layout.addWidget(QLabel("Orden:"))
        self.order_combo = QComboBox()
        self.order_combo.addItems(["nombre", "fecha modificación", "manual (lista)"])
        opt_layout.addWidget(self.order_combo)
        opt_layout.addStretch()
        layout.addWidget(opt_group)

        list_group = QGroupBox("Archivos Markdown a unir")
        list_layout = QVBoxLayout(list_group)
        self.flist = FileListWidget(extensions=["md"])
        list_layout.addWidget(self.flist)
        layout.addWidget(list_group)

        self.run_btn = QPushButton("⚙   Unir MD → un solo archivo")
        self.run_btn.setProperty("class", "action")
        self.run_btn.setFixedHeight(42)
        self.run_btn.clicked.connect(self._run)
        layout.addWidget(self.run_btn)

    def _run(self):
        files = self.flist.files
        if not files:
            self.log.log("⚠  Sin archivos.")
            return
        order = self.order_combo.currentText()
        if order == "nombre":
            files = sorted(files, key=lambda p: os.path.basename(p).lower())
        elif order == "fecha modificación":
            files = sorted(files, key=os.path.getmtime)

        dest, _ = QFileDialog.getSaveFileName(self, "Guardar como", "",
                                               "Markdown (*.md)")
        if not dest:
            return
        self._run_in_thread(self._merge, files, dest)

    @staticmethod
    def _merge(emit, files, dest):
        emit(f"\n▶  Uniendo {len(files)} archivo(s) MD…")
        try:
            input_dir = os.path.dirname(files[0]) if files else ""
            build_consolidated_md(files, dest, input_dir)
            emit(f"✔  Guardado en: {dest}")
        except Exception as e:
            emit(f"✗  Error: {e}")


class FolderBatchTab(BaseTab):
    def __init__(self, log, parent=None):
        super().__init__(log, parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # Carpeta origen
        src_group = QGroupBox("Carpeta origen")
        src_layout = QHBoxLayout(src_group)
        self.src_edit = QLineEdit()
        self.src_edit.setPlaceholderText("Seleccioná una carpeta…")
        self.src_edit.setReadOnly(True)
        src_layout.addWidget(self.src_edit)
        src_btn = QPushButton("📁 Elegir")
        src_btn.setFixedWidth(90)
        src_btn.clicked.connect(self._pick_src)
        src_layout.addWidget(src_btn)
        layout.addWidget(src_group)

        # Carpeta destino
        dst_group = QGroupBox("Carpeta destino")
        dst_layout = QHBoxLayout(dst_group)
        self.dst_same = QCheckBox("Misma que origen")
        self.dst_same.setChecked(True)
        self.dst_same.toggled.connect(self._toggle_dst)
        dst_layout.addWidget(self.dst_same)
        self.dst_edit = QLineEdit()
        self.dst_edit.setPlaceholderText("Carpeta destino…")
        self.dst_edit.setReadOnly(True)
        self.dst_edit.setEnabled(False)
        dst_layout.addWidget(self.dst_edit)
        self.dst_btn = QPushButton("📁 Elegir")
        self.dst_btn.setFixedWidth(90)
        self.dst_btn.setEnabled(False)
        self.dst_btn.clicked.connect(self._pick_dst)
        dst_layout.addWidget(self.dst_btn)
        layout.addWidget(dst_group)

        # Operación
        op_group = QGroupBox("Operación")
        op_layout = QHBoxLayout(op_group)
        op_layout.addWidget(QLabel("Operación:"))
        self.op_combo = QComboBox()
        self.op_combo.addItems([
            "DOCX → MD  (un .md por cada .docx)",
            "DOCX → MD  (unir todo en un .md)",
            "MD → unir todo en un .md",
        ])
        op_layout.addWidget(self.op_combo)
        op_layout.addSpacing(20)
        self.rec_check = QCheckBox("Recursivo (incluir subcarpetas)")
        op_layout.addWidget(self.rec_check)
        op_layout.addStretch()
        layout.addWidget(op_group)

        layout.addStretch()

        self.run_btn = QPushButton("⚙   Procesar carpeta")
        self.run_btn.setProperty("class", "action")
        self.run_btn.setFixedHeight(42)
        self.run_btn.clicked.connect(self._run)
        layout.addWidget(self.run_btn)

    def _pick_src(self):
        folder = QFileDialog.getExistingDirectory(self, "Carpeta origen")
        if folder:
            self.src_edit.setText(folder)
            if self.dst_same.isChecked():
                self.dst_edit.setText(folder)

    def _pick_dst(self):
        folder = QFileDialog.getExistingDirectory(self, "Carpeta destino")
        if folder:
            self.dst_edit.setText(folder)

    def _toggle_dst(self, checked):
        self.dst_edit.setEnabled(not checked)
        self.dst_btn.setEnabled(not checked)
        if checked and self.src_edit.text():
            self.dst_edit.setText(self.src_edit.text())

    def _run(self):
        src = self.src_edit.text()
        if not src:
            self.log.log("⚠  Elegí una carpeta origen.")
            return
        dst = src if self.dst_same.isChecked() else self.dst_edit.text()
        op = self.op_combo.currentIndex()
        rec = self.rec_check.isChecked()

        # Para ops que necesitan guardar archivo, pedir destino aquí (main thread)
        dest = None
        if op in (1, 2):
            folder_name = Path(src).name
            dest, _ = QFileDialog.getSaveFileName(
                self, "Guardar archivo consolidado",
                os.path.join(src, f"{folder_name}Consolidado.md"),
                "Markdown (*.md)")
            if not dest:
                return

        self._run_in_thread(self._process, src, dst, op, rec, dest)

    def _process(self, emit, src, dst, op, recursive, dest):
        if not check_pandoc() and op in (0, 1):
            emit("✗  pandoc no encontrado. Instalalo con: brew install pandoc")
            return
        pattern = "**/" if recursive else ""
        emit(f"\n▶  Carpeta: {src}")

        if op == 0:  # DOCX → MD individual
            files = glob.glob(os.path.join(src, pattern + "*.docx"), recursive=recursive)
            emit(f"   {len(files)} archivos .docx encontrados")
            ok = 0
            for path in files:
                success, _, msg = docx_to_md_file(path, dst if dst != src else None)
                emit(f"  {msg}")
                if success:
                    ok += 1
            emit(f"✔  {ok}/{len(files)} convertidos.")

        elif op == 1:  # DOCX → MD unir
            files = sorted(glob.glob(os.path.join(src, pattern + "*.docx"), recursive=recursive))
            emit(f"   {len(files)} archivos .docx encontrados")
            import tempfile, shutil
            tmp_dir = tempfile.mkdtemp()
            md_files = []
            for path in files:
                emit(f"  →  {os.path.basename(path)}")
                success, md_path, msg = docx_to_md_file(path, tmp_dir)
                emit(f"     {msg}")
                if success:
                    md_files.append(md_path)
            if md_files:
                build_consolidated_md(md_files, dest, src)
                emit(f"✔  Guardado en: {dest}")
            else:
                emit("✗  No se pudo convertir ningún archivo.")
            shutil.rmtree(tmp_dir, ignore_errors=True)

        elif op == 2:  # MD → unir
            files = sorted(glob.glob(os.path.join(src, pattern + "*.md"), recursive=recursive))
            emit(f"   {len(files)} archivos .md encontrados")
            if files:
                build_consolidated_md(files, dest, src)
                emit(f"✔  Guardado en: {dest}")
            else:
                emit("⚠  No se encontraron archivos .md.")


# ════════════════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MD Toolkit")
        self.resize(900, 680)
        self.setMinimumSize(720, 520)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ── Header ──────────────────────────────────────────────────────────
        header = QFrame()
        header.setFixedHeight(60)
        header.setStyleSheet("background-color: #11111b; border-bottom: 1px solid #313244;")
        h_layout = QHBoxLayout(header)
        h_layout.setContentsMargins(20, 0, 20, 0)

        title = QLabel("📄  MD Toolkit")
        title.setStyleSheet("color: #89b4fa; font-size: 20px; font-weight: 700;")
        h_layout.addWidget(title)

        subtitle = QLabel("DOCX → Markdown  •  Conversión  •  Unión  •  Carpetas")
        subtitle.setStyleSheet("color: #6c7086; font-size: 12px; margin-left: 12px;")
        h_layout.addWidget(subtitle)
        h_layout.addStretch()

        # Checkbox debug
        self.debug_check = QCheckBox("🔍 Debug")
        self.debug_check.setChecked(False)
        self.debug_check.setStyleSheet("""
            QCheckBox { color: #6c7086; font-size: 11px; }
            QCheckBox:checked { color: #f9e2af; }
            QCheckBox::indicator { width: 14px; height: 14px; }
        """)
        self.debug_check.toggled.connect(self._toggle_debug)
        h_layout.addWidget(self.debug_check)

        main_layout.addWidget(header)

        # ── Cuerpo ──────────────────────────────────────────────────────────
        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(12, 12, 12, 8)
        body_layout.setSpacing(8)

        # Log (compartido entre tabs)
        self.log_panel = LogPanel()

        # Tabs
        self.tabs = QTabWidget()
        tab_defs = [
            ("  DOCX → MD  ", DocxToMdTab),
            ("  Unir DOCX→MD  ", MergeDocxTab),
            ("  Unir MD  ", MergeMdTab),
            ("  Carpeta  ", FolderBatchTab),
        ]
        for name, cls in tab_defs:
            tab = cls(log=self.log_panel)
            self.tabs.addTab(tab, name)

        # Splitter tabs / log
        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.addWidget(self.tabs)
        splitter.addWidget(self.log_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        splitter.setHandleWidth(4)

        body_layout.addWidget(splitter)
        main_layout.addWidget(body)

        # Status bar
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("MD Toolkit listo  •  Arrastrá archivos a las listas o usá los botones")

        # Conectar debug al log panel
        DebugConfig.set_log(self.log_panel.log)

        self.log_panel.log("MD Toolkit iniciado. Seleccioná archivos o una carpeta para comenzar.")

    def _toggle_debug(self, enabled: bool):
        DebugConfig.enabled = enabled
        state = "activado ✓" if enabled else "desactivado"
        self.log_panel.log(f"🔍 Modo debug {state}")
        if enabled:
            self.log_panel.log("   Cada acción mostrará detalles internos en este log.")


# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyleSheet(STYLESHEET)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
