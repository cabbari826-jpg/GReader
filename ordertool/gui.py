from __future__ import annotations
from pathlib import Path
from decimal import Decimal

from PySide6.QtCore import Qt, QThread, Signal, QStandardPaths, QUrl, QSettings, QProcess
from PySide6.QtGui import QPixmap, QDesktopServices
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QGroupBox,
    QPushButton, QLabel, QLineEdit, QFileDialog, QListWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QProgressBar, QDoubleSpinBox, QCheckBox,
    QTabWidget, QMessageBox, QDialog, QDialogButtonBox, QComboBox,
    QAbstractItemView, QFrame,
)

from .pipeline import process_batch, write_audit_log
from .history import HistoryStore
from .exports import write_order_workbook, build_photo_library, build_validation_report
from .utils import unique_path, safe_name
from .updater import check_for_update, download_update, UpdateError
from .version import __version__


class PdfDropList(QListWidget):
    filesDropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setMinimumHeight(100)
        self.setToolTip("Drop one or more PDF order documents here")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        self.dragEnterEvent(event)

    def dropEvent(self, event):
        paths = [
            url.toLocalFile() for url in event.mimeData().urls()
            if url.isLocalFile() and url.toLocalFile().lower().endswith(".pdf")
        ]
        if paths:
            self.filesDropped.emit(paths)
        event.acceptProposedAction()


class ProcessingWorker(QThread):
    progress = Signal(int, int, str)
    completed = Signal(object)
    failed = Signal(str)

    def __init__(self, pdfs, price_lists, output_dir, photo_dir, tolerance, kit_check,
                 discount_percent, reprocess, history_db):
        super().__init__()
        self.args = (pdfs, price_lists, output_dir, photo_dir, Decimal(str(tolerance)),
                     kit_check, Decimal(str(discount_percent)), reprocess, history_db)

    def run(self):
        try:
            result = process_batch(
                self.args[0], self.args[1], self.args[2], photo_dir=self.args[3],
                tolerance=self.args[4], check_kit_multiple=self.args[5],
                retail_discount_percent=self.args[6],
                allow_reprocess=self.args[7], history_db=self.args[8],
                progress=lambda i, n, message: self.progress.emit(i, n, message),
            )
            self.completed.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


class UpdateCheckWorker(QThread):
    completed = Signal(object)
    failed = Signal(str)

    def run(self):
        try:
            self.completed.emit(check_for_update(__version__))
        except Exception as exc:
            self.failed.emit(str(exc))


class UpdateDownloadWorker(QThread):
    progress = Signal(int, int)
    completed = Signal(str)
    failed = Signal(str)

    def __init__(self, release):
        super().__init__()
        self.release = release

    def run(self):
        try:
            path = download_update(
                self.release,
                progress=lambda received, total: self.progress.emit(received, total),
            )
            self.completed.emit(str(path))
        except Exception as exc:
            self.failed.emit(str(exc))


class PreviewDialog(QDialog):
    def __init__(self, order, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Order {order.order_number} - Extraction preview")
        self.resize(1050, 600)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            f"Order {order.order_number} | {order.order_date or 'Date unavailable'} | "
            f"{len(order.products)} product lines | {order.validation_status}"
        ))
        table = QTableWidget(len(order.products), 9)
        table.setHorizontalHeaderLabels([
            "Article Code", "Style Code", "Description", "Qty", "PDF Unit Price",
            "Retail Price", "Expected Bulk Price", "Photo Status", "Validation",
        ])
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        for row, product in enumerate(order.products):
            values = [
                product.article_code, product.style_code, product.description,
                str(product.quantity or ""), str(product.unit_price or ""),
                str(product.retail_price or ""), str(product.expected_unit_price or ""),
                product.photo_status, product.validation_status,
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column == 8:
                    item.setBackground({
                        "VERIFIED": Qt.green, "DISCREPANCY": Qt.red, "REVIEW": Qt.yellow
                    }.get(product.validation_status, Qt.lightGray))
                table.setItem(row, column, item)
        layout.addWidget(table)
        notes = order.order_issues + order.extraction_warnings
        if notes:
            label = QLabel("Audit notes: " + "; ".join(notes))
            label.setWordWrap(True)
            layout.addWidget(label)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


class PhotoReviewDialog(QDialog):
    """Resolve uncertain extracted images against a manually selected product SKU."""

    def __init__(self, order, parent=None):
        super().__init__(parent)
        self.order = order
        self.applied_count = 0
        self.setWindowTitle(f"Product photo review - Order {order.order_number}")
        self.resize(900, 560)
        layout = QVBoxLayout(self)
        intro = QLabel(
            "Confirm the product SKU for each uncertain image. Leave a row unassigned "
            "when the photograph cannot be identified confidently."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)
        entries = []
        for product in order.products:
            paths = product.photo_paths or ([product.photo_path] if product.photo_path else [])
            names = product.photo_filenames or ([product.photo_filename] if product.photo_filename else [])
            if product.photo_status == "REVIEW":
                entries.extend((product, path, names[i] if i < len(names) else Path(path).name)
                               for i, path in enumerate(paths))
        self.table = QTableWidget(len(entries), 4)
        self.table.setHorizontalHeaderLabels(["Photo Preview", "Source Row SKU", "Image Filename", "Confirmed SKU"])
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.setColumnWidth(0, 130)
        self.table.setColumnWidth(1, 150)
        self.table.setColumnWidth(3, 180)
        self.table.verticalHeader().setDefaultSectionSize(95)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        sku_options = [product.sku for product in order.products if product.sku]
        for row, (source, image_path, filename) in enumerate(entries):
            preview = QLabel()
            preview.setAlignment(Qt.AlignCenter)
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                preview.setPixmap(pixmap.scaled(110, 75, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                preview.setText("Preview unavailable")
            self.table.setCellWidget(row, 0, preview)
            self.table.setItem(row, 1, QTableWidgetItem(source.sku or "Unknown"))
            self.table.setItem(row, 2, QTableWidgetItem(filename))
            combo = QComboBox()
            combo.addItem("Leave unassigned", "")
            for sku in sku_options:
                combo.addItem(sku, sku)
            self.table.setCellWidget(row, 3, combo)
            entries[row] = (source, image_path, combo)
        self._entries = entries
        layout.addWidget(self.table)
        if not entries:
            layout.addWidget(QLabel("No uncertain extracted images are available for reassignment."))
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._apply)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _apply(self):
        sku_map = {product.sku: product for product in self.order.products if product.sku}
        for source, image_path, combo in self._entries:
            target_sku = combo.currentData()
            if not target_sku or target_sku not in sku_map:
                continue
            target = sku_map[target_sku]
            image = Path(image_path)
            desired = image.with_name(f"{safe_name(target_sku)}{image.suffix.lower()}")
            destination = image
            if image.name != desired.name:
                candidate = unique_path(desired)
                try:
                    image.rename(candidate)
                    destination = candidate
                except OSError:
                    destination = image
            paths = target.photo_paths or ([target.photo_path] if target.photo_path else [])
            names = target.photo_filenames or ([target.photo_filename] if target.photo_filename else [])
            if str(destination) not in paths:
                paths.append(str(destination))
                names.append(destination.name)
            target.photo_paths, target.photo_filenames = paths, names
            target.photo_path, target.photo_filename = paths[0], names[0]
            target.photo_status = "MATCHED" if len(paths) == 1 else "MULTIPLE"
            target.photo_note = "Manually verified by user"
            if source is not target:
                source_paths = source.photo_paths or ([source.photo_path] if source.photo_path else [])
                remaining = [path for path in source_paths if path != image_path]
                source.photo_paths = remaining
                source.photo_filenames = [Path(path).name for path in remaining]
                source.photo_path = remaining[0] if remaining else None
                source.photo_filename = Path(source.photo_path).name if source.photo_path else ""
                source.photo_status = ("MULTIPLE" if len(remaining) > 1 else "MATCHED") if remaining else "MISSING"
                source.photo_note = "Photo reassigned by user"
            self.applied_count += 1
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Order Converter & Validator")
        self.resize(1280, 850)
        self.pdf_paths = []
        self.batch_orders = []
        self.batch_result = None
        app_data = QStandardPaths.writableLocation(QStandardPaths.AppLocalDataLocation)
        self.app_data_dir = Path(app_data or (Path.home() / ".pdf_order_validator"))
        self.app_data_dir.mkdir(parents=True, exist_ok=True)
        self.history = HistoryStore(self.app_data_dir / "history.sqlite")
        self.settings = QSettings("GReader", "PDFOrderValidator")
        self.worker = None
        self.update_worker = None
        self.download_worker = None
        self.pending_release = None
        self._build_ui()
        self._refresh_history()
        self._refresh_dashboard()


    def check_updates(self):
        if self.update_worker and self.update_worker.isRunning():
            return
        self.update_button.setEnabled(False)
        self.update_button.setText("Checking…")
        self.update_worker = UpdateCheckWorker(self)
        self.update_worker.completed.connect(self.on_update_checked)
        self.update_worker.failed.connect(self.on_update_failed)
        self.update_worker.start()

    def on_update_checked(self, release):
        self.update_button.setEnabled(True)
        self.update_button.setText("Check for updates")
        if release is None:
            QMessageBox.information(self, "No updates available",
                                    f"You are using the latest version ({__version__}).")
            return
        self.pending_release = release
        message = (
            f"Version {release.version} is available.\n"
            f"Installed version: {__version__}\n\n"
            "The installer will be downloaded from this application's GitHub Releases page "
            "and its SHA-256 checksum verified before it is opened.\n\n"
            "Download this update now?"
        )
        if release.notes:
            message += "\n\nRelease notes:\n" + release.notes[:2500]
        answer = QMessageBox.question(self, "Application update available", message,
                                      QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if answer == QMessageBox.Yes:
            self.download_latest_update()

    def on_update_failed(self, error):
        self.update_button.setEnabled(True)
        self.update_button.setText("Check for updates")
        QMessageBox.warning(self, "Update check failed", str(error))

    def download_latest_update(self):
        if not self.pending_release:
            return
        self.update_button.setEnabled(False)
        self.update_button.setText("Downloading update…")
        self.progress.setValue(0)
        self.progress_text.setText(f"Downloading version {self.pending_release.version}")
        self.download_worker = UpdateDownloadWorker(self.pending_release)
        self.download_worker.progress.connect(self.on_update_download_progress)
        self.download_worker.completed.connect(self.on_update_downloaded)
        self.download_worker.failed.connect(self.on_update_download_failed)
        self.download_worker.start()

    def on_update_download_progress(self, received, total):
        if total > 0:
            self.progress.setValue(int(received * 100 / total))
            self.progress_text.setText(f"Downloading update: {received // 1024} / {total // 1024} KB")

    def on_update_downloaded(self, installer_path):
        self.update_button.setEnabled(True)
        self.update_button.setText("Check for updates")
        self.progress.setValue(100)
        self.progress_text.setText("Update downloaded and checksum verified")
        answer = QMessageBox.question(
            self, "Update ready",
            "The installer was downloaded and its SHA-256 checksum verified. "
            "The application must close before installation. Start the installer now?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes,
        )
        if answer == QMessageBox.Yes:
            started, _pid = QProcess.startDetached(installer_path, [])
            if started:
                self.close()
            else:
                QMessageBox.warning(self, "Could not start installer",
                                    f"Run the installer manually: {installer_path}")

    def on_update_download_failed(self, error):
        self.update_button.setEnabled(True)
        self.update_button.setText("Check for updates")
        self.progress_text.setText("Update download failed")
        QMessageBox.critical(self, "Update failed", str(error))

    def _build_ui(self):
        root = QWidget()
        layout = QVBoxLayout(root)
        header = QHBoxLayout()
        title = QLabel("PDF Order Converter & Validator")
        title.setStyleSheet("font-size:20pt;font-weight:700;color:#24364B")
        header.addWidget(title)
        header.addStretch()
        self.update_button = QPushButton("Check for updates")
        self.update_button.clicked.connect(self.check_updates)
        header.addWidget(self.update_button)
        layout.addLayout(header)
        subtitle = QLabel(
            "Local PDF extraction • Article Code matching • Product photos in Excel cells • "
            "Bulk price validation"
        )
        subtitle.setStyleSheet("color:#637487")
        layout.addWidget(subtitle)
        self.kpi_labels = {}
        cards = QHBoxLayout()
        for key, label in (
            ("processed", "Files processed"), ("converted", "Converted"),
            ("review", "Needs review"), ("discrepancies", "Discrepancies"),
        ):
            frame = QFrame()
            frame.setFrameShape(QFrame.StyledPanel)
            card_layout = QVBoxLayout(frame)
            number = QLabel("0")
            number.setStyleSheet("font-size:20pt;font-weight:700")
            card_layout.addWidget(number)
            card_layout.addWidget(QLabel(label))
            self.kpi_labels[key] = number
            cards.addWidget(frame)
        layout.addLayout(cards)
        self.tabs = QTabWidget()
        workspace = QWidget()
        self._build_workspace(workspace)
        history = QWidget()
        self._build_history(history)
        self.tabs.addTab(workspace, "Workspace")
        self.tabs.addTab(history, "Processing History")
        layout.addWidget(self.tabs)
        self.setCentralWidget(root)
        self.setStyleSheet("""
            QMainWindow,QWidget{background:#F4F6F9;color:#24364B;font-family:'Segoe UI';font-size:10pt}
            QGroupBox{background:white;border:1px solid #DCE3EB;border-radius:8px;margin-top:10px;padding:10px;font-weight:600}
            QPushButton{background:#2F5597;color:white;padding:8px 13px;border:0;border-radius:6px;font-weight:600}
            QPushButton:disabled{background:#AAB6C3}
            QLineEdit,QListWidget,QTableWidget,QDoubleSpinBox{background:white;border:1px solid #CCD6E0;border-radius:5px;padding:5px}
            QHeaderView::section{background:#E9EFF5;color:#2F455B;font-weight:600;padding:7px;border:0}
        """)

    @staticmethod
    def _browse_button(text, callback):
        button = QPushButton(text)
        button.clicked.connect(callback)
        return button

    def _build_workspace(self, parent):
        layout = QVBoxLayout(parent)
        files_group = QGroupBox("Order documents")
        grid = QGridLayout(files_group)
        self.pdf_list = PdfDropList()
        self.pdf_list.filesDropped.connect(self.add_pdf_files)
        grid.addWidget(self.pdf_list, 0, 0, 4, 2)
        grid.addWidget(self._browse_button("Import PDF Orders", self.choose_pdfs), 0, 2)
        grid.addWidget(self._browse_button("Remove selected", self.remove_selected), 1, 2)
        grid.addWidget(self._browse_button("Clear list", self.clear_files), 2, 2)
        layout.addWidget(files_group)

        settings = QGroupBox("Reference files and output")
        grid = QGridLayout(settings)
        self.price_edit = QLineEdit()
        self.price_edit.setPlaceholderText("Select one or more official price lists")
        grid.addWidget(QLabel("Official price list(s)"), 0, 0)
        grid.addWidget(self.price_edit, 0, 1)
        grid.addWidget(self._browse_button("Browse…", self.choose_price_lists), 0, 2)
        self.output_edit = QLineEdit(str(Path.home() / "Documents" / "PDFOrderValidator"))
        grid.addWidget(QLabel("Output folder"), 1, 0)
        grid.addWidget(self.output_edit, 1, 1)
        grid.addWidget(self._browse_button("Browse…", self.choose_output), 1, 2)
        self.photo_edit = QLineEdit()
        self.photo_edit.setPlaceholderText("Optional; defaults to Output/Product_Photos")
        grid.addWidget(QLabel("Product photo folder"), 2, 0)
        grid.addWidget(self.photo_edit, 2, 1)
        grid.addWidget(self._browse_button("Browse…", self.choose_photo_folder), 2, 2)
        self.tolerance = QDoubleSpinBox()
        self.tolerance.setDecimals(2)
        self.tolerance.setRange(0, 100)
        self.tolerance.setSingleStep(.01)
        self.tolerance.setValue(.01)
        self.kit_check = QCheckBox("Check Kit Q.ty multiples")
        self.kit_check.setChecked(True)
        self.duplicate_check = QCheckBox("Reprocess duplicate files")
        grid.addWidget(QLabel("Rounding tolerance"), 3, 0)
        row = QHBoxLayout()
        row.addWidget(self.tolerance)
        row.addWidget(self.kit_check)
        row.addWidget(self.duplicate_check)
        row.addStretch()
        grid.addLayout(row, 3, 1, 1, 2)
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setDecimals(1)
        self.discount_spin.setRange(0, 100)
        self.discount_spin.setSuffix("%")
        self.discount_spin.setValue(float(self.settings.value("retail_discount_percent", 60.0)))
        self.discount_spin.valueChanged.connect(
            lambda value: self.settings.setValue("retail_discount_percent", value)
        )
        grid.addWidget(QLabel("Bulk discount from Retail Price"), 4, 0)
        grid.addWidget(self.discount_spin, 4, 1)
        discount_hint = QLabel("60% discount means expected purchase price = 40% of Retail Price.")
        discount_hint.setStyleSheet("color:#637487")
        grid.addWidget(discount_hint, 4, 2)
        layout.addWidget(settings)

        buttons = QHBoxLayout()
        self.process_button = QPushButton("Start Conversion & Validation")
        self.process_button.clicked.connect(self.start_processing)
        self.preview_button = QPushButton("Preview order")
        self.preview_button.clicked.connect(self.preview_selected)
        self.photo_button = QPushButton("Review photo matches")
        self.photo_button.clicked.connect(self.review_photos)
        self.excel_button = QPushButton("Open order Excel")
        self.excel_button.clicked.connect(self.open_selected_excel)
        self.report_button = QPushButton("Open validation report")
        self.report_button.clicked.connect(self.open_report)
        self.folder_button = QPushButton("Open output folder")
        self.folder_button.clicked.connect(lambda: self.open_path(self.output_edit.text()))
        for button in (self.process_button, self.preview_button, self.photo_button,
                       self.excel_button, self.report_button, self.folder_button):
            buttons.addWidget(button)
        layout.addLayout(buttons)
        progress = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress_text = QLabel("Ready")
        progress.addWidget(self.progress, 3)
        progress.addWidget(self.progress_text, 2)
        layout.addLayout(progress)
        self.results = QTableWidget(0, 7)
        self.results.setHorizontalHeaderLabels([
            "Source PDF", "Order", "Lines", "Conversion", "Validation", "Workbook", "Photos to review",
        ])
        self.results.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.results.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.results.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.results.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.results)

    def _build_history(self, parent):
        layout = QVBoxLayout(parent)
        self.history_table = QTableWidget(0, 8)
        self.history_table.setHorizontalHeaderLabels([
            "Processed", "Filename", "Order", "Conversion", "Validation",
            "Issues", "Workbook", "Validation Report",
        ])
        self.history_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.history_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.history_table)
        buttons = QHBoxLayout()
        open_excel = QPushButton("Open selected workbook")
        open_excel.clicked.connect(self.open_history_excel)
        open_report = QPushButton("Open selected report")
        open_report.clicked.connect(self.open_history_report)
        buttons.addWidget(open_excel)
        buttons.addWidget(open_report)
        layout.addLayout(buttons)

    def add_pdf_files(self, paths):
        for path in paths:
            resolved = str(Path(path).resolve())
            if Path(resolved).is_file() and resolved.lower().endswith(".pdf") and resolved not in self.pdf_paths:
                self.pdf_paths.append(resolved)
                self.pdf_list.addItem(resolved)

    def choose_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select PDF orders", "", "PDF files (*.pdf)")
        self.add_pdf_files(paths)

    def remove_selected(self):
        for item in self.pdf_list.selectedItems():
            row = self.pdf_list.row(item)
            path = self.pdf_list.takeItem(row).text()
            if path in self.pdf_paths:
                self.pdf_paths.remove(path)

    def clear_files(self):
        self.pdf_paths.clear()
        self.pdf_list.clear()

    def choose_price_lists(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Select official price list(s)", "", "Price lists (*.xlsx *.xlsm *.csv *.pdf)"
        )
        if paths:
            self.price_edit.setText(";".join(paths))

    def choose_output(self):
        path = QFileDialog.getExistingDirectory(self, "Select output folder")
        if path:
            self.output_edit.setText(path)

    def choose_photo_folder(self):
        path = QFileDialog.getExistingDirectory(self, "Select product photo folder")
        if path:
            self.photo_edit.setText(path)

    def start_processing(self):
        if not self.pdf_paths:
            QMessageBox.information(self, "PDF files required", "Import one or more PDF order documents.")
            return
        price_lists = [part.strip() for part in self.price_edit.text().split(";") if part.strip()]
        if not price_lists:
            QMessageBox.information(self, "Price list required", "Select an official price list.")
            return
        if any(not Path(path).is_file() for path in price_lists):
            QMessageBox.warning(self, "Price list unavailable", "One or more selected price lists cannot be found.")
            return
        output = self.output_edit.text().strip()
        if not output:
            QMessageBox.warning(self, "Output folder required", "Choose an output directory.")
            return
        self._set_busy(True)
        self.results.setRowCount(0)
        self.progress.setValue(0)
        self.worker = ProcessingWorker(
            list(self.pdf_paths), price_lists, output, self.photo_edit.text().strip() or None,
            self.tolerance.value(), self.kit_check.isChecked(), self.discount_spin.value(),
            self.duplicate_check.isChecked(), str(self.history.path),
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.completed.connect(self.on_completed)
        self.worker.failed.connect(self.on_failed)
        self.worker.start()

    def _set_busy(self, busy):
        self.process_button.setEnabled(not busy)
        self.progress_text.setText("Processing…" if busy else "Ready")

    def on_progress(self, current, total, message):
        self.progress.setValue(int(current * 100 / max(1, total)))
        self.progress_text.setText(message)

    def on_completed(self, result):
        self._set_busy(False)
        self.batch_result = result
        self.batch_orders = result.orders
        self._populate_results(result)
        self._refresh_history()
        self._refresh_dashboard()
        self.progress.setValue(100)
        if result.errors:
            message = "\n".join(f"{item['filename']}: {item['error']}" for item in result.errors[:8])
            QMessageBox.warning(self, "Some files need attention", message)
        else:
            QMessageBox.information(
                self, "Batch completed",
                f"Successfully converted {len(result.orders)} order(s). "
                "Review yellow statuses and the validation report before use.",
            )

    def on_failed(self, error):
        self._set_busy(False)
        QMessageBox.critical(self, "Processing failed", error)

    def _populate_results(self, result):
        self.results.setRowCount(0)
        for order in result.orders:
            row = self.results.rowCount()
            self.results.insertRow(row)
            values = [
                Path(order.source_pdf).name, order.order_number, str(len(order.products)),
                order.conversion_status, order.validation_status, order.order_workbook,
                str(sum(p.photo_status in ("REVIEW", "MISSING") for p in order.products)),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if column == 4:
                    item.setBackground({
                        "VERIFIED": Qt.green, "ERRORS": Qt.red, "REVIEW": Qt.yellow,
                    }.get(order.validation_status, Qt.lightGray))
                self.results.setItem(row, column, item)
        for error in result.errors:
            row = self.results.rowCount()
            self.results.insertRow(row)
            self.results.setItem(row, 0, QTableWidgetItem(error["filename"]))
            self.results.setItem(row, 3, QTableWidgetItem("FAILED"))
            self.results.setItem(row, 4, QTableWidgetItem(error["error"]))

    def _selected_order(self):
        row = self.results.currentRow()
        return self.batch_orders[row] if 0 <= row < len(self.batch_orders) else None

    def preview_selected(self):
        order = self._selected_order()
        if order:
            PreviewDialog(order, self).exec()
        else:
            QMessageBox.information(self, "Select an order", "Select a processed order first.")

    def review_photos(self):
        order = self._selected_order()
        if not order:
            QMessageBox.information(self, "Select an order", "Select a processed order first.")
            return
        dialog = PhotoReviewDialog(order, self)
        if dialog.exec() != QDialog.Accepted or not dialog.applied_count:
            return
        output_root = Path(self.output_edit.text().strip())
        report_dir = output_root / "Validation_Reports"
        log_dir = output_root / "Processing_Logs"
        try:
            # Only photo-only review can clear. Extraction and numerical issues remain.
            for item in self.batch_orders:
                for product in item.products:
                    hard = any(issue.startswith(("PRICE_", "LINE_TOTAL_ERROR", "KIT_QTY_MISMATCH"))
                               for issue in product.issues)
                    if hard:
                        product.validation_status = "DISCREPANCY"
                    elif product.issues or product.photo_status in ("REVIEW", "MISSING"):
                        product.validation_status = "REVIEW"
                    else:
                        product.validation_status = "VERIFIED"
                if any(p.validation_status == "DISCREPANCY" for p in item.products) or item.order_issues:
                    item.validation_status = "ERRORS"
                elif any(p.validation_status == "REVIEW" for p in item.products) or item.extraction_warnings:
                    item.validation_status = "REVIEW"
                else:
                    item.validation_status = "VERIFIED"
                if item is order:
                    item.order_workbook = write_order_workbook(item, Path(item.order_workbook).parent)
                write_audit_log(item, log_dir)
            records = self._collect_photo_records(self.batch_orders)
            self.batch_result.photo_library = build_photo_library(records, output_root)
            self.batch_result.validation_report = build_validation_report(
                self.batch_orders, report_dir, Decimal(str(self.tolerance.value()))
            )
            for item in self.batch_orders:
                self.history.record(item, self.batch_result.validation_report, self.batch_result.photo_library)
            self._populate_results(self.batch_result)
            self._refresh_history()
            self._refresh_dashboard()
            QMessageBox.information(
                self, "Photo mappings saved",
                f"Applied {dialog.applied_count} association(s). The order workbook, "
                "audit log, photo library, and validation report were regenerated.",
            )
        except Exception as exc:
            QMessageBox.critical(self, "Could not save photo mappings", str(exc))

    @staticmethod
    def _collect_photo_records(orders):
        records = []
        for order in orders:
            for product in order.products:
                paths = product.photo_paths or ([product.photo_path] if product.photo_path else [])
                names = product.photo_filenames or ([product.photo_filename] if product.photo_filename else [])
                if not paths:
                    records.append({
                        "sku": product.article_code, "article_code": product.article_code,
                        "style_code": product.style_code, "description": product.description, "filename": "",
                        "source_pdf": Path(order.source_pdf).name, "extraction_status": "MISSING",
                        "verification_status": "MISSING", "page": product.page, "image_path": "",
                    })
                else:
                    for index, image_path in enumerate(paths):
                        records.append({
                            "sku": product.article_code if product.photo_status in ("MATCHED", "MULTIPLE") else "",
                            "article_code": product.article_code if product.photo_status in ("MATCHED", "MULTIPLE") else "",
                            "style_code": product.style_code,
                            "description": product.description,
                            "filename": names[index] if index < len(names) else Path(image_path).name,
                            "source_pdf": Path(order.source_pdf).name, "extraction_status": "EXTRACTED",
                            "verification_status": product.photo_status, "page": product.page,
                            "image_path": image_path,
                        })
        return records

    def open_selected_excel(self):
        order = self._selected_order()
        if order:
            self.open_path(order.order_workbook)
        else:
            QMessageBox.information(self, "Select an order", "Select a processed order first.")

    def open_report(self):
        path = self.batch_result.validation_report if self.batch_result else ""
        if path:
            self.open_path(path)
        else:
            QMessageBox.information(self, "Report unavailable", "Run a batch to generate a report.")

    def open_path(self, path):
        if path and Path(path).exists():
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(path).resolve())))
        else:
            QMessageBox.warning(self, "File unavailable", f"Could not find: {path}")

    def _refresh_history(self):
        if not hasattr(self, "history_table"):
            return
        self.history_table.setRowCount(0)
        self._history_rows = self.history.list_recent()
        for record in self._history_rows:
            row = self.history_table.rowCount()
            self.history_table.insertRow(row)
            values = [
                record.get("processed_at"), record.get("filename"), record.get("order_number"),
                record.get("conversion_status"), record.get("validation_status"),
                record.get("discrepancy_count"), record.get("output_file"),
                record.get("validation_report"),
            ]
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                item.setToolTip(str(value or ""))
                self.history_table.setItem(row, column, item)

    def _selected_history(self):
        row = self.history_table.currentRow()
        return self._history_rows[row] if 0 <= row < len(self._history_rows) else None

    def open_history_excel(self):
        record = self._selected_history()
        if record:
            self.open_path(record.get("output_file"))

    def open_history_report(self):
        record = self._selected_history()
        if record:
            self.open_path(record.get("validation_report"))

    def _refresh_dashboard(self):
        records = self.history.list_recent()
        self.kpi_labels["processed"].setText(str(len(records)))
        self.kpi_labels["converted"].setText(str(sum(r.get("conversion_status") == "SUCCESS" for r in records)))
        self.kpi_labels["review"].setText(str(sum(r.get("validation_status") == "REVIEW" for r in records)))
        self.kpi_labels["discrepancies"].setText(
            str(sum(int(r.get("discrepancy_count") or 0) for r in records))
        )
