from __future__ import annotations

import logging
import threading
from pathlib import Path

from PySide6.QtCore import QDate, QEvent, Qt, QTimer, Signal
from PySide6.QtWidgets import (QComboBox, QCompleter, QDateEdit, QFileDialog, QHBoxLayout, QLabel,
                              QMessageBox, QPushButton, QVBoxLayout, QWidget)

from .cst_workflow import export_cst_batch, fetch_facilities
from .editor_window import PdfEditorWindow
from .cst_submission import discard_batch, inspect_batch, submit_batch
from .outlook_intake import TECHNICIANS

FACILITIES_CACHE = Path.home() / "AppData" / "Local" / "PDFSplitter" / "Facilities.txt"


class FacilityComboBox(QComboBox):
    """Tab accepts the highlighted suggestion, or the first filtered match."""
    def __init__(self, facilities: list[str]) -> None:
        super().__init__()
        self.setEditable(True)
        self.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.setMinimumContentsLength(20)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.addItems(facilities)
        self.setCurrentIndex(-1)
        self.lineEdit().setPlaceholderText("Search facility; Tab accepts first match…")
        completer = QCompleter(facilities, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        self.setCompleter(completer)
        self.installEventFilter(self)
        self.lineEdit().installEventFilter(self)
        completer.popup().installEventFilter(self)
        self.currentTextChanged.connect(self.setToolTip)

    def accept_suggestion(self) -> None:
        text = self.currentText().strip()
        if not text:
            return
        completer = self.completer()
        popup = completer.popup()
        selected = popup.currentIndex()
        if popup.isVisible() and selected.isValid():
            self.setEditText(str(selected.data()))
        elif self.findText(text, Qt.MatchExactly) < 0:
            completer.setCompletionPrefix(text)
            if completer.setCurrentRow(0):
                self.setEditText(completer.currentCompletion())
        popup.hide()

    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Tab, Qt.Key_Backtab):
            self.accept_suggestion()
            self.setFocus()
            self.focusNextPrevChild(event.key() == Qt.Key_Tab and not (event.modifiers() & Qt.ShiftModifier))
            return True
        return super().eventFilter(obj, event)


class CstEditorWindow(PdfEditorWindow):
    batch_exported = Signal(str)
    submission_progress = Signal(str)
    submission_finished = Signal(bool, str)

    def __init__(self) -> None:
        self.facilities: list[str] = []
        self.section_facilities: dict[int, str] = {}
        self.section_visit_types: dict[int, str] = {}
        self.facility_inputs: dict[int, QComboBox] = {}
        self.intake_locked = False
        self.submitting = False
        # Runs on the submission thread once a batch is fully received; returns a note for the user.
        self.on_received = lambda source: ""
        super().__init__()
        self.setWindowTitle("CST — PDF Page Editor")
        self.save_btn.hide()
        self.replace_original_checkbox.hide()
        self.delete_source_checkbox.hide()
        self.hint_label.setText("Click page images to mark section starts. Page 1 always starts a section. "
                                "Set the shared document date and choose one facility per section.")
        self.output_label.setText("Facility for each section:")

        panel = QWidget()
        controls = QVBoxLayout(panel)
        controls.addWidget(QLabel("Document date (all sections):"))
        self.document_date = QDateEdit(QDate.currentDate())
        self.document_date.setDisplayFormat("MM/dd/yyyy")
        self.document_date.setCalendarPopup(True)
        date_row = QHBoxLayout()
        for text, days in (("◀", -1), ("▶", 1)):
            step = QPushButton(text)
            step.setObjectName(f"date_step_{days}")
            step.setFixedWidth(36)
            step.setFocusPolicy(Qt.NoFocus)  # Keep typing focus where it is.
            step.clicked.connect(lambda _, d=days: self.document_date.setDate(self.document_date.date().addDays(d)))
            date_row.addWidget(step)
        date_row.insertWidget(1, self.document_date, 1)
        controls.addLayout(date_row)
        controls.addWidget(QLabel("Technician (set automatically for emailed PDFs):"))
        self.technician = QComboBox()
        self.technician.addItems(TECHNICIANS)
        self.technician.setFocusPolicy(Qt.ClickFocus)
        controls.addWidget(self.technician)
        self.facilities_label = QLabel()
        self.facilities_label.setWordWrap(True)
        controls.addWidget(self.facilities_label)
        self.intake_note = QLabel("Select split starts, choose a facility for each section, then export.")
        self.intake_note.setWordWrap(True)
        controls.addWidget(self.intake_note)
        self.right_panel.layout().insertWidget(0, panel)
        self.submit_button = QPushButton("Export && submit to RMRR")
        self.submit_button.clicked.connect(lambda: self._export(send=True))
        self.submit_button.setToolTip("Send each section using the selected technician, date, facility, and type")
        self.right_panel.layout().addWidget(self.submit_button)
        self.retry_button = QPushButton("Submit / retry saved CST batch…")
        self.retry_button.clicked.connect(self._pick_submission_batch)
        self.right_panel.layout().addWidget(self.retry_button)
        self.submission_status = QLabel("")
        self.submission_status.setWordWrap(True)
        self.right_panel.layout().addWidget(self.submission_status)
        self.submission_progress.connect(self.submission_status.setText)
        self.submission_finished.connect(self._submission_finished)

    def pick_pdf(self) -> None:
        self._pick_pdf()

    def set_aside(self) -> None:
        self._clear_loaded_pdf()
        self.intake_locked = False

    def _restore_window_header(self) -> None:
        super()._restore_window_header()
        self.setWindowTitle("CST — " + self.windowTitle())

    def load_pdf(self, path: str | Path) -> bool:
        if self.submitting:
            self.bring_to_front()
            return False
        if self.metadata is not None:
            if self.intake_locked:
                self.bring_to_front()
                QMessageBox.information(self, "CST intake in progress", "Export the current queued PDF before opening another CST document.")
                return False
            answer = QMessageBox.question(self, "Another CST PDF is open",
                "Replace this document and discard its unexported date and facility choices?",
                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
            if answer != QMessageBox.Yes:
                return False
        self._load_facilities()
        result = super().load_pdf(path)
        if result:
            self.document_date.setDate(QDate.currentDate())
            self._restore_window_header()
            self.intake_note.setText("Select split starts and a facility per section.")
            QTimer.singleShot(0, self.document_date.setFocus)
        return result

    def _load_facilities(self) -> None:
        # Refreshed per document so the choices always match what the RMRR form accepts.
        try:
            self.facilities = fetch_facilities(FACILITIES_CACHE)
            self.facilities_label.setText(f"{len(self.facilities)} facilities loaded from the RMRR app list")
        except Exception as exc:
            self.facilities = []
            self.facilities_label.setText(f"Facilities unavailable: {exc}. Reopen the PDF when online.")
            logging.warning("CST facilities could not be loaded: %s", exc)

    def _reset_loaded_document_state(self) -> None:
        super()._reset_loaded_document_state()
        self.section_facilities = {}
        self.section_visit_types = {}
        self.facility_inputs = {}

    def _toggle_split_start(self, page: int) -> None:
        super()._toggle_split_start(page)
        if page in self.split_starts:  # A new section is ready to type into.
            self.facility_inputs[page].setFocus()

    def _refresh_sections_ui(self) -> None:
        if self.metadata is None:
            return
        self._clear_section_names_ui()
        self.section_layout.takeAt(0)  # Base reset inserts a stretch; keep fields at the top.
        self.facility_inputs = {}
        for index, (start, end, page) in enumerate(self._ordered_sections(), 1):
            block = QWidget()
            layout = QVBoxLayout(block)
            layout.addWidget(QLabel(f"Section {index}: pages {start}–{end}"))
            combo = FacilityComboBox(self.facilities)
            combo.setEditText(self.section_facilities.get(page, ""))
            combo.currentTextChanged.connect(lambda text, p=page: self.section_facilities.__setitem__(p, text))
            visit_type = QComboBox()
            visit_type.addItems(["Tracking", "Delivery"])
            visit_type.setCurrentText(self.section_visit_types.get(page, "Tracking"))
            visit_type.setFocusPolicy(Qt.ClickFocus)  # Tab goes facility to facility.
            visit_type.currentTextChanged.connect(lambda text, p=page: self.section_visit_types.__setitem__(p, text))
            row = QHBoxLayout()
            row.addWidget(combo, 1)
            row.addWidget(visit_type)
            layout.addLayout(row)
            self.section_layout.addWidget(block)
            self.facility_inputs[page] = combo
        self.section_layout.addStretch()
        previous = self.document_date
        for combo in self.facility_inputs.values():
            self.setTabOrder(previous, combo)
            previous = combo
        self.setTabOrder(previous, self.submit_button)
        self.setTabOrder(self.submit_button, self.retry_button)
        for page in self.page_order:
            self._update_item_visual(page)

    def _export(self, send: bool = False) -> None:
        if self.metadata is None:
            QMessageBox.information(self, "No CST PDF", "Open a PDF first.")
            return
        sections = [(start, end, self.section_facilities.get(page, ""), self.section_visit_types.get(page, "Tracking"))
                    for start, end, page in self._ordered_sections()]
        if not sections or any(facility not in self.facilities for _, _, facility, _ in sections):
            QMessageBox.warning(self, "Select facilities", "Every section needs an exact facility name from the loaded list.")
            return
        # Asked once, and again only if that folder goes away.
        chosen = self.settings.value("cst/output_dir", "", str)
        if not chosen or not Path(chosen).is_dir():
            chosen = QFileDialog.getExistingDirectory(self, "Choose CST export folder",
                                                      str(self.metadata.source_path.parent))
            if not chosen:
                return
        try:
            batch = export_cst_batch(self.metadata.source_path, self.password,
                self.document_date.date().toPython(), sections, self.facilities,
                self.page_order, self.page_rotations, Path(chosen),
                technician=self.technician.currentText())
        except Exception as exc:
            logging.exception("CST batch export failed")
            QMessageBox.critical(self, "CST export failed", str(exc))
            return
        self.settings.setValue("cst/output_dir", chosen)
        self.settings.setValue("cst/last_batch", str(batch / "batch.json"))
        self._clear_loaded_pdf()
        if send:
            self._start_submission(batch / "batch.json")
        else:
            QMessageBox.information(self, "CST batch ready",
                f"Exported {len(sections)} PDF(s) and their date/facility record to:\n{batch}\n\n"
                "Use Submit / retry saved CST batch to send it.")
        self.batch_exported.emit(str(batch))

    def _pick_submission_batch(self) -> None:
        chosen, _ = QFileDialog.getOpenFileName(self, "Select CST batch.json",
            self.settings.value("cst/last_batch", "", str), "CST batch (batch.json)")
        if chosen:
            try:
                manifest, prepared = inspect_batch(Path(chosen))
            except Exception as exc:
                QMessageBox.warning(self, "Invalid CST batch", str(exc))
                return
            answer = QMessageBox.question(self, "Submit saved CST batch?",
                f"Send {len(prepared)} section(s) to RMRR as {manifest['technician']} for "
                f"{manifest['document_date']}? Already received sections will be skipped.",
                QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
            if answer == QMessageBox.Yes:
                self._start_submission(Path(chosen))

    def _start_submission(self, manifest: Path) -> None:
        if self.submitting:
            return
        self.submitting = True
        self.settings.setValue("cst/last_batch", str(manifest))
        self.centralWidget().setEnabled(False)
        self.submission_status.setText("Preparing RMRR submission…")
        def work():
            try:
                count = submit_batch(manifest, progress=self.submission_progress.emit)
                note = self.on_received(discard_batch(manifest))
                self.submission_finished.emit(True, f"RMRR received all {count} section(s); local copies removed.{note}")
            except Exception as exc:
                logging.exception("CST submission did not finish")
                self.submission_finished.emit(False,
                    f"Submission needs attention: {exc}\nYour batch is saved at {manifest.parent}. "
                    "Use Submit / retry saved CST batch to continue.")
        self.submission_worker = threading.Thread(target=work, name="CST submission", daemon=True)
        self.submission_worker.start()

    def _submission_finished(self, success: bool, message: str) -> None:
        self.submitting = False
        self.centralWidget().setEnabled(True)
        self.submission_status.setText(message)
        if success:
            # The tray shows the confirmation. Stay open only if a document is loaded.
            if self.metadata is None:
                self.hide()
        else:
            QMessageBox.warning(self, "CST submission needs attention", message)
