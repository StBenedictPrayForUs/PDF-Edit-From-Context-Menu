from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import QEvent, QObject, QSettings, QSize, Qt, QTimer, Signal
from PySide6.QtGui import QPixmap, QTransform
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSlider,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from . import pdf_ops


class ClickableLabel(QLabel):
    clicked = Signal()

    def mouseReleaseEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(event)


class PageRowWidget(QWidget):
    split_toggled = Signal(int)
    rotate_left = Signal(int)
    rotate_right = Signal(int)
    move_up = Signal(int)
    move_down = Signal(int)

    def __init__(self, page: int, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.page = page

        outer = QVBoxLayout(self)
        outer.setContentsMargins(8, 8, 8, 8)
        outer.setSpacing(6)

        self.header = QLabel(f"Page {page}")
        self.header.setAlignment(Qt.AlignHCenter)
        self.header.setStyleSheet("color: palette(window-text);")
        outer.addWidget(self.header)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(10)

        row.addStretch(1)

        self.image_label = ClickableLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setStyleSheet(
            "background: palette(base); border: 2px solid palette(mid); border-radius: 4px;"
        )
        self.image_label.setCursor(Qt.PointingHandCursor)
        self.image_label.clicked.connect(lambda: self.split_toggled.emit(self.page))
        row.addWidget(self.image_label)

        controls = QVBoxLayout()
        controls.setContentsMargins(0, 0, 0, 0)
        controls.setSpacing(6)

        self.move_up_btn = QPushButton("↑")
        self.move_up_btn.setFixedSize(36, 30)
        self.move_up_btn.setToolTip("Move page up")
        self.move_up_btn.setFocusPolicy(Qt.NoFocus)
        self.move_up_btn.clicked.connect(self._emit_move_up)

        self.move_down_btn = QPushButton("↓")
        self.move_down_btn.setFixedSize(36, 30)
        self.move_down_btn.setToolTip("Move page down")
        self.move_down_btn.setFocusPolicy(Qt.NoFocus)
        self.move_down_btn.clicked.connect(self._emit_move_down)

        self.rotate_left_btn = QPushButton()
        self.rotate_left_btn.setText("Rotate -90")
        self.rotate_left_btn.setFixedSize(90, 30)
        self.rotate_left_btn.setToolTip("Rotate left 90")
        self.rotate_left_btn.setFocusPolicy(Qt.NoFocus)
        self.rotate_left_btn.clicked.connect(lambda: self.rotate_left.emit(self.page))

        self.rotate_right_btn = QPushButton()
        self.rotate_right_btn.setText("Rotate +90")
        self.rotate_right_btn.setFixedSize(90, 30)
        self.rotate_right_btn.setToolTip("Rotate right 90")
        self.rotate_right_btn.setFocusPolicy(Qt.NoFocus)
        self.rotate_right_btn.clicked.connect(lambda: self.rotate_right.emit(self.page))

        controls.addWidget(self.move_up_btn)
        controls.addWidget(self.move_down_btn)
        controls.addWidget(self.rotate_left_btn)
        controls.addWidget(self.rotate_right_btn)
        controls.addStretch(1)
        row.addLayout(controls)

        row.addStretch(1)
        outer.addLayout(row)

    def set_thumbnail(self, pixmap: QPixmap, width: int, height: int) -> None:
        self.image_label.setFixedSize(width + 8, height + 8)
        self.image_label.setPixmap(pixmap)

    def set_state(
        self,
        output_position: int,
        split_start: bool,
        rotation: int,
        can_move_up: bool,
        can_move_down: bool,
    ) -> None:
        tags: list[str] = []
        if split_start:
            tags.append("Split Start")
        if rotation:
            tags.append(f"R{rotation}")

        suffix = f" | {' | '.join(tags)}" if tags else ""
        self.header.setText(f"Page {output_position}{suffix}")
        self.move_up_btn.setEnabled(can_move_up)
        self.move_down_btn.setEnabled(can_move_down)

        if split_start:
            self.setStyleSheet("background: #0b6e0b; border: 1px solid #1ca31c; border-radius: 6px;")
        else:
            self.setStyleSheet("background: transparent; border: 1px solid palette(mid); border-radius: 6px;")

    def _emit_move_up(self) -> None:
        # Defer list mutation until the button click event finishes inside Qt.
        QTimer.singleShot(0, lambda: self.move_up.emit(self.page))

    def _emit_move_down(self) -> None:
        # Defer list mutation until the button click event finishes inside Qt.
        QTimer.singleShot(0, lambda: self.move_down.emit(self.page))


class PdfEditorWindow(QMainWindow):
    DELETE_SOURCE_SETTING_KEY = "split/delete_source_after_export"

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("PDF Splitter")
        self.resize(1280, 820)
        self.settings = QSettings("RMRR", "PDF Splitter")

        self.metadata: pdf_ops.PdfMetadata | None = None
        self.password: str | None = None
        self.split_starts: set[int] = set()
        self.page_rotations: dict[int, int] = {}
        self.page_order: list[int] = []
        self.section_name_overrides: dict[int, str] = {}
        self.page_items: dict[int, QListWidgetItem] = {}
        self.page_rows: dict[int, PageRowWidget] = {}
        self.base_pixmaps: dict[int, QPixmap] = {}
        self.section_name_inputs: dict[int, QLineEdit] = {}
        self._loading_items = False
        self._pending_page_item_pages: list[int] = []
        self._pending_thumbnail_pages: list[int] = []
        self._active_load_token = 0
        self.thumbnail_width = 520

        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)

        top_bar = QHBoxLayout()
        open_button = QPushButton("Open PDF")
        open_button.clicked.connect(self._pick_pdf)
        self.file_label = QLabel("No PDF loaded")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        top_bar.addWidget(open_button)
        top_bar.addWidget(self.file_label, 1)
        layout.addLayout(top_bar)

        action_bar = QHBoxLayout()
        action_bar.addWidget(QLabel("Preview Size"))
        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(180)
        self.zoom_slider.setMaximum(520)
        self.zoom_slider.setValue(self.zoom_slider.maximum())
        self.zoom_slider.setFixedWidth(180)
        self.zoom_slider.valueChanged.connect(self._on_zoom_changed)
        action_bar.addWidget(self.zoom_slider)
        page_up_btn = QPushButton("Page Up")
        page_up_btn.clicked.connect(lambda: self._page_scroll(-1))
        action_bar.addWidget(page_up_btn)
        page_down_btn = QPushButton("Page Down")
        page_down_btn.clicked.connect(lambda: self._page_scroll(1))
        action_bar.addWidget(page_down_btn)

        self.delete_source_checkbox = QCheckBox("Delete source after split")
        self.delete_source_checkbox.setChecked(self._load_delete_source_setting())
        self.delete_source_checkbox.toggled.connect(self._save_delete_source_setting)
        action_bar.addWidget(self.delete_source_checkbox)
        action_bar.addStretch(1)

        export_btn = QPushButton("Export Splits")
        export_btn.clicked.connect(self._export)
        action_bar.addWidget(export_btn)
        layout.addLayout(action_bar)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        self.page_list = QListWidget()
        self.page_list.setSelectionMode(QAbstractItemView.NoSelection)
        self.page_list.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.page_list.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.page_list.setUniformItemSizes(False)
        self.page_list.setSpacing(8)
        self.page_list.viewport().installEventFilter(self)
        self._apply_thumbnail_size()
        splitter.addWidget(self.page_list)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        hint = QLabel(
            "How to use:\n"
            "- Click the page image to mark split starts.\n"
            "- Page 1 is always the first section start.\n"
            "- Use the arrow buttons to move pages up or down for export.\n"
            "- Use the rotate buttons beside each page to rotate."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: palette(window-text);")

        right_layout.addWidget(hint)
        output_label = QLabel("Output section names:")
        output_label.setStyleSheet("color: palette(window-text);")
        right_layout.addWidget(output_label)

        self.section_scroll = QScrollArea()
        self.section_scroll.setWidgetResizable(True)
        self.section_container = QWidget()
        self.section_layout = QVBoxLayout(self.section_container)
        self.section_layout.addStretch(1)
        self.section_scroll.setWidget(self.section_container)
        right_layout.addWidget(self.section_scroll, 1)

        splitter.addWidget(right_panel)
        right_panel.setMinimumWidth(420)
        splitter.setSizes([860, 420])
        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)
        layout.addWidget(splitter, 1)

        self.setCentralWidget(root)

    def bring_to_front(self) -> None:
        if self.isMinimized():
            self.setWindowState(self.windowState() & ~Qt.WindowMinimized)

        self.show()
        self.raise_()
        self.activateWindow()

        self.setWindowFlag(Qt.WindowStaysOnTopHint, True)
        self.show()
        self.raise_()
        self.activateWindow()
        self.setWindowFlag(Qt.WindowStaysOnTopHint, False)
        self.show()
        self.raise_()
        self.activateWindow()

    def eventFilter(self, obj: QObject, event: QEvent) -> bool:  # type: ignore[override]
        if obj is self.page_list.viewport() and event.type() == QEvent.Wheel:
            bar = self.page_list.verticalScrollBar()
            pixel = event.pixelDelta().y()
            if pixel:
                bar.setValue(bar.value() - int(pixel * 0.8))
                return True

            delta = event.angleDelta().y()
            if delta:
                notches = delta / 120.0
                bar.setValue(bar.value() - int(notches * 30))
                return True
        return super().eventFilter(obj, event)

    def _page_scroll(self, direction: int) -> None:
        bar = self.page_list.verticalScrollBar()
        step = max(120, int(bar.pageStep() * 0.85))
        bar.setValue(bar.value() + (step * direction))

    def _on_zoom_changed(self, value: int) -> None:
        self.thumbnail_width = int(value)
        self._apply_thumbnail_size()
        if self.metadata is None:
            return
        self._pending_thumbnail_pages = list(range(1, self.metadata.page_count + 1))
        load_token = self._active_load_token
        QTimer.singleShot(0, lambda: self._load_next_thumbnail(load_token))

    def _thumbnail_height(self) -> int:
        return int(self.thumbnail_width * 1.35)

    def _item_row_height(self) -> int:
        return self._thumbnail_height() + 74

    def _apply_thumbnail_size(self) -> None:
        self.page_list.setIconSize(QSize(self.thumbnail_width, self._thumbnail_height()))
        for page, item in self.page_items.items():
            item.setSizeHint(QSize(1, self._item_row_height()))
            row = self.page_rows[page]
            display = self._display_pixmap_for_page(page)
            row.set_thumbnail(display, self.thumbnail_width, self._thumbnail_height())

    def _pick_pdf(self) -> None:
        chosen, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF Files (*.pdf)")
        if chosen:
            self.load_pdf(chosen)

    def _restore_window_header(self) -> None:
        if self.metadata is None:
            self.setWindowTitle("PDF Splitter")
            self.file_label.setText("No PDF loaded")
            return

        self.setWindowTitle(f"PDF Splitter - {self.metadata.source_path.name}")
        self.file_label.setText(str(self.metadata.source_path))

    def load_pdf(self, path: str | Path) -> bool:
        path = str(Path(path).resolve())
        logging.info("Editor requested to load PDF: %s", path)
        password: str | None = None
        self.setWindowTitle("PDF Splitter - Loading...")
        self.file_label.setText(path)
        self.bring_to_front()
        QApplication.processEvents()

        while True:
            try:
                logging.info("Loading PDF metadata...")
                metadata = pdf_ops.load_pdf_metadata(path, password=password)
                logging.info("Metadata loaded. Page count=%s", metadata.page_count)
                break
            except pdf_ops.PasswordRequiredError:
                entered, ok = QInputDialog.getText(
                    self,
                    "Password Required",
                    "This PDF is protected. Enter password:",
                    QLineEdit.Password,
                )
                if not ok:
                    self._restore_window_header()
                    return False
                password = entered
            except pdf_ops.InvalidPasswordError:
                entered, ok = QInputDialog.getText(
                    self,
                    "Invalid Password",
                    "Password was incorrect. Try again:",
                    QLineEdit.Password,
                )
                if not ok:
                    self._restore_window_header()
                    return False
                password = entered
            except Exception as exc:
                self._restore_window_header()
                QMessageBox.critical(self, "Failed to Open PDF", str(exc))
                return False

        self.metadata = metadata
        self.password = password
        self.file_label.setText(str(metadata.source_path))
        self.setWindowTitle(f"PDF Splitter - {metadata.source_path.name}")

        self._reset_loaded_document_state()
        self.page_order = list(range(1, metadata.page_count + 1))
        load_token = self._active_load_token

        self._loading_items = True
        self.page_list.clear()
        self._pending_page_item_pages = list(range(1, metadata.page_count + 1))
        self._pending_thumbnail_pages = list(range(1, metadata.page_count + 1))
        QTimer.singleShot(0, lambda: self._create_next_page_batch(load_token))
        return True

    def _load_delete_source_setting(self) -> bool:
        value = self.settings.value(self.DELETE_SOURCE_SETTING_KEY, False, bool)
        return bool(value)

    def _save_delete_source_setting(self, checked: bool) -> None:
        self.settings.setValue(self.DELETE_SOURCE_SETTING_KEY, checked)

    def _clear_section_names_ui(self) -> None:
        while self.section_layout.count():
            child = self.section_layout.takeAt(0)
            widget = child.widget()
            if widget is not None:
                widget.deleteLater()
        self.section_layout.addStretch(1)
        self.section_name_inputs = {}

    def _reset_loaded_document_state(self) -> None:
        self.split_starts = set()
        self.page_rotations = {}
        self.page_order = []
        self.section_name_overrides = {}
        self.page_items = {}
        self.page_rows = {}
        self.base_pixmaps = {}
        self._pending_page_item_pages = []
        self._pending_thumbnail_pages = []
        self._active_load_token += 1
        self._loading_items = False
        self.page_list.clear()
        self._clear_section_names_ui()

    def _clear_loaded_pdf(self) -> None:
        self.metadata = None
        self.password = None
        self._reset_loaded_document_state()
        self._restore_window_header()

    def _create_next_page_batch(self, load_token: int) -> None:
        if load_token != self._active_load_token:
            return
        if self.metadata is None:
            return

        batch_size = 20
        created = 0
        while self._pending_page_item_pages and created < batch_size:
            page = self._pending_page_item_pages.pop(0)
            item, row = self._create_page_row(page)

            pixmap = QPixmap(self.thumbnail_width, self._thumbnail_height())
            pixmap.fill(Qt.lightGray)

            self.base_pixmaps[page] = pixmap
            self.page_items[page] = item
            self.page_rows[page] = row

            self._update_item_visual(page)
            created += 1

        if self._pending_page_item_pages:
            QTimer.singleShot(0, lambda: self._create_next_page_batch(load_token))
            return

        self._loading_items = False
        self._sync_page_list_order()
        self._refresh_sections_ui()
        logging.info("Editor finished loading page rows for %s (%s pages)", self.metadata.source_path, self.metadata.page_count)
        QTimer.singleShot(0, lambda: self._load_next_thumbnail(load_token))

    def _load_next_thumbnail(self, load_token: int) -> None:
        if load_token != self._active_load_token:
            return
        if self.metadata is None or not self._pending_thumbnail_pages:
            return

        page = self._pending_thumbnail_pages.pop(0)
        try:
            image = pdf_ops.render_page_thumbnail(
                self.metadata.source_path,
                page - 1,
                self.password,
                max_width=self.thumbnail_width,
            )
            pixmap = QPixmap.fromImage(image)
            self.base_pixmaps[page] = pixmap
            self._update_item_visual(page)
        except Exception:
            logging.exception("Failed to render thumbnail for page %s", page)

        QTimer.singleShot(0, lambda: self._load_next_thumbnail(load_token))

    def _page_position(self, page: int) -> int:
        return self.page_order.index(page) + 1

    def _current_split_positions(self) -> set[int]:
        split_positions: set[int] = set()
        for page in self.split_starts:
            position = self._page_position(page)
            if position > 1:
                split_positions.add(position)
        return split_positions

    def _ordered_sections(self) -> list[tuple[int, int, int]]:
        sections = pdf_ops.compute_sections(len(self.page_order), self._current_split_positions())
        return [(start, end, self.page_order[start - 1]) for start, end in sections]

    def _create_page_row(self, page: int) -> tuple[QListWidgetItem, PageRowWidget]:
        item = QListWidgetItem()
        item.setSizeHint(QSize(1, self._item_row_height()))
        self.page_list.addItem(item)

        row = PageRowWidget(page)
        row.split_toggled.connect(self._toggle_split_start)
        row.move_up.connect(lambda p=page: self._move_page(p, -1))
        row.move_down.connect(lambda p=page: self._move_page(p, 1))
        row.rotate_left.connect(lambda p=page: self._rotate_page(p, -90))
        row.rotate_right.connect(lambda p=page: self._rotate_page(p, 90))
        self.page_list.setItemWidget(item, row)
        return item, row

    def _sync_page_list_order(self) -> None:
        scroll_bar = self.page_list.verticalScrollBar()
        scroll_value = scroll_bar.value()

        self.page_list.setUpdatesEnabled(False)
        try:
            self.page_list.clear()
            self.page_items = {}
            self.page_rows = {}

            for page in self.page_order:
                item, row = self._create_page_row(page)
                self.page_items[page] = item
                self.page_rows[page] = row
                self._update_item_visual(page)
        finally:
            self.page_list.setUpdatesEnabled(True)

        for page in self.page_order:
            self._update_item_visual(page)

        QTimer.singleShot(0, lambda: scroll_bar.setValue(scroll_value))

    def _display_pixmap_for_page(self, page: int) -> QPixmap:
        base = self.base_pixmaps[page]
        rotation = self.page_rotations.get(page, 0)
        if not rotation:
            transformed = base
        else:
            transform = QTransform()
            transform.rotate(rotation)
            transformed = base.transformed(transform, Qt.SmoothTransformation)

        return transformed.scaled(
            self.thumbnail_width,
            self._thumbnail_height(),
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation,
        )

    def _toggle_split_start(self, page: int) -> None:
        if self._loading_items or self.metadata is None:
            return
        if self._page_position(page) == 1:
            return

        if page in self.split_starts:
            self.split_starts.remove(page)
        else:
            self.split_starts.add(page)

        self._update_item_visual(page)
        self._refresh_sections_ui()

    def _rotate_page(self, page: int, delta: int) -> None:
        if self._loading_items or self.metadata is None:
            return

        current = self.page_rotations.get(page, 0)
        self.page_rotations[page] = (current + delta) % 360
        self._update_item_visual(page)

    def _move_page(self, page: int, direction: int) -> None:
        if self._loading_items or self.metadata is None:
            return

        current_index = self.page_order.index(page)
        target_index = current_index + direction
        if target_index < 0 or target_index >= len(self.page_order):
            return

        logging.info("Moving page %s from position %s to %s", page, current_index + 1, target_index + 1)
        self.page_order[current_index], self.page_order[target_index] = (
            self.page_order[target_index],
            self.page_order[current_index],
        )
        self._sync_page_list_order()
        self._refresh_sections_ui()

    def _update_item_visual(self, page: int) -> None:
        row = self.page_rows[page]
        position = self._page_position(page)
        split = position == 1 or page in self.split_starts
        rotation = self.page_rotations.get(page, 0)
        row.set_state(
            position,
            split,
            rotation,
            can_move_up=position > 1,
            can_move_down=position < len(self.page_order),
        )
        display = self._display_pixmap_for_page(page)
        row.set_thumbnail(display, self.thumbnail_width, self._thumbnail_height())

    def _default_section_name(self, section_index: int) -> str:
        base = self.metadata.base_name if self.metadata else "document"
        return f"{base}_part_{section_index:02d}"

    def _refresh_sections_ui(self) -> None:
        if self.metadata is None:
            return

        self._clear_section_names_ui()
        sections = self._ordered_sections()

        for index, (start, end, start_page) in enumerate(sections, start=1):
            default_name = self._default_section_name(index)
            existing = self.section_name_overrides.get(start_page, default_name)

            block = QWidget()
            block_layout = QVBoxLayout(block)
            block_layout.setContentsMargins(6, 6, 6, 6)

            label = QLabel(f"Section {index}: pages {start}-{end}")
            label.setStyleSheet("color: palette(window-text);")
            name_input = QLineEdit(existing)
            name_input.setMinimumWidth(360)
            name_input.setPlaceholderText(default_name)
            name_input.textChanged.connect(lambda txt, p=start_page: self._on_section_name_changed(p, txt))

            block_layout.addWidget(label)
            block_layout.addWidget(name_input)
            block.setStyleSheet(
                "background: palette(base); "
                "border: 1px solid palette(mid); "
                "border-radius: 6px; "
                "color: palette(text);"
            )

            self.section_layout.addWidget(block)
            self.section_name_inputs[start_page] = name_input

        self.section_layout.addStretch(1)

        for page in self.page_order:
            self._update_item_visual(page)

    def _on_section_name_changed(self, start_page: int, text: str) -> None:
        self.section_name_overrides[start_page] = text

    def _export(self) -> None:
        if self.metadata is None:
            QMessageBox.information(self, "Nothing to Export", "Load a PDF first.")
            return

        sections = self._ordered_sections()
        if not sections:
            QMessageBox.information(self, "Nothing to Export", "This PDF has no pages to export.")
            return
        names: dict[int, str] = {}
        split_positions = self._current_split_positions()
        for idx, (start, _, start_page) in enumerate(sections, start=1):
            default_name = self._default_section_name(idx)
            text = self.section_name_inputs.get(start_page).text() if start_page in self.section_name_inputs else ""
            names[start] = text or default_name

        try:
            outputs = pdf_ops.split_pdf(
                source_path=self.metadata.source_path,
                password=self.password,
                split_starts=split_positions,
                section_names=names,
                page_rotations=self.page_rotations,
                page_order=self.page_order,
                delete_source=self.delete_source_checkbox.isChecked(),
                output_dir=self.metadata.source_path.parent,
            )
        except Exception as exc:
            QMessageBox.critical(self, "Export Failed", str(exc))
            return

        output_dir = self.metadata.source_path.parent
        delete_source = self.delete_source_checkbox.isChecked()
        self._clear_loaded_pdf()
        msg = f"Wrote {len(outputs)} file(s) to:\n{output_dir}"
        if delete_source:
            msg += "\n\nSource PDF deleted."
        QMessageBox.information(self, "Export Complete", msg)

    def resizeEvent(self, event) -> None:  # type: ignore[override]
        super().resizeEvent(event)
        self._apply_thumbnail_size()


def main() -> int:
    app = QApplication([])
    window = PdfEditorWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
