import os

from PySide6.QtCore import QObject, QThread, QUrl, Signal, Slot
from PySide6.QtGui import QDesktopServices, QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    BodyLabel,
    CaptionLabel,
    CardWidget,
    CheckBox,
    FluentIcon,
    FluentWindow,
    IndeterminateProgressBar,
    IndeterminateProgressRing,
    LineEdit,
    MessageBox,
    NavigationItemPosition,
    PrimaryPushButton,
    ProgressBar,
    PushButton,
    SubtitleLabel,
    Theme,
    TitleLabel,
    setTheme,
    setThemeColor,
)


def _success_message(result):
    message = "Done."
    if result.get("images") is not None:
        replaced, missed = result["images"]
        message = f"Done. {replaced} tables and equations kept as images"
        if missed:
            message += f", {missed} left as converted text"
        message += "."
    if result.get("polish") is not None:
        counts = result["polish"]
        message += (f" Restored {counts['spaces']} spaces,"
                    f" removed {counts['page breaks dropped']} page breaks.")
    return message


def _clear_saved_output(page):
    page._last_output_path = None
    page.open_button.setEnabled(False)


def _open_saved_output(page):
    destination = page._last_output_path
    if not destination or not os.path.isfile(destination):
        _clear_saved_output(page)
        MessageBox(
            "Cannot open document", f"The saved DOCX is no longer available: {destination}",
            page).exec()
    elif not QDesktopServices.openUrl(QUrl.fromLocalFile(destination)):
        MessageBox(
            "Cannot open document",
            f"Your system could not open {destination}. Choose a default app for DOCX files.",
            page).exec()


class FileDropZone(CardWidget):
    """Fluent card that accepts drops for one file extension."""

    fileDropped = Signal(str)
    browseRequested = Signal()

    def __init__(self, file_type, extension, parent=None):
        super().__init__(parent)
        self.file_type = file_type
        self.extension = extension.lower().lstrip(".")
        self.setAcceptDrops(True)
        self.setObjectName(f"{self.extension}DropZone")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(6)

        top_row = QHBoxLayout()
        top_row.addWidget(SubtitleLabel(f"Input {file_type}", self))
        top_row.addStretch(1)
        self.browse_button = PushButton(f"Browse {file_type}", self)
        self.browse_button.clicked.connect(self.browseRequested)
        top_row.addWidget(self.browse_button)
        layout.addLayout(top_row)

        self.drop_label = BodyLabel(f"Drop a {file_type} file here or browse.", self)
        self.drop_label.setWordWrap(True)
        layout.addWidget(self.drop_label)
        self.path_label = CaptionLabel(f"No {file_type} selected", self)
        self.path_label.setWordWrap(True)
        layout.addWidget(self.path_label)

    def file_paths(self, mime_data):
        return [url.toLocalFile() for url in mime_data.urls()
                if url.isLocalFile()
                and url.toLocalFile().lower().endswith("." + self.extension)]

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.file_paths(event.mimeData())
        if paths:
            self.fileDropped.emit(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def set_selected_file(self, path):
        self.path_label.setText(path)


class ConversionWorker(QObject):
    """Worker object moved to a QThread; all UI updates use its signals."""

    statusChanged = Signal(str)
    progressChanged = Signal(int)
    warningRaised = Signal(str)
    succeeded = Signal(object)
    failed = Signal(str)
    finished = Signal()

    def __init__(self, converter, *args, report_warnings=False, report_progress=False,
                 **kwargs):
        super().__init__()
        self.converter = converter
        self.args = args
        self.kwargs = kwargs
        self.report_warnings = report_warnings
        self.report_progress = report_progress

    @Slot()
    def run(self):
        try:
            kwargs = dict(self.kwargs)
            kwargs["on_status"] = self.statusChanged.emit
            if self.report_warnings:
                kwargs["on_warning"] = self.warningRaised.emit
            if self.report_progress:
                kwargs["on_progress"] = self.progressChanged.emit
            result = self.converter(*self.args, **kwargs)
        except Exception as exc:
            self.failed.emit(str(exc))
        else:
            self.succeeded.emit(result)
        finally:
            self.finished.emit()


def _start_worker(page, converter, *args, report_warnings=False, report_progress=False,
                  **kwargs):
    _clear_saved_output(page)
    page._set_busy(True)
    thread = QThread(page)
    worker = ConversionWorker(
        converter, *args, report_warnings=report_warnings,
        report_progress=report_progress, **kwargs)
    worker.moveToThread(thread)
    thread.started.connect(worker.run)
    worker.statusChanged.connect(page.status_label.setText)
    if report_warnings:
        worker.warningRaised.connect(page._conversion_warning)
    if report_progress:
        worker.progressChanged.connect(page._conversion_progress)
    worker.succeeded.connect(page._conversion_succeeded)
    worker.failed.connect(page._conversion_failed)
    worker.finished.connect(thread.quit)
    worker.finished.connect(worker.deleteLater)
    thread.finished.connect(page._thread_finished)
    thread.finished.connect(thread.deleteLater)
    page._thread = thread
    page._worker = worker
    thread.start()


class PdfConverterPage(QWidget):
    def __init__(self, converter, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.setObjectName("pdfConverterPage")
        self.input_path = None
        self._last_output_path = None
        self._active_destination = None
        self._thread = None
        self._worker = None

        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)

        heading = QVBoxLayout()
        heading.setSpacing(2)
        heading.addWidget(TitleLabel("PDF to DOCX", self))
        description = CaptionLabel(
            "Convert locally while keeping document content editable.", self)
        description.setWordWrap(True)
        heading.addWidget(description)
        layout.addLayout(heading)

        self.drop_zone = FileDropZone("PDF", "pdf", self)
        self.drop_zone.fileDropped.connect(self.set_input)
        self.drop_zone.browseRequested.connect(self._browse_pdf)
        layout.addWidget(self.drop_zone)

        output_card = CardWidget(self)
        output_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        output_layout = QVBoxLayout(output_card)
        output_layout.setContentsMargins(14, 12, 14, 12)
        output_layout.setSpacing(6)
        output_layout.addWidget(SubtitleLabel("Output document", output_card))
        output_row = QHBoxLayout()
        self.output_edit = LineEdit(output_card)
        self.output_edit.setPlaceholderText("Choose where to save the DOCX")
        output_row.addWidget(self.output_edit, 1)
        self.output_browse_button = PushButton("Save as", output_card)
        self.output_browse_button.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_browse_button)
        output_layout.addLayout(output_row)
        layout.addWidget(output_card)

        options_card = CardWidget(self)
        options_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        options_layout = QVBoxLayout(options_card)
        options_layout.setContentsMargins(14, 12, 14, 12)
        options_layout.setSpacing(6)
        options_layout.addWidget(SubtitleLabel("Conversion options", options_card))

        self.images_option = CheckBox(
            "Keep tables and equations as images", options_card)
        self.images_option.setChecked(True)
        options_layout.addWidget(self.images_option)
        images_help = CaptionLabel(
            "Preserves their layout while body text stays editable.", options_card)
        images_help.setWordWrap(True)
        options_layout.addWidget(images_help)

        self.tidy_option = CheckBox(
            "Tidy the document for reading and editing", options_card)
        self.tidy_option.setChecked(True)
        options_layout.addWidget(self.tidy_option)
        tidy_help = CaptionLabel(
            "Substitutes missing fonts, restores spaces, and lets text flow.",
            options_card)
        tidy_help.setWordWrap(True)
        options_layout.addWidget(tidy_help)
        layout.addWidget(options_card)

        footer = QHBoxLayout()
        footer.setSpacing(8)
        self.status_label = BodyLabel("Ready", self)
        self.status_label.setWordWrap(True)
        footer.addWidget(self.status_label, 1)
        self.open_button = PushButton("Open DOCX", self)
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(lambda: _open_saved_output(self))
        footer.addWidget(self.open_button)
        self.convert_button = PrimaryPushButton("Convert", self)
        self.convert_button.setEnabled(False)
        self.convert_button.clicked.connect(self.start_conversion)
        footer.addWidget(self.convert_button)
        layout.addLayout(footer)

        self.progress = IndeterminateProgressBar(self, start=False)
        self.progress.hide()
        layout.addWidget(self.progress)
        layout.addStretch(1)
        self.output_edit.textChanged.connect(lambda: _clear_saved_output(self))

    def set_input(self, path):
        if not path or not path.lower().endswith(".pdf"):
            return False
        _clear_saved_output(self)
        self.input_path = path
        self.drop_zone.set_selected_file(path)
        self.output_edit.setText(os.path.splitext(path)[0] + ".docx")
        self.convert_button.setEnabled(True)
        self.status_label.setText("Ready to convert")
        return True

    def set_output(self, path):
        if path:
            self.output_edit.setText(path)

    def file_paths(self, mime_data):
        return self.drop_zone.file_paths(mime_data)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.file_paths(event.mimeData())
        if paths:
            self.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def _browse_pdf(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a PDF file", "", "PDF files (*.pdf)")
        if path:
            self.set_input(path)

    def _browse_output(self):
        current = self.output_edit.text().strip()
        initial = current or (
            os.path.splitext(self.input_path)[0] + ".docx"
            if self.input_path else "output.docx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save DOCX as", initial, "Word documents (*.docx)")
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.set_output(path)

    def start_conversion(self):
        if not self.input_path:
            return
        destination = self.output_edit.text().strip()
        if not destination:
            destination = os.path.splitext(self.input_path)[0] + ".docx"
        if not destination.lower().endswith(".docx"):
            destination += ".docx"
        self.output_edit.setText(destination)
        self._active_destination = os.path.abspath(destination)

        self.status_label.setText("Converting PDF…")
        _start_worker(
            self, self.converter,
            self.input_path, destination,
            snip=self.images_option.isChecked(),
            tidy=self.tidy_option.isChecked())

    def _set_busy(self, busy):
        self.setAcceptDrops(not busy)
        self.drop_zone.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.output_browse_button.setEnabled(not busy)
        self.images_option.setEnabled(not busy)
        self.tidy_option.setEnabled(not busy)
        self.convert_button.setEnabled(not busy and self.input_path is not None)
        self.open_button.setEnabled(
            not busy and self._last_output_path is not None
            and os.path.isfile(self._last_output_path))
        if busy:
            self.progress.start()
        else:
            self.progress.stop()
        self.progress.setVisible(busy)

    def _conversion_succeeded(self, result):
        self._last_output_path = self._active_destination
        self.status_label.setText(_success_message(result))

    def _conversion_failed(self, message):
        self.status_label.setText("Conversion failed")
        MessageBox("Conversion failed", message, self).exec()

    @Slot()
    def _thread_finished(self):
        self._worker = None
        self._thread = None
        self._set_busy(False)

    def closeEvent(self, event):
        if self._thread and self._thread.isRunning():
            self.status_label.setText("Please wait for the conversion to finish.")
            event.ignore()
            return
        event.accept()


def _default_tex_output(source_path):
    return os.path.splitext(source_path)[0] + "_from_tex.docx"


class TexConverterPage(QWidget):
    def __init__(self, converter, parent=None):
        super().__init__(parent)
        self.converter = converter
        self.input_path = None
        self._last_output_path = None
        self._warnings = []
        self._thread = None
        self._worker = None
        self.setObjectName("texConverterPage")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(10)
        heading = QVBoxLayout()
        heading.setSpacing(2)
        heading.addWidget(TitleLabel("TEX to DOCX", self))
        description = CaptionLabel(
            "Convert a TeX source with its figures and bibliography.", self)
        description.setWordWrap(True)
        heading.addWidget(description)
        layout.addLayout(heading)

        self.drop_zone = FileDropZone("TeX source", "tex", self)
        self.drop_zone.fileDropped.connect(self.set_input)
        self.drop_zone.browseRequested.connect(self._browse_tex)
        layout.addWidget(self.drop_zone)

        output_card = CardWidget(self)
        output_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        output_layout = QVBoxLayout(output_card)
        output_layout.setContentsMargins(14, 12, 14, 12)
        output_layout.setSpacing(6)
        output_layout.addWidget(SubtitleLabel("Output document", output_card))
        output_row = QHBoxLayout()
        self.output_edit = LineEdit(output_card)
        self.output_edit.setPlaceholderText("Choose where to save the DOCX")
        output_row.addWidget(self.output_edit, 1)
        self.output_browse_button = PushButton("Save as", output_card)
        self.output_browse_button.clicked.connect(self._browse_output)
        output_row.addWidget(self.output_browse_button)
        output_layout.addLayout(output_row)
        self.images_option = CheckBox(
            "Insert tables and equations as images", output_card)
        self.images_option.setChecked(False)
        output_layout.addWidget(self.images_option)
        images_help = CaptionLabel(
            "Renders them with LaTeX so they look exactly like the PDF. "
            "Leave unchecked for editable Word tables and equations.", output_card)
        images_help.setWordWrap(True)
        output_layout.addWidget(images_help)
        layout.addWidget(output_card)

        footer = QHBoxLayout()
        footer.setSpacing(8)
        self.status_label = BodyLabel("Ready", self)
        self.status_label.setWordWrap(True)
        footer.addWidget(self.status_label, 1)
        self.open_button = PushButton("Open DOCX", self)
        self.open_button.setEnabled(False)
        self.open_button.clicked.connect(lambda: _open_saved_output(self))
        footer.addWidget(self.open_button)
        self.convert_button = PrimaryPushButton("Convert", self)
        self.convert_button.setEnabled(False)
        self.convert_button.clicked.connect(self.start_conversion)
        footer.addWidget(self.convert_button)
        layout.addLayout(footer)

        progress_row = QHBoxLayout()
        self.progress = ProgressBar(self)
        self.progress.setRange(0, 100)
        self.progress.hide()
        progress_row.addWidget(self.progress, 1)
        self.percentage_label = CaptionLabel("Stages completed: 0%", self)
        self.percentage_label.hide()
        progress_row.addWidget(self.percentage_label)
        self.activity_ring = IndeterminateProgressRing(self, start=False)
        self.activity_ring.setFixedSize(20, 20)
        self.activity_ring.setStrokeWidth(2)
        self.activity_ring.hide()
        progress_row.addWidget(self.activity_ring)
        layout.addLayout(progress_row)
        layout.addStretch(1)
        self.output_edit.textChanged.connect(lambda: _clear_saved_output(self))

    def set_input(self, path):
        if not path or not path.lower().endswith(".tex"):
            return False
        _clear_saved_output(self)
        self.input_path = path
        self.drop_zone.set_selected_file(path)
        self.output_edit.setText(_default_tex_output(path))
        self.convert_button.setEnabled(True)
        self.status_label.setText("Ready to convert")
        return True

    def set_output(self, path):
        if path:
            self.output_edit.setText(path)

    def file_paths(self, mime_data):
        return self.drop_zone.file_paths(mime_data)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.file_paths(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        paths = self.file_paths(event.mimeData())
        if paths:
            self.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def _browse_tex(self):
        initial = os.path.dirname(self.input_path) if self.input_path else ""
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose a LaTeX source", initial, "LaTeX files (*.tex)")
        if path:
            self.set_input(path)

    def _browse_output(self):
        initial = self.output_edit.text().strip() or (
            _default_tex_output(self.input_path) if self.input_path
            else "output_from_tex.docx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Save DOCX as", initial, "Word documents (*.docx)")
        if path:
            if not path.lower().endswith(".docx"):
                path += ".docx"
            self.set_output(path)

    def start_conversion(self):
        if not self.input_path:
            return
        destination = self.output_edit.text().strip()
        if not destination:
            destination = _default_tex_output(self.input_path)
        if not destination.lower().endswith(".docx"):
            destination += ".docx"
        self.output_edit.setText(destination)
        self._warnings.clear()
        self.status_label.setText("Converting LaTeX to Word...")
        image_options = ({"tables_as_images": True, "equations_as_images": True}
                         if self.images_option.isChecked() else {})
        _start_worker(
            self, self.converter, self.input_path, destination,
            report_warnings=True, report_progress=True, **image_options)

    def _set_busy(self, busy):
        self.setAcceptDrops(not busy)
        self.drop_zone.setEnabled(not busy)
        self.output_edit.setEnabled(not busy)
        self.output_browse_button.setEnabled(not busy)
        self.images_option.setEnabled(not busy)
        self.convert_button.setEnabled(not busy and self.input_path is not None)
        self.open_button.setEnabled(
            not busy and self._last_output_path is not None
            and os.path.isfile(self._last_output_path))
        if busy:
            self.progress.setError(False)
            self._conversion_progress(0)
            self.progress.show()
            self.percentage_label.show()
            self.activity_ring.start()
        else:
            self.activity_ring.stop()
        self.activity_ring.setVisible(busy)

    @Slot(int)
    def _conversion_progress(self, percent):
        percent = max(0, min(100, percent))
        self.progress.setValue(percent)
        self.percentage_label.setText(f"Stages completed: {percent}%")

    def _conversion_succeeded(self, destination):
        self._last_output_path = os.path.abspath(destination)
        self._conversion_progress(100)
        self.activity_ring.stop()
        self.activity_ring.hide()
        if self._warnings:
            self.status_label.setText(f"Saved with warnings: {destination}")
            MessageBox(
                "Saved with warnings", "\n".join(self._warnings), self).exec()
        else:
            self.status_label.setText(f"Done. Saved to {destination}")

    def _conversion_warning(self, message):
        if message not in self._warnings:
            self._warnings.append(message)

    def _conversion_failed(self, message):
        self.progress.setError(True)
        self.activity_ring.stop()
        self.activity_ring.hide()
        self.status_label.setText("Conversion failed")
        MessageBox("Conversion failed", message, self).exec()

    @Slot()
    def _thread_finished(self):
        self._worker = None
        self._thread = None
        self._set_busy(False)


class ConverterWindow(FluentWindow):
    def __init__(self, converter, latex_converter=None, parent=None):
        super().__init__(parent)
        if latex_converter is None:
            from latex_to_docx import convert_latex_to_docx

            latex_converter = convert_latex_to_docx

        setTheme(Theme.AUTO)
        setThemeColor("#0078D4")
        self.setWindowTitle("PDF and LaTeX to Word Converter")
        self.setMinimumWidth(760)
        self.resize(880, 560)
        self.setAcceptDrops(True)

        self.pdf_page = PdfConverterPage(converter, self)
        self.tex_page = TexConverterPage(latex_converter, self)
        self.addSubInterface(
            self.pdf_page, FluentIcon.DOCUMENT, "PDF to Word",
            position=NavigationItemPosition.TOP)
        self.addSubInterface(
            self.tex_page, FluentIcon.CODE, "TEX to DOCX",
            position=NavigationItemPosition.TOP)

    def set_input(self, path):
        """Keep the existing PDF preload entry point for CLI GUI startup."""
        return self.pdf_page.set_input(path)

    def dragEnterEvent(self, event: QDragEnterEvent):
        page = self.stackedWidget.currentWidget()
        if page.file_paths(event.mimeData()) and not page._thread:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        self.dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        page = self.stackedWidget.currentWidget()
        paths = page.file_paths(event.mimeData())
        if paths and not page._thread:
            page.set_input(paths[0])
            event.acceptProposedAction()
        else:
            event.ignore()

    def closeEvent(self, event):
        pages = (self.pdf_page, self.tex_page)
        active = [page for page in pages
                  if page._thread and page._thread.isRunning()]
        if active:
            active[0].status_label.setText(
                "Please wait for the conversion to finish.")
            event.ignore()
            return
        event.accept()
