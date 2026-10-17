import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import (
    QAbstractAnimation, QEventLoop, QMimeData, QPoint, QPointF, Qt, QTimer, QUrl,
)
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication

from fluent_ui import (
    ConversionWorker,
    ConverterWindow,
    FileDropZone,
    _success_message,
)


class FluentUiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def wait_for_worker(self, page):
        loop = QEventLoop()
        poll = QTimer()
        poll.timeout.connect(lambda: loop.quit() if page._thread is None else None)
        poll.start(10)
        QTimer.singleShot(3000, loop.quit)
        loop.exec()
        poll.stop()
        self.assertIsNone(page._thread)

    def test_conversion_options_default_to_enabled_and_wait_for_pdf(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        page = window.pdf_page

        self.assertEqual(window.stackedWidget.count(), 2)
        self.assertTrue(page.images_option.isChecked())
        self.assertTrue(page.tidy_option.isChecked())
        self.assertFalse(page.convert_button.isEnabled())
        self.assertFalse(page.open_button.isEnabled())
        self.assertFalse(window.tex_page.open_button.isEnabled())

    def test_tex_selection_sets_safe_default_output_and_rejects_other_files(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        page = window.tex_page

        self.assertFalse(page.set_input("paper.pdf"))
        self.assertTrue(page.set_input("/tmp/paper.tex"))
        self.assertEqual(page.output_edit.text(), "/tmp/docx/paper_from_tex.docx")
        self.assertTrue(page.convert_button.isEnabled())

    def test_progress_bar_animates_only_while_conversion_is_busy(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        page = window.pdf_page

        self.assertFalse(page.progress.isStarted())
        page._set_busy(True)
        self.assertFalse(page.progress.isHidden())
        self.assertTrue(page.progress.isStarted())
        start_position = page.progress.shortPos

        loop = QEventLoop()
        QTimer.singleShot(250, loop.quit)
        loop.exec()
        self.assertNotEqual(page.progress.shortPos, start_position)

        page._set_busy(False)
        self.assertTrue(page.progress.isHidden())
        self.assertFalse(page.progress.isStarted())

    def test_success_message_reports_the_correct_page_break_count(self):
        self.assertEqual(
            _success_message({
                "images": (12, 2),
                "polish": {"spaces": 8, "page breaks dropped": 5},
            }),
            "Done. 12 tables and equations kept as images, 2 left as converted "
            "text. Restored 8 spaces, removed 5 page breaks.")

    def test_tex_activity_animates_while_stage_percentage_is_unchanged(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        page = window.tex_page
        window.switchTo(page)
        window.show()
        self.app.processEvents()
        page._set_busy(True)
        page._conversion_progress(14)
        start_angle = page.activity_ring.startAngle
        self.assertEqual(page.activity_ring.aniGroup.state(), QAbstractAnimation.Running)
        loop = QEventLoop()
        poll = QTimer()
        poll.timeout.connect(
            lambda: loop.quit() if page.activity_ring.startAngle != start_angle else None)
        poll.start(20)
        QTimer.singleShot(2000, loop.quit)
        loop.exec()
        poll.stop()

        self.assertNotEqual(page.activity_ring.startAngle, start_angle)
        self.assertEqual(page.progress.value(), 14)
        self.assertEqual(page.percentage_label.text(), "Stages completed: 14%")
        page._set_busy(False)
        self.assertFalse(page.progress.isHidden())
        self.assertTrue(page.activity_ring.isHidden())
        self.assertEqual(page.activity_ring.aniGroup.state(), QAbstractAnimation.Stopped)

    def test_pdf_selection_sets_default_output_and_rejects_other_files(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        page = window.pdf_page

        self.assertFalse(page.set_input("paper.docx"))
        self.assertTrue(window.set_input("/tmp/paper.pdf"))
        self.assertEqual(page.output_edit.text(), "/tmp/paper.docx")
        self.assertTrue(page.convert_button.isEnabled())

        page.output_edit.setText("/tmp/custom.docx")
        self.assertEqual(page.output_edit.text(), "/tmp/custom.docx")

    def test_drop_accepts_pdf_and_uses_the_first_pdf_url(self):
        zone = FileDropZone("PDF", "pdf")
        self.addCleanup(zone.close)
        mime_data = QMimeData()
        mime_data.setUrls([
            QUrl.fromLocalFile("/tmp/readme.txt"),
            QUrl.fromLocalFile("/tmp/paper.PDF"),
        ])
        dropped = []
        zone.fileDropped.connect(dropped.append)

        drag = QDragEnterEvent(
            QPoint(5, 5), Qt.CopyAction, mime_data,
            Qt.LeftButton, Qt.NoModifier)
        zone.dragEnterEvent(drag)
        self.assertTrue(drag.isAccepted())

        drop = QDropEvent(
            QPointF(5, 5), Qt.CopyAction, mime_data,
            Qt.LeftButton, Qt.NoModifier)
        zone.dropEvent(drop)
        self.assertTrue(drop.isAccepted())
        self.assertEqual(dropped, ["/tmp/paper.PDF"])

    def test_drop_accepts_tex_and_uses_the_first_tex_url(self):
        zone = FileDropZone("TeX source", "tex")
        self.addCleanup(zone.close)
        mime_data = QMimeData()
        mime_data.setUrls([
            QUrl.fromLocalFile("/tmp/readme.txt"),
            QUrl.fromLocalFile("/tmp/paper.TeX"),
        ])
        dropped = []
        zone.fileDropped.connect(dropped.append)

        drop = QDropEvent(
            QPointF(5, 5), Qt.CopyAction, mime_data,
            Qt.LeftButton, Qt.NoModifier)
        zone.dropEvent(drop)

        self.assertTrue(drop.isAccepted())
        self.assertEqual(dropped, ["/tmp/paper.TeX"])

    def test_window_drop_targets_the_selected_converter_page(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        for page, extension in ((window.pdf_page, "PDF"), (window.tex_page, "TeX")):
            with self.subTest(extension=extension):
                window.switchTo(page)
                source = "/tmp/paper." + extension
                mime_data = QMimeData()
                mime_data.setUrls([QUrl.fromLocalFile(source)])
                drag = QDragEnterEvent(
                    QPoint(5, 5), Qt.CopyAction, mime_data,
                    Qt.LeftButton, Qt.NoModifier)
                window.dragEnterEvent(drag)
                self.assertTrue(drag.isAccepted())
                drop = QDropEvent(
                    QPointF(5, 5), Qt.CopyAction, mime_data,
                    Qt.LeftButton, Qt.NoModifier)
                window.dropEvent(drop)
                self.assertTrue(drop.isAccepted())
                self.assertEqual(page.input_path, source)

    def test_worker_emits_status_and_success(self):
        results = []
        statuses = []
        finished = []

        def convert(src, dst, snip, tidy, on_status):
            on_status("Tidying")
            return {"images": None, "polish": None}

        worker = ConversionWorker(convert, "input.pdf", "output.docx", True, True)
        worker.succeeded.connect(results.append)
        worker.statusChanged.connect(statuses.append)
        worker.finished.connect(lambda: finished.append(True))
        worker.run()

        self.assertEqual(results, [{"images": None, "polish": None}])
        self.assertEqual(statuses, ["Tidying"])
        self.assertEqual(finished, [True])

    def test_worker_emits_original_error_text(self):
        errors = []
        finished = []

        def fail(*args, **kwargs):
            raise RuntimeError("page conversion failed")

        worker = ConversionWorker(fail, "input.pdf", "output.docx", True, True)
        worker.failed.connect(errors.append)
        worker.finished.connect(lambda: finished.append(True))
        worker.run()

        self.assertEqual(errors, ["page conversion failed"])
        self.assertEqual(finished, [True])

    def test_tex_worker_emits_warning_signal(self):
        warnings = []

        def convert(src, dst, on_status, on_warning):
            on_warning("Table 2: header recovery needs review")
            return dst

        worker = ConversionWorker(
            convert, "paper.tex", "paper.docx", report_warnings=True)
        worker.warningRaised.connect(warnings.append)
        worker.run()

        self.assertEqual(warnings, ["Table 2: header recovery needs review"])

    def test_tex_worker_emits_progress_signal(self):
        progress = []

        def convert(src, dst, on_status, on_progress):
            on_progress(14)
            on_progress(100)
            return dst

        worker = ConversionWorker(
            convert, "paper.tex", "paper.docx", report_progress=True)
        worker.progressChanged.connect(progress.append)
        worker.run()

        self.assertEqual(progress, [14, 100])

    def test_tex_page_shows_saved_with_warnings_after_worker_signal(self):
        def convert(src, dst, on_status, on_warning, on_progress, **options):
            Path(dst).touch()
            on_progress(50)
            on_warning("Figure 3: image path is missing.png")
            on_progress(100)
            return dst

        window = ConverterWindow(lambda *args, **kwargs: {}, convert)
        self.addCleanup(window.close)
        page = window.tex_page
        page.set_input("/tmp/paper.tex")

        with tempfile.TemporaryDirectory() as directory, patch("fluent_ui.MessageBox") as message_box:
            page.set_output(os.path.join(directory, "warned.docx"))
            page.start_conversion()
            self.wait_for_worker(page)
            self.assertTrue(page.open_button.isEnabled())

        self.assertTrue(page.status_label.text().startswith("Saved with warnings"))
        self.assertEqual(page._warnings, ["Figure 3: image path is missing.png"])
        self.assertEqual(page.progress.value(), 100)
        self.assertFalse(page.progress.isHidden())
        self.assertTrue(page.activity_ring.isHidden())
        message_box.assert_called_once()

    def test_tex_failure_keeps_last_progress_and_next_run_resets_it(self):
        attempts = []

        def convert(src, dst, on_status, on_warning, on_progress, **options):
            attempts.append(True)
            on_progress(21)
            if len(attempts) == 1:
                raise RuntimeError("Pandoc failed")
            Path(dst).touch()
            on_progress(100)
            return dst

        window = ConverterWindow(lambda *args, **kwargs: {}, convert)
        self.addCleanup(window.close)
        page = window.tex_page
        page.set_input("paper.tex")
        with tempfile.TemporaryDirectory() as directory, patch("fluent_ui.MessageBox"):
            page.set_output(os.path.join(directory, "retry.docx"))
            page.start_conversion()
            self.wait_for_worker(page)
            self.assertEqual(page.progress.value(), 21)
            self.assertTrue(page.progress.isError())
            self.assertTrue(page.activity_ring.isHidden())
            self.assertFalse(page.open_button.isEnabled())

            page.start_conversion()
            self.assertEqual(page.progress.value(), 0)
            self.assertFalse(page.progress.isError())
            self.assertFalse(page.open_button.isEnabled())
            self.wait_for_worker(page)
            self.assertEqual(page.progress.value(), 100)
            self.assertTrue(page.open_button.isEnabled())

    def test_both_pages_open_the_completed_destination_and_invalidate_edits(self):
        def convert_pdf(src, dst, snip, tidy, on_status):
            Path(dst).touch()
            return {"images": None, "polish": None}

        def convert_tex(src, dst, on_status, on_warning, on_progress, **options):
            Path(dst).touch()
            on_progress(100)
            return dst

        window = ConverterWindow(convert_pdf, convert_tex)
        self.addCleanup(window.close)
        with tempfile.TemporaryDirectory() as directory:
            for page, extension in ((window.pdf_page, "pdf"), (window.tex_page, "tex")):
                with self.subTest(extension=extension):
                    page.set_input("paper." + extension)
                    destination = os.path.join(directory, f"{extension} output with spaces.docx")
                    page.set_output(destination)
                    page.start_conversion()
                    self.assertFalse(page.open_button.isEnabled())
                    self.wait_for_worker(page)
                    self.assertTrue(page.open_button.isEnabled())
                    with patch("fluent_ui.QDesktopServices.openUrl", return_value=True) as opener:
                        page.open_button.click()
                    self.assertEqual(opener.call_args.args[0].toLocalFile(), destination.replace("\\", "/"))

                    page.output_edit.setText(destination + ".edited.docx")
                    self.assertFalse(page.open_button.isEnabled())
                    self.assertIsNone(page._last_output_path)
                    page._last_output_path = destination
                    page._set_busy(False)
                    page.set_input("another." + extension)
                    self.assertFalse(page.open_button.isEnabled())
                    self.assertIsNone(page._last_output_path)

    def test_open_docx_reports_missing_files_and_failed_system_launches(self):
        window = ConverterWindow(lambda *args, **kwargs: {})
        self.addCleanup(window.close)
        with tempfile.TemporaryDirectory() as directory:
            for page in (window.pdf_page, window.tex_page):
                destination = os.path.join(directory, "saved.docx")
                Path(destination).touch()
                page._last_output_path = destination
                page._set_busy(False)
                with patch("fluent_ui.QDesktopServices.openUrl", return_value=False) as opener, \
                        patch("fluent_ui.MessageBox") as message_box:
                    page.open_button.click()
                    opener.assert_called_once()
                    message_box.assert_called_once()
                    self.assertIn(destination, message_box.call_args.args[1])

                os.remove(destination)
                with patch("fluent_ui.QDesktopServices.openUrl") as opener, \
                        patch("fluent_ui.MessageBox") as message_box:
                    page.open_button.click()
                    opener.assert_not_called()
                    message_box.assert_called_once()
                    self.assertFalse(page.open_button.isEnabled())
                    self.assertIsNone(page._last_output_path)

    def test_window_finishes_conversion_from_background_thread(self):
        def convert(src, dst, snip, tidy, on_status):
            on_status("Processing")
            return {"images": None, "polish": None}

        window = ConverterWindow(convert)
        self.addCleanup(window.close)
        page = window.pdf_page
        window.set_input("/tmp/paper.pdf")
        page.start_conversion()
        self.wait_for_worker(page)

        self.assertIsNone(page._thread)
        self.assertEqual(page.status_label.text(), "Done.")
        self.assertTrue(page.convert_button.isEnabled())


if __name__ == "__main__":
    unittest.main()
