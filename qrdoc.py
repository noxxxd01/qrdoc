"""
PDF Viewer with thumbnails, bulk QR placement, and large-PDF optimizations

Dependencies:
    pip install PyQt5 pymupdf Pillow qrcode[pil]

Run:
    python pdf_viewer.py

What changed:
- Asynchronous thumbnail generation (QThread) with a floating progress dialog: "Generating thumbnails..."
- Export uses a modal QProgressDialog (blocking) with Cancel option; export loop checks for cancel and aborts cleanly.
- Prompts the user to skip thumbnails automatically if the PDF has many pages.
- Thumbnails are not all kept in memory permanently; they are created and added to the UI then released.

Usage:
- Paste links (one per line) in the right dock.
- Click "Bulk QR Code Create" and draw a rectangle on the visible page.
- Click "Export" to save a new PDF with QRs embedded. During export you can cancel.

"""

import sys
import io
from functools import partial

from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QScrollArea,
    QFrame,
    QSizePolicy,
    QSlider,
    QTextEdit,
    QMessageBox,
    QProgressDialog,
)
from PyQt5.QtGui import QPixmap, QImage, QPainter, QPen
from PyQt5.QtCore import Qt, QRect, QThread, pyqtSignal

import fitz  # PyMuPDF
from PIL import Image
import qrcode


def pixmap_from_fitz_page(page, zoom=1.0):
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    mode = "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    qt_pix = QPixmap()
    qt_pix.loadFromData(buf.getvalue(), format=b"PNG")
    return qt_pix


class ThumbnailWorker(QThread):
    progress = pyqtSignal(int, int)  # current, total
    produced = pyqtSignal(int, QPixmap)  # index, pixmap
    finished_signal = pyqtSignal()

    def __init__(self, doc, thumb_max_height=120, parent=None):
        super().__init__(parent)
        self.doc = doc
        self.thumb_max_height = thumb_max_height
        self._running = True

    def run(self):
        total = self.doc.page_count
        for i in range(total):
            if not self._running:
                break
            try:
                page = self.doc.load_page(i)
                rect = page.rect
                base_height = rect.height
                zoom = (self.thumb_max_height / base_height) if base_height > 0 else 0.2
                thumb_pix = pixmap_from_fitz_page(page, zoom=zoom)
            except Exception:
                thumb_pix = QPixmap(80, self.thumb_max_height)
                thumb_pix.fill(Qt.lightGray)

            self.produced.emit(i, thumb_pix)
            self.progress.emit(i + 1, total)

        self.finished_signal.emit()

    def stop(self):
        self._running = False


class SelectableLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._start = None
        self._end = None
        self.selection = None
        self._dragging = False
        self._drag_offset = (0, 0)

    def mousePressEvent(self, event):
        if not self.pixmap():
            return
        if event.button() == Qt.LeftButton:
            if self.selection and self._point_in_selection(event.pos()):
                # Start dragging
                self._dragging = True
                nx, ny, nw, nh = self.selection
                pw, ph = self.pixmap().width(), self.pixmap().height()
                x = int(nx * pw)
                y = int(ny * ph)
                self._drag_offset = (event.pos().x() - x, event.pos().y() - y)
            else:
                # Start drawing new rectangle
                self._start = event.pos()
                self._end = self._start
                self.selection = None
            self.update()

    def mouseMoveEvent(self, event):
        if self._dragging:
            # Move the selection
            nx, ny, nw, nh = self.selection
            pw, ph = self.pixmap().width(), self.pixmap().height()
            w, h = int(nw * pw), int(nh * ph)
            new_x = event.pos().x() - self._drag_offset[0]
            new_y = event.pos().y() - self._drag_offset[1]

            # clamp to pixmap boundaries
            new_x = max(0, min(new_x, pw - w))
            new_y = max(0, min(new_y, ph - h))

            self.selection = (new_x / pw, new_y / ph, nw, nh)
            self.update()

        elif self._start is not None:
            self._end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if self._dragging:
            self._dragging = False
            self.update()
        elif self._start is not None:
            self._end = event.pos()
            self._finalize_selection()
            self._start = None
            self._end = None
            self.update()

    def _point_in_selection(self, point):
        if self.selection is None:
            return False
        nx, ny, nw, nh = self.selection
        pw, ph = self.pixmap().width(), self.pixmap().height()
        x = int(nx * pw)
        y = int(ny * ph)
        w = int(nw * pw)
        h = int(nh * ph)
        return x <= point.x() <= x + w and y <= point.y() <= y + h

    def _finalize_selection(self):
        if self._start is None or self._end is None or not self.pixmap():
            self.selection = None
            return

        pw, ph = self.pixmap().width(), self.pixmap().height()
        x1 = max(0, min(self._start.x(), self._end.x()))
        y1 = max(0, min(self._start.y(), self._end.y()))
        x2 = max(0, max(self._start.x(), self._end.x()))
        y2 = max(0, max(self._start.y(), self._end.y()))

        # enforce square selection
        size = min(x2 - x1, y2 - y1)
        if size <= 2:
            self.selection = None
            return

        self.selection = (x1 / pw, y1 / ph, size / pw, size / ph)
    
    def paintEvent(self, event):
        super().paintEvent(event)

        if not self.pixmap():
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw currently drawn rectangle
        if self._start and self._end:
            pen = QPen(Qt.red, 2, Qt.SolidLine)
            painter.setPen(pen)
            brush = Qt.transparent
            painter.setBrush(brush)
            rect = QRect(self._start, self._end)
            painter.drawRect(rect.normalized())

        # Draw finalized selection
        if self.selection:
            nx, ny, nw, nh = self.selection
            pw, ph = self.pixmap().width(), self.pixmap().height()
            x = int(nx * pw)
            y = int(ny * ph)
            w = int(nw * pw)
            h = int(nh * ph)
            pen = QPen(Qt.green, 2, Qt.SolidLine)
            painter.setPen(pen)
            painter.setBrush(Qt.transparent)
            painter.drawRect(x, y, w, h)



class PDFViewer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("QRDoc - PDF Viewer with Bulk QR Placement")
        self.resize(1200, 800)

        self.doc = None
        self.current_page_index = 0
        self.zoom = 1.0

        central = QWidget()
        self.setCentralWidget(central)
        main_h = QHBoxLayout(central)

        # Left viewer
        left = QWidget()
        left_v = QVBoxLayout(left)

        controls = QHBoxLayout()
        btn_open = QPushButton("Open PDF")
        btn_open.clicked.connect(self.open_pdf)
        controls.addWidget(btn_open)

        btn_prev = QPushButton("Prev Page")
        btn_prev.clicked.connect(self.prev_page)
        controls.addWidget(btn_prev)

        btn_next = QPushButton("Next Page")
        btn_next.clicked.connect(self.next_page)
        controls.addWidget(btn_next)

        btn_zoom_in = QPushButton("Zoom +")
        btn_zoom_in.clicked.connect(self.zoom_in)
        controls.addWidget(btn_zoom_in)

        btn_zoom_out = QPushButton("Zoom -")
        btn_zoom_out.clicked.connect(self.zoom_out)
        controls.addWidget(btn_zoom_out)

        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setRange(25, 400)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(self.zoom_slider_changed)
        controls.addWidget(self.zoom_slider)

        left_v.addLayout(controls)

        # Main page display
        self.page_scroll = QScrollArea()
        self.page_scroll.setWidgetResizable(True)
        self.page_label = SelectableLabel()
        self.page_label.setAlignment(Qt.AlignCenter)
        self.page_label.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        page_container = QWidget()
        page_layout = QVBoxLayout(page_container)
        page_layout.addWidget(self.page_label, alignment=Qt.AlignCenter)

        self.page_scroll.setWidget(page_container)
        left_v.addWidget(self.page_scroll, stretch=1)

        # Thumbnail strip with placeholder
        self.thumbs_container = QWidget()
        self.thumbs_hbox = QHBoxLayout(self.thumbs_container)
        self.thumbs_hbox.setContentsMargins(5, 5, 5, 5)
        self.thumbs_hbox.setSpacing(5)

        self.thumbs_scroll = QScrollArea()
        self.thumbs_scroll.setWidgetResizable(True)
        self.thumbs_scroll.setFixedHeight(140)
        self.thumbs_scroll.setWidget(self.thumbs_container)

        left_v.addWidget(self.thumbs_scroll)

        main_h.addWidget(left, stretch=3)

        # Right dock
        dock = QFrame()
        dock.setFrameShape(QFrame.StyledPanel)
        dock.setFixedWidth(350)
        dock_v = QVBoxLayout(dock)

        dock_v.addWidget(QLabel("Paste links (one per line):"))
        self.links_text = QTextEdit()
        dock_v.addWidget(self.links_text)

        self.btn_bulk = QPushButton("Bulk QR Code Create")
        self.btn_bulk.clicked.connect(self.bulk_create_prompt)
        dock_v.addWidget(self.btn_bulk)

        self.btn_export = QPushButton("Export (Save modified PDF)")
        self.btn_export.clicked.connect(self.export_pdf)
        self.btn_export.setEnabled(False)
        dock_v.addWidget(self.btn_export)

        dock_v.addStretch()

        main_h.addWidget(dock, stretch=1)

        # state
        self.selection = None
        self.qr_images = []
        self.thumb_worker = None
        self.thumb_dialog = None

    # ---------------------- PDF Loading & Rendering ----------------------
    def open_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open PDF", "", "PDF Files (*.pdf);;All Files (*)")
        if not path:
            return
        try:
            self.doc = fitz.open(path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error opening PDF:\n{e}")
            return

        self.current_page_index = 0
        self.zoom = self.zoom_slider.value() / 100.0
        self.selection = None
        self.qr_images = []
        self.btn_export.setEnabled(False)
        self.render_current_page()

        # If doc large, prompt to skip thumbnails
        if self.doc.page_count > 200:
            resp = QMessageBox.question(
                self,
                "Large PDF",
                f"This PDF has {self.doc.page_count} pages. Generating thumbnails may be slow.\n\n"
                "Do you want to skip thumbnail generation?\n(Press No to generate thumbnails in background)",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes,
            )
            if resp == QMessageBox.Yes:
                # clear existing thumbs
                for i in reversed(range(self.thumbs_hbox.count())):
                    w = self.thumbs_hbox.itemAt(i).widget()
                    if w:
                        w.setParent(None)
                return

        # otherwise, generate thumbnails asynchronously with floating dialog
        self.start_thumbnail_worker()

    def render_current_page(self):
        if not self.doc:
            self.page_label.setText("No document loaded")
            return
        page = self.doc.load_page(self.current_page_index)
        pix = pixmap_from_fitz_page(page, zoom=self.zoom)
        self.page_label.setPixmap(pix)
        self.page_label.resize(pix.width(), pix.height())
        if self.page_label.selection != self.selection:
            self.page_label.selection = self.selection
            self.page_label.update()

    def start_thumbnail_worker(self):
        # clear existing
        for i in reversed(range(self.thumbs_hbox.count())):
            w = self.thumbs_hbox.itemAt(i).widget()
            if w:
                w.setParent(None)

        self.thumb_dialog = QProgressDialog("Generating thumbnails...", "Cancel", 0, self.doc.page_count, self)
        # non-modal floating dialog
        self.thumb_dialog.setWindowModality(Qt.NonModal)
        self.thumb_dialog.setWindowTitle("Generating thumbnails")
        self.thumb_dialog.show()

        self.thumb_worker = ThumbnailWorker(self.doc, thumb_max_height=120)
        self.thumb_worker.produced.connect(self.on_thumb_produced)
        self.thumb_worker.progress.connect(self.on_thumb_progress)
        self.thumb_worker.finished_signal.connect(self.on_thumb_finished)
        self.thumb_worker.start()

    def on_thumb_produced(self, index, pix):
        # Build thumbnail widget quickly and release pixmap reference if needed
        lbl = QLabel()
        lbl.setPixmap(pix)
        lbl.setToolTip(f"Page {index + 1}")
        lbl.mousePressEvent = partial(self.on_thumb_click, index=index)

        container = QFrame()
        container.setFrameShape(QFrame.StyledPanel)
        c_layout = QVBoxLayout(container)
        c_layout.setContentsMargins(2, 2, 2, 2)
        c_layout.addWidget(lbl)
        self.thumbs_hbox.addWidget(container)

        # allow Qt to free the pix variable when out of scope

    def on_thumb_progress(self, current, total):
        if self.thumb_dialog:
            self.thumb_dialog.setValue(current)
            if self.thumb_dialog.wasCanceled():
                if self.thumb_worker:
                    self.thumb_worker.stop()

    def on_thumb_finished(self):
        if self.thumb_dialog:
            self.thumb_dialog.close()
            self.thumb_dialog = None
        self.thumb_worker = None
        self.thumbs_hbox.addStretch()

    def build_thumbnails(self):
        # kept for API compatibility; use start_thumbnail_worker instead
        self.start_thumbnail_worker()

    def on_thumb_click(self, event, index):
        self.current_page_index = index
        self.render_current_page()

    def prev_page(self):
        if not self.doc:
            return
        self.current_page_index = max(0, self.current_page_index - 1)
        self.render_current_page()

    def next_page(self):
        if not self.doc:
            return
        self.current_page_index = min(self.doc.page_count - 1, self.current_page_index + 1)
        self.render_current_page()

    # ---------------------- Zoom controls ----------------------
    def zoom_in(self):
        val = self.zoom_slider.value()
        val = min(val + 25, self.zoom_slider.maximum())
        self.zoom_slider.setValue(val)

    def zoom_out(self):
        val = self.zoom_slider.value()
        val = max(val - 25, self.zoom_slider.minimum())
        self.zoom_slider.setValue(val)

    def zoom_slider_changed(self, value):
        self.zoom = value / 100.0
        self.render_current_page()

    # ---------------------- Bulk QR Creation Flow ----------------------
    def bulk_create_prompt(self):
        if not self.doc:
            QMessageBox.warning(self, "No PDF", "Please open a PDF first.")
            return

        text = self.links_text.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "No links", "Please paste one or more links (one per line).")
            return
        links = [line.strip() for line in text.splitlines() if line.strip()]
        num_links = len(links)
        num_pages = self.doc.page_count

        if num_links != num_pages:
            resp = QMessageBox.question(
                self,
                "Link/Page count mismatch",
                f"The number of links ({num_links}) does not match the number of PDF pages ({num_pages}).\n\n"
                "Proceed and place QR codes for the first min(links,pages)?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if resp == QMessageBox.No:
                return

        count = min(num_links, num_pages)
        self.qr_images = []
        for i in range(count):
            qr = qrcode.make(links[i])
            if qr.mode != 'RGB':
                qr = qr.convert('RGB')
            self.qr_images.append(qr)

        QMessageBox.information(self, "Place QR", "Now drag a rectangle on the PDF page to choose where QR codes should be placed on each page.")

        self.selection = None
        self.page_label.selection = None
        self.page_label.update()

        # watch selection and enable export when set
        from PyQt5.QtCore import QTimer

        def check_selection():
            if self.page_label.selection is not None:
                self.selection = self.page_label.selection
                self.btn_export.setEnabled(True)
                timer.stop()

        timer = QTimer(self)
        timer.setInterval(300)
        timer.timeout.connect(check_selection)
        timer.start()

    # ---------------------- Export (write PDF with embedded QRs) ----------------------
    def export_pdf(self):
        if not self.doc:
            QMessageBox.warning(self, "No PDF", "Please open a PDF first.")
            return
        if not self.qr_images or not self.selection:
            QMessageBox.warning(self, "Incomplete", "Please run Bulk QR Code Create and draw a rectangle on the page first.")
            return

        out_path, _ = QFileDialog.getSaveFileName(self, "Save PDF as", "", "PDF Files (*.pdf);;All Files (*)")
        if not out_path:
            return

        # create a copy of the original by reopening the file path if available; otherwise copy pages
        doc = None
        if hasattr(self.doc, 'name') and getattr(self.doc, 'name'):
            try:
                doc = fitz.open(self.doc.name)
            except Exception:
                doc = fitz.open()
                for p in self.doc:
                    doc.insert_pdf(self.doc, from_page=p.number, to_page=p.number)
        else:
            doc = fitz.open()
            for p in self.doc:
                doc.insert_pdf(self.doc, from_page=p.number, to_page=p.number)

        count = min(len(self.qr_images), self.doc.page_count)

        # modal progress dialog with cancel
        pdlg = QProgressDialog("Embedding QR codes...", "Cancel", 0, count, self)
        pdlg.setWindowModality(Qt.ApplicationModal)
        pdlg.setWindowTitle("Exporting PDF")
        pdlg.show()

        try:
            border_ratio = 0.05
            for i in range(count):
                if pdlg.wasCanceled():
                    QMessageBox.information(self, "Cancelled", "Export cancelled by user.")
                    return

                page = doc.load_page(i)
                rect = page.rect  # PDF page coordinates

                # Use self.selection (normalized 0-1) to map to PDF page coordinates
                nx, ny, nw, nh = self.selection  # <-- must be here, inside the loop
                x0 = rect.x0 + nx * rect.width
                y0 = rect.y0 + ny * rect.height
                x1 = x0 + nw * rect.width
                y1 = y0 + nh * rect.height

                # Apply border
                border_x = (x1 - x0) * border_ratio
                border_y = (y1 - y0) * border_ratio
                x0 += border_x
                y0 += border_y
                x1 -= border_x
                y1 -= border_y

                qr_img = self.qr_images[i]
                buf = io.BytesIO()
                qr_img.save(buf, format="PNG")
                img_bytes = buf.getvalue()

                page.insert_image(fitz.Rect(x0, y0, x1, y1), stream=img_bytes)

                pdlg.setValue(i + 1)

            doc.save(out_path)
            QMessageBox.information(self, "Saved", f"Saved new PDF with QR codes to:\n{out_path}")

            self.selection = None
            self.page_label.selection = None
            self.qr_images = []
            self.page_label.update()
            self.btn_export.setEnabled(False)
            
        except Exception as e:
            QMessageBox.critical(self, "Export failed", f"Failed during export:\n{e}")


def main():
    app = QApplication(sys.argv)
    viewer = PDFViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
