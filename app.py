
import sys
import os
import shutil
import tempfile
import json
from collections import deque

# PDF Handling
import fitz  # PyMuPDF

# Word/Excel exports
from docx import Document
import openpyxl
import pandas as pd
import tempfile

# PyQt5 imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QFileDialog,
    QAction, QToolBar, QMessageBox, QLabel, QVBoxLayout,
    QScrollArea, QWidget, QLineEdit, QDialog, QHBoxLayout,
    QSpinBox, QComboBox, QPushButton, QStackedWidget,
    QColorDialog, QTextEdit, QMenu, QMenuBar, QInputDialog,
    QTableWidget, QTableWidgetItem, QTextBrowser
)
from PyQt5.QtGui import (
    QCursor, QFont, QColor, QPixmap, QImage
)
from PyQt5.QtCore import Qt


###############################################################################
# EditDialog: A dialog to choose “replace text” or “highlight text”
###############################################################################
class EditDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit PDF")
        self.setModal(True)

        layout = QVBoxLayout()

        # Page number selection
        page_layout = QHBoxLayout()
        page_layout.addWidget(QLabel("Page Number:"))
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        page_layout.addWidget(self.page_spin)
        layout.addLayout(page_layout)

        # Edit operation selection
        edit_select_layout = QHBoxLayout()
        edit_select_layout.addWidget(QLabel("Edit Operation:"))
        self.edit_combo = QComboBox()
        self.edit_combo.addItems(["Replace Text", "Highlight Text", "Delete Text"])
        edit_select_layout.addWidget(self.edit_combo)
        layout.addLayout(edit_select_layout)

        # Stacked widget for different edit operations
        self.stacked_widget = QStackedWidget()

        # Replace Text widget
        replace_widget = QWidget()
        replace_layout = QVBoxLayout()
        self.old_text = QLineEdit()
        self.old_text.setPlaceholderText("Text to replace")
        self.new_text = QLineEdit()
        self.new_text.setPlaceholderText("New text")
        replace_layout.addWidget(self.old_text)
        replace_layout.addWidget(self.new_text)
        replace_btn = QPushButton("Replace")
        replace_btn.clicked.connect(self.replace_text)
        replace_layout.addWidget(replace_btn)
        replace_widget.setLayout(replace_layout)
        self.stacked_widget.addWidget(replace_widget)

        # Highlight Text widget
        highlight_widget = QWidget()
        highlight_layout = QVBoxLayout()
        self.highlight_text = QLineEdit()
        self.highlight_text.setPlaceholderText("Text to highlight")
        highlight_layout.addWidget(self.highlight_text)
        highlight_btn = QPushButton("Highlight")
        highlight_btn.clicked.connect(self.highlight_text_action)
        highlight_layout.addWidget(highlight_btn)
        highlight_widget.setLayout(highlight_layout)
        self.stacked_widget.addWidget(highlight_widget)

        # Delete Text widget
        delete_widget = QWidget()
        delete_layout = QVBoxLayout()
        self.delete_text = QLineEdit()
        self.delete_text.setPlaceholderText("Text to delete")
        delete_layout.addWidget(self.delete_text)
        delete_btn = QPushButton("Delete")
        delete_btn.clicked.connect(self.delete_text_action)
        delete_layout.addWidget(delete_btn)
        delete_widget.setLayout(delete_layout)
        self.stacked_widget.addWidget(delete_widget)

        layout.addWidget(self.stacked_widget)

        # Connect combo box to stacked widget
        self.edit_combo.currentIndexChanged.connect(self.stacked_widget.setCurrentIndex)

        self.setLayout(layout)
        self.pdf_widget = None

    def set_pdf_widget(self, widget):
        self.pdf_widget = widget
        # Set pageSpin max
        if widget and widget.doc:
            self.page_spin.setMaximum(len(widget.doc))

    def replace_text(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        old_txt = self.old_text.text()
        new_txt = self.new_text.text()

        if old_txt.strip() and new_txt.strip():
            success = self.pdf_widget.edit_text(page_num, old_txt, new_txt)
            if success:
                QMessageBox.information(self, "Success", "Text replaced!")
            else:
                QMessageBox.warning(self, "Failure", "Text replace failed.")

    def highlight_text_action(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        text = self.highlight_text.text().strip()
        if text:
            success = self.pdf_widget.add_highlight(page_num, text)
            if success:
                QMessageBox.information(self, "Success", "Text highlighted!")
            else:
                QMessageBox.warning(self, "Failure", "Highlight failed.")

    def delete_text_action(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        text = self.delete_text.text().strip()
        if text:
            success = self.pdf_widget.delete_text(page_num, text)
            if success:
                QMessageBox.information(self, "Success", "Text deleted!")
            else:
                QMessageBox.warning(self, "Failure", "Text deletion failed.")


###############################################################################
# PageWidget: Container for each PDF page’s rendered image
###############################################################################
class PageWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.layout)

        # Create label for PDF page
        self.page_label = QLabel()
        self.page_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.page_label)

        # Context menu
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.page_label.setContextMenuPolicy(Qt.CustomContextMenu)
        self.page_label.customContextMenuRequested.connect(self.show_context_menu)

    def show_context_menu(self, position):
        # Right-click context menu for the page
        menu = QMenu(self)
        close_page_action = menu.addAction("Close This Page")

        parent = self.parent()
        while parent and not isinstance(parent, PDFViewWidget):
            parent = parent.parent()

        if parent and isinstance(parent, PDFViewWidget):
            try:
                page_index = parent.page_widgets.index(self)
                close_page_action.triggered.connect(
                    lambda: parent.remove_page(page_index)
                )
            except ValueError:
                pass

        menu.exec_(self.mapToGlobal(position))


###############################################################################
# PDFViewWidget: Displays a single PDF (multiple pages) & supports editing
###############################################################################
class PDFViewWidget(QWidget):
    """
    Displays one PDF file in multiple pages (PageWidgets).
    Allows in-place text editing where the user clicks.
    """
    def __init__(self, pdf_file_path, parent=None):
        super().__init__(parent)

        self.original_pdf_path = pdf_file_path
        self.doc = None
        self.zoom = 1.0
        self.edit_mode = False
        self.text_placement_mode = False
        self.current_page_index = None
        self.undo_stack = deque(maxlen=50)  # Store last 50 actions
        self.redo_stack = deque(maxlen=50)

        # Map from GUI font names to built-in PDF fonts (to fix "need font file" error).
        self.font_map = {
            "Arial": "helv",              # "Helvetica" base-14
            "Times New Roman": "Times-Roman",
            "Courier": "Courier",
            "Verdana": "helv",            # Using helv as fallback
        }

        # Default style: we store the user's "friendly" name, then map to base-14
        self.text_style = {
            'font': 'Arial',
            'size': 12,
            'color': (0, 0, 0)  # black in RGB (0-1)
        }

        # Create a temporary working copy of the PDF
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="pdf_viewer_")
            temp_filename = os.path.basename(pdf_file_path)
            self.working_pdf_path = os.path.join(self.temp_dir, f"working_{temp_filename}")

            shutil.copy2(pdf_file_path, self.working_pdf_path)
            self.doc = fitz.open(self.working_pdf_path)

            if not self.doc or self.doc.page_count == 0:
                raise ValueError("PDF file is invalid or has no pages.")
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Could not open PDF file:\n{str(e)}")
            self.cleanup_temp_files()
            return

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(main_layout)

        # Scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        main_layout.addWidget(self.scroll_area)

        # Container for pages
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setAlignment(Qt.AlignHCenter)
        self.content_widget.setLayout(self.content_layout)
        self.scroll_area.setWidget(self.content_widget)

        self.page_widgets = []
        self.load_pages()

        # QLineEdit for placing new text
        self.current_text_edit = QLineEdit(self)
        # Make it visible with a semi-transparent background
        self.current_text_edit.setStyleSheet("""
            QLineEdit {
                border: 1px solid #0078d7;
                background: rgba(255, 255, 255, 0.9);
                color: black;
                font-size: 14px;
                padding: 2px;
            }
        """)
        self.current_text_edit.hide()
        self.current_text_edit.returnPressed.connect(self.finish_text_edit)
        
        # Add right-click menu for text editing
        self.current_text_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        self.current_text_edit.customContextMenuRequested.connect(self.show_text_edit_menu)
        # Override focus out to finalize or discard
        orig_focus_out = self.current_text_edit.focusOutEvent

        def custom_focus_out(event):
            typed = self.current_text_edit.text().strip()
            if typed:
                self.finish_text_edit()
            else:
                self.current_text_edit.hide()
            orig_focus_out(event)

        self.current_text_edit.focusOutEvent = custom_focus_out

    def show_text_edit_menu(self, position):
        """Show context menu for text editing."""
        menu = QMenu()
        delete_action = menu.addAction("Delete")
        delete_action.triggered.connect(self.delete_current_text)
        menu.exec_(self.current_text_edit.mapToGlobal(position))

    def delete_current_text(self):
        """Delete the currently selected text with precise boundary."""
        if not hasattr(self.current_text_edit, 'original_word') or not self.current_text_edit.original_word:
            return

        try:
            page = self.doc[self.current_page_index]
            word = self.current_text_edit.original_word
            
            # Create a precise rectangle for just this word
            rect = fitz.Rect(word[0], word[1], word[2], word[3])
            
            # Add redaction with exact boundaries
            annot = page.add_redact_annot(rect)
            annot.set_rect(rect)  # Ensure rect is exact
            page.apply_redactions()
            
            self.doc.saveIncr()
            self.show_page(self.current_page_index)
            self.current_text_edit.hide()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not delete text:\n{str(e)}")

        # Zoom buttons at the bottom left
        self.zoom_layout = QHBoxLayout()
        self.zoom_in_button = QPushButton("+")
        self.zoom_in_button.clicked.connect(self.zoom_in)
        self.zoom_out_button = QPushButton("-")
        self.zoom_out_button.clicked.connect(self.zoom_out)
        self.zoom_layout.addWidget(self.zoom_out_button)
        self.zoom_layout.addWidget(self.zoom_in_button)
        main_layout.addLayout(self.zoom_layout)

    def load_pages(self):
        """Load and render all pages from the PDF doc."""
        for i in range(len(self.doc)):
            page_widget = PageWidget(self.content_widget)
            self.content_layout.addWidget(page_widget)
            self.content_layout.addSpacing(20)

            self.page_widgets.append(page_widget)
            self.show_page(i)

    def show_page(self, index):
        """Render a single PDF page and display it in the PageWidget."""
        if not self.doc:
            return
        try:
            page = self.doc.load_page(index)
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = QImage(
                pix.samples, pix.width, pix.height,
                pix.stride, QImage.Format_RGB888
            )
            pixmap = QPixmap.fromImage(img)

            page_widget = self.page_widgets[index]
            page_widget.page_label.setPixmap(pixmap)
            page_widget.page_label.adjustSize()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not load page {index+1}:\n{str(e)}")

    def remove_page(self, index):
        """Remove a page from the PDF and update the UI."""
        try:
            page_widget = self.page_widgets.pop(index)
            self.content_layout.removeWidget(page_widget)
            page_widget.deleteLater()

            self.doc.delete_page(index)
            self.doc.saveIncr()

            # Re-render the remaining pages
            for i in range(len(self.page_widgets)):
                self.show_page(i)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not remove page {index+1}:\n{str(e)}")

    def rotate_all_pages(self, degrees=90):
        """Rotate all pages in the PDF."""
        try:
            for page in self.doc:
                page.set_rotation(degrees)
            self.doc.save(self.working_pdf_path)
            for i in range(len(self.page_widgets)):
                self.show_page(i)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not rotate pages:\n{str(e)}")

    def cleanup_temp_files(self):
        """Clean up temporary files/folders."""
        try:
            if self.doc:
                self.doc.close()
            if hasattr(self, 'temp_dir') and os.path.isdir(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as ex:
            print(f"Cleanup error: {ex}")

    def closeEvent(self, event):
        self.cleanup_temp_files()
        super().closeEvent(event)

    def save_as(self, new_path):
        """Save working PDF to a new location."""
        try:
            shutil.copy2(self.working_pdf_path, new_path)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error Saving", f"Could not save:\n{str(e)}")
            return False

    # -- Export to Word/Excel (naive text-based)
    def export_to_word(self, docx_path):
        """Export text from each page into a .docx document."""
        try:
            doc = Document()
            for p in range(len(self.doc)):
                page = self.doc.load_page(p)
                text = page.get_text("text")
                doc.add_paragraph(f"--- Page {p+1} ---")
                doc.add_paragraph(text)
                doc.add_page_break()
            doc.save(docx_path)
            return True
        except Exception as e:
            QMessageBox.warning(self, "Export Failed", f"Word export failed:\n{str(e)}")
            return False

    def export_to_excel(self, xlsx_path):
        """Export text from each page into an .xlsx workbook (one sheet per page)."""
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Page1"
            for p in range(len(self.doc)):
                if p > 0:
                    ws = wb.create_sheet(title=f"Page{p+1}")
                page = self.doc.load_page(p)
                text = page.get_text("text")
                lines = text.splitlines()
                row = 1
                for line in lines:
                    ws.cell(row=row, column=1, value=line)
                    row += 1
            wb.save(xlsx_path)
            return True
        except Exception as e:
            QMessageBox.warning(self, "Export Failed", f"Excel export failed:\n{str(e)}")
            return False

    # ---------------------
    # Text Editing Methods
    # ---------------------
    def edit_text(self, page_number, old_text, new_text):
        """Search & replace text on a given page. (Simple approach)"""
        if not self.doc:
            return False
        try:
            page = self.doc[page_number]
            matches = page.search_for(old_text)
            if matches:
                # Convert user font name to base-14
                fontname = self.font_map.get(self.text_style['font'], "helv")
                for rect in matches:
                    # Redact out old text
                    page.add_redact_annot(rect)
                    page.apply_redactions()
                    # Insert new text in roughly the same position
                    x0, y0, x1, y1 = rect
                    height = y1 - y0
                    page.insert_text(
                        (x0, y0 + height * 0.8),
                        new_text,
                        fontsize=self.text_style['size'],
                        fontname=fontname,
                        color=self.text_style['color']
                    )
                self.doc.saveIncr()
                self.show_page(page_number)
                return True
            return False
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to edit text:\n{str(e)}")
            return False

    def add_highlight(self, page_number, text):
        """Highlight all occurrences of `text` on a given page."""
        if not self.doc:
            return False
        try:
            page = self.doc[page_number]
            matches = page.search_for(text)
            for rect in matches:
                annot = page.add_highlight_annot(rect)
                annot.update()
            self.doc.saveIncr()
            self.show_page(page_number)
            return True
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to highlight text:\n{str(e)}")
            return False

    def delete_text(self, page_number, text):
        """Delete specific occurrences of `text` on a given page with improved precision."""
        if not self.doc:
            return False
        try:
            page = self.doc[page_number]
            # Get exact word locations with more context
            words = page.get_text("words")
            matches = []
            
            # Find exact matches only
            for word in words:
                if word[4].strip() == text.strip():  # Compare exact text
                    rect = fitz.Rect(word[0], word[1], word[2], word[3])
                    matches.append(rect)
            
            if matches:
                for rect in matches:
                    # Add small padding to avoid overlapping with nearby text
                    rect.x0 -= 1
                    rect.x1 += 1
                    page.add_redact_annot(rect)
                page.apply_redactions()
                self.doc.saveIncr()
                self.show_page(page_number)
                return True
            return False
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to delete text:\n{str(e)}")
            return False

    def find_text(self, page_number, text):
        """Find all occurrences of text on a given page without highlighting."""
        if not self.doc:
            return False
        try:
            page = self.doc[page_number]
            matches = page.search_for(text)
            return len(matches) > 0
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to find text:\n{str(e)}")
            return False




    # ---------------------
    # Adding New Text via Mouse Click
    # ---------------------
    def set_text_style(self, style):
        """Set user's chosen font family, size, color (RGB in 0..1)."""
        self.text_style = style

    def start_edit_mode(self):
        self.edit_mode = True
        self.text_placement_mode = True
        self.setCursor(Qt.IBeamCursor)
        # Override mousePressEvent on each page
        for i, pg_widget in enumerate(self.page_widgets):
            pg_widget.page_label.setMouseTracking(True)
            pg_widget.page_label.mousePressEvent = lambda e, idx=i: self.handle_edit_click(e, idx)

    def stop_edit_mode(self):
        self.edit_mode = False
        self.text_placement_mode = False
        self.setCursor(Qt.ArrowCursor)
        # Reset mousePressEvent
        for pg_widget in self.page_widgets:
            pg_widget.page_label.mousePressEvent = None

    def show_text_edit_menu(self, position):
        """Show context menu for text editing."""
        menu = QMenu()
        delete_action = menu.addAction("Delete")
        delete_action.triggered.connect(self.delete_current_text)
        menu.exec_(self.current_text_edit.mapToGlobal(position))

    def handle_edit_click(self, event, page_idx):
        """Enhanced click handler for text editing."""
        if not self.edit_mode:
            return

        self.current_page_index = page_idx
        pos = event.pos()
        
        # Calculate PDF coordinates with proper scaling
        page = self.doc[page_idx]
        page_rect = page.rect
        scale_x = page_rect.width / self.page_widgets[page_idx].page_label.width()
        scale_y = page_rect.height / self.page_widgets[page_idx].page_label.height()
        
        pdf_x = pos.x() * scale_x
        pdf_y = pos.y() * scale_y

        # Check if there's existing text at click position
        words = page.get_text("words")
        clicked_word = None
        for word in words:
            # word tuple format: (x0, y0, x1, y1, word, block_no, line_no, word_no)
            if (word[0] <= pdf_x <= word[2]) and (word[1] <= pdf_y <= word[3]):
                clicked_word = word
                break

        page_widget = self.page_widgets[page_idx]
        global_pos = page_widget.page_label.mapToGlobal(pos)
        widget_pos = self.mapFromGlobal(global_pos)

        # Position and show text edit
        self.current_text_edit.move(widget_pos)
        if clicked_word:
            # Pre-fill with existing text if clicked on word
            self.current_text_edit.setText(clicked_word[4])  # word[4] contains the text
            self.current_text_edit.selectAll()
            # Store original word info for deletion
            self.current_text_edit.original_word = clicked_word
        else:
            self.current_text_edit.clear()
            self.current_text_edit.original_word = None

        self.current_text_edit.resize(180, 25)
        self.current_text_edit.show()
        self.current_text_edit.setFocus()
        self.current_text_edit.pdf_position = (pdf_x, pdf_y)

    def delete_current_text(self):
        """Delete the currently selected text."""
        if not hasattr(self.current_text_edit, 'original_word') or not self.current_text_edit.original_word:
            return

        try:
            page = self.doc[self.current_page_index]
            word = self.current_text_edit.original_word
            rect = fitz.Rect(word[0], word[1], word[2], word[3])
            
            # Create and apply redaction
            page.add_redact_annot(rect)
            page.apply_redactions()
            
            self.doc.saveIncr()
            self.show_page(self.current_page_index)
            self.current_text_edit.hide()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not delete text:\n{str(e)}")

    def finish_text_edit(self):
        """Enhanced method to finish text editing."""
        if not self.current_text_edit.isVisible() or self.current_page_index is None:
            return

        new_text = self.current_text_edit.text().strip()
        original_text = self.current_text_edit.original_word[4] if hasattr(self.current_text_edit, 'original_word') and self.current_text_edit.original_word else None

        # If text hasn't changed, just hide the editor without making changes
        if original_text and new_text == original_text:
            self.current_text_edit.hide()
            return

        if not new_text:
            if hasattr(self.current_text_edit, 'original_word') and self.current_text_edit.original_word:
                self.delete_current_text()
            self.current_text_edit.hide()
            return

        try:
            page = self.doc[self.current_page_index]
            
            # If replacing existing text, delete it first
            if hasattr(self.current_text_edit, 'original_word') and self.current_text_edit.original_word:
                word = self.current_text_edit.original_word
                rect = fitz.Rect(word[0], word[1], word[2], word[3])
                page.add_redact_annot(rect)
                page.apply_redactions()

            # Insert new text
            fontname = self.font_map.get(self.text_style['font'], "helv")
            page.insert_text(
                self.current_text_edit.pdf_position,
                new_text,
                fontsize=self.text_style['size'],
                fontname=fontname,
                color=self.text_style['color']
            )
            
            self.doc.saveIncr()
            self.show_page(self.current_page_index)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not insert text:\n{str(e)}")
        finally:
            self.current_text_edit.hide()

    def zoom_in(self):
        """Zoom in on the current PDF."""
        self.zoom *= 1.2
        for i in range(len(self.page_widgets)):
            self.show_page(i)

    def zoom_out(self):
        """Zoom out on the current PDF."""
        self.zoom /= 1.2
        for i in range(len(self.page_widgets)):
            self.show_page(i)


###############################################################################
# Main Window: Contains menubar, toolbars, and QTabWidget for multiple PDFs
###############################################################################
class PDFViewerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Universal File Viewer & Editor")
        self.resize(1200, 800)

        # Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.setCentralWidget(self.tab_widget)

        self.current_color = (0, 0, 0)  # default black
        self.create_menus()
        self.create_toolbars()

        # We'll keep track of whether dark mode is ON or OFF
        self.dark_mode_enabled = False

    def create_menus(self):
        menubar = self.menuBar()

        # File Menu
        file_menu = menubar.addMenu("File")

        open_action = QAction("Open PDF...", self)
        open_action.triggered.connect(self.open_pdf)
        file_menu.addAction(open_action)

        # Save As sub-menu
        save_as_menu = QMenu("Save As", self)
        save_as_pdf_action = QAction("PDF", self)
        save_as_pdf_action.triggered.connect(self.save_pdf_as)
        save_as_word_action = QAction("Word (Docx)", self)
        save_as_word_action.triggered.connect(self.save_pdf_as_word)
        save_as_excel_action = QAction("Excel (XLSX)", self)
        save_as_excel_action.triggered.connect(self.save_pdf_as_excel)
        save_as_menu.addAction(save_as_pdf_action)
        save_as_menu.addAction(save_as_word_action)
        save_as_menu.addAction(save_as_excel_action)
        file_menu.addMenu(save_as_menu)

        close_action = QAction("Close Current File", self)
        close_action.triggered.connect(self.close_current_file)
        file_menu.addAction(close_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Edit Menu
        edit_menu = menubar.addMenu("Edit")

        edit_pdf_action = QAction("Edit PDF...", self)
        edit_pdf_action.triggered.connect(self.edit_pdf)
        edit_menu.addAction(edit_pdf_action)

        rotate_action = QAction("Rotate All Pages 90°", self)
        rotate_action.triggered.connect(lambda: self.rotate_all(90))
        edit_menu.addAction(rotate_action)

        # View Menu (Dark Mode Toggle)
        view_menu = menubar.addMenu("View")
        dark_mode_action = QAction("Toggle Dark Mode", self)
        dark_mode_action.triggered.connect(self.toggle_dark_mode)
        view_menu.addAction(dark_mode_action)

    def create_toolbars(self):
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)

        # Toggle text edit mode
        self.edit_mode_action = QAction("Text Edit Mode", self)
        self.edit_mode_action.setCheckable(True)
        self.edit_mode_action.triggered.connect(self.toggle_edit_mode)
        toolbar.addAction(self.edit_mode_action)

        # Font combo
        self.font_combo = QComboBox()
        # Provide the same keys as our PDFViewWidget.font_map
        self.font_combo.addItems(["Arial", "Times New Roman", "Courier", "Verdana"])
        self.font_combo.setCurrentText("Arial")
        self.font_combo.currentIndexChanged.connect(self.update_text_style)
        toolbar.addWidget(self.font_combo)

        # Font size combo
        self.size_combo = QComboBox()
        sizes = ["8","9","10","11","12","14","16","18","20","22","24","28","32","36","48","72"]
        self.size_combo.addItems(sizes)
        self.size_combo.setCurrentText("12")
        self.size_combo.currentTextChanged.connect(self.update_text_style)
        toolbar.addWidget(self.size_combo)

        # Text color combo
        self.color_combo = QComboBox()
        self.color_combo.addItems(["Black", "Red", "Blue", "Green", "Custom..."])
        self.color_combo.currentTextChanged.connect(self.handle_color_selection)
        toolbar.addWidget(self.color_combo)

        # Add a "Find Text" button
        find_action = QAction("Find Text", self)
        find_action.triggered.connect(self.find_text_dialog)
        toolbar.addAction(find_action)


    # ---------------------------------------------------------------------
    # FILE MENU ACTIONS
    # ---------------------------------------------------------------------
    def open_pdf(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Open File(s)",
            "",
            "All Supported Files (*.pdf *.xlsx *.xls *.csv *.json *.txt);;PDF Files (*.pdf);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;JSON Files (*.json);;Text Files (*.txt);;All Files (*)",
            options=options
        )
        if files:
            for f in files:
                self.add_file_tab(f)

    def add_file_tab(self, file_path):
        try:
            extension = os.path.splitext(file_path)[1].lower()
            viewer = None
            
            if extension == '.pdf':
                viewer = PDFViewWidget(file_path)
            elif extension in ['.xlsx', '.xls']:
                viewer = self.create_excel_viewer(file_path)
            elif extension == '.csv':
                viewer = self.create_csv_viewer(file_path)
            elif extension == '.json':
                viewer = self.create_json_viewer(file_path)
            elif extension == '.txt':
                viewer = self.create_text_viewer(file_path)
            else:
                QMessageBox.warning(self, "Unsupported Format", "This file format is not supported.")
                return

            if viewer:
                filename = os.path.basename(file_path)
                self.tab_widget.addTab(viewer, filename)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open file:\n{str(e)}")


    def create_excel_viewer(self, file_path):
        """Create an Excel viewer widget that supports images."""
        try:
            # Create a container widget
            container = QWidget()
            layout = QVBoxLayout(container)
            
            # Create table for data
            table = QTableWidget()
            df = pd.read_excel(file_path)
            table.setRowCount(df.shape[0])
            table.setColumnCount(df.shape[1])
            table.setHorizontalHeaderLabels(df.columns)

            # Load data and handle potential embedded images
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            
            for row in range(df.shape[0]):
                for col in range(df.shape[1]):
                    cell = ws.cell(row=row+1, column=col+1)
                    
                    # Check if cell contains image
                    if cell._hyperlink is not None and cell._hyperlink.target:
                        try:
                            image_path = cell._hyperlink.target
                            label = QLabel()
                            pixmap = QPixmap(image_path)
                            label.setPixmap(pixmap.scaled(100, 100, Qt.KeepAspectRatio))
                            table.setCellWidget(row, col, label)
                        except:
                            item = QTableWidgetItem(str(df.iloc[row, col]))
                            table.setItem(row, col, item)
                    else:
                        item = QTableWidgetItem(str(df.iloc[row, col]))
                        table.setItem(row, col, item)

            layout.addWidget(table)
            return container
        except Exception as e:
            QMessageBox.warning(self, "Excel Error", f"Could not read Excel file:\n{str(e)}")
            return None

    def create_csv_viewer(self, file_path):
        """Create a CSV viewer widget."""
        try:
            df = pd.read_csv(file_path)
            table = QTableWidget()
            table.setRowCount(df.shape[0])
            table.setColumnCount(df.shape[1])
            table.setHorizontalHeaderLabels(df.columns)

            for row in range(df.shape[0]):
                for col in range(df.shape[1]):
                    item = QTableWidgetItem(str(df.iloc[row, col]))
                    table.setItem(row, col, item)

            return table
        except Exception as e:
            QMessageBox.warning(self, "CSV Error", f"Could not read CSV file:\n{str(e)}")
            return None

    def create_json_viewer(self, file_path):
        """Create a JSON viewer widget."""
        try:
            with open(file_path, 'r') as f:
                data = json.load(f)
            
            text_browser = QTextBrowser()
            text_browser.setPlainText(json.dumps(data, indent=2))
            return text_browser
        except Exception as e:
            QMessageBox.warning(self, "JSON Error", f"Could not read JSON file:\n{str(e)}")
            return None

    def create_text_viewer(self, file_path):
        """Create a text viewer widget."""
        try:
            with open(file_path, 'r') as f:
                text = f.read()
            
            text_browser = QTextBrowser()
            text_browser.setPlainText(text)
            return text_browser
        except Exception as e:
            QMessageBox.warning(self, "Text Error", f"Could not read text file:\n{str(e)}")
            return None

    def create_word_viewer(self, file_path):
        """Create a Word document viewer that properly handles text and images."""
        try:
            # Create scrollable text widget
            text_widget = QTextEdit()
            text_widget.setReadOnly(True)
            
            # Load the document
            doc = Document(file_path)
            
            # Create a temporary directory for images
            temp_dir = tempfile.mkdtemp()
            
            try:
                # Process paragraphs and images
                for paragraph in doc.paragraphs:
                    text_widget.append(paragraph.text)
                    
                    # Handle inline images
                    for run in paragraph.runs:
                        if 'w:drawing' in run._element.xml or 'w:pict' in run._element.xml:
                            # Extract image
                            for element in run._element.findall('.//v:imagedata', run._element.nsmap):
                                image_rid = element.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                if image_rid:
                                    image_part = doc.part.related_parts[image_rid]
                                    image_bytes = image_part.blob
                                    
                                    # Save image temporarily
                                    img_path = os.path.join(temp_dir, f"img_{image_rid}.png")
                                    with open(img_path, 'wb') as img_file:
                                        img_file.write(image_bytes)
                                    
                                    # Insert image into text widget
                                    cursor = text_widget.textCursor()
                                    image_format = text_widget.textCursor().charFormat()
                                    image = QImage(img_path)
                                    if not image.isNull():
                                        # Scale image if too large
                                        if image.width() > 800:
                                            image = image.scaledToWidth(800, Qt.SmoothTransformation)
                                        text_widget.document().addResource(
                                            QTextDocument.ImageResource,
                                            img_path,
                                            image
                                        )
                                        cursor.insertImage(img_path)
                                        text_widget.append("")  # Add newline after image
                
                return text_widget
                
            finally:
                # Cleanup temporary files
                try:
                    shutil.rmtree(temp_dir)
                except:
                    pass
                    
        except Exception as e:
            QMessageBox.warning(self, "Word Error", f"Could not read Word file:\n{str(e)}")
            return None

    def close_tab(self, index):
        widget = self.tab_widget.widget(index)
        if widget:
            if hasattr(widget, 'cleanup_temp_files'):
                widget.cleanup_temp_files()
            self.tab_widget.removeTab(index)

    def close_current_file(self):
        idx = self.tab_widget.currentIndex()
        if idx >= 0:
            self.close_tab(idx)

    def save_pdf_as(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget or not hasattr(current_widget, "save_as"):
            QMessageBox.warning(self, "No File", "Open a file first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save As...", "", 
            "PDF Files (*.pdf);;Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json);;Text Files (*.txt);;All Files (*)"
        )
        if file_path:
            try:
                if isinstance(current_widget, PDFViewWidget):
                    success = current_widget.save_as(file_path)
                elif isinstance(current_widget, QTableWidget):
                    # For Excel and CSV
                    model = current_widget.model()
                    df = pd.DataFrame(
                        [[current_widget.item(row, col).text() for col in range(current_widget.columnCount())] 
                         for row in range(current_widget.rowCount())],
                        columns=[current_widget.horizontalHeaderItem(col).text() for col in range(current_widget.columnCount())]
                    )
                    
                    ext = os.path.splitext(file_path)[1].lower()
                    if ext in ['.xlsx', '.xls']:
                        df.to_excel(file_path, index=False)
                    elif ext == '.csv':
                        df.to_csv(file_path, index=False)
                    success = True
                elif isinstance(current_widget, QTextBrowser):
                    # For JSON and Text files
                    with open(file_path, 'w') as f:
                        f.write(current_widget.toPlainText())
                    success = True
                else:
                    success = False

                if success:
                    QMessageBox.information(self, "Saved", "File saved successfully!")
                else:
                    QMessageBox.warning(self, "Save Failed", "Could not save the file.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not save file:\n{str(e)}")

    def save_pdf_as_word(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget or not hasattr(current_widget, "export_to_word"):
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save As Word (Docx)", "", "Word Documents (*.docx);;All Files (*)"
        )
        if file_path:
            success = current_widget.export_to_word(file_path)
            if success:
                QMessageBox.information(self, "Exported", "Exported to Word successfully!")

    def save_pdf_as_excel(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget or not hasattr(current_widget, "export_to_excel"):
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save As Excel (XLSX)", "", "Excel Files (*.xlsx);;All Files (*)"
        )
        if file_path:
            success = current_widget.export_to_excel(file_path)
            if success:
                QMessageBox.information(self, "Exported", "Exported to Excel successfully!")

    # ---------------------------------------------------------------------
    # EDIT MENU ACTIONS
    # ---------------------------------------------------------------------
    def edit_pdf(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return

        dlg = EditDialog(self)
        dlg.set_pdf_widget(current_widget)
        dlg.exec_()

    def rotate_all(self, degrees):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget or not hasattr(current_widget, "rotate_all_pages"):
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return
        current_widget.rotate_all_pages(degrees)

    # ---------------------------------------------------------------------
    # VIEW MENU ACTIONS
    # ---------------------------------------------------------------------
    def toggle_dark_mode(self):
        """Toggle dark mode for the entire app."""
        self.dark_mode_enabled = not self.dark_mode_enabled
        if self.dark_mode_enabled:
            self.apply_dark_theme()
        else:
            self.apply_light_theme()

    def apply_dark_theme(self):
        dark_style = """
            QMainWindow {
                background-color: #2b2b2b;
            }
            QWidget {
                color: #dddddd;
                background-color: #3c3c3c;
            }
            QLabel {
                color: #ffffff;
            }
            QLineEdit, QTextEdit, QSpinBox, QComboBox {
                background-color: #2b2b2b;
                color: #ffffff;
                border: 1px solid #5a5a5a;
            }
            QToolBar {
                background-color: #3c3c3c;
            }
            QMenuBar {
                background-color: #3c3c3c;
            }
            QMenuBar::item {
                background: #3c3c3c;
                color: #ffffff;
            }
            QMenu {
                background-color: #3c3c3c;
                color: #ffffff;
            }
            QPushButton {
                background-color: #5a5a5a;
                color: #ffffff;
            }
        """
        self.setStyleSheet(dark_style)

    def apply_light_theme(self):
        self.setStyleSheet("")

    # ---------------------------------------------------------------------
    # TOOLBAR ACTIONS
    # ---------------------------------------------------------------------
    def toggle_edit_mode(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            self.edit_mode_action.setChecked(False)
            return

        if self.edit_mode_action.isChecked():
            current_widget.start_edit_mode()
        else:
            current_widget.stop_edit_mode()

    def update_text_style(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            return

        font_family = self.font_combo.currentText()
        try:
            size = float(self.size_combo.currentText())
        except ValueError:
            size = 12
        color = self.current_color

        style = {
            'font': font_family,
            'size': size,
            'color': color
        }
        if hasattr(current_widget, "set_text_style"):
            current_widget.set_text_style(style)

    def handle_color_selection(self, color_name):
        color_map = {
            "Black": (0, 0, 0),
            "Red": (1, 0, 0),
            "Blue": (0, 0, 1),
            "Green": (0, 1, 0)
        }
        if color_name == "Custom...":
            c = QColorDialog.getColor()
            if c.isValid():
                self.current_color = (c.red() / 255, c.green() / 255, c.blue() / 255)
                new_label = f"Custom ({c.name()})"
                if self.color_combo.findText(new_label) == -1:
                    self.color_combo.insertItem(0, new_label)
                self.color_combo.setCurrentText(new_label)
                color_map[new_label] = self.current_color
        else:
            self.current_color = color_map.get(color_name, (0, 0, 0))
        self.update_text_style()

    def find_text_dialog(self):
        """Simple dialog to search text in the current PDF (without highlighting)."""
        current_widget = self.tab_widget.currentWidget()
        if not current_widget or not current_widget.doc:
            QMessageBox.warning(self, "No PDF", "Open a PDF first.")
            return

        text, ok = QInputDialog.getText(self, "Find Text", "Enter text to find:")
        if ok and text.strip():
            found_pages = []
            for page_idx in range(current_widget.doc.page_count):
                if current_widget.find_text(page_idx, text.strip()):
                    found_pages.append(str(page_idx + 1))
            
            if found_pages:
                QMessageBox.information(
                    self, 
                    "Found", 
                    f"Text found on page(s): {', '.join(found_pages)}"
                )
            else:
                QMessageBox.information(self, "Not Found", "Text not found in document.")





def main():
    app = QApplication(sys.argv)
    viewer = PDFViewerApp()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
