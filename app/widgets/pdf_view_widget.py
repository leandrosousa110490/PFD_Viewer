import fitz
import os
import shutil
from PyQt5.QtWidgets import (QWidget, QLabel, QVBoxLayout, QScrollArea, 
                            QMenu, QMessageBox, QTabWidget, QInputDialog, QLineEdit)
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt
import tempfile

class PageWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.layout)

        # Create label for PDF page
        self.page_label = QLabel()
        self.page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.page_label)
        
        # Enable mouse tracking for manual edit
        self.page_label.setMouseTracking(True)
        self.page_label.mousePressEvent = self.on_mouse_press
        
        # Enable context menu
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.page_label.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.page_label.customContextMenuRequested.connect(self.show_context_menu)

    def on_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            # Find PDFViewWidget parent
            parent = self.parent()
            while parent and not isinstance(parent, PDFViewWidget):
                parent = parent.parent()
                
            if parent:
                try:
                    page_number = parent.page_widgets.index(self)
                    # Convert click position to PDF coordinates
                    pos = event.pos()
                    scale = 1.0 / parent.zoom
                    pdf_x = pos.x() * scale
                    pdf_y = pos.y() * scale
                    text, position = parent.start_manual_edit(page_number, (pdf_x, pdf_y))
                    
                    # Show edit dialog
                    new_text, ok = QInputDialog.getText(
                        self,
                        "Edit Text",
                        "Edit text at this position:",
                        QLineEdit.Normal,
                        text
                    )
                    if ok and new_text:
                        parent.add_text(page_number, new_text, position)
                except ValueError:
                    pass

    def show_context_menu(self, position):
        menu = QMenu(self)
        close_page_action = menu.addAction("Close Page")
        
        # Find the PDFViewWidget parent
        parent = self.parent()
        while parent and not isinstance(parent, PDFViewWidget):
            parent = parent.parent()
            
        if parent:
            try:
                page_number = parent.page_widgets.index(self)
                close_page_action.triggered.connect(lambda: parent.remove_page(page_number))
                menu.exec_(self.mapToGlobal(position))
            except ValueError:
                pass

class PDFViewWidget(QWidget):
    def __init__(self, pdf_file_path, parent=None):
        super().__init__(parent)
        self.original_pdf_path = pdf_file_path
        
        # Create a temporary working copy of the PDF
        try:
            # Create temp directory if it doesn't exist
            self.temp_dir = tempfile.mkdtemp(prefix="pdf_viewer_")
            
            # Create working copy name
            temp_filename = f"working_copy_{os.path.basename(pdf_file_path)}"
            self.working_pdf_path = os.path.join(self.temp_dir, temp_filename)
            
            # Copy original to working copy
            shutil.copy2(self.original_pdf_path, self.working_pdf_path)
            
            # Open the working copy
            self.doc = fitz.open(self.working_pdf_path)
            
            # Additional validation
            if not self.doc:
                raise ValueError("Could not open PDF file")
            
            if not isinstance(self.doc, fitz.Document):
                raise ValueError("Not a valid PDF document")
                
            if self.doc.page_count <= 0:
                raise ValueError("PDF file has no pages")
                
        except Exception as e:
            self.cleanup_temp_files()
            QMessageBox.critical(None, "Error Opening PDF", 
                               f"Could not open PDF file:\n{str(e)}")
            self.doc = None
            return

        # Create main layout
        self.layout = QVBoxLayout()
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(self.layout)

        # Create scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Create content widget for scroll area
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout()
        self.content_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.content_widget.setLayout(self.content_layout)

        # Create page widgets for each page
        self.page_widgets = []
        self.zoom = 1.5
        
        # Only load pages if we have a valid document
        if self.doc and self.doc.page_count > 0:
            self.load_all_pages()

        # Set content widget to scroll area and add to main layout
        self.scroll_area.setWidget(self.content_widget)
        self.layout.addWidget(self.scroll_area)

    def show_page(self, page_number):
        if not self.doc:
            return
            
        try:
            page = self.doc.load_page(page_number)
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)  # Disable alpha to avoid format issues
            
            # Use RGB format consistently
            qimage = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimage)
            
            self.page_widgets[page_number].page_label.setPixmap(pixmap)
            self.page_widgets[page_number].page_label.adjustSize()
        except Exception as e:
            QMessageBox.warning(None, "Error Loading Page", 
                              f"Could not load page {page_number + 1}:\n{str(e)}")

    def load_all_pages(self):
        for page_num in range(len(self.doc)):
            try:
                # Create page widget
                page_widget = PageWidget(self.content_widget)
                
                # Add some spacing between pages
                if page_num > 0:
                    self.content_layout.addSpacing(20)
                
                self.content_layout.addWidget(page_widget)
                self.page_widgets.append(page_widget)
                
                # Load the page
                self.show_page(page_num)
            except Exception as e:
                QMessageBox.warning(None, "Error Loading Page", 
                                  f"Could not load page {page_num + 1}:\n{str(e)}")

    def remove_page(self, page_number):
        try:
            # Remove the widget from layout and list
            page_widget = self.page_widgets.pop(page_number)
            self.content_layout.removeWidget(page_widget)
            page_widget.deleteLater()
            
            # Update the working copy
            self.doc.delete_page(page_number)
            self.doc.saveIncr()  # Use saveIncr() instead of save()
            
            # Refresh the remaining pages
            for i, widget in enumerate(self.page_widgets):
                self.show_page(i)
            
        except Exception as e:
            QMessageBox.warning(None, "Error Removing Page", 
                              f"Could not remove page {page_number + 1}:\n{str(e)}")

    def rotate_all_pages(self, degrees=90):
        try:
            for page in self.doc:
                page.set_rotation(degrees)
            # Save changes to working copy
            self.doc.save(self.working_pdf_path)
            # Re-render all pages
            for i in range(len(self.page_widgets)):
                self.show_page(i)
        except Exception as e:
            QMessageBox.warning(None, "Error Rotating Pages", 
                              f"Could not rotate pages:\n{str(e)}")

    def cleanup_temp_files(self):
        """Clean up temporary files when widget is destroyed"""
        try:
            if hasattr(self, 'doc') and self.doc:
                self.doc.close()
            if hasattr(self, 'temp_dir') and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir)
        except Exception as e:
            print(f"Error cleaning up temporary files: {e}")

    def closeEvent(self, event):
        """Override close event to clean up temp files"""
        self.cleanup_temp_files()
        super().closeEvent(event)

    def __del__(self):
        """Destructor to ensure cleanup"""
        self.cleanup_temp_files()

    def save_as(self, new_path):
        """Save the current state to a new file"""
        try:
            shutil.copy2(self.working_pdf_path, new_path)
            return True
        except Exception as e:
            QMessageBox.critical(None, "Error Saving", 
                               f"Could not save PDF file:\n{str(e)}")
            return False

    def close_file(self):
        """Close the current PDF file"""
        # Find the tab widget and close the current tab
        parent = self.parent()
        while parent and not isinstance(parent, QTabWidget):
            parent = parent.parent()
            
        if parent:
            index = parent.indexOf(self)
            if index >= 0:
                parent.tabCloseRequested.emit(index)

    def edit_text(self, page_number, old_text, new_text):
        """Edit text content in the PDF"""
        try:
            page = self.doc[page_number]
            # Search for text instances and replace them
            text_instances = page.search_for(old_text)
            if text_instances:
                for inst in text_instances:
                    # Get original position and font info
                    text_info = page.get_text("dict", clip=inst)
                    if "blocks" in text_info and text_info["blocks"]:
                        block = text_info["blocks"][0]
                        if "lines" in block and block["lines"]:
                            line = block["lines"][0]
                            if "spans" in line and line["spans"]:
                                span = line["spans"][0]
                                font_size = span["size"]
                                font_name = span["font"]
                                color = span.get("color", (0, 0, 0))
                                
                                # Calculate text position
                                x0, y0, x1, y1 = inst
                                text_width = x1 - x0
                                text_height = y1 - y0
                                
                                # Remove old text
                                page.add_redact_annot(inst)
                                page.apply_redactions()
                                
                                # Insert new text with preserved formatting and position
                                page.insert_text(
                                    (x0, y0 + text_height * 0.8),  # Adjust Y position
                                    new_text,
                                    fontsize=font_size,
                                    fontname=font_name,
                                    color=color
                                )
                
                self.doc.saveIncr()
                self.show_page(page_number)
                return True
            return False
        except Exception as e:
            QMessageBox.warning(None, "Error Editing Text", 
                              f"Could not edit text on page {page_number + 1}:\n{str(e)}")
            return False

    def start_manual_edit(self, page_number, position):
        """Start manual text editing at clicked position"""
        try:
            page = self.doc[page_number]
            # Get text near clicked position
            text_info = page.get_text("dict", clip=(position[0]-50, position[1]-20, position[0]+50, position[1]+20))
            if "blocks" in text_info and text_info["blocks"]:
                return text_info["blocks"][0].get("text", ""), position
            return "", position
        except Exception as e:
            QMessageBox.warning(None, "Error", f"Could not start manual edit:\n{str(e)}")
            return "", position

    def add_text(self, page_number, text, position=(100, 100)):
        """Add new text to the PDF"""
        try:
            page = self.doc[page_number]
            page.insert_text(position, text)
            self.doc.saveIncr()
            self.show_page(page_number)
            return True
        except Exception as e:
            QMessageBox.warning(None, "Error Adding Text", 
                              f"Could not add text to page {page_number + 1}:\n{str(e)}")
            return False

    def add_highlight(self, page_number, text):
        """Highlight text in the PDF"""
        try:
            page = self.doc[page_number]
            text_instances = page.search_for(text)
            for inst in text_instances:
                highlight = page.add_highlight_annot(inst)
                highlight.update()
            self.doc.saveIncr()
            self.show_page(page_number)
            return True
        except Exception as e:
            QMessageBox.warning(None, "Error Highlighting Text", 
                              f"Could not highlight text on page {page_number + 1}:\n{str(e)}")
            return False

    def add_annotation(self, page_number, text, position=(100, 100)):
        """Add a text annotation (comment) to the PDF"""
        try:
            page = self.doc[page_number]
            annot = page.add_text_annot(position, text)
            annot.update()
            self.doc.saveIncr()
            self.show_page(page_number)
            return True
        except Exception as e:
            QMessageBox.warning(None, "Error Adding Annotation", 
                              f"Could not add annotation to page {page_number + 1}:\n{str(e)}")
            return False
