from PyQt5.QtWidgets import (
    QMainWindow, QFileDialog, QAction, QTabWidget, QToolBar, QMessageBox
)
from .widgets.pdf_view_widget import PDFViewWidget

class PDFViewerApp(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("My PDF Viewer & Editor")
        self.resize(1000, 700)

        # Main Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)  # Enable close buttons on tabs
        self.tab_widget.tabCloseRequested.connect(self.close_tab)  # Connect close signal
        self.setCentralWidget(self.tab_widget)

        # Setup Toolbars / Menus
        self.setup_menus_and_toolbars()

    def setup_menus_and_toolbars(self):
        # File menu
        open_action = QAction("Open PDF...", self)
        open_action.triggered.connect(self.open_pdf)
        
        save_as_action = QAction("Save As...", self)
        save_as_action.triggered.connect(self.save_pdf_as)
        
        edit_action = QAction("Edit PDF...", self)
        edit_action.triggered.connect(self.edit_pdf)

        # Toolbar
        toolbar = QToolBar("Main Toolbar")
        toolbar.addAction(open_action)
        toolbar.addAction(save_as_action)
        toolbar.addAction(edit_action)
        self.addToolBar(toolbar)

    def open_pdf(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(
            self, 
            "Open PDF(s)", 
            "", 
            "PDF Files (*.pdf);;All Files (*)", 
            options=options
        )
        if files:
            for pdf_file in files:
                self.add_pdf_tab(pdf_file)

    def add_pdf_tab(self, pdf_file_path):
        try:
            pdf_viewer = PDFViewWidget(pdf_file_path)
            if not pdf_viewer.doc:  # Check if document loaded successfully
                return  # Don't add tab if document failed to load
            
            filename = pdf_file_path.split('/')[-1]
            self.tab_widget.addTab(pdf_viewer, filename)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not open PDF file:\n{str(e)}")

    def edit_pdf(self):
        # Get the currently active tab’s PDF
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            QMessageBox.warning(self, "No PDF Open", "Please open a PDF first.")
            return

        # A simple example: Rotate pages or add annotation
        # This could be replaced with a more sophisticated UI
        current_widget.rotate_all_pages(90)  # rotate 90 degrees
        QMessageBox.information(self, "Edit Complete", "PDF was rotated by 90 degrees.")

    def save_pdf_as(self):
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            QMessageBox.warning(self, "No PDF Open", "Please open a PDF first.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save PDF As...",
            "",
            "PDF Files (*.pdf);;All Files (*)",
            options=options
        )
        
        if file_path:
            if current_widget.save_as(file_path):
                QMessageBox.information(self, "Success", "PDF saved successfully!")

    def close_tab(self, index):
        """Close the tab at the given index"""
        widget = self.tab_widget.widget(index)
        if widget:
            widget.cleanup_temp_files()  # Clean up temp files
            self.tab_widget.removeTab(index)
