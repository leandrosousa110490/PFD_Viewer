from PyQt5.QtWidgets import (
    QMainWindow, QFileDialog, QAction, QTabWidget, QToolBar, QMessageBox,
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QSpinBox, QTextEdit, QInputDialog, QWidget, QComboBox, QStackedWidget
)
from PyQt5.QtCore import Qt
from .widgets.pdf_view_widget import PDFViewWidget, PageWidget

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
        self.edit_combo.addItems([
            "Replace Text",
            "Add New Text",
            "Highlight Text",
            "Add Annotation"
        ])
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
        replace_btn.clicked.connect(self.edit_text)
        replace_layout.addWidget(replace_btn)
        replace_widget.setLayout(replace_layout)
        self.stacked_widget.addWidget(replace_widget)

        # Add New Text widget
        add_text_widget = QWidget()
        add_text_layout = QVBoxLayout()
        self.new_text_content = QLineEdit()
        self.new_text_content.setPlaceholderText("Text to add")
        add_text_layout.addWidget(self.new_text_content)
        add_text_btn = QPushButton("Add")
        add_text_btn.clicked.connect(self.add_text)
        add_text_layout.addWidget(add_text_btn)
        add_text_widget.setLayout(add_text_layout)
        self.stacked_widget.addWidget(add_text_widget)

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

        # Add Annotation widget
        annot_widget = QWidget()
        annot_layout = QVBoxLayout()
        self.annotation_text = QTextEdit()
        self.annotation_text.setPlaceholderText("Annotation text")
        annot_layout.addWidget(self.annotation_text)
        annot_btn = QPushButton("Add Annotation")
        annot_btn.clicked.connect(self.add_annotation)
        annot_layout.addWidget(annot_btn)
        annot_widget.setLayout(annot_layout)
        self.stacked_widget.addWidget(annot_widget)

        layout.addWidget(self.stacked_widget)
        
        # Connect combo box to stacked widget
        self.edit_combo.currentIndexChanged.connect(self.stacked_widget.setCurrentIndex)

        # Add manual edit instructions
        manual_edit_label = QLabel(
            "Click anywhere on the PDF to manually edit text at that position"
        )
        manual_edit_label.setWordWrap(True)
        layout.addWidget(manual_edit_label)

        self.setLayout(layout)
        self.pdf_widget = None

    def set_pdf_widget(self, widget):
        self.pdf_widget = widget
        if widget and widget.doc:
            self.page_spin.setMaximum(len(widget.doc))

    def edit_text(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        old = self.old_text.text()
        new = self.new_text.text()
        if old and new:
            self.pdf_widget.edit_text(page_num, old, new)

    def add_text(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        text = self.new_text_content.text()
        if text:
            self.pdf_widget.add_text(page_num, text)

    def highlight_text_action(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        text = self.highlight_text.text()
        if text:
            self.pdf_widget.add_highlight(page_num, text)

    def add_annotation(self):
        if not self.pdf_widget:
            return
        page_num = self.page_spin.value() - 1
        text = self.annotation_text.toPlainText()
        if text:
            self.pdf_widget.add_annotation(page_num, text)

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
        current_widget = self.tab_widget.currentWidget()
        if not current_widget:
            QMessageBox.warning(self, "No PDF Open", "Please open a PDF first.")
            return

        dialog = EditDialog(self)
        dialog.set_pdf_widget(current_widget)
        dialog.exec_()

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
