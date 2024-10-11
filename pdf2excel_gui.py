import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QListWidget, QFileDialog, QProgressBar, QLabel,
                             QAbstractItemView, QComboBox, QMessageBox, QInputDialog, QLineEdit,
                             QCheckBox)  # Add QCheckBox to imports
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QDesktopServices, QPixmap, QPainter, QColor, QFont, QKeyEvent
from pdf2excel import convert_pdf_to_excel, auto_adjust_columns, setup_logging
import pandas as pd
from datetime import datetime
import time
import logging

VERSION = "1.0.0"

# Translations dictionary
translations = {
    'Français': {
        'window_title': "Convertisseur PDF vers Excel",
        'add_files': "Ajouter des fichiers",
        'remove_selected': "Enlever la sélection",
        'convert': "Convertir",
        'select_pdf_files': "Choisir les fichiers PDF",
        'select_output_folder': "Choisir le dossier de sortie",
        'new_files_added': "{} nouveau(x) fichier(s) ajouté(s). {} fichier(s) déjà dans la liste.",
        'files_added': "{} nouveau(x) fichier(s) ajouté(s).",
        'add_pdf_files': "SVP ajouter des fichiers PDF à convertir.",
        'converting': "Conversion en cours...",
        'conversion_success': "Conversion terminée avec succès!",
        'error': "Erreur: {}",
        'language': "Langue",
        'about': "À propos",
        'about_title': "À propos de PDF vers Excel",
        'about_text': "Fait avec amour par Lev! Fini le gaspillage de temps ;-)",
        'custom_filename': "Nom de fichier personnalisé",
        'enter_filename': "Entrez le nom du fichier Excel (sans extension)",
        'default_filename': "Utiliser le nom par défaut",
        'duplicate_files_title': "Fichiers en double",
        'duplicate_files_msg': "Les fichiers suivants sont déjà dans la liste:\n{}",
        'files_overwritten': "{} fichier(s) déjà dans la liste ont été remplacés.",
        'no_new_files': "Aucun nouveau fichier ajouté.",
        'files_removed': "Fichier(s) supprimé(s)",
        'replace_and_add': "Remplacer et ajouter de nouveaux fichiers",
        'add_new_only': "Ajouter uniquement de nouveaux fichiers",
        'operation_cancelled': "Opération annulée",
        'enable_logging': "Activer la journalisation"
    },
    'English': {
        'window_title': "PDF to Excel Converter",
        'add_files': "Add Files",
        'remove_selected': "Remove Selected",
        'convert': "Convert",
        'select_pdf_files': "Select PDF Files",
        'select_output_folder': "Select Output Folder",
        'new_files_added': "{} new file(s) added. {} file(s) already in the list.",
        'files_added': "{} new file(s) added.",
        'add_pdf_files': "Please add PDF files to convert.",
        'converting': "Converting...",
        'conversion_success': "Conversion completed successfully!",
        'error': "Error: {}",
        'language': "Language",
        'about': "About",
        'about_title': "About PDF to Excel",
        'about_text': "Made with love by Lev! No more wasted time ;-)",
        'custom_filename': "Custom filename",
        'enter_filename': "Enter the Excel filename (without extension)",
        'default_filename': "Use default filename",
        'duplicate_files_title': "Duplicate Files",
        'duplicate_files_msg': "The following files are already in the list:\n{}",
        'files_overwritten': "{} files already in the list have been replaced.",
        'no_new_files': "No new files added.",
        'files_removed': "File(s) removed",
        'replace_and_add': "Replace & Add New Files",
        'add_new_only': "Add New Files Only",
        'operation_cancelled': "Operation cancelled",
        'enable_logging': "Enable logging"
    }
}

class DragDropListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)
        
        # Enable rubber band selection
        self.setSelectionRectVisible(True)
        self.parent = parent  # Store reference to parent

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().endswith('.pdf')]
        self.parent.add_new_files(files)  # Use the parent's add_new_files method

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key_Delete:
            self.parent.remove_files()  # Call the parent's remove_files method
        elif event.key() == Qt.Key_A and event.modifiers() & Qt.ControlModifier:
            self.selectAll()  # Select all items
        else:
            super().keyPressEvent(event)  # Handle other key events normally

class ConversionThread(QThread):
    progress_update = pyqtSignal(int)
    conversion_complete = pyqtSignal()
    error_occurred = pyqtSignal(str)

    def __init__(self, pdf_files, output_dir, merge_files, custom_filename=None, enable_logging=False):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir
        self.merge_files = merge_files
        self.custom_filename = custom_filename
        self.enable_logging = enable_logging

    def run(self):
        try:
            for progress in convert_pdf_to_excel(self.pdf_files, self.output_dir, self.merge_files, self.custom_filename, self.enable_logging):
                self.progress_update.emit(progress)
            
            # Add a small delay before emitting completion signal
            time.sleep(2)
            self.conversion_complete.emit()
        except Exception as e:
            self.error_occurred.emit(str(e))

class PDFToExcelGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.language = 'Français'  # Default language
        self.enable_logging = False  # Disable logging by default
        self.setWindowTitle(translations[self.language]['window_title'])
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.setup_ui()

    def setup_ui(self):
        # Top bar with Language and About
        top_bar = QHBoxLayout()
        
        # Language selection
        lang_layout = QHBoxLayout()
        self.lang_label = QLabel(translations[self.language]['language'])
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(['Français', 'English'])
        self.lang_combo.setCurrentText(self.language)
        self.lang_combo.currentTextChanged.connect(self.change_language)
        lang_layout.addWidget(self.lang_label)
        lang_layout.addWidget(self.lang_combo)
        top_bar.addLayout(lang_layout)
        
        # About button
        self.about_btn = QPushButton(translations[self.language]['about'])
        self.about_btn.clicked.connect(self.show_about)
        top_bar.addWidget(self.about_btn)
        
        self.layout.addLayout(top_bar)

        # File selection
        self.file_list = DragDropListWidget(self)  # Pass self as parent
        self.layout.addWidget(self.file_list)

        # Buttons
        button_layout = QHBoxLayout()
        self.add_files_btn = QPushButton(translations[self.language]['add_files'])
        self.remove_files_btn = QPushButton(translations[self.language]['remove_selected'])
        self.convert_btn = QPushButton(translations[self.language]['convert'])
        button_layout.addWidget(self.add_files_btn)
        button_layout.addWidget(self.remove_files_btn)
        button_layout.addWidget(self.convert_btn)
        self.layout.addLayout(button_layout)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.layout.addWidget(self.progress_bar)

        # Status label
        self.status_label = QLabel()
        self.layout.addWidget(self.status_label)

        # Connect signals
        self.add_files_btn.clicked.connect(self.add_files)
        self.remove_files_btn.clicked.connect(self.remove_files)
        self.convert_btn.clicked.connect(self.start_conversion)

    def change_language(self, new_language):
        if new_language != self.language:
            self.language = new_language
            self.setWindowTitle(translations[self.language]['window_title'])
            self.lang_label.setText(translations[self.language]['language'])
            self.about_btn.setText(translations[self.language]['about'])
            self.add_files_btn.setText(translations[self.language]['add_files'])
            self.remove_files_btn.setText(translations[self.language]['remove_selected'])
            self.convert_btn.setText(translations[self.language]['convert'])
            
            # Update any visible status message
            current_status = self.status_label.text()
            for key, value in translations[self.language].items():
                if value == current_status:
                    self.status_label.setText(translations[self.language][key])
                    break
            
            # Force update of the UI
            self.update()
            QApplication.processEvents()

    def show_about(self):
        about_box = QMessageBox(self)
        about_box.setWindowTitle(translations[self.language]['about_title'])
        about_box.setText(f"{translations[self.language]['about_text']}\n\nVersion: {VERSION}")
        about_box.setInformativeText('<a href="https://github.com/LevSky22/PDF2Excel_AddressConverter">https://github.com/LevSky22/PDF2Excel_AddressConverter</a>')
        about_box.setTextFormat(Qt.RichText)
        about_box.setTextInteractionFlags(Qt.TextBrowserInteraction)

        # Add logging checkbox to the about box
        logging_checkbox = QCheckBox(translations[self.language]['enable_logging'])
        logging_checkbox.setChecked(self.enable_logging)
        about_box.setCheckBox(logging_checkbox)

        about_box.exec_()

        # Update the logging state based on the checkbox
        self.enable_logging = logging_checkbox.isChecked()

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, translations[self.language]['select_pdf_files'], "", "PDF Files (*.pdf)")
        self.add_new_files(files)

    def add_new_files(self, files):
        current_files = set(self.file_list.item(i).text() for i in range(self.file_list.count()))
        new_files = [file for file in files if file not in current_files]
        duplicate_files = [file for file in files if file in current_files]
        
        if duplicate_files:
            duplicate_msg = "\n".join(os.path.basename(f) for f in duplicate_files)
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle(translations[self.language]['duplicate_files_title'])
            msg_box.setText(translations[self.language]['duplicate_files_msg'].format(duplicate_msg))
            replace_button = msg_box.addButton(translations[self.language]['replace_and_add'], QMessageBox.ActionRole)
            add_new_button = msg_box.addButton(translations[self.language]['add_new_only'], QMessageBox.ActionRole)
            cancel_button = msg_box.addButton(QMessageBox.Cancel)
            
            msg_box.exec_()
            
            if msg_box.clickedButton() == replace_button:
                # Remove duplicates from the list
                for file in duplicate_files:
                    items = self.file_list.findItems(file, Qt.MatchExactly)
                    for item in items:
                        self.file_list.takeItem(self.file_list.row(item))
                # Add all files (including the "duplicates")
                self.file_list.addItems(files)
                self.status_label.setText(translations[self.language]['files_overwritten'].format(len(files)))
            elif msg_box.clickedButton() == add_new_button:
                # Add only new files
                self.file_list.addItems(new_files)
                if new_files:
                    self.status_label.setText(translations[self.language]['new_files_added'].format(len(new_files), len(duplicate_files)))
                else:
                    self.status_label.setText(translations[self.language]['no_new_files'])
            else:  # Cancel button clicked
                self.status_label.setText(translations[self.language]['operation_cancelled'])
        else:
            # No duplicates, add all files
            self.file_list.addItems(files)
            self.status_label.setText(translations[self.language]['files_added'].format(len(files)))

    def remove_files(self):
        for item in self.file_list.selectedItems():
            self.file_list.takeItem(self.file_list.row(item))
        self.status_label.setText(translations[self.language]['files_removed'])

    def start_conversion(self):
        if self.file_list.count() == 0:
            self.status_label.setText(translations[self.language]['add_pdf_files'])
            return

        output_dir = QFileDialog.getExistingDirectory(self, translations[self.language]['select_output_folder'])
        if not output_dir:
            return

        pdf_files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        merge_files = len(pdf_files) > 1  # Merge if there's more than one file

        # Ask for custom filename
        custom_filename, ok = QInputDialog.getText(self, translations[self.language]['custom_filename'],
                                                   translations[self.language]['enter_filename'],
                                                   QLineEdit.Normal,
                                                   translations[self.language]['default_filename'])
        
        if not ok or custom_filename == translations[self.language]['default_filename']:
            custom_filename = None

        # Use self.enable_logging instead of checkbox
        self.conversion_thread = ConversionThread(pdf_files, output_dir, merge_files, custom_filename, self.enable_logging)
        self.conversion_thread.progress_update.connect(self.update_progress)
        self.conversion_thread.conversion_complete.connect(self.conversion_finished)
        self.conversion_thread.error_occurred.connect(self.show_error)

        self.conversion_thread.start()
        self.convert_btn.setEnabled(False)
        self.status_label.setText(translations[self.language]['converting'])

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def conversion_finished(self):
        self.status_label.setText(translations[self.language]['conversion_success'])
        self.convert_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        # Add a slight delay before resetting the progress bar
        QTimer.singleShot(1000, self.reset_progress_bar)

    def reset_progress_bar(self):
        self.progress_bar.setValue(0)

    def show_error(self, error_message):
        self.status_label.setText(translations[self.language]['error'].format(error_message))
        self.convert_btn.setEnabled(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFToExcelGUI()
    window.show()
    sys.exit(app.exec_())