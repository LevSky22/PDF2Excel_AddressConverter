import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QListWidget, QFileDialog, QProgressBar, QLabel,
                             QAbstractItemView, QComboBox, QMessageBox, QInputDialog, QLineEdit,
                             QCheckBox, QDialog, QFormLayout, QDialogButtonBox, QFrame)  # Add QCheckBox and QFrame to imports
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QDesktopServices, QPixmap, QPainter, QColor, QFont, QKeyEvent, QIcon
from pdf2excel import convert_pdf_to_excel, auto_adjust_columns, setup_logging
import pandas as pd
from datetime import datetime
import time
import logging
import json
import ctypes

VERSION = "1.1"

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
        'custom_filename': "Nom du fichier",
        'enter_filename': "Modifier le nom du fichier si désiré (sans extension)",
        'default_filename': "Utiliser le nom par défaut",
        'duplicate_files_title': "Fichiers en double",
        'duplicate_files_msg': "Les fichiers suivants sont déjà dans la liste:\n{}",
        'files_overwritten': "{} fichier(s) déjà dans la liste ont été remplacés.",
        'no_new_files': "Aucun nouveau fichier ajouté.",
        'files_removed': "Fichier(s) supprimé(s)",
        'replace_and_add': "Remplacer et ajouter de nouveaux fichiers",
        'add_new_only': "Ajouter uniquement de nouveaux fichiers",
        'operation_cancelled': "Opération annulée",
        'enable_logging': "Activer la journalisation",
        'column_settings': "Paramètres des colonnes",
        'merge_names_checkbox': "Fusionner Prénom/Nom",
        'merged_column_name': "Nom de la colonne fusionnée",
        'default_value': "Valeur par défaut",
        'first_name': "Prénom",
        'last_name': "Nom",
        'address': "Adresse",
        'city': "Ville",
        'province': "Province",
        'postal_code': "Code postal",
        'column_settings_title': "Paramètres des colonnes",
        'column_name': "Nom de colonne",
        'save_preset': "Sauvegarder préréglage",
        'load_preset': "Charger préréglage",
        'preset_name': "Nom du préréglage",
        'enter_preset_name': "Entrez le nom du préréglage",
        'preset_saved': "Préréglage sauvegardé",
        'select_preset': "Sélectionner préréglage",
        'delete_preset': "Supprimer préréglage",
        'confirm_delete': "Êtes-vous sûr de vouloir supprimer ce préréglage?",
        'file_format': "Format du fichier",
        'select_format': "Sélectionner le format",
        'excel_format': "Excel (.xlsx)",
        'csv_format': "CSV (.csv)",
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
        'custom_filename': "Filename",
        'enter_filename': "Edit filename if desired (without extension)",
        'default_filename': "Use default filename",
        'duplicate_files_title': "Duplicate Files",
        'duplicate_files_msg': "The following files are already in the list:\n{}",
        'files_overwritten': "{} files already in the list have been replaced.",
        'no_new_files': "No new files added.",
        'files_removed': "File(s) removed",
        'replace_and_add': "Replace & Add New Files",
        'add_new_only': "Add New Files Only",
        'operation_cancelled': "Operation cancelled",
        'enable_logging': "Enable logging",
        'column_settings': "Column Settings",
        'merge_names_checkbox': "Merge First/Last Name",
        'merged_column_name': "Merged Column Name",
        'default_value': "Default Value",
        'first_name': "First Name",
        'last_name': "Last Name",
        'address': "Address",
        'city': "City",
        'province': "Province",
        'postal_code': "Postal Code",
        'column_settings_title': "Column Settings",
        'column_name': "Column Name",
        'save_preset': "Save Preset",
        'load_preset': "Load Preset",
        'preset_name': "Preset Name",
        'enter_preset_name': "Enter preset name",
        'preset_saved': "Preset saved",
        'select_preset': "Select Preset",
        'delete_preset': "Delete Preset",
        'confirm_delete': "Are you sure you want to delete this preset?",
        'file_format': "File Format",
        'select_format': "Select format",
        'excel_format': "Excel (.xlsx)",
        'csv_format': "CSV (.csv)",
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
    conversion_complete = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, pdf_files, output_dir, merge_files, custom_filename=None, enable_logging=False):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir
        self.merge_files = merge_files
        self.custom_filename = custom_filename
        self.enable_logging = enable_logging
        self.column_names = None
        self.merge_names = False
        self.merged_name = None
        self.default_values = None
        self.file_format = 'xlsx'  # Add default format

    def run(self):
        try:
            output_file = None
            for progress in convert_pdf_to_excel(
                self.pdf_files, 
                self.output_dir, 
                self.merge_files, 
                self.custom_filename, 
                self.enable_logging,
                self.column_names,
                self.merge_names,
                self.merged_name,
                self.default_values,
                self.file_format  # Pass the file format
            ):
                if isinstance(progress, str):
                    output_file = progress
                else:
                    self.progress_update.emit(progress)
            
            time.sleep(2)
            self.conversion_complete.emit(output_file)
        except Exception as e:
            self.error_occurred.emit(str(e))

class ColumnSettingsDialog(QDialog):
    def __init__(self, current_columns, merge_names=False, merged_name="Full Name", default_values=None, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle(translations[self.parent.language]['column_settings_title'])
        self.setMinimumWidth(500)
        
        # Main layout with margins
        layout = QVBoxLayout()
        layout.setSpacing(15)  # Increase spacing between sections
        layout.setContentsMargins(20, 20, 20, 20)  # Add margins around the dialog
        
        # Create a group for name merge settings
        merge_group = QFrame()
        merge_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        merge_group_layout = QVBoxLayout(merge_group)
        merge_group_layout.setSpacing(10)
        
        # Merge checkbox with some styling
        self.merge_checkbox = QCheckBox(translations[self.parent.language]['merge_names_checkbox'])
        self.merge_checkbox.setStyleSheet("""
            QCheckBox {
                font-weight: bold;
                padding: 5px;
            }
        """)
        self.merge_checkbox.setChecked(merge_names)
        self.merge_checkbox.stateChanged.connect(self.on_merge_changed)
        merge_group_layout.addWidget(self.merge_checkbox)
        
        # Merged name settings in a horizontal layout
        merged_settings = QHBoxLayout()
        merged_settings.setSpacing(10)
        
        # Column name input
        name_layout = QVBoxLayout()
        name_label = QLabel(translations[self.parent.language]['merged_column_name'])
        name_label.setStyleSheet("font-weight: normal;")
        self.merged_name_input = QLineEdit(merged_name)
        self.merged_name_input.setEnabled(merge_names)
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.merged_name_input)
        merged_settings.addLayout(name_layout)
        
        # Default value input
        default_layout = QVBoxLayout()
        default_label = QLabel(translations[self.parent.language]['default_value'])
        default_label.setStyleSheet("font-weight: normal;")
        self.merged_default_value = QLineEdit(default_values.get(merged_name, "À l'occupant") if default_values else "À l'occupant")
        self.merged_default_value.setEnabled(merge_names)
        default_layout.addWidget(default_label)
        default_layout.addWidget(self.merged_default_value)
        merged_settings.addLayout(default_layout)
        
        merge_group_layout.addLayout(merged_settings)
        layout.addWidget(merge_group)
        
        # Add separator with margin
        layout.addSpacing(10)
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)
        layout.addSpacing(10)
        
        # Store original column structure
        self.original_columns = {
            'First Name': 'first_name',
            'Last Name': 'last_name',
            'Address': 'address',
            'City': 'city',
            'Province': 'province',
            'Postal Code': 'postal_code'
        }
        
        # Grid layout for column settings
        grid_layout = QFormLayout()
        grid_layout.setSpacing(10)
        self.column_inputs = {}
        self.default_inputs = {}
        
        # Use original column keys for iteration
        for original_key in self.original_columns.keys():
            col_layout = QHBoxLayout()
            col_layout.setSpacing(10)
            
            # Column name input with label
            name_layout = QVBoxLayout()
            name_label = QLabel(translations[self.parent.language]['column_name'])
            # Get the current custom name if it exists, otherwise use original key
            current_name = current_columns.get(original_key, original_key)
            name_input = QLineEdit(current_name)
            name_input.setMinimumWidth(150)
            name_layout.addWidget(name_label)
            name_layout.addWidget(name_input)
            col_layout.addLayout(name_layout)
            self.column_inputs[original_key] = name_input
            
            # Default value input with label
            default_layout = QVBoxLayout()
            default_label = QLabel(translations[self.parent.language]['default_value'])
            default_input = QLineEdit(default_values.get(current_name, "") if default_values else "")
            default_input.setMinimumWidth(150)
            default_layout.addWidget(default_label)
            default_layout.addWidget(default_input)
            col_layout.addLayout(default_layout)
            self.default_inputs[original_key] = default_input
            
            # Disable First/Last name inputs if merged
            if original_key in ['First Name', 'Last Name']:
                name_input.setEnabled(not merge_names)
                default_input.setEnabled(not merge_names)
            
            # Use original key for translation lookup
            translated_label = translations[self.parent.language][self.original_columns[original_key]]
            grid_layout.addRow(translated_label, col_layout)
        
        layout.addLayout(grid_layout)
        
        # Add buttons with some spacing
        layout.addSpacing(15)
        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel,
            Qt.Horizontal, self)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        buttons.setStyleSheet("""
            QPushButton {
                min-width: 80px;
                padding: 5px;
            }
        """)
        layout.addWidget(buttons)
        
        # Add preset controls at the top
        preset_layout = QHBoxLayout()
        
        # Preset selector
        self.preset_combo = QComboBox()
        self.load_presets()  # Load available presets
        self.preset_combo.currentTextChanged.connect(self.load_preset)
        preset_layout.addWidget(self.preset_combo)
        
        # Save preset button
        save_preset_btn = QPushButton(translations[self.parent.language]['save_preset'])
        save_preset_btn.clicked.connect(self.save_preset)
        preset_layout.addWidget(save_preset_btn)
        
        # Delete preset button
        delete_preset_btn = QPushButton(translations[self.parent.language]['delete_preset'])
        delete_preset_btn.clicked.connect(self.delete_preset)
        preset_layout.addWidget(delete_preset_btn)
        
        layout.insertLayout(0, preset_layout)  # Add at the top of the dialog
        
        # Add separator after presets
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.insertWidget(1, separator)

        self.setLayout(layout)

        self.initial_load = True  # Add this flag

    def on_merge_changed(self, state):
        is_merged = state == Qt.Checked
        self.merged_name_input.setEnabled(is_merged)
        self.merged_default_value.setEnabled(is_merged)
        
        # Toggle First/Last name inputs
        for key in ['First Name', 'Last Name']:
            if key in self.column_inputs:
                self.column_inputs[key].setEnabled(not is_merged)
                self.default_inputs[key].setEnabled(not is_merged)

    def get_settings(self):
        merge_names = self.merge_checkbox.isChecked()
        merged_name = self.merged_name_input.text() if merge_names else None
        
        column_names = {}
        default_values = {}
        
        if merge_names:
            column_names[merged_name] = merged_name
            default_values[merged_name] = self.merged_default_value.text()
        
        # Add other columns using original keys
        for original_key in self.original_columns.keys():
            if original_key not in ['First Name', 'Last Name'] or not merge_names:
                column_name = self.column_inputs[original_key].text()
                column_names[original_key] = column_name
                default_values[column_name] = self.default_inputs[original_key].text()
        
        return {
            'merge_names': merge_names,
            'merged_name': merged_name,
            'column_names': column_names,
            'default_values': default_values
        }

    def load_presets(self):
        self.preset_combo.clear()
        self.preset_combo.addItem("")  # Empty option
        try:
            with open('column_presets.json', 'r', encoding='utf-8') as f:
                presets = json.load(f)
                for preset_name in presets.keys():
                    self.preset_combo.addItem(preset_name)
        except FileNotFoundError:
            pass

    def save_preset(self):
        name, ok = QInputDialog.getText(
            self,
            translations[self.parent.language]['preset_name'],
            translations[self.parent.language]['enter_preset_name']
        )
        
        if ok and name:
            settings = {
                'merge_names': self.merge_checkbox.isChecked(),
                'merged_name': self.merged_name_input.text(),
                'merged_default': self.merged_default_value.text(),
                'columns': {
                    key: {
                        'name': self.column_inputs[key].text(),
                        'default': self.default_inputs[key].text()
                    }
                    for key in self.column_inputs
                }
            }
            
            try:
                with open('column_presets.json', 'r', encoding='utf-8') as f:
                    presets = json.load(f)
            except FileNotFoundError:
                presets = {}
            
            presets[name] = settings
            
            with open('column_presets.json', 'w', encoding='utf-8') as f:
                json.dump(presets, f, ensure_ascii=False, indent=2)
            
            self.load_presets()
            self.preset_combo.setCurrentText(name)
            QMessageBox.information(self, "", translations[self.parent.language]['preset_saved'])

    def load_preset(self, preset_name):
        if not preset_name:
            # Reset to default values
            self.merge_checkbox.setChecked(False)
            self.merged_name_input.setText("Full Name")
            self.merged_default_value.setText("À l'occupant")
            
            # Reset column names and defaults to original values
            default_columns = {
                'First Name': ('First Name', 'À'),
                'Last Name': ('Last Name', "l'occupant"),
                'Address': ('Address', ''),
                'City': ('City', ''),
                'Province': ('Province', ''),
                'Postal Code': ('Postal Code', '')
            }
            
            for key, (name, default) in default_columns.items():
                if key in self.column_inputs:
                    self.column_inputs[key].setText(name)
                    self.default_inputs[key].setText(default)
            return
            
        try:
            with open('column_presets.json', 'r', encoding='utf-8') as f:
                presets = json.load(f)
                if preset_name in presets:
                    settings = presets[preset_name]
                    
                    # Apply merge settings
                    self.merge_checkbox.setChecked(settings['merge_names'])
                    self.merged_name_input.setText(settings['merged_name'])
                    self.merged_default_value.setText(settings['merged_default'])
                    
                    # Apply column settings
                    for key, values in settings['columns'].items():
                        if key in self.column_inputs:
                            self.column_inputs[key].setText(values['name'])
                            self.default_inputs[key].setText(values['default'])
        except FileNotFoundError:
            pass

    def delete_preset(self):
        preset_name = self.preset_combo.currentText()
        if not preset_name:
            return
            
        reply = QMessageBox.question(
            self,
            "",
            translations[self.parent.language]['confirm_delete'],
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with open('column_presets.json', 'r', encoding='utf-8') as f:
                    presets = json.load(f)
                
                if preset_name in presets:
                    del presets[preset_name]
                    
                    with open('column_presets.json', 'w', encoding='utf-8') as f:
                        json.dump(presets, f, ensure_ascii=False, indent=2)
                    
                    self.load_presets()
                    self.preset_combo.setCurrentText("")
            except FileNotFoundError:
                pass

class PDFToExcelGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Windows-specific taskbar icon fix
        if sys.platform == 'win32':
            myappid = 'levsky.pdf2excel.gui.1.0'  # arbitrary string
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
        if hasattr(sys, '_MEIPASS'):
            # Running as exe
            icon_path = os.path.join(sys._MEIPASS, 'F2E.ico')
        else:
            # Running as script
            icon_path = 'F2E.ico'
        
        # Set icon in multiple places
        self.setWindowIcon(QIcon(icon_path))
        QApplication.setWindowIcon(QIcon(icon_path))
        
        # Force Windows to refresh the icon
        if sys.platform == 'win32':
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowMinMaxButtonsHint)
            self.setWindowFlags(self.windowFlags() | Qt.WindowMinMaxButtonsHint)
        
        self.language = 'Français'
        self.enable_logging = False
        self.merge_names = False
        self.merged_name = "Full Name"
        self.column_names = {
            'First Name': 'First Name',
            'Last Name': 'Last Name',
            'Address': 'Address',
            'City': 'City',
            'Province': 'Province',
            'Postal Code': 'Postal Code'
        }
        self.default_values = {}
        self.current_preset = ""  # Add this line to store current preset name
        self.setWindowTitle(translations[self.language]['window_title'])
        self.setGeometry(100, 100, 600, 400)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.setup_ui()

        self.last_output_file = None  # Add this line to store the last output file path

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

        # Settings layout
        settings_layout = QHBoxLayout()
        self.column_settings_btn = QPushButton(translations[self.language]['column_settings'])
        self.column_settings_btn.clicked.connect(self.show_column_settings)
        settings_layout.addWidget(self.column_settings_btn)
        self.layout.addLayout(settings_layout)

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
            
            # Update Column Settings button text
            self.column_settings_btn.setText(translations[self.language]['column_settings'])
            
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
            self.status_label.setText(translations[self.language]['operation_cancelled'])
            return

        # Add format selection dialog
        format_dialog = QDialog(self)
        format_dialog.setWindowTitle(translations[self.language]['file_format'])
        layout = QVBoxLayout()
        
        format_combo = QComboBox()
        format_combo.addItems([
            translations[self.language]['excel_format'],
            translations[self.language]['csv_format']
        ])
        layout.addWidget(QLabel(translations[self.language]['select_format']))
        layout.addWidget(format_combo)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(format_dialog.accept)
        buttons.rejected.connect(format_dialog.reject)
        layout.addWidget(buttons)
        
        format_dialog.setLayout(layout)
        
        if format_dialog.exec_() != QDialog.Accepted:
            self.status_label.setText(translations[self.language]['operation_cancelled'])
            return
            
        file_format = 'xlsx' if format_combo.currentIndex() == 0 else 'csv'

        pdf_files = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        merge_files = len(pdf_files) > 1

        # Generate default filename
        current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        if merge_files:
            default_filename = f'merged_output_{current_time}'
        else:
            default_filename = f'{os.path.splitext(os.path.basename(pdf_files[0]))[0]}_{current_time}'
            if len(pdf_files) > 1:
                default_filename += '_1'

        custom_filename, ok = QInputDialog.getText(
            self, 
            translations[self.language]['custom_filename'],
            translations[self.language]['enter_filename'],
            QLineEdit.Normal,
            default_filename
        )
        
        if not ok:
            self.status_label.setText(translations[self.language]['operation_cancelled'])
            return

        if not custom_filename:
            custom_filename = default_filename

        self.conversion_thread = ConversionThread(pdf_files, output_dir, merge_files, 
                                                custom_filename, self.enable_logging)
        self.conversion_thread.column_names = self.column_names
        self.conversion_thread.merge_names = self.merge_names
        self.conversion_thread.merged_name = self.merged_name
        self.conversion_thread.default_values = self.default_values
        self.conversion_thread.file_format = file_format  # Add file format
        
        self.conversion_thread.progress_update.connect(self.update_progress)
        self.conversion_thread.conversion_complete.connect(self.conversion_finished)
        self.conversion_thread.error_occurred.connect(self.show_error)

        self.conversion_thread.start()
        self.convert_btn.setEnabled(False)
        self.status_label.setText(translations[self.language]['converting'])

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def conversion_finished(self, output_file):
        self.status_label.setText(translations[self.language]['conversion_success'])
        self.convert_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        
        # Store and open the output file
        self.last_output_file = output_file
        if self.last_output_file and os.path.exists(self.last_output_file):
            if sys.platform == 'win32':
                os.startfile(self.last_output_file)
            else:
                QDesktopServices.openUrl(QUrl.fromLocalFile(self.last_output_file))
        
        # Add a slight delay before resetting the progress bar
        QTimer.singleShot(1000, self.reset_progress_bar)

    def reset_progress_bar(self):
        self.progress_bar.setValue(0)

    def show_error(self, error_message):
        self.status_label.setText(translations[self.language]['error'].format(error_message))
        self.convert_btn.setEnabled(True)

    def show_column_settings(self):
        dialog = ColumnSettingsDialog(
            self.column_names, 
            self.merge_names, 
            self.merged_name,
            self.default_values,
            self
        )
        # Set the current preset in the combo box
        if hasattr(self, 'current_preset'):
            dialog.preset_combo.setCurrentText(self.current_preset)
            
        if dialog.exec_() == QDialog.Accepted:
            settings = dialog.get_settings()
            self.merge_names = settings['merge_names']
            self.merged_name = settings['merged_name']
            self.column_names = settings['column_names']
            self.default_values = settings['default_values']
            # Store the selected preset name
            self.current_preset = dialog.preset_combo.currentText()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set app icon globally
    if hasattr(sys, '_MEIPASS'):
        icon_path = os.path.join(sys._MEIPASS, 'F2E.ico')
    else:
        icon_path = 'F2E.ico'
    app.setWindowIcon(QIcon(icon_path))
    
    window = PDFToExcelGUI()
    window.show()
    sys.exit(app.exec_())