import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QListWidget, QFileDialog, QProgressBar, QLabel,
                             QAbstractItemView, QComboBox, QMessageBox, QInputDialog, QLineEdit,
                             QCheckBox, QDialog, QFormLayout, QDialogButtonBox, QFrame, QDateEdit,
                             QScrollArea)  # Add QScrollArea
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl, QDate  # Add QDate
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QDesktopServices, QPixmap, QPainter, QColor, QFont, QKeyEvent, QIcon
from pdf2excel import convert_pdf_to_excel, auto_adjust_columns, setup_logging
import pandas as pd
from datetime import datetime
import time
import logging
import json
import ctypes
from quebec_regions_mapping import get_shore_region

VERSION = "1.5"

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
        'merge_address': "Fusionner les champs d'adresse",
        'merged_address_name': "Nom de la colonne fusionnée",
        'address_separator': "Séparateur d'adresse",
        'province_default': "Province par défaut",
        'extract_apartment': "Extraire les numéros d'appartement",
        'apartment_column_name': "Nom de la colonne d'appartement",
        'filter_apartments': "Exclure les adresses avec appartements",
        'include_apartment_column': "Inclure la colonne d'appartement",
        'include_phone': "Inclure numéro de téléphone",
        'phone_column_name': "Nom de la colonne téléphone",
        'phone_default': "Numéro par défaut",
        'include_date': "Inclure la date",
        'date_column_name': "Nom de la colonne date",
        'date_value': "Sélectionner la date",
        'filter_by_region': "Filtrer par région",
        'region_settings': "Paramètres des régions",
        'branch_id': "ID de succursale",
        'north_shore': "Rive Nord",
        'south_shore': "Rive Sud",
        'montreal': "Montréal",
        'laval': "Laval",
        'longueuil': "Longueuil",
        'unknown': "Inconnu",
        'use_custom_sectors': "Utiliser des secteurs personnalisés",
        'remove_accents': "Retirer les accents",
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
        'merge_address': "Merge Address Fields",
        'merged_address_name': "Merged Column Name",
        'address_separator': "Address Separator",
        'province_default': "Default Province",
        'extract_apartment': "Extract Apartment Numbers",
        'apartment_column_name': "Apartment Column Name",
        'filter_apartments': "Exclude addresses with apartments",
        'include_apartment_column': "Include Apartment Column",
        'include_phone': "Include phone number",
        'phone_column_name': "Phone column name",
        'phone_default': "Default number",
        'include_date': "Include date",
        'date_column_name': "Date column name",
        'date_value': "Select date",
        'filter_by_region': "Filter by region",
        'region_settings': "Region settings",
        'branch_id': "Branch ID",
        'north_shore': "North Shore",
        'south_shore': "South Shore",
        'montreal': "Montreal",
        'laval': "Laval",
        'longueuil': "Longueuil",
        'unknown': "Unknown",
        'use_custom_sectors': "Use Custom Sectors",
        'remove_accents': "Remove accented characters",
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
        self.file_format = 'xlsx'
        self.merge_address = False
        self.merged_address_name = "Complete Address"
        self.address_separator = ", "
        self.province_default = "QC"
        self.extract_apartment = False
        self.apartment_column_name = "Apartment"
        self.filter_apartments = False
        self.include_apartment_column = True
        self.include_phone = False
        self.phone_column_name = "Phone"
        self.phone_default = ""
        self.include_date = False
        self.date_column_name = "Date"
        self.date_value = None
        self.filter_by_region = False
        self.region_branch_ids = None
        self.use_custom_sectors = False
        self.custom_sector_ids = {}
        self.remove_accents = False

    def run(self):
        try:
            if self.enable_logging:
                logging.info(f"ConversionThread starting with remove_accents={self.remove_accents}")
                
            # Initialize logging if enabled
            if self.enable_logging:
                # Reset any existing handlers
                logging.getLogger().handlers = []
                logging.disable(logging.NOTSET)  # Enable logging
                self.log_file = setup_logging()  # Create new log file
                logging.info("Starting PDF conversion process")
                logging.info(f"Processing files: {self.pdf_files}")
                logging.info(f"Output directory: {self.output_dir}")
                logging.info(f"Merge files: {self.merge_files}")
                logging.info(f"Custom filename: {self.custom_filename}")
                logging.info(f"Custom sectors enabled: {self.use_custom_sectors}")

            output_file = None
            # Only apply apartment filtering if explicitly set to True
            should_filter = (self.extract_apartment and 
                            hasattr(self, 'filter_apartments') and 
                            self.filter_apartments is True)

            if self.enable_logging:
                logging.info(f"Settings: extract_apartment={self.extract_apartment}, "
                            f"filter_apartments={should_filter}, "
                            f"merge_address={self.merge_address}, "
                            f"use_custom_sectors={self.use_custom_sectors}")

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
                self.file_format,
                self.merge_address,
                self.merged_address_name,
                self.address_separator,
                self.province_default,
                self.extract_apartment,
                self.apartment_column_name,
                should_filter,
                self.include_apartment_column,
                self.include_phone,
                self.phone_default,
                self.include_date,
                self.date_value,
                self.filter_by_region,
                self.region_branch_ids,
                use_custom_sectors=self.use_custom_sectors,  # Add custom sectors parameter
                remove_accents=self.remove_accents  # Add remove_accents parameter
            ):
                if isinstance(progress, str):
                    output_file = progress
                    if self.enable_logging:
                        logging.info(f"Created output file: {output_file}")
                else:
                    self.progress_update.emit(progress)
                    if self.enable_logging:
                        logging.info(f"Progress: {progress}%")

            time.sleep(2)
            if self.enable_logging:
                logging.info("Conversion completed successfully")
            self.conversion_complete.emit(output_file)

        except Exception as e:
            if self.enable_logging:
                logging.error(f"Error during conversion: {str(e)}", exc_info=True)
            self.error_occurred.emit(str(e))

class ColumnSettingsDialog(QDialog):
    def __init__(self, current_columns, merge_names=False, merged_name="Full Name", default_values=None, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle(translations[self.parent.language]['column_settings_title'])
        self.setMinimumWidth(500)
        
        # Create main layout
        main_layout = QVBoxLayout(self)
        
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        
        # Create container widget for scroll area
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Initialize column inputs dictionaries
        self.column_inputs = {}
        self.default_inputs = {}
        
        # Store original column structure
        self.original_columns = {
            'First Name': 'first_name',
            'Last Name': 'last_name',
            'Address': 'address',
            'City': 'city',
            'Province': 'province',
            'Postal Code': 'postal_code'
        }
        
        # Create column input fields
        column_group = QFrame()
        column_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        column_layout = QVBoxLayout(column_group)
        
        # Create input fields for each column
        for original_key in self.original_columns.keys():
            row_layout = QHBoxLayout()
            
            # Column name input
            name_layout = QVBoxLayout()
            name_label = QLabel(translations[self.parent.language]['column_name'])
            current_name = current_columns.get(original_key, original_key)
            name_input = QLineEdit(current_name)
            name_layout.addWidget(name_label)
            name_layout.addWidget(name_input)
            row_layout.addLayout(name_layout)
            self.column_inputs[original_key] = name_input
            
            # Default value input
            default_layout = QVBoxLayout()
            default_label = QLabel(translations[self.parent.language]['default_value'])
            default_input = QLineEdit(default_values.get(current_name, ""))
            default_layout.addWidget(default_label)
            default_layout.addWidget(default_input)
            row_layout.addLayout(default_layout)
            self.default_inputs[original_key] = default_input
            
            # Add row to column layout
            field_label = QLabel(translations[self.parent.language][self.original_columns[original_key]])
            field_label.setStyleSheet("font-weight: bold;")
            column_layout.addWidget(field_label)
            column_layout.addLayout(row_layout)
            column_layout.addSpacing(10)
        
        layout.addWidget(column_group)
        
        # Add preset controls at the top (outside scroll area)
        preset_layout = QHBoxLayout()
        self.preset_combo = QComboBox()
        self.load_presets()
        self.preset_combo.currentTextChanged.connect(self.load_preset)
        preset_layout.addWidget(self.preset_combo)
        
        save_preset_btn = QPushButton(translations[self.parent.language]['save_preset'])
        save_preset_btn.clicked.connect(self.save_preset)
        preset_layout.addWidget(save_preset_btn)
        
        delete_preset_btn = QPushButton(translations[self.parent.language]['delete_preset'])
        delete_preset_btn.clicked.connect(self.delete_preset)
        preset_layout.addWidget(delete_preset_btn)
        
        main_layout.addLayout(preset_layout)
        
        # Add separator after presets
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(separator)
        
        # Add name merge settings
        name_group = QFrame()
        name_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        name_group_layout = QVBoxLayout(name_group)
        
        self.merge_checkbox = QCheckBox(translations[self.parent.language]['merge_names_checkbox'])
        self.merge_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.merge_checkbox.stateChanged.connect(self.on_merge_changed)
        name_group_layout.addWidget(self.merge_checkbox)
        
        merged_settings = QHBoxLayout()
        self.merged_name_input = QLineEdit(merged_name)
        self.merged_default_value = QLineEdit(default_values.get(merged_name, "À l'occupant") if default_values else "À l'occupant")
        merged_settings.addWidget(QLabel(translations[self.parent.language]['merged_column_name']))
        merged_settings.addWidget(self.merged_name_input)
        merged_settings.addWidget(QLabel(translations[self.parent.language]['default_value']))
        merged_settings.addWidget(self.merged_default_value)
        name_group_layout.addLayout(merged_settings)
        layout.addWidget(name_group)
        
        # Add address merge settings
        address_group = QFrame()
        address_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        address_group_layout = QVBoxLayout(address_group)
        
        self.merge_address_checkbox = QCheckBox(translations[self.parent.language]['merge_address'])
        self.merge_address_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.merge_address_checkbox.stateChanged.connect(self.on_merge_address_changed)
        address_group_layout.addWidget(self.merge_address_checkbox)
        
        address_settings = QHBoxLayout()
        self.merged_address_input = QLineEdit("Complete Address")
        self.address_separator_input = QLineEdit(", ")
        self.province_default_input = QLineEdit("QC")
        
        address_name_layout = QVBoxLayout()
        address_name_layout.addWidget(QLabel(translations[self.parent.language]['merged_column_name']))
        address_name_layout.addWidget(self.merged_address_input)
        address_settings.addLayout(address_name_layout)
        
        separator_layout = QVBoxLayout()
        separator_layout.addWidget(QLabel(translations[self.parent.language]['address_separator']))
        separator_layout.addWidget(self.address_separator_input)
        address_settings.addLayout(separator_layout)
        
        province_layout = QVBoxLayout()
        province_layout.addWidget(QLabel(translations[self.parent.language]['province_default']))
        province_layout.addWidget(self.province_default_input)
        address_settings.addLayout(province_layout)
        
        address_group_layout.addLayout(address_settings)
        layout.addWidget(address_group)
        
        # Add apartment settings
        apartment_group = QFrame()
        apartment_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        apartment_group_layout = QVBoxLayout(apartment_group)
        
        self.extract_apartment_checkbox = QCheckBox(translations[self.parent.language]['extract_apartment'])
        self.extract_apartment_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.extract_apartment_checkbox.stateChanged.connect(self.on_extract_apartment_changed)
        apartment_group_layout.addWidget(self.extract_apartment_checkbox)
        
        apartment_settings = QVBoxLayout()
        apartment_name_layout = QHBoxLayout()
        self.apartment_name_input = QLineEdit("Apartment")
        apartment_name_layout.addWidget(QLabel(translations[self.parent.language]['apartment_column_name']))
        apartment_name_layout.addWidget(self.apartment_name_input)
        apartment_settings.addLayout(apartment_name_layout)
        
        self.include_apartment_checkbox = QCheckBox(translations[self.parent.language]['include_apartment_column'])
        self.filter_apartments_checkbox = QCheckBox(translations[self.parent.language]['filter_apartments'])
        apartment_settings.addWidget(self.include_apartment_checkbox)
        apartment_settings.addWidget(self.filter_apartments_checkbox)
        
        apartment_group_layout.addLayout(apartment_settings)
        layout.addWidget(apartment_group)
        
        # Add phone settings
        phone_group = QFrame()
        phone_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        phone_group_layout = QVBoxLayout(phone_group)
        
        self.include_phone_checkbox = QCheckBox(translations[self.parent.language]['include_phone'])
        self.include_phone_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.include_phone_checkbox.stateChanged.connect(self.on_phone_changed)
        phone_group_layout.addWidget(self.include_phone_checkbox)
        
        phone_settings = QHBoxLayout()
        self.phone_name_input = QLineEdit("Phone")
        self.phone_default_input = QLineEdit("")
        phone_settings.addWidget(QLabel(translations[self.parent.language]['phone_column_name']))
        phone_settings.addWidget(self.phone_name_input)
        phone_settings.addWidget(QLabel(translations[self.parent.language]['phone_default']))
        phone_settings.addWidget(self.phone_default_input)
        phone_group_layout.addLayout(phone_settings)
        layout.addWidget(phone_group)
        
        # Add date settings
        date_group = QFrame()
        date_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        date_group_layout = QVBoxLayout(date_group)
        
        self.include_date_checkbox = QCheckBox(translations[self.parent.language]['include_date'])
        self.include_date_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.include_date_checkbox.stateChanged.connect(self.on_date_changed)
        date_group_layout.addWidget(self.include_date_checkbox)
        
        date_settings = QHBoxLayout()
        self.date_name_input = QLineEdit("Date")
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDate(QDate.currentDate())
        date_settings.addWidget(QLabel(translations[self.parent.language]['date_column_name']))
        date_settings.addWidget(self.date_name_input)
        date_settings.addWidget(QLabel(translations[self.parent.language]['date_value']))
        date_settings.addWidget(self.date_picker)
        date_group_layout.addLayout(date_settings)
        layout.addWidget(date_group)
        
        # Add region settings
        region_group = QFrame()
        region_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        region_group_layout = QVBoxLayout(region_group)
        
        self.filter_region_checkbox = QCheckBox(translations[self.parent.language]['filter_by_region'])
        self.filter_region_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.filter_region_checkbox.stateChanged.connect(self.on_region_filter_changed)
        region_group_layout.addWidget(self.filter_region_checkbox)
        
        # Create region branch ID inputs
        self.region_inputs = {}
        regions = {
            'flyer_north_shore': ('north_shore', 'flyer_north_shore'),
            'flyer_south_shore': ('south_shore', 'flyer_south_shore'),
            'flyer_montreal': ('montreal', 'flyer_montreal'),
            'flyer_laval': ('laval', 'flyer_laval'),
            'flyer_longueuil': ('longueuil', 'flyer_longueuil'),
            'flyer_unknown': ('unknown', 'flyer_unknown')
        }
        
        # Add regular region inputs
        for region_key, (translation_key, default_id) in regions.items():
            row_layout = QHBoxLayout()
            row_layout.addWidget(QLabel(translations[self.parent.language][translation_key]))
            branch_input = QLineEdit()
            branch_input.setPlaceholderText(default_id)
            branch_input.setText(default_id)
            branch_input.setEnabled(False)
            self.region_inputs[region_key] = branch_input
            row_layout.addWidget(branch_input)
            region_group_layout.addLayout(row_layout)
        
        # Add separator
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        region_group_layout.addWidget(separator)
        
        # Add custom sectors option
        self.use_custom_sectors_checkbox = QCheckBox(translations[self.parent.language].get('use_custom_sectors', 'Use Custom Sectors'))
        self.use_custom_sectors_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        self.use_custom_sectors_checkbox.stateChanged.connect(self.on_custom_sectors_changed)
        region_group_layout.addWidget(self.use_custom_sectors_checkbox)
        
        # Create custom sector inputs
        self.sector_inputs = {}
        sectors = {
            'flyer_chateauguay': ('chateauguay_region', 'flyer_chateauguay'),
            # More sectors can be added here later
        }
        
        for sector_key, (translation_key, default_id) in sectors.items():
            row_layout = QHBoxLayout()
            row_layout.addWidget(QLabel(translations[self.parent.language].get(translation_key, sector_key)))
            branch_input = QLineEdit()
            branch_input.setPlaceholderText(default_id)
            branch_input.setText(default_id)
            branch_input.setEnabled(False)
            self.sector_inputs[sector_key] = branch_input
            row_layout.addWidget(branch_input)
            region_group_layout.addLayout(row_layout)
        
        layout.addWidget(region_group)
        
        # Set initial states
        self.on_merge_changed(merge_names)
        self.on_merge_address_changed(False)
        self.on_extract_apartment_changed(False)
        self.on_phone_changed(False)
        self.on_date_changed(False)
        
        # Set the container as the scroll area widget
        scroll.setWidget(container)
        
        # Add scroll area to main layout
        main_layout.addWidget(scroll)
        
        # Add buttons at the bottom (outside scroll area)
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
        main_layout.addWidget(buttons)
        
        # Set size of dialog
        self.resize(600, 800)  # Adjust these values as needed

        # Add remove accents option
        accent_group = QFrame()
        accent_group.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        accent_group_layout = QVBoxLayout(accent_group)
        
        self.remove_accents_checkbox = QCheckBox(translations[self.parent.language]['remove_accents'])
        self.remove_accents_checkbox.setStyleSheet("QCheckBox { font-weight: bold; padding: 5px; }")
        # Set initial state from parent's current setting
        self.remove_accents_checkbox.setChecked(getattr(parent, 'remove_accents', False))
        accent_group_layout.addWidget(self.remove_accents_checkbox)
        layout.addWidget(accent_group)

    def load_presets(self):
        """Load available presets from file"""
        self.preset_combo.clear()
        self.preset_combo.addItem("")  # Empty option
        try:
            with open('column_presets.json', 'r', encoding='utf-8') as f:
                presets = json.load(f)
                for preset_name in presets.keys():
                    self.preset_combo.addItem(preset_name)
        except FileNotFoundError:
            pass

    def load_preset(self, preset_name):
        """Load a specific preset"""
        if not preset_name:
            self.reset_to_defaults()
            return
        
        try:
            with open('column_presets.json', 'r', encoding='utf-8') as f:
                presets = json.load(f)
                if preset_name in presets:
                    settings = presets[preset_name]
                    
                    # Update apartment settings first - ensure defaults if not specified
                    extract_apt = settings.get('extract_apartment', False)
                    self.extract_apartment_checkbox.setChecked(extract_apt)
                    self.on_extract_apartment_changed(extract_apt)  # Explicitly call the handler
                    
                    if extract_apt:
                        self.apartment_name_input.setText(settings.get('apartment_column_name', 'Apartment'))
                        # Only set these if they're explicitly True in the settings
                        self.include_apartment_checkbox.setChecked(settings.get('include_apartment_column', True))
                        self.filter_apartments_checkbox.setChecked(settings.get('filter_apartments', False))
                    
                    # First update merge checkboxes only
                    self.merge_checkbox.setChecked(settings.get('merge_names', False))
                    self.merge_address_checkbox.setChecked(settings.get('merge_address', False))
                    
                    # Update all input fields regardless of merge state
                    for key in ['First Name', 'Last Name']:
                        if key in self.column_inputs and key in settings.get('column_names', {}):
                            self.column_inputs[key].setText(settings['column_names'][key])
                            if key in settings.get('default_values', {}):
                                self.default_inputs[key].setText(settings['default_values'][key])
                    
                    for key in ['Address', 'City', 'Province', 'Postal Code']:
                        if key in self.column_inputs and key in settings.get('column_names', {}):
                            self.column_inputs[key].setText(settings['column_names'][key])
                            if key in settings.get('default_values', {}):
                                self.default_inputs[key].setText(settings['default_values'][key])
                    
                    # Update merged fields
                    self.merged_name_input.setText(settings.get('merged_name', 'Full Name'))
                    self.merged_default_value.setText(settings.get('merged_default', " l'occupant"))
                    self.merged_address_input.setText(settings.get('merged_address_name', 'Complete Address'))
                    self.address_separator_input.setText(settings.get('address_separator', ', '))
                    self.province_default_input.setText(settings.get('province_default', 'QC'))
                    
                    # Update other settings without forcing their enabled/disabled states
                    self.include_phone_checkbox.setChecked(settings.get('include_phone', False))
                    self.phone_name_input.setText(settings.get('phone_column_name', 'Phone'))
                    self.phone_default_input.setText(settings.get('phone_default', ''))
                    
                    self.include_date_checkbox.setChecked(settings.get('include_date', False))
                    self.date_name_input.setText(settings.get('date_column_name', 'Date'))
                    if settings.get('date_value'):
                        self.date_picker.setDate(QDate.fromString(settings['date_value'], 'yyyy-MM-dd'))
                    
                    # Only force update of name and address merge states
                    self.on_merge_changed(settings.get('merge_names', False))
                    self.on_merge_address_changed(settings.get('merge_address', False))
                    
                    # Enable input fields for other settings based on their checkboxes
                    self.phone_name_input.setEnabled(self.include_phone_checkbox.isChecked())
                    self.phone_default_input.setEnabled(self.include_phone_checkbox.isChecked())
                    
                    self.date_name_input.setEnabled(self.include_date_checkbox.isChecked())
                    self.date_picker.setEnabled(self.include_date_checkbox.isChecked())
                    
                    # Update region filter settings
                    self.filter_region_checkbox.setChecked(settings.get('filter_by_region', False))
                    region_branch_ids = settings.get('region_branch_ids', {})
                    
                    # Update default regions dictionary to include flyer_unknown
                    default_regions = {
                        'flyer_north_shore': 'flyer_north_shore',
                        'flyer_south_shore': 'flyer_south_shore',
                        'flyer_montreal': 'flyer_montreal',
                        'flyer_laval': 'flyer_laval',
                        'flyer_longueuil': 'flyer_longueuil',
                        'flyer_unknown': 'flyer_unknown'
                    }
                    
                    # Update region inputs and keep them enabled if filter is checked
                    is_filter_enabled = settings.get('filter_by_region', False)
                    for region_key, input_field in self.region_inputs.items():
                        # Use get() method to provide a fallback if key doesn't exist
                        default_value = default_regions.get(region_key, region_key)
                        input_field.setText(region_branch_ids.get(region_key, default_value))
                        input_field.setEnabled(is_filter_enabled)
                    
                    # Load custom sectors settings
                    self.use_custom_sectors_checkbox.setChecked(settings.get('use_custom_sectors', False))
                    custom_sector_ids = settings.get('custom_sector_ids', {})
                    for sector, input_field in self.sector_inputs.items():
                        if sector in custom_sector_ids:
                            input_field.setText(custom_sector_ids[sector])
                    
                    # Update enabled states
                    self.on_custom_sectors_changed(settings.get('use_custom_sectors', False))
                
        except FileNotFoundError:
            pass

    def save_preset(self):
        """Save current settings as a preset"""
        name, ok = QInputDialog.getText(
            self,
            translations[self.parent.language]['preset_name'],
            translations[self.parent.language]['enter_preset_name']
        )
        
        if ok and name:
            settings = self.get_settings()
            
            try:
                with open('column_presets.json', 'r', encoding='utf-8') as f:
                    presets = json.load(f)
            except FileNotFoundError:
                presets = {}
            
            presets[name] = settings
            
            # Save with proper UTF-8 encoding and ensure_ascii=False
            with open('column_presets.json', 'w', encoding='utf-8') as f:
                json.dump(presets, f, ensure_ascii=False, indent=2)
            
            self.load_presets()
            self.preset_combo.setCurrentText(name)
            QMessageBox.information(self, "", translations[self.parent.language]['preset_saved'])

    def delete_preset(self):
        """Delete the selected preset"""
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

    def on_merge_changed(self, state):
        """Handle merge names checkbox state change"""
        # Convert Qt.Checked/Qt.Unchecked to boolean if needed
        is_checked = state if isinstance(state, bool) else state == Qt.Checked
        
        # Set enabled state of merged name fields
        self.merged_name_input.setEnabled(is_checked)
        self.merged_default_value.setEnabled(is_checked)
        
        # Only disable individual name fields, don't clear them
        for key in ['First Name', 'Last Name']:
            if key in self.column_inputs:
                self.column_inputs[key].setEnabled(not is_checked)
                self.default_inputs[key].setEnabled(not is_checked)

    def on_merge_address_changed(self, state):
        """Handle merge address checkbox state change"""
        # Convert Qt.Checked/Qt.Unchecked to boolean if needed
        is_checked = state if isinstance(state, bool) else state == Qt.Checked
        
        # Set enabled state of merged address fields
        self.merged_address_input.setEnabled(is_checked)
        self.address_separator_input.setEnabled(is_checked)
        self.province_default_input.setEnabled(is_checked)
        
        # Only disable individual address fields, don't clear them
        address_fields = ['Address', 'City', 'Province', 'Postal Code']
        for field in address_fields:
            if field in self.column_inputs:
                self.column_inputs[field].setEnabled(not is_checked)
                self.default_inputs[field].setEnabled(not is_checked)

    def on_extract_apartment_changed(self, state):
        """Handle extract apartment checkbox state change"""
        is_checked = state if isinstance(state, bool) else state == Qt.Checked
        
        # Enable/disable apartment-related inputs
        self.apartment_name_input.setEnabled(is_checked)
        self.include_apartment_checkbox.setEnabled(is_checked)
        self.filter_apartments_checkbox.setEnabled(is_checked)
        
        # Reset values if unchecked
        if not is_checked:
            self.include_apartment_checkbox.setChecked(False)
            self.filter_apartments_checkbox.setChecked(False)
            self.apartment_name_input.setText("Apartment")

    def on_phone_changed(self, state):
        is_checked = state == Qt.Checked
        self.phone_name_input.setEnabled(is_checked)
        self.phone_default_input.setEnabled(is_checked)

    def on_date_changed(self, state):
        is_checked = state == Qt.Checked
        self.date_name_input.setEnabled(is_checked)
        self.date_picker.setEnabled(is_checked)

    def on_region_filter_changed(self, state):
        """Handle region filter checkbox state change"""
        is_checked = state == Qt.Checked
        for input_field in self.region_inputs.values():
            input_field.setEnabled(is_checked)
            if is_checked and not input_field.text():
                # Set default value if empty when enabled
                region_key = [k for k, v in self.region_inputs.items() if v == input_field][0]
                input_field.setText(region_key)

    def on_custom_sectors_changed(self, state):
        """Handle custom sectors checkbox state change"""
        is_checked = state == Qt.Checked
        for input_field in self.sector_inputs.values():
            input_field.setEnabled(is_checked)
            if is_checked and not input_field.text():
                # Set default value if empty when enabled
                sector_key = [k for k, v in self.sector_inputs.items() if v == input_field][0]
                input_field.setText(sector_key)

    def get_settings(self):
        """Get all settings from the dialog"""
        # Get the merged name default value first
        merged_default = self.merged_default_value.text().strip()
        
        settings = {
            'merge_names': self.merge_checkbox.isChecked(),
            'merged_name': self.merged_name_input.text().strip(),
            'merged_default': merged_default,  # Store this separately
            'column_names': {
                key: self.column_inputs[key].text().strip() 
                for key in self.original_columns.keys()
            },
            'default_values': {}
        }
        
        # Handle default values based on merge state
        if settings['merge_names']:
            settings['default_values'][settings['merged_name']] = merged_default
        else:
            # Store individual field defaults
            for key in self.original_columns.keys():
                col_name = self.column_inputs[key].text().strip()
                default_val = self.default_inputs[key].text().strip()
                if default_val:  # Only store non-empty defaults
                    settings['default_values'][col_name] = default_val
        
        # Add other settings...
        settings.update({
            'merge_address': self.merge_address_checkbox.isChecked(),
            'merged_address_name': self.merged_address_input.text().strip(),
            'address_separator': self.address_separator_input.text(),
            'province_default': self.province_default_input.text().strip(),
            'extract_apartment': self.extract_apartment_checkbox.isChecked(),
            'apartment_column_name': self.apartment_name_input.text().strip(),
            'filter_apartments': self.filter_apartments_checkbox.isChecked(),
            'include_apartment_column': self.include_apartment_checkbox.isChecked(),
            'include_phone': self.include_phone_checkbox.isChecked(),
            'phone_column_name': self.phone_name_input.text().strip(),
            'phone_default': self.phone_default_input.text(),
            'include_date': self.include_date_checkbox.isChecked(),
            'date_column_name': self.date_name_input.text().strip(),
            'date_value': self.date_picker.date().toString('yyyy-MM-dd') if self.include_date_checkbox.isChecked() else None,
            'filter_by_region': self.filter_region_checkbox.isChecked(),
            'region_branch_ids': {
                region: input_field.text()
                for region, input_field in self.region_inputs.items()
                if self.filter_region_checkbox.isChecked()
            }
        })
        
        # Add custom sectors settings
        settings.update({
            'use_custom_sectors': self.use_custom_sectors_checkbox.isChecked(),
            'custom_sector_ids': {
                sector: input_field.text()
                for sector, input_field in self.sector_inputs.items()
                if self.use_custom_sectors_checkbox.isChecked()
            }
        })
        
        # Update region settings to include custom sectors
        settings.update({
            'filter_by_region': self.filter_region_checkbox.isChecked(),
            'region_branch_ids': {
                region: input_field.text()
                for region, input_field in self.region_inputs.items()
                if self.filter_region_checkbox.isChecked() and not self.use_custom_sectors_checkbox.isChecked()
            }
        })
        
        # Add remove accents option
        settings['remove_accents'] = self.remove_accents_checkbox.isChecked()
        
        return settings

    def reset_to_defaults(self):
        """Reset all fields to their default values"""
        # Reset checkboxes
        self.merge_checkbox.setChecked(False)
        self.merge_address_checkbox.setChecked(False)
        self.extract_apartment_checkbox.setChecked(False)
        self.include_phone_checkbox.setChecked(False)
        self.include_date_checkbox.setChecked(False)
        
        # Reset column names to defaults
        default_columns = {
            'First Name': 'First Name',
            'Last Name': 'Last Name',
            'Address': 'Address',
            'City': 'City',
            'Province': 'Province',
            'Postal Code': 'Postal Code'
        }
        
        for key, value in default_columns.items():
            if key in self.column_inputs:
                self.column_inputs[key].setText(value)
                self.default_inputs[key].setText("")
        
        # Reset merged fields
        self.merged_name_input.setText("Full Name")
        self.merged_default_value.setText("À l'occupant")
        self.merged_address_input.setText("Complete Address")
        self.address_separator_input.setText(", ")
        self.province_default_input.setText("QC")
        
        # Reset other fields
        self.apartment_name_input.setText("Apartment")
        self.filter_apartments_checkbox.setChecked(False)
        self.include_apartment_checkbox.setChecked(True)
        self.phone_name_input.setText("Phone")
        self.phone_default_input.setText("")
        self.date_name_input.setText("Date")
        self.date_picker.setDate(QDate.currentDate())
        
        # Reset region settings
        self.filter_region_checkbox.setChecked(False)
        default_regions = {
            'flyer_north_shore': 'flyer_north_shore',
            'flyer_south_shore': 'flyer_south_shore',
            'flyer_montreal': 'flyer_montreal',
            'flyer_laval': 'flyer_laval',
            'flyer_longueuil': 'flyer_longueuil',
            'flyer_unknown': 'flyer_unknown'
        }
        for region_key, default_id in default_regions.items():
            if region_key in self.region_inputs:
                self.region_inputs[region_key].setText(default_id)
                self.region_inputs[region_key].setEnabled(False)

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
        self.log_file = None  # Add this line to store log file path
        
        # Initialize logging in disabled state
        logging.getLogger().handlers = []
        logging.disable(logging.CRITICAL)
        
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
        self.merge_address = False  # Add new property
        self.extract_apartment = False
        self.apartment_column_name = "Apartment"
        self.filter_apartments = False
        self.include_apartment_column = True
        self.include_phone = False
        self.phone_column_name = "Phone"
        self.phone_default = ""
        self.include_date = False
        self.date_column_name = "Date"
        self.date_value = None
        self.filter_by_region = False
        self.region_branch_ids = {}
        self.remove_accents = False  # Initialize remove_accents setting

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
        # Create custom dialog
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle(translations[self.language]['about_title'])
        layout = QVBoxLayout()
        
        # Add text labels
        text_label = QLabel(f"{translations[self.language]['about_text']}\n\nVersion: {VERSION}")
        layout.addWidget(text_label)
        
        # Add link
        link_label = QLabel('<a href="https://github.com/LevSky22/PDF2Excel_AddressConverter">https://github.com/LevSky22/PDF2Excel_AddressConverter</a>')
        link_label.setOpenExternalLinks(True)
        layout.addWidget(link_label)
        
        # Add checkbox
        logging_checkbox = QCheckBox(translations[self.language]['enable_logging'])
        logging_checkbox.setChecked(self.enable_logging)
        layout.addWidget(logging_checkbox)
        
        # Add OK button
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(about_dialog.accept)
        layout.addWidget(button_box)
        
        about_dialog.setLayout(layout)
        
        if about_dialog.exec_() == QDialog.Accepted:
            try:
                new_logging_state = bool(logging_checkbox.isChecked())
                if new_logging_state != self.enable_logging:
                    self.enable_logging = new_logging_state
                    if self.enable_logging:
                        # Just set the flag, don't create log file yet
                        logging.disable(logging.NOTSET)
                    else:
                        # Disable logging when unchecked
                        if logging.getLogger().handlers:
                            logging.getLogger().handlers = []
                        logging.disable(logging.CRITICAL)
            except Exception as e:
                QMessageBox.warning(self, "Warning", f"Error setting logging state: {str(e)}")

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
        if self.merge_names:
            self.conversion_thread.default_values = {
                self.merged_name: self.default_values.get(self.merged_name, "À l'occupant")
            }
        else:
            self.conversion_thread.default_values = {
                k: v for k, v in self.default_values.items() 
                if k in self.column_names.values()
            }
        self.conversion_thread.file_format = file_format
        
        # Add all address merge settings
        self.conversion_thread.merge_address = self.merge_address
        self.conversion_thread.merged_address_name = self.column_names.get('Address', 'Complete Address')
        self.conversion_thread.address_separator = getattr(self, 'address_separator', ', ')
        self.conversion_thread.province_default = getattr(self, 'province_default', 'QC')
        
        # Fix apartment filtering settings
        self.conversion_thread.extract_apartment = self.extract_apartment
        self.conversion_thread.apartment_column_name = self.apartment_column_name
        # Explicitly set filter_apartments from the instance variable
        self.conversion_thread.filter_apartments = self.filter_apartments
        self.conversion_thread.include_apartment_column = self.include_apartment_column
        
        # Add phone and date settings
        self.conversion_thread.include_phone = self.include_phone
        self.conversion_thread.phone_column_name = self.phone_column_name
        self.conversion_thread.phone_default = self.phone_default
        self.conversion_thread.include_date = self.include_date
        self.conversion_thread.date_column_name = self.date_column_name
        self.conversion_thread.date_value = self.date_value
        
        # Add custom sectors settings
        self.conversion_thread.use_custom_sectors = getattr(self, 'use_custom_sectors', False)
        self.conversion_thread.custom_sector_ids = getattr(self, 'custom_sector_ids', {})
        
        # Update region settings
        self.conversion_thread.filter_by_region = self.filter_by_region
        if self.use_custom_sectors:
            self.conversion_thread.region_branch_ids = self.custom_sector_ids
        else:
            self.conversion_thread.region_branch_ids = self.region_branch_ids
        
        # Update remove_accents setting - add explicit logging
        logging.info(f"Setting remove_accents in conversion thread to: {self.remove_accents}")
        self.conversion_thread.remove_accents = self.remove_accents  # Make sure to use instance variable
        
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
        
        # Set current settings explicitly
        dialog.extract_apartment_checkbox.setChecked(self.extract_apartment)
        dialog.filter_apartments_checkbox.setChecked(self.filter_apartments)
        dialog.include_apartment_checkbox.setChecked(self.include_apartment_column)
        dialog.merge_address_checkbox.setChecked(self.merge_address)
        dialog.remove_accents_checkbox.setChecked(self.remove_accents)  # Set current state
        
        # Force update handlers
        dialog.on_extract_apartment_changed(self.extract_apartment)
        dialog.on_merge_address_changed(self.merge_address)
        
        if dialog.exec_() == QDialog.Accepted:
            settings = dialog.get_settings()
            self.merge_names = settings['merge_names']
            self.merged_name = settings['merged_name']
            self.column_names = settings['column_names']
            
            # Fix default values handling
            if self.merge_names:
                # When merging names, ensure we store the default value for the merged column
                self.default_values = {
                    self.merged_name: settings['default_values'].get(self.merged_name, "À l'occupant")
                }
            else:
                # When using separate fields, store defaults for First/Last name
                self.default_values = {
                    self.column_names['First Name']: settings['default_values'].get(self.column_names['First Name'], 'À'),
                    self.column_names['Last Name']: settings['default_values'].get(self.column_names['Last Name'], "l'occupant")
                }
            self.merge_address = settings['merge_address']
            self.merged_address_name = settings['merged_address_name']
            self.address_separator = settings['address_separator']
            self.province_default = settings['province_default']
            
            # Update apartment settings more explicitly
            self.extract_apartment = settings.get('extract_apartment', False)
            self.apartment_column_name = settings.get('apartment_column_name', 'Apartment')
            # Directly set filter_apartments from settings
            self.filter_apartments = settings.get('filter_apartments', False)
            self.include_apartment_column = settings.get('include_apartment_column', True)
            
            # Log the settings for debugging
            if self.enable_logging:
                logging.info(f"Apartment settings: extract={self.extract_apartment}, "
                            f"filter={self.filter_apartments}, "
                            f"include_column={self.include_apartment_column}")
            
            if self.extract_apartment and self.include_apartment_column:
                self.column_names['Apartment'] = self.apartment_column_name
            elif 'Apartment' in self.column_names:
                del self.column_names['Apartment']
            
            self.current_preset = dialog.preset_combo.currentText()
            
            # Phone settings
            self.include_phone = settings['include_phone']
            self.phone_column_name = settings['phone_column_name']
            self.phone_default = settings['phone_default']
            if self.include_phone:
                self.column_names['Phone'] = self.phone_column_name
            elif 'Phone' in self.column_names:
                del self.column_names['Phone']
            
            # Date settings
            self.include_date = settings['include_date']
            self.date_column_name = settings['date_column_name']
            self.date_value = settings['date_value']
            if self.include_date:
                self.column_names['Date'] = self.date_column_name
            elif 'Date' in self.column_names:
                del self.column_names['Date']
            
            # Region settings
            self.filter_by_region = settings['filter_by_region']
            self.region_branch_ids = settings['region_branch_ids']
            
            # Update custom sectors settings
            self.use_custom_sectors = settings.get('use_custom_sectors', False)
            self.custom_sector_ids = settings.get('custom_sector_ids', {})
            
            # Update region settings
            self.filter_by_region = settings.get('filter_by_region', False)
            if self.use_custom_sectors:
                self.region_branch_ids = self.custom_sector_ids
            else:
                self.region_branch_ids = settings.get('region_branch_ids', {})
            self.remove_accents = settings.get('remove_accents', False)
            logging.info(f"Updated remove_accents setting in GUI to: {self.remove_accents}")

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