# ==========================================
# CONFIGURATION
# ==========================================
#
# Need to install:
# - PyQt6 (pip install PyQt6) - GUI
# - pandas (pip install pandas)
# - pillow (pip install pillow)
# - pytesseract (pip install pytesseract)
# - pdf2image (pip install pdf2image)
# - Tesseract OCR (installer available at https://github.com/UB-Mannheim/tesseract/wiki)
# - Poppler (binary available at https://github.com/oschwartz10612/poppler-windows/releases/)
#
import sys
import os
import csv
import json
import shutil
import time
import uuid
import re
from datetime import datetime
from PIL import Image

# External libraries for OCR/PDF
try:
    import pytesseract
    from pdf2image import convert_from_path
except ImportError:
    print("OCR libraries not found. Install pytesseract and pdf2image for full functionality.")
    pytesseract = None
    convert_from_path = None

# PyQt6 Imports
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QTableWidget, QTableWidgetItem, QPushButton, 
    QLineEdit, QLabel, QFileDialog, QMenu, QSplitter, 
    QTabWidget, QTextEdit, QFormLayout, QDialog,
    QMessageBox, QInputDialog, QAbstractItemView, QCheckBox, QFrame,
    QHeaderView, QWidgetAction
)
from PyQt6.QtGui import QPixmap, QAction, QColor, QDesktopServices, QPalette, QCursor
from PyQt6.QtCore import Qt, QSize, QUrl, QEvent, QTimer

# ==========================================
# CONFIGURATION & CONSTANTS
# ==========================================

# WINDOWS USERS: Configure Tesseract/Poppler if needed
# if pytesseract: pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# POPPLER_PATH = r'C:\Program Files\poppler-23.11.0\Library\bin' 
POPPLER_PATH = None 

DEFAULT_CSV_NAME = "contacts.csv"
IMG_FOLDER_NAME = "card_images"
CONFIG_FILE = "config.txt"

CSV_HEADERS = [
    "ID", "First Name", "Last Name", "Company", "Job Title", 
    "E-mail Address", "Mobile Phone", "Business Phone", 
    "Address", "Notes Data", "Image Data"
]

ALL_AVAILABLE_COLS = [
    "First Name", "Last Name", "Company", "Job Title",
    "E-mail Address", "Mobile Phone", "Business Phone", "Address"
]

DEFAULT_CONFIG = {
    "theme": "Dark",
    "working_directory": os.getcwd(),
    "visible_columns": ["First Name", "Last Name", "Company", "E-mail Address", "Mobile Phone"],
    "show_images": True,
    "show_directory_bar": False,
    "colors_dark": {
        "window": "#353535", "window_text": "#ffffff", "base": "#252525", "text": "#ffffff",
        "button": "#353535", "button_text": "#ffffff", "highlight": "#2a82da", "highlight_text": "#000000",
        "input_bg": "#454545"
    },
    "colors_light": {
        "window": "#f0f0f0", "window_text": "#000000", "base": "#ffffff", "text": "#000000",
        "button": "#e0e0e0", "button_text": "#000000", "highlight": "#308cc6", "highlight_text": "#ffffff",
        "input_bg": "#ffffff"
    }
}

# ==========================================
# HELPER CLASSES
# ==========================================

class AspectRatioLabel(QLabel):
    """ Custom Label to keep aspect ratio of pixmap when resizing """
    def __init__(self, parent=None, double_click_callback=None):
        super().__init__(parent)
        self.setScaledContents(False)
        self._pixmap = None
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.double_click_callback = double_click_callback

    def setPixmap(self, p):
        self._pixmap = p
        super().setPixmap(p)
        self.updateScaled()

    def resizeEvent(self, e):
        self.updateScaled()
        super().resizeEvent(e)
    
    def mouseDoubleClickEvent(self, e):
        if self.double_click_callback:
            self.double_click_callback()
        super().mouseDoubleClickEvent(e)

    def updateScaled(self):
        if self._pixmap and not self._pixmap.isNull():
            size = self.size()
            if size.width() > 0 and size.height() > 0:
                scaled = self._pixmap.scaled(size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                super().setPixmap(scaled)

class PopupDialog(QDialog):
    """ A Frameless Popup that closes on focus loss """
    def __init__(self, parent):
        super().__init__(parent, Qt.WindowType.Popup)
        self.setStyleSheet("QDialog { border: 1px solid #888; }")
        
    def focusOutEvent(self, event):
        self.close()
        super().focusOutEvent(event)

class FilterCheckBox(QWidget):
    """ Widget for Persistent Menu Actions """
    def __init__(self, text, checked, callback):
        super().__init__()
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 2, 5, 2)
        self.chk = QCheckBox(text)
        self.chk.setChecked(checked)
        self.chk.toggled.connect(callback)
        layout.addWidget(self.chk)

class ContactEditor(QDialog):
    def __init__(self, parent, contact_data=None):
        super().__init__(parent)
        self.parent_app = parent
        
        if contact_data:
            self.data = json.loads(json.dumps(contact_data))
        else:
            self.data = {k: "" for k in CSV_HEADERS}

        if "ID" not in self.data or not self.data["ID"]: self.data["ID"] = str(uuid.uuid4())
        if not isinstance(self.data.get("Image Data"), list): self.data["Image Data"] = []
        if not isinstance(self.data.get("Notes Data"), list): self.data["Notes Data"] = []

        fname = self.data.get('First Name', 'New')
        self.setWindowTitle(f"Edit Contact - {fname}")
        self.resize(1100, 700)
        self.setup_ui()
        self.apply_local_theme()

    def apply_local_theme(self):
        if self.parent_app.config["theme"] == "Light":
            # IMPROVEMENT 2: White background for Context Menus in Light Mode
            self.setStyleSheet("""
                QMenu { background-color: white; border: 1px solid #ccc; color: black; }
                QMenu::item { background-color: transparent; padding: 4px 20px; }
                QMenu::item:selected { background-color: #e0e0e0; }
                QTabWidget::pane { border: 1px solid #C2C7CB; }
                QTabBar::tab { background: #E0E0E0; color: black; padding: 5px; }
                QTabBar::tab:selected { background: white; font-weight: bold; border-bottom: 2px solid #308cc6; }
            """)
        else:
            self.setStyleSheet("""
                QTabBar::tab { background: #454545; color: white; padding: 5px; }
                QTabBar::tab:selected { background: #666666; font-weight: bold; border-bottom: 2px solid #2a82da; }
            """)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False) 
        layout.addWidget(splitter)

        # --- LEFT SIDE (Images) ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        self.img_tabs = QTabWidget()
        self.img_tabs.setTabsClosable(False) 
        self.img_tabs.tabBarDoubleClicked.connect(self.rename_img_tab)
        self.img_tabs.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.img_tabs.customContextMenuRequested.connect(lambda pos: self.show_tab_menu(pos, self.img_tabs, "img"))
        self.img_tabs.setMinimumWidth(50) 
        
        left_layout.addWidget(self.img_tabs)
        
        btn_add_img = QPushButton("Add Image/PDF")
        btn_add_img.clicked.connect(self.add_image)
        left_layout.addWidget(btn_add_img)
        
        splitter.addWidget(left_widget)

        # --- RIGHT SIDE (Form + Notes) ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        form_layout = QFormLayout()
        self.inputs = {}
        fields = ["First Name", "Last Name", "Company", "Job Title", 
                  "E-mail Address", "Mobile Phone", "Business Phone", "Address"]
        
        for f in fields:
            le = QLineEdit(str(self.data.get(f, "")))
            self.inputs[f] = le
            form_layout.addRow(f + ":", le)
            
        right_layout.addLayout(form_layout)

        right_layout.addWidget(QLabel("Notes:"))
        self.note_tabs = QTabWidget()
        self.note_tabs.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.note_tabs.customContextMenuRequested.connect(lambda pos: self.show_tab_menu(pos, self.note_tabs, "note"))
        self.note_tabs.tabBarDoubleClicked.connect(self.rename_note_tab)
        
        self.note_tabs.currentChanged.connect(self.handle_note_tab_change)
        
        right_layout.addWidget(self.note_tabs)
        splitter.addWidget(right_widget)
        
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        # --- BOTTOM BUTTONS ---
        btn_box = QHBoxLayout()
        btn_delete = QPushButton("Delete Contact")
        btn_delete.setStyleSheet("background-color: #d32f2f; color: white; font-weight: bold;")
        btn_delete.clicked.connect(self.delete_contact)
        
        btn_save = QPushButton("Save")
        btn_save.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        btn_save.clicked.connect(self.save_contact)
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)

        if any(c["ID"] == self.data["ID"] for c in self.parent_app.contacts):
            btn_box.addWidget(btn_delete)
        btn_box.addStretch()
        btn_box.addWidget(btn_save)
        btn_box.addWidget(self.btn_cancel)
        
        layout.addLayout(btn_box)

        self.load_images()
        self.load_notes()

        # Default focus
        self.btn_cancel.setFocus()

    def load_images(self):
        self.img_tabs.clear()
        images = self.data.get("Image Data", [])
        if not images:
            self.img_tabs.addTab(QLabel("No Images"), "None")
        
        for img in images:
            path = img.get("path", "")
            name = img.get("name", "Img")
            lbl = AspectRatioLabel()
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            if os.path.exists(path):
                lbl.setPixmap(QPixmap(path))
            else:
                lbl.setText("Image Missing")
            self.img_tabs.addTab(lbl, name)

    def load_notes(self):
        self.note_tabs.blockSignals(True)
        self.note_tabs.clear()
        notes = self.data.get("Notes Data", [])
        if not notes: 
            notes = [{"name": "General", "content": ""}]
            self.data["Notes Data"] = notes
        
        for note in notes:
            txt = QTextEdit()
            txt.setText(note.get("content", ""))
            self.note_tabs.addTab(txt, note.get("name", "Note"))
        
        self.note_tabs.addTab(QWidget(), "+")
        self.note_tabs.blockSignals(False)

    def add_image(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Images/PDFs", "", "Images (*.png *.jpg *.jpeg);;PDF (*.pdf)")
        if not files: return
        
        base_name, ok = QInputDialog.getText(self, "Tab Name", "Enter new image tab name:", text="Doc")
        if not ok: base_name = "Doc"

        new_entries = []
        for f in files:
            if f.lower().endswith(".pdf") and convert_from_path:
                try:
                    pil_imgs = convert_from_path(f, poppler_path=POPPLER_PATH)
                    for i, img in enumerate(pil_imgs):
                        suffix = f" ({i+1})" if len(pil_imgs) > 1 else ""
                        fname = f"doc_{int(time.time())}_{i}.jpg"
                        save_path = os.path.join(self.parent_app.config["working_directory"], IMG_FOLDER_NAME, fname)
                        img.save(save_path, "JPEG")
                        new_entries.append({"name": f"{base_name}{suffix}", "path": save_path})
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"PDF Error: {str(e)}")
            else:
                fname = f"img_{int(time.time())}_{os.path.basename(f)}"
                save_path = os.path.join(self.parent_app.config["working_directory"], IMG_FOLDER_NAME, fname)
                shutil.copy2(f, save_path)
                new_entries.append({"name": base_name, "path": save_path})

        if len(files) > 1 and not files[0].lower().endswith(".pdf"):
             for i, entry in enumerate(new_entries):
                 entry["name"] = f"{base_name} ({i+1})"

        self.data["Image Data"].extend(new_entries)
        self.load_images()
        self.img_tabs.setCurrentIndex(self.img_tabs.count()-1)

    def handle_note_tab_change(self, index):
        if index == self.note_tabs.count() - 1:
            self.save_current_notes_to_data()
            new_name = datetime.now().strftime("%m/%d/%Y")
            self.data["Notes Data"].append({"name": new_name, "content": ""})
            self.load_notes()
            self.note_tabs.setCurrentIndex(len(self.data["Notes Data"]) - 1)

    def save_current_notes_to_data(self):
        notes = []
        for i in range(self.note_tabs.count() - 1):
            name = self.note_tabs.tabText(i)
            widget = self.note_tabs.widget(i)
            notes.append({"name": name, "content": widget.toPlainText()})
        self.data["Notes Data"] = notes

    def rename_img_tab(self, index):
        old_name = self.img_tabs.tabText(index)
        new_name, ok = QInputDialog.getText(self, "Rename Tab", "New Name:", text=old_name)
        if ok and new_name:
            self.img_tabs.setTabText(index, new_name)
            if index < len(self.data["Image Data"]):
                self.data["Image Data"][index]["name"] = new_name

    def rename_note_tab(self, index):
        if index == self.note_tabs.count() - 1: return
        old_name = self.note_tabs.tabText(index)
        new_name, ok = QInputDialog.getText(self, "Rename Tab", "New Name:", text=old_name)
        if ok and new_name:
            self.note_tabs.setTabText(index, new_name)
            self.save_current_notes_to_data()

    def show_tab_menu(self, pos, tab_widget, type_):
        index = tab_widget.tabBar().tabAt(pos)
        if index == -1: return
        if type_ == "note" and index == tab_widget.count() - 1: return

        menu = QMenu()
        action_rename = menu.addAction("Rename")
        action_delete = menu.addAction("Delete")
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
            }
            QMenu::item::selected {
                background-color: #f3f3f3;
                color: black;
            }
        """)
        action = menu.exec(tab_widget.mapToGlobal(pos))
        
        if action == action_rename:
            if type_ == "img": self.rename_img_tab(index)
            else: self.rename_note_tab(index)
        elif action == action_delete:
            if type_ == "img": self.delete_img_tab(index)
            else: 
                self.save_current_notes_to_data()
                del self.data["Notes Data"][index]
                self.load_notes()

    def delete_img_tab(self, index):
        if index < len(self.data["Image Data"]):
            path = self.data["Image Data"][index]["path"]
            if os.path.exists(path):
                try: os.remove(path)
                except: pass
            del self.data["Image Data"][index]
            self.load_images()

    def delete_contact(self):
        if QMessageBox.question(self, 'Confirm Delete', "Are you sure you want to delete this contact?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.parent_app.delete_contact_by_id(self.data["ID"])
            self.close()

    def save_contact(self):
        for field, widget in self.inputs.items():
            self.data[field] = widget.text().strip()
        self.save_current_notes_to_data()
        self.parent_app.save_contact_data(self.data)
        self.close()


class RolodexApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Rolodex")
        self.resize(1200, 800)
        
        self.config = DEFAULT_CONFIG.copy()
        self.load_config()
        
        self.contacts = []
        self.active_filters = {} 

        self.ensure_directories()
        self.apply_theme()
        
        self.setup_ui()
        self.load_data()
        self.refresh_table()
        
        # IMPROVEMENT 3: Auto-fit columns at startup
        for i in range(2, self.table.columnCount()):
            self.table.resizeColumnToContents(i)
        self.table.setColumnWidth(0, 50) 
        self.table.setColumnWidth(1, 100) # Improvement 9: Less default width for Image

    # ==========================
    # DATA HANDLING (CRITICAL)
    # ==========================
    def load_data(self):
        path = os.path.join(self.config["working_directory"], DEFAULT_CSV_NAME)
        self.contacts = []
        if os.path.exists(path):
            with open(path, mode='r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try: row["Image Data"] = json.loads(row.get("Image Data", "[]"))
                    except: row["Image Data"] = []
                    try: row["Notes Data"] = json.loads(row.get("Notes Data", "[]"))
                    except: row["Notes Data"] = []
                    self.contacts.append(row)
    
    def save_data_to_disk(self):
        path = os.path.join(self.config["working_directory"], DEFAULT_CSV_NAME)
        export_list = []
        for c in self.contacts:
            copy_c = c.copy()
            copy_c["Image Data"] = json.dumps(c.get("Image Data", []))
            copy_c["Notes Data"] = json.dumps(c.get("Notes Data", []))
            export_list.append(copy_c)
            
        with open(path, mode='w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
            writer.writeheader()
            writer.writerows(export_list)

    def save_contact_data(self, contact_data):
        cid = contact_data["ID"]
        existing = next((i for i, c in enumerate(self.contacts) if c["ID"] == cid), None)
        if existing is not None:
            self.contacts[existing] = contact_data
        else:
            self.contacts.append(contact_data)
        self.save_data_to_disk()
        self.refresh_table()

    def delete_contact_by_id(self, cid):
        self.contacts = [c for c in self.contacts if c["ID"] != cid]
        self.save_data_to_disk()
        self.refresh_table()

    def delete_selected(self):
        ids = self.get_selected_ids()
        if not ids: return
        reply = QMessageBox.question(self, 'Delete', f"Delete {len(ids)} contacts?", 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.contacts = [c for c in self.contacts if c["ID"] not in ids]
            self.save_data_to_disk()
            self.refresh_table()

    def edit_selected(self):
        ids = self.get_selected_ids()
        for cid in ids:
            contact = next((c for c in self.contacts if c["ID"] == cid), None)
            if contact:
                self.open_editor_data(contact)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    saved = json.load(f)
                    self.config.update(saved)
            except: pass

    def save_config(self):
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=4)

    def closeEvent(self, event):
        self.save_config()
        super().closeEvent(event)

    def ensure_directories(self):
        img_dir = os.path.join(self.config["working_directory"], IMG_FOLDER_NAME)
        if not os.path.exists(img_dir): os.makedirs(img_dir)

    def apply_theme(self):
        is_dark = self.config["theme"] == "Dark"
        palette = QPalette()
        c = self.config["colors_dark"] if is_dark else self.config["colors_light"]
        
        palette.setColor(QPalette.ColorRole.Window, QColor(c["window"]))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(c["window_text"]))
        palette.setColor(QPalette.ColorRole.Base, QColor(c["base"]))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(c["window"]))
        palette.setColor(QPalette.ColorRole.Text, QColor(c["text"]))
        palette.setColor(QPalette.ColorRole.Button, QColor(c["button"]))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(c["button_text"]))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(c["highlight"]))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(c["highlight_text"]))
        
        QApplication.instance().setPalette(palette)
        
        style = """
            QPushButton::menu-indicator { 
                width: 0px; 
                image: none; 
            }
        """
        
        if not is_dark:
            style += """
                QMenu { background-color: white; border: 1px solid #ccc; color: black; margin: 2px; }
                QMenu::item:selected { background-color: #666666; }
                QCheckBox { color: black; }
                QDialog { background-color: #f0f0f0; color: black; }
                QLabel { color: black; }
                QLineEdit { background-color: white; color: black; }
            """
        
        self.setStyleSheet(style)
        
        if hasattr(self, 'search_bar'):
             self.search_bar.setStyleSheet(f"background: {c['input_bg']}; color: {c['text']}; border: 1px solid #888;")
             if hasattr(self, 'lbl_dir'):
                 self.lbl_dir.setStyleSheet(f"background: {c['window']}; color: {c['text']}; border: none;")

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- TOOLBAR ---
        toolbar = QHBoxLayout()
        
        btn_add = QPushButton("+")
        btn_add.setStyleSheet("""
            QPushButton { 
                background-color: #4CAF50; color: white; font-weight: bold; font-size: 16px; 
                text-align: center; padding: 0px; margin: 0px;
            }
            QPushButton::menu-indicator { image: none; }
        """)
        btn_add.setFixedSize(40, 40)
        
        add_menu = QMenu(self)
        # Improvement 1: Wide menu
        add_menu.setMinimumWidth(200)
        add_menu.addAction("Manual Creation", lambda: self.add_new_contact())
        add_menu.addAction("From Image", lambda: self.add_from_file(False))
        add_menu.addAction("From PDF", lambda: self.add_from_file(True))
        btn_add.setMenu(add_menu)
        toolbar.addWidget(btn_add)

        toolbar.addWidget(QLabel("Search:"))
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Filter contacts...")
        self.search_bar.textChanged.connect(self.refresh_table)
        
        c = self.config["colors_dark"] if self.config["theme"] == "Dark" else self.config["colors_light"]
        self.search_bar.setStyleSheet(f"background: {c['input_bg']}; color: {c['text']}; border: 1px solid #888;")
        
        toolbar.addWidget(self.search_bar)

        self.lbl_dir_label = QLabel("|  Dir:")
        self.lbl_dir = QLineEdit(self.config["working_directory"])
        self.lbl_dir.setReadOnly(True)
        self.lbl_dir.setStyleSheet(f"background: {c['window']}; color: {c['text']}; border: none;")
        
        toolbar.addWidget(self.lbl_dir_label)
        toolbar.addWidget(self.lbl_dir)
        
        if not self.config.get("show_directory_bar", False):
            self.lbl_dir_label.hide()
            self.lbl_dir.hide()

        toolbar.addStretch()

        self.btn_del_selected = QPushButton("Delete Selected")
        self.btn_del_selected.setStyleSheet("background-color: #e53935; color: white; font-weight: bold;")
        self.btn_del_selected.clicked.connect(self.delete_selected)
        self.btn_del_selected.hide()
        toolbar.addWidget(self.btn_del_selected)

        self.btn_edit_selected = QPushButton("Edit Selected")
        self.btn_edit_selected.setStyleSheet("background-color: #757575; color: white; font-weight: bold;")
        self.btn_edit_selected.clicked.connect(self.edit_selected)
        self.btn_edit_selected.hide()
        toolbar.addWidget(self.btn_edit_selected)

        self.btn_settings = QPushButton("☰")
        self.btn_settings.setFixedSize(40, 40)
        self.btn_settings.setStyleSheet("QPushButton::menu-indicator { image: none; }")
        
        self.settings_menu = QMenu(self)
        self.btn_settings.setMenu(self.settings_menu)
        self.settings_menu.aboutToShow.connect(self.populate_settings_menu)
        
        toolbar.addWidget(self.btn_settings)
        main_layout.addLayout(toolbar)

        # --- TABLE ---
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)
        self.table.horizontalHeader().sectionResized.connect(self.on_column_resized)
        
        self.table.itemDoubleClicked.connect(self.on_double_click)
        self.table.itemChanged.connect(self.on_item_changed)
        main_layout.addWidget(self.table)
        
        self.apply_theme()

    def populate_settings_menu(self):
        self.settings_menu.clear()
        
        act_cols = QAction("Column visibility...", self.settings_menu)
        act_cols.triggered.connect(self.open_column_popup)
        self.settings_menu.addAction(act_cols)

        self.settings_menu.addSeparator()

        act_dir = QAction("Directory...", self.settings_menu)
        act_dir.triggered.connect(self.open_directory_popup)
        self.settings_menu.addAction(act_dir)

        self.settings_menu.addSeparator()

        is_dark = self.config["theme"] == "Dark"
        theme_name = "Light" if is_dark else "Dark"
        symbol = "☀" if is_dark else "☾" 
        act_theme = QAction(f"Theme: {theme_name} {symbol}", self.settings_menu)
        act_theme.triggered.connect(self.toggle_theme)
        self.settings_menu.addAction(act_theme)

    # --- POPUPS ---
    def open_column_popup(self):
        popup = PopupDialog(self)
        popup.setWindowTitle("Columns")
        layout = QVBoxLayout(popup)
        
        chk_img = QCheckBox("Show Images")
        chk_img.setChecked(self.config["show_images"])
        chk_img.toggled.connect(self.toggle_images)
        layout.addWidget(chk_img)
        
        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine); layout.addWidget(line)

        for col in ALL_AVAILABLE_COLS:
            chk = QCheckBox(col)
            chk.setChecked(col in self.config["visible_columns"])
            chk.toggled.connect(lambda checked, c=col: self.toggle_column(c, checked))
            layout.addWidget(chk)
            
        cursor_pos = QCursor.pos()
        popup.move(cursor_pos.x(), cursor_pos.y())
        popup.show()

    def open_directory_popup(self):
        popup = PopupDialog(self)
        popup.setWindowTitle("Directory")
        screen_width = QApplication.primaryScreen().size().width()
        popup.setMinimumWidth(min(800, screen_width - 100))
        
        layout = QVBoxLayout(popup)
        layout.addWidget(QLabel("Current Directory:"))
        
        self.dir_edit_popup = QLineEdit(self.config["working_directory"])
        layout.addWidget(self.dir_edit_popup)
        
        chk_show = QCheckBox("Display on toolbar")
        chk_show.setChecked(self.config.get("show_directory_bar", False))
        chk_show.toggled.connect(self.toggle_directory_bar)
        layout.addWidget(chk_show)
        
        btn = QPushButton("Choose Directory...")
        btn.clicked.connect(lambda: self.browse_directory(popup))
        layout.addWidget(btn)
        
        cursor_pos = QCursor.pos()
        popup.move(cursor_pos.x() - 400, cursor_pos.y())
        popup.show()

    # --- ACTIONS ---
    def toggle_theme(self):
        self.config["theme"] = "Light" if self.config["theme"] == "Dark" else "Dark"
        self.apply_theme()
        self.save_config()

    def toggle_images(self, checked):
        self.config["show_images"] = checked
        self.table.setColumnHidden(1, not checked)
        self.adjust_row_heights()
        self.save_config()

    def toggle_directory_bar(self, checked):
        self.config["show_directory_bar"] = checked
        if checked:
            self.lbl_dir_label.show()
            self.lbl_dir.show()
        else:
            self.lbl_dir_label.hide()
            self.lbl_dir.hide()
        self.save_config()

    def toggle_column(self, col_name, checked):
        if checked:
            if col_name not in self.config["visible_columns"]:
                self.config["visible_columns"].append(col_name)
        else:
            if col_name in self.config["visible_columns"]:
                self.config["visible_columns"].remove(col_name)
        self.refresh_table_structure()
        self.save_config()

    def browse_directory(self, popup):
        d = QFileDialog.getExistingDirectory(self, "Select Directory")
        if d:
            self.config["working_directory"] = d
            self.dir_edit_popup.setText(d)
            self.lbl_dir.setText(d)
            self.ensure_directories()
            self.load_data()
            self.refresh_table()
            self.save_config()
            popup.close()

    def on_header_clicked(self, logicalIndex):
        if logicalIndex == 0:
            self.toggle_select_all()
        elif logicalIndex > 1:
            header_item = self.table.horizontalHeaderItem(logicalIndex)
            col_name = header_item.text()
            
            menu = QMenu(self)
            menu.addAction("Sort Ascending", lambda: self.sort_table(logicalIndex, Qt.SortOrder.AscendingOrder))
            menu.addAction("Sort Descending", lambda: self.sort_table(logicalIndex, Qt.SortOrder.DescendingOrder))
            
            filter_menu = menu.addMenu("Filter")
            
            values = set()
            for row in range(self.table.rowCount()):
                item = self.table.item(row, logicalIndex)
                if item: values.add(item.text())
            
            filter_menu.addAction("Clear Filter", lambda: self.clear_filter(col_name))
            filter_menu.addSeparator()
            
            sorted_vals = sorted(list(values))
            current_filters = self.active_filters.get(col_name, [])
            all_allowed = col_name not in self.active_filters
            
            for val in sorted_vals:
                lbl = val if val else "(Blank)"
                # IMPROVEMENT 4: Persistent Filter Menu using QWidgetAction
                is_checked = all_allowed or (val in current_filters)
                
                # Callback wrapper to pass context
                def make_callback(c, v):
                    return lambda checked: self.toggle_filter(c, v, checked)
                
                chk_widget = FilterCheckBox(lbl, is_checked, make_callback(col_name, val))
                action = QWidgetAction(filter_menu)
                action.setDefaultWidget(chk_widget)
                filter_menu.addAction(action)
                
            menu.exec(QCursor.pos())
            self.table.setSortingEnabled(False)

    def sort_table(self, col_index, order):
        self.table.setSortingEnabled(True)
        self.table.sortItems(col_index, order)
        self.table.setSortingEnabled(False)

    def toggle_filter(self, col_name, value, checked):
        if col_name not in self.active_filters:
            col_idx = -1
            for i in range(self.table.columnCount()):
                if self.table.horizontalHeaderItem(i).text() == col_name:
                    col_idx = i
                    break
            if col_idx == -1: return

            all_vals = set()
            for row in range(self.table.rowCount()):
                item = self.table.item(row, col_idx)
                if item: all_vals.add(item.text())
            self.active_filters[col_name] = list(all_vals)
        
        if checked:
            if value not in self.active_filters[col_name]:
                self.active_filters[col_name].append(value)
        else:
            if value in self.active_filters[col_name]:
                self.active_filters[col_name].remove(value)
        
        self.refresh_table_data()

    def clear_filter(self, col_name):
        if col_name in self.active_filters:
            del self.active_filters[col_name]
        self.refresh_table_data()

    def toggle_select_all(self):
        if self.table.rowCount() == 0: return
        
        first_widget = self.table.cellWidget(0, 0)
        if first_widget:
            first_chk = first_widget.findChild(QCheckBox)
            new_state = not first_chk.isChecked()
            
            for row in range(self.table.rowCount()):
                widget = self.table.cellWidget(row, 0)
                if widget:
                    chk = widget.findChild(QCheckBox)
                    chk.setChecked(new_state)

    def refresh_table_structure(self):
        headers = ["✔", "Image"] + self.config["visible_columns"]
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        
        self.table.setColumnHidden(1, not self.config["show_images"])
        self.table.setColumnWidth(0, 50)
        self.table.setColumnWidth(1, 150)
        self.refresh_table_data()

    def refresh_table(self):
        if self.table.columnCount() != 2 + len(self.config["visible_columns"]):
            self.refresh_table_structure()
        else:
            self.refresh_table_data()

    def refresh_table_data(self):
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        self.table.clearContents()
        self.table.setSortingEnabled(False) 
        
        query = self.search_bar.text().lower()
        
        filtered = []
        for c in self.contacts:
            if query and query not in "".join([str(v) for v in c.values()]).lower():
                continue
            
            match = True
            for col, allowed in self.active_filters.items():
                val = c.get(col, "")
                if val not in allowed:
                    match = False
                    break
            if match:
                filtered.append(c)
        
        self.table.setRowCount(len(filtered))
        
        for row, contact in enumerate(filtered):
            # 0: Checkbox
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0,0,0,0)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk = QCheckBox()
            chk.setProperty("cid", contact["ID"])
            chk.stateChanged.connect(self.update_batch_buttons)
            chk_layout.addWidget(chk)
            self.table.setCellWidget(row, 0, chk_widget)
            
            # 1: Image
            # IMPROVEMENT 1: Double click image logic passed to label
            img_label = AspectRatioLabel(double_click_callback=lambda cid=contact["ID"]: self.open_editor_by_id(cid))
            img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            img_data = contact.get("Image Data", [])
            if img_data and os.path.exists(img_data[0]["path"]):
                pix = QPixmap(img_data[0]["path"])
                if not pix.isNull():
                    img_label.setPixmap(pix)
            self.table.setCellWidget(row, 1, img_label)
            
            # Data Cols
            for col_idx, key in enumerate(self.config["visible_columns"]):
                val = contact.get(key, "")
                item = QTableWidgetItem(str(val))
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                if key == "E-mail Address" and val:
                    item.setForeground(QColor("blue"))
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)
                self.table.setItem(row, 2 + col_idx, item)

        self.adjust_row_heights()
        self.update_batch_buttons()
        self.table.setSortingEnabled(True)
        self.table.blockSignals(False)

    def on_column_resized(self, logicalIndex, oldSize, newSize):
        if logicalIndex == 1 and self.config["show_images"]:
            self.adjust_row_heights()

    def adjust_row_heights(self):
        if not self.config["show_images"]:
            height = 30
        else:
            width = self.table.columnWidth(1)
            height = int(width / 1.58)
            if height < 40: height = 40
            
        for i in range(self.table.rowCount()):
            self.table.setRowHeight(i, height)

    def on_item_changed(self, item):
        pass

    def get_selected_ids(self):
        ids = []
        for row in range(self.table.rowCount()):
            widget = self.table.cellWidget(row, 0)
            if widget:
                chk = widget.findChild(QCheckBox)
                if chk and chk.isChecked():
                    ids.append(chk.property("cid"))
        return ids

    def update_batch_buttons(self):
        ids = self.get_selected_ids()
        if ids:
            self.btn_del_selected.show()
            self.btn_edit_selected.show()
            self.btn_del_selected.setText(f"Delete Selected ({len(ids)})")
            self.btn_edit_selected.setText(f"Edit Selected ({len(ids)})")
        else:
            self.btn_del_selected.hide()
            self.btn_edit_selected.hide()

    def add_new_contact(self):
        self.open_editor_data(None)

    def add_from_file(self, is_pdf):
        files, _ = QFileDialog.getOpenFileNames(self, "Select File", "", 
                                                "PDF (*.pdf)" if is_pdf else "Images (*.png *.jpg *.jpeg)")
        if not files: return
        
        new_data = {k: "" for k in CSV_HEADERS}
        new_data["ID"] = str(uuid.uuid4())
        new_data["Image Data"] = []
        new_data["Notes Data"] = []
        base_name = "Import"
        
        for f in files:
            if is_pdf:
                try:
                    imgs = convert_from_path(f, poppler_path=POPPLER_PATH)
                    for i, img in enumerate(imgs):
                        fname = f"doc_{int(time.time())}_{i}.jpg"
                        path = os.path.join(self.config["working_directory"], IMG_FOLDER_NAME, fname)
                        img.save(path, "JPEG")
                        new_data["Image Data"].append({"name": f"{base_name} {i+1}", "path": path})
                        if i == 0 and pytesseract:
                             text = pytesseract.image_to_string(img)
                             new_data = self.heuristic_parse(text, new_data)
                             new_data["Notes Data"].append({"name": "OCR", "content": text})
                except: pass
            else:
                fname = f"img_{int(time.time())}_{os.path.basename(f)}"
                path = os.path.join(self.config["working_directory"], IMG_FOLDER_NAME, fname)
                shutil.copy2(f, path)
                new_data["Image Data"].append({"name": base_name, "path": path})
                try:
                    if pytesseract:
                        img = Image.open(path)
                        text = pytesseract.image_to_string(img)
                        new_data = self.heuristic_parse(text, new_data)
                        new_data["Notes Data"].append({"name": "OCR", "content": text})
                except: pass

        self.open_editor_data(new_data)

    def heuristic_parse(self, text, data):
        emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', text)
        if emails: data["E-mail Address"] = emails[0]
        phones = re.findall(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
        if phones: data["Mobile Phone"] = phones[0]
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if lines:
            parts = lines[0].split()
            if len(parts) > 1:
                data["First Name"] = parts[0]
                data["Last Name"] = parts[-1]
            else:
                data["First Name"] = lines[0]
        return data

    def open_editor_data(self, data):
        editor = ContactEditor(self, data)
        editor.exec()

    def open_editor_by_id(self, cid):
        contact = next((c for c in self.contacts if c["ID"] == cid), None)
        if contact:
            self.open_editor_data(contact)

    def on_double_click(self, item):
        row = item.row()
        col = item.column()
        
        # Get ID from checkbox widget in col 0
        widget = self.table.cellWidget(row, 0)
        # Fix: Widget might be none if clicked while loading
        if not widget: return

        chk = widget.findChild(QCheckBox)
        cid = chk.property("cid")
        
        # 0=Check, 1=Image
        if col == 1:
             self.open_editor_by_id(cid)
             return

        if col >= 2:
            key = self.config["visible_columns"][col-2]
            if key == "E-mail Address" and item.text():
                 QDesktopServices.openUrl(QUrl(f"mailto:{item.text()}"))
                 return
        
        self.open_editor_by_id(cid)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RolodexApp()
    window.show()
    sys.exit(app.exec())
