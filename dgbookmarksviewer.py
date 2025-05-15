# -*- coding: utf-8 -*-
import sys
import os
import json
import shutil
from xml.etree import ElementTree as ET
import logging
import subprocess
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QListWidget, QListWidgetItem,
    QLabel, QDialog, QPushButton, QFileDialog, QMenu, QMessageBox, QTextEdit, QSplitter,
    QAbstractItemView, QMenuBar, QSlider, QMainWindow, QSystemTrayIcon, QStyledItemDelegate, QStyle,
    QRadioButton, QComboBox, QButtonGroup, QDialogButtonBox
)
from PySide6.QtCore import (
    Qt, QSize, QPoint, QSettings, QStandardPaths, QRect, Signal, Slot,
    QItemSelectionModel # <<<--- ADD THIS
)
from PySide6.QtGui import (
    QAction, QColor, QFont, QGuiApplication, QIcon, QPainter, QTextDocument, QFontMetrics,
    QKeyEvent, QCursor, QTextCharFormat, QTextCursor
)

# --- Application Metadata and Data Folder Setup ---
APP_NAME = "DGBookmarksViewer"
ORG_NAME = "MyLocalScripts"
QApplication.setOrganizationName(ORG_NAME)
QApplication.setApplicationName(APP_NAME)
APP_DATA_DIR = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
if not APP_DATA_DIR:
    # Fallback if AppDataLocation is not writable/available
    if getattr(sys, 'frozen', False): # PyInstaller bundle
        APP_DATA_DIR = os.path.join(os.path.dirname(sys.executable), f".{APP_NAME}_data")
    else: # Running as script
        APP_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), f".{APP_NAME}_data")
    logging.warning(f"AppDataLocation fallback: {APP_DATA_DIR}")
else:
    logging.info(f"AppDataLocation: {APP_DATA_DIR}")

LOG_DIR = os.path.join(APP_DATA_DIR, "logs")
CONFIG_DIR = os.path.join(APP_DATA_DIR, "configs")
BOOKMARKS_COPY_DIR = os.path.join(APP_DATA_DIR, "bookmarks")
ICON_DIR = os.path.join(APP_DATA_DIR, "icons")
HELP_DIR = os.path.join(APP_DATA_DIR, "help")

# Create directories if they don't exist
for directory in [APP_DATA_DIR, LOG_DIR, CONFIG_DIR, BOOKMARKS_COPY_DIR, ICON_DIR, HELP_DIR]:
    try:
        os.makedirs(directory, exist_ok=True)
    except OSError as e:
        logging.error(f"Failed to create directory {directory}: {e}")

LOG_FILE = os.path.join(LOG_DIR, "bookmark_viewer.log")
BOOKMARK_ACTIONS_LOG = os.path.join(LOG_DIR, "bookmark_actions.log")
SETTINGS_FILE = os.path.join(CONFIG_DIR, "settings.json")
USAGE_COUNTS_FILE = os.path.join(CONFIG_DIR, "usage_counts.json")
LAST_BOOKMARKS_COPY = os.path.join(BOOKMARKS_COPY_DIR, "last_bookmarks_copy.xml")
MAIN_ICON_FILE = os.path.join(ICON_DIR, "app_icon.ico")
TRAY_ICON_FILE = os.path.join(ICON_DIR, "tray_icon.ico")
HELP_LOCATIONS_FILE = os.path.join(HELP_DIR, "help_locations.txt")

# Define default DataGrip path (adjust if necessary)
DEFAULT_DATAGRIP_PATH = r"C:\Users\cfriedberg\AppData\Local\JetBrains\DataGrip 2024.1.4\bin\datagrip64.exe"

# --- Logging Setup ---
def setup_logging():
    log_format = '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            datefmt=date_format,
            handlers=[
                logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8'),
                logging.StreamHandler(sys.stdout) # Also log to console
            ],
            force=True # Override any existing config
        )
    except Exception as e:
        print(f"FATAL: Logging setup failed: {e}", file=sys.stderr)
        # No logging available here, just print

    logging.info("="*20 + f" {APP_NAME} Start " + "="*20)
    logging.info(f"Application Data Directory: {APP_DATA_DIR}")
    logging.info(f"Main Icon Path (Expected): {MAIN_ICON_FILE}")
    logging.info(f"Tray Icon Path (Expected): {TRAY_ICON_FILE}")

setup_logging() # Initialize logging immediately

# --- Configuration Management ---
class AppSettings:
    def __init__(self):
        self.settings_path = SETTINGS_FILE
        self.settings = {}
        self.load_settings()

    def load_settings(self):
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                logging.info(f"Settings loaded from {self.settings_path}")
            except json.JSONDecodeError as e:
                logging.error(f"Error decoding settings JSON from {self.settings_path}: {e}", exc_info=True)
                self.settings = {} # Reset to defaults on decode error
            except Exception as e:
                logging.error(f"Error loading settings from {self.settings_path}: {e}", exc_info=True)
                self.settings = {}
        else:
            logging.info(f"Settings file not found at {self.settings_path}. Using defaults.")
            self.settings = {}

        # Ensure essential settings have defaults
        if 'datagrip_path' not in self.settings:
            self.settings['datagrip_path'] = DEFAULT_DATAGRIP_PATH if os.path.exists(DEFAULT_DATAGRIP_PATH) else None
        if 'transparency' not in self.settings:
            self.settings['transparency'] = 1.0 # Default to fully opaque
        if 'font_size' not in self.settings:
            self.settings['font_size'] = "12" # Default font size

    def save_settings(self):
        try:
            os.makedirs(os.path.dirname(self.settings_path), exist_ok=True)
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)
            logging.info(f"Settings saved to {self.settings_path}")
        except Exception as e:
            logging.error(f"Error saving settings to {self.settings_path}: {e}", exc_info=True)

    def get(self, key, default=None):
        return self.settings.get(key, default)

    def set(self, key, value):
        self.settings[key] = value

# --- Usage Count Management ---
class UsageCounts:
    def __init__(self):
        self.counts_path = USAGE_COUNTS_FILE
        self.counts = {}
        self.load_counts()

    def load_counts(self):
        if os.path.exists(self.counts_path):
            try:
                with open(self.counts_path, 'r', encoding='utf-8') as f:
                    self.counts = json.load(f)
                logging.info(f"Usage counts loaded from {self.counts_path}")
            except json.JSONDecodeError as e:
                logging.error(f"Error decoding usage counts JSON from {self.counts_path}: {e}", exc_info=True)
                self.counts = {} # Reset on decode error
            except Exception as e:
                logging.error(f"Error loading usage counts from {self.counts_path}: {e}", exc_info=True)
                self.counts = {}
        else:
            logging.info(f"Usage counts file not found at {self.counts_path}. Starting fresh.")
            self.counts = {}

    def save_counts(self):
        try:
            os.makedirs(os.path.dirname(self.counts_path), exist_ok=True)
            with open(self.counts_path, 'w', encoding='utf-8') as f:
                json.dump(self.counts, f, indent=4, ensure_ascii=False)
            logging.info(f"Usage counts saved to {self.counts_path}")
        except Exception as e:
            logging.error(f"Error saving usage counts to {self.counts_path}: {e}", exc_info=True)

    def increment_count(self, bid):
        if bid is None:
            logging.warning("Attempted to increment count for None bookmark ID.")
            return
        bid_str = str(bid) # Ensure key is string for JSON compatibility
        self.counts[bid_str] = self.counts.get(bid_str, 0) + 1
        logging.debug(f"Incremented count for '{bid_str}' to {self.counts[bid_str]}")

    def get_count(self, bid):
        if bid is None:
            return 0
        return self.counts.get(str(bid), 0)

    def clear_counts(self):
        self.counts = {}
        logging.info("All usage counts cleared.")
        # Note: save_counts() needs to be called explicitly after clearing if persistence is desired immediately.

# --- Helper Functions ---
def parse_bookmarks_xml(file_path):
    bookmarks = []
    if not os.path.exists(file_path):
        logging.warning(f"Bookmarks XML file not found: {file_path}")
        return [] # Return empty list if file doesn't exist

    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        # Find the BookmarkManager component, or search from root if not found
        bm_comp = root.find("./component[@name='BookmarkManager']")
        search_root = bm_comp if bm_comp is not None else root
        bm_states = search_root.findall(".//BookmarkState") # Find all BookmarkState elements anywhere under search_root

        if not bm_states:
            logging.warning(f"No 'BookmarkState' elements found within the XML structure in {file_path}.")
            return []

        for state in bm_states:
            url, line, desc = '', '', ''
            # Try finding description via option first
            desc_opt = state.find('option[@name="description"]')
            if desc_opt is not None:
                desc = desc_opt.get('value', '')

            # Try finding url/line via attributes element (newer format?)
            attrs = state.find('attributes')
            if attrs is not None:
                url_el = attrs.find("entry[@key='url']")
                line_el = attrs.find("entry[@key='line']")
                url = url_el.get('value') if url_el is not None else ''
                line = line_el.get('value') if line_el is not None else ''
            else: # Fallback to option elements (older format?)
                url_opt = state.find('option[@name="url"]')
                line_opt = state.find('option[@name="line"]')
                url = url_opt.get('value','') if url_opt is not None else ''
                line = line_opt.get('value','') if line_opt is not None else ''

            # Validate essential data
            if not url or not line or not desc:
                logging.warning(f"Skipping incomplete bookmark entry: URL='{url}', Line='{line}', Desc='{desc}'")
                continue

            title = desc # Use description as the primary display title
            fname = os.path.basename(url.replace("file://", "")) if url else "Unknown File" # Extract filename
            full_text = f"{desc} (File: {fname}, Line: {line})" # For logging/tooltip maybe
            bid = f"{url}|{line}" # Create a unique ID based on URL and line

            bookmarks.append({
                'url': url,
                'line': line,
                'description': desc,
                'title': title,
                'full_text': full_text,
                'id': bid,
                'count': 0 # Initialize count, will be updated later
            })

        logging.info(f"Successfully parsed {len(bookmarks)} bookmarks from {file_path}")
        return bookmarks

    except ET.ParseError as e:
        logging.error(f"XML Parse Error in {file_path}: {e}")
        QMessageBox.critical(None, "XML Parsing Error", f"Failed to parse the bookmarks XML file:\n{e}\n\nFile: {file_path}")
        return []
    except Exception as e:
        logging.error(f"Unexpected error parsing XML file {file_path}: {e}", exc_info=True)
        QMessageBox.critical(None, "XML Parsing Error", f"An unexpected error occurred while parsing the XML file:\n{e}")
        return []

def generate_help_locations_text():
    """Generates a formatted string detailing important file locations."""
    help_text = f"""
{APP_NAME} File Locations:
==============================
Application Data Folder:
{os.path.normpath(APP_DATA_DIR)}

Log File:
{os.path.normpath(LOG_FILE)}

Bookmark Actions Log:
{os.path.normpath(BOOKMARK_ACTIONS_LOG)}

Configuration File (Settings):
{os.path.normpath(SETTINGS_FILE)}

Usage Counts File:
{os.path.normpath(USAGE_COUNTS_FILE)}

Copied Bookmarks File (Last Loaded):
{os.path.normpath(LAST_BOOKMARKS_COPY)}

Icon Files Directory:
{os.path.normpath(ICON_DIR)}
(Place app_icon.ico and tray_icon.ico here)

Note: Icons must be in .ico format. If missing or invalid,
system defaults or fallbacks will be attempted.
"""
    try:
        os.makedirs(os.path.dirname(HELP_LOCATIONS_FILE), exist_ok=True)
        with open(HELP_LOCATIONS_FILE, 'w', encoding='utf-8') as f:
            f.write(help_text)
        logging.info(f"Help locations file generated/updated: {HELP_LOCATIONS_FILE}")
    except Exception as e:
        logging.error(f"Error writing help locations file {HELP_LOCATIONS_FILE}: {e}", exc_info=True)
    return help_text

# --- Custom Item Delegate for Fixed Height List Items ---
class BookmarkDelegate(QStyledItemDelegate):
    ITEM_PADDING = 8 # Padding around content within the item rect
    COUNT_BOX_WIDTH = 40 # Fixed width for the usage count box
    TEXT_SPACING = 8 # Space between count box and title text
    FIXED_ITEM_HEIGHT_FACTOR = 1.5 # Multiplier for font height to get base item height

    def calculate_fixed_height(self, font):
        """Calculates the fixed item height based on font metrics."""
        fm = QFontMetrics(font)
        # Base height on font, add padding for top/bottom
        return int((fm.height() * self.FIXED_ITEM_HEIGHT_FACTOR) + (self.ITEM_PADDING * 1.5))

    def paint(self, painter: QPainter, option, index):
        # Get bookmark data from the item's UserRole
        bm = index.data(Qt.ItemDataRole.UserRole)
        if not bm or not isinstance(bm, dict):
            # If data is missing or not a dict, fall back to default painting
            super().paint(painter, option, index)
            return

        count = bm.get('count', 0)
        title = bm.get('title', 'No Title')
        fixed_h = self.calculate_fixed_height(option.font) # Use calculated height

        painter.save() # Save painter state

        # Determine background and text colors based on state
        palette = option.palette
        bg_col = palette.base().color() # Default background
        txt_col = palette.text().color() # Default text

        if option.state & QStyle.StateFlag.State_Selected:
            bg_col = palette.highlight().color()
            txt_col = palette.highlightedText().color()
        elif index.row() % 2 == 0: # Alternating row colors
             bg_col = QColor("#3c3c3c") # Slightly darker
        else:
             bg_col = QColor("#4a4a4a") # Slightly lighter

        # Add hover effect if not selected
        if option.state & QStyle.StateFlag.State_MouseOver and not (option.state & QStyle.StateFlag.State_Selected):
            bg_col = bg_col.lighter(115) # Lighten on hover

        # Fill background and draw bottom border
        painter.fillRect(option.rect, bg_col)
        painter.setPen(palette.midlight().color()) # Use a subtle color for the line
        painter.drawLine(option.rect.bottomLeft(), option.rect.bottomRight())

        # Calculate rects for count box and text area
        fm = option.fontMetrics
        count_h = fm.height() + 4 # Make count box slightly taller than text
        # Center count box vertically
        count_rect = QRect(option.rect.left() + self.ITEM_PADDING // 2,
                           option.rect.top() + (option.rect.height() - count_h) // 2,
                           self.COUNT_BOX_WIDTH, count_h)

        text_l = count_rect.right() + self.TEXT_SPACING
        text_rect = QRect(text_l, option.rect.top(),
                          option.rect.width() - text_l - self.ITEM_PADDING // 2,
                          option.rect.height())
        if text_rect.width() < 0: text_rect.setWidth(0) # Prevent negative width

        # Draw count box background and text
        painter.fillRect(count_rect, QColor("#666666")) # Dark grey background for count
        painter.setPen(QColor("#ffffff")) # White text for count
        painter.setFont(option.font)
        painter.drawText(count_rect, Qt.AlignmentFlag.AlignCenter, str(count))

        # Draw the main bookmark title (elided if too long)
        painter.setPen(txt_col) # Use the determined text color
        painter.setFont(option.font)
        elided_txt = fm.elidedText(title, Qt.TextElideMode.ElideRight, text_rect.width())
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, elided_txt)

        painter.restore() # Restore painter state

    def sizeHint(self, option, index):
        # Return the calculated fixed height and base width
        fixed_h = self.calculate_fixed_height(option.font)
        base_w = super().sizeHint(option, index).width() # Get default width hint
        return QSize(base_w, fixed_h)

# --- Custom List Widget for Key Press Handling ---
class CustomListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setUniformItemSizes(True) # Optimization for fixed item sizes
        self.setLayoutMode(QListWidget.LayoutMode.Batched) # Performance hint
        self.setBatchSize(100) # How many items to process in a batch

        # Store reference to parent window if possible, for callbacks
        if isinstance(parent, FloatingBookmarksWindow):
            self.parent_window = parent
        else:
            self.parent_window = None
            logging.warning("CustomListWidget initialized without a valid FloatingBookmarksWindow parent.")

    def keyPressEvent(self, event: QKeyEvent):
        if not self.parent_window:
            super().keyPressEvent(event) # Default behavior if no parent window
            return

        key = event.key()
        current_item = self.currentItem()

        if key == Qt.Key.Key_Space:
            # Spacebar updates preview pane
            if current_item:
                self.parent_window.update_preview_pane(current_item)
                event.accept() # Consume the event
            else:
                super().keyPressEvent(event) # Default if no item selected
        elif key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            # Enter/Return triggers the main action (copy SQL)
            if current_item:
                self.parent_window.handle_item_action(current_item)
                event.accept() # Consume the event
            else:
                super().keyPressEvent(event) # Default if no item selected
        else:
            # Let QListWidget handle other keys (arrows, page up/down, etc.)
            super().keyPressEvent(event)

# --- Tray Copy Dialog ---
class TrayCopyDialog(QDialog):
    bookmark_selected_for_copy = Signal(dict) # Signal emitting the selected bookmark data

    def __init__(self, bookmarks_data, current_font, parent=None):
        super().__init__(parent)
        self.bookmarks = bookmarks_data
        self.current_font = current_font

        self.setWindowTitle("Copy Bookmark SQL")
        # Try to set a specific icon, fallback to theme
        copy_icon = QIcon.fromTheme("edit-copy", QIcon())
        if not copy_icon.isNull(): self.setWindowIcon(copy_icon)

        self.setMinimumSize(450, 400)
        # Make it a popup that stays on top relative to its parent
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.Popup)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        layout.addWidget(QLabel("Double-click a bookmark to copy its SQL:"))

        self.list_widget = QListWidget()
        self.list_widget.setFont(self.current_font)
        self.list_widget.setUniformItemSizes(True)
        self.list_widget.setLayoutMode(QListWidget.LayoutMode.Batched)
        # Use the same delegate as the main window for consistency
        self.list_widget.setItemDelegate(BookmarkDelegate(self.list_widget))
        layout.addWidget(self.list_widget, 1) # Allow list to stretch

        self.populate_list()
        self.list_widget.itemDoubleClicked.connect(self.handle_item_double_clicked)

        # Buttons at the bottom
        btn_layout = QHBoxLayout()
        close_btn = QPushButton("Cancel")
        close_btn.clicked.connect(self.reject) # Closes the dialog without emitting signal
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def populate_list(self):
        """Fills the list widget with bookmark items."""
        self.list_widget.clear()
        if not self.bookmarks:
            self.list_widget.addItem("No bookmarks available.")
            self.list_widget.setEnabled(False) # Disable if empty
            return

        self.list_widget.setEnabled(True)
        for bm in self.bookmarks:
            if isinstance(bm, dict):
                item = QListWidgetItem()
                item.setData(Qt.ItemDataRole.UserRole, bm) # Store full data
                # Delegate handles text/display
                self.list_widget.addItem(item)
            else:
                logging.warning(f"Skipping invalid item in bookmarks_data for TrayCopyDialog: {bm}")

    @Slot(QListWidgetItem)
    def handle_item_double_clicked(self, item: QListWidgetItem):
        """Emits the selected bookmark data and closes the dialog."""
        bm_data = item.data(Qt.ItemDataRole.UserRole)
        if bm_data and isinstance(bm_data, dict):
            logging.info(f"Bookmark selected via TrayCopyDialog: {bm_data.get('title', 'N/A')}")
            self.bookmark_selected_for_copy.emit(bm_data)
            self.accept() # Closes the dialog successfully
        else:
            logging.warning("Invalid item double-clicked in TrayCopyDialog, data missing or incorrect type.")

    def exec(self):
        """Overrides exec to position the dialog near the cursor."""
        cursor_pos = QCursor.pos()
        screen_geo = QGuiApplication.primaryScreen().availableGeometry()
        hint = self.sizeHint()
        # Use hint if valid, otherwise current size
        w = hint.width() if hint.isValid() else self.width()
        h = hint.height() if hint.isValid() else self.height()

        # Position below and slightly right of cursor
        x, y = cursor_pos.x(), cursor_pos.y() + 20

        # Adjust if popup goes off-screen
        if x + w > screen_geo.right():
            x = screen_geo.right() - w
        if x < screen_geo.left():
            x = screen_geo.left()
        if y + h > screen_geo.bottom():
            # If it goes off bottom, position above cursor instead
            y = cursor_pos.y() - h - 10
        if y < screen_geo.top():
            y = screen_geo.top()

        self.move(x, y)
        return super().exec()

# --- Main Application Window ---
class FloatingBookmarksWindow(QMainWindow):
    """Main window class for the DataGrip Bookmarks Viewer."""
    def __init__(self, settings: AppSettings, usage_counts: UsageCounts):
        super().__init__()
        self.settings = settings
        self.usage_counts = usage_counts
        self.bookmarks = [] # Holds the raw loaded bookmark data {dict}
        self.sorted_bookmarks_cache = [] # Holds currently sorted/filtered list for UI/tray {dict}
        # Try to load path of last used bookmarks file copy
        self.loaded_file_path = self.settings.get('loaded_copy_path', LAST_BOOKMARKS_COPY)
        self.context_item = None # Stores the item clicked for context menu
        self.tray_icon = None # QSystemTrayIcon instance
        self.tray_menu = None # QMenu for tray icon

        self.setWindowTitle(APP_NAME)
        # Default to a standard window frame initially
        self.setWindowFlags(Qt.WindowType.Window)

        self._load_and_apply_geometry()
        self._apply_transparency()
        self.setMinimumSize(600, 400) # Reasonable minimum size

        self.setup_ui() # Create UI elements
        self.create_actions() # Create QActions
        self.add_actions_to_menus() # Populate menubar
        self.init_context_menu() # Setup right-click menu for list
        self.connect_signals() # Connect UI signals to slots
        self.apply_styles() # Apply CSS-like styles

        # Load initial data and setup tray icon
        self.load_bookmarks(self.loaded_file_path) # Load data, updates UI & saves path
        self.init_tray_icon() # Setup tray icon and its menu

        self.show() # Make the window visible
        logging.info("Main window initialized and shown.")

    def _load_and_apply_geometry(self):
        """Loads and applies window geometry from settings."""
        geo_data_hex = self.settings.get('window_geometry')
        restored = False
        if geo_data_hex:
            try:
                geo_data = bytes.fromhex(geo_data_hex)
                if self.restoreGeometry(geo_data):
                    logging.info("Window geometry restored from settings.")
                    restored = True
                else:
                    logging.warning("Failed to restore geometry from saved data (invalid?).")
            except (ValueError, TypeError) as e:
                logging.warning(f"Error decoding geometry data from settings: {e}", exc_info=True)
        else:
            logging.info("No saved window geometry found in settings.")

        if not restored:
            # Default geometry if not restored
            logging.info("Setting default window geometry.")
            self.setGeometry(300, 300, 1000, 600) # Position and size

    def _apply_transparency(self):
        """Applies window transparency from settings."""
        try:
            transparency = float(self.settings.get('transparency', 1.0))
            # Clamp between 0.1 (mostly transparent) and 1.0 (fully opaque)
            transparency = max(0.1, min(1.0, transparency))
            self.setWindowOpacity(transparency)
            logging.info(f"Window transparency set to {transparency*100:.0f}%")
        except (ValueError, TypeError) as e:
            logging.warning(f"Invalid transparency value in settings. Using default (1.0). Error: {e}")
            self.setWindowOpacity(1.0)
            self.settings.set('transparency', 1.0) # Correct setting value

    def setup_ui(self):
        """Creates and arranges the main UI elements."""
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(6)

        # --- Menubar ---
        self.menu_bar = self.menuBar()
        self.file_menu = self.menu_bar.addMenu("&File")
        self.view_menu = self.menu_bar.addMenu("&View") # Added View Menu
        self.help_menu = self.menu_bar.addMenu("&Help")

        # --- Top Layout (Search, Font, Count) ---
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # --- Search Box ---
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search Bookmarks...")
        top_layout.addWidget(self.search_box, 2) # Give search box more stretch factor

        # --- Search Options ---
        s_opt_layout = QHBoxLayout()
        s_opt_layout.addWidget(QLabel("Search:"))
        self.search_group = QButtonGroup(self)
        self.search_title_radio = QRadioButton("Title")
        self.search_syntax_radio = QRadioButton("Syntax")
        self.search_both_radio = QRadioButton("Both")
        for btn in [self.search_title_radio, self.search_syntax_radio, self.search_both_radio]:
            self.search_group.addButton(btn)
            s_opt_layout.addWidget(btn)
        self.search_both_radio.setChecked(True) # Default to searching both
        top_layout.addLayout(s_opt_layout)
        top_layout.addStretch(1) # Add some space before font settings

        # --- Font Size Combo Box (CORRECTED AREA from previous response) ---
        fnt_layout = QHBoxLayout()
        fnt_layout.addWidget(QLabel("Font:"))
        self.font_size_combo = QComboBox()
        sizes = ["8", "9", "10", "11", "12", "13", "14", "16"] # Available font sizes
        self.font_size_combo.addItems(sizes)

        # Get the saved font size, default to "12"
        def_fnt = self.settings.get('font_size', "12")

        # Ensure the default is one of the available sizes
        if def_fnt not in sizes:
            logging.warning(f"Saved font size '{def_fnt}' is not in the available list {sizes}. Falling back to 12.")
            def_fnt = "12" # Reset to a valid default

        # Set the current text and add the combo box to the layout
        self.font_size_combo.setCurrentText(def_fnt)
        self.font_size_combo.setMinimumWidth(50) # Give it a minimum width
        fnt_layout.addWidget(self.font_size_combo)
        top_layout.addLayout(fnt_layout)
        # --- End Font Size Combo Box ---

        # --- Bookmark Count ---
        self.bookmark_count_label = QLabel("Bookmarks: 0")
        self.bookmark_count_label.setObjectName("bookmark_count_label") # For styling
        self.bookmark_count_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.bookmark_count_label.setMinimumWidth(100) # Ensure space for text
        top_layout.addWidget(self.bookmark_count_label)

        main_layout.addLayout(top_layout)

        # --- Splitter, List, Preview ---
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(self.splitter, 1) # Allow splitter to take remaining vertical space

        # Left side: List View container
        list_cont = QWidget()
        list_lay = QVBoxLayout(list_cont)
        list_lay.setContentsMargins(0, 0, 0, 0) # No margins for the list container
        self.bookmark_list = CustomListWidget(self) # Use our custom list widget
        self.bookmark_list.setItemDelegate(BookmarkDelegate(self.bookmark_list)) # Apply custom delegate
        list_lay.addWidget(self.bookmark_list)
        # Label shown when list is empty
        self.no_bookmarks_label = QLabel("No bookmarks loaded.")
        self.no_bookmarks_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        list_lay.addWidget(self.no_bookmarks_label)
        self.no_bookmarks_label.hide() # Initially hidden
        self.splitter.addWidget(list_cont)

        # Right side: Preview Pane
        self.preview_pane = QTextEdit()
        self.preview_pane.setReadOnly(True)
        self.preview_pane.setPlaceholderText("Select a bookmark to preview its SQL content...")
        self.splitter.addWidget(self.preview_pane)

        # Restore splitter state or set default split
        splitter_state_hex = self.settings.get('splitter_state')
        restored_split = False
        if splitter_state_hex:
            try:
                splitter_state = bytes.fromhex(splitter_state_hex)
                if self.splitter.restoreState(splitter_state):
                    logging.info("Splitter state restored.")
                    restored_split = True
                else:
                    logging.warning("Failed to restore splitter state (invalid?).")
            except (ValueError, TypeError) as e:
                logging.warning(f"Error decoding splitter state: {e}", exc_info=True)

        if not restored_split:
            # Default split ratio (e.g., 40% list, 60% preview)
            logging.info("Setting default splitter sizes.")
            total_width = max(self.width(), self.minimumWidth()) # Use current or minimum width
            self.splitter.setSizes([int(total_width * 0.4), int(total_width * 0.6)])

        # --- Bottom Layout (Buttons) ---
        btm_layout = QHBoxLayout()
        btm_layout.setSpacing(10)
        # Move Up/Down buttons (functionality currently basic - UI only)
        self.up_button = QPushButton("Move Up")
        self.down_button = QPushButton("Move Down")
        # Temporarily disable Move buttons as underlying data move isn't robustly implemented yet
        self.up_button.setEnabled(False)
        self.down_button.setEnabled(False)
        # btm_layout.addWidget(self.up_button) # Uncomment when move implemented
        # btm_layout.addWidget(self.down_button)# Uncomment when move implemented
        btm_layout.addStretch() # Push buttons to the right

        self.hide_button = QPushButton("Hide to Tray")
        self.hide_button.setObjectName("hide_button") # For styling
        btm_layout.addWidget(self.hide_button)

        self.custom_close_button = QPushButton("Close") # Renamed for clarity vs. standard close
        self.custom_close_button.setObjectName("custom_close_button") # For styling
        btm_layout.addWidget(self.custom_close_button)

        main_layout.addLayout(btm_layout)

        # Update font size initially based on combo box's current value
        # Do this *after* all widgets using the font are created
        self.update_font_size(self.font_size_combo.currentText())

    def create_actions(self):
        """Create QActions for menus and potentially toolbars."""
        # --- File Menu Actions ---
        open_icon = QIcon.fromTheme("document-open", QIcon(os.path.join(ICON_DIR, "open_file.png"))) # Fallback icon
        self.open_file_action = QAction(open_icon, "&Open Bookmarks XML...", self)
        # **FIX:** Use Qt.KeyboardModifier for Control key
        self.open_file_action.setShortcut(Qt.KeyboardModifier.ControlModifier | Qt.Key.Key_O)
        self.open_file_action.setStatusTip("Open a DataGrip bookmarks XML file")

        folder_icon = QIcon.fromTheme("folder", QIcon())
        self.choose_sql_location_action = QAction(folder_icon, "Set &SQL Root Directory...", self)
        self.choose_sql_location_action.setStatusTip("Set the root directory where your SQL files are located")

        exec_icon = QIcon.fromTheme("application-x-executable", QIcon())
        self.choose_datagrip_action = QAction(exec_icon, "Set DataGrip &Executable...", self)
        self.choose_datagrip_action.setStatusTip("Set the path to the DataGrip executable (future use)")
        self.choose_datagrip_action.setEnabled(False) # Not currently used

        exit_icon = QIcon.fromTheme("application-exit", QIcon(os.path.join(ICON_DIR, "exit.png"))) # Fallback icon
        self.exit_action = QAction(exit_icon, "E&xit", self)
        # **FIX:** Use Qt.KeyboardModifier for Control key
        self.exit_action.setShortcut(Qt.KeyboardModifier.ControlModifier | Qt.Key.Key_Q)
        self.exit_action.setStatusTip("Exit the application")

        # --- View Menu Actions ---
        settings_icon = QIcon.fromTheme("preferences-system", QIcon())
        self.choose_favicon_action = QAction(settings_icon, "Set App &Icon...", self)
        self.choose_favicon_action.setStatusTip("Choose the main application icon (.ico)")
        self.choose_tray_icon_action = QAction(settings_icon, "Set &Tray Icon...", self)
        self.choose_tray_icon_action.setStatusTip("Choose the system tray icon (.ico)")

        clear_icon = QIcon.fromTheme("edit-clear", QIcon())
        self.clear_counts_action = QAction(clear_icon,"&Clear Usage Counts", self)
        self.clear_counts_action.setStatusTip("Reset all bookmark usage counts to zero")

        self.bookmark_count_menu_action = QAction("Bookmarks: 0", self) # Placeholder, updated dynamically
        self.bookmark_count_menu_action.setEnabled(False) # Not clickable

        # Transparency Submenu/Action
        self.transparency_menu = QMenu("Transparency", self) # This will be added to View menu
        opacity_icon = QIcon.fromTheme("preferences-desktop-theme", QIcon())
        self.transparency_slider_action = QAction(opacity_icon, "Set Transparency...", self)
        self.transparency_slider_action.setStatusTip("Adjust window transparency level")
        self.transparency_menu.addAction(self.transparency_slider_action)

        # --- Help Menu Actions ---
        help_icon = QIcon.fromTheme("help-contents", QIcon())
        self.help_locations_action = QAction(help_icon, "Show File &Locations", self)
        self.help_locations_action.setStatusTip("Show the locations of application data, logs, and configuration files")

        # --- Tray Menu Actions (also used by main window hide/show) ---
        self.show_window_action = QAction(QIcon.fromTheme("window-maximize"), "&Show Window", self)
        self.hide_window_action = QAction(QIcon.fromTheme("window-minimize"), "&Hide Window", self)

        copy_icon = QIcon.fromTheme("edit-copy", QIcon())
        self.tray_copy_sql_action = QAction(copy_icon, "&Copy SQL...", self)
        self.tray_copy_sql_action.setStatusTip("Select a bookmark from a list to copy its SQL")

    def add_actions_to_menus(self):
        """Populate the main menubar."""
        # File Menu
        self.file_menu.addAction(self.open_file_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.choose_sql_location_action)
        self.file_menu.addAction(self.choose_datagrip_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)

        # View Menu
        self.view_menu.addAction(self.choose_favicon_action)
        self.view_menu.addAction(self.choose_tray_icon_action)
        self.view_menu.addMenu(self.transparency_menu)
        self.view_menu.addSeparator()
        self.view_menu.addAction(self.clear_counts_action)
        self.view_menu.addSeparator()
        self.view_menu.addAction(self.bookmark_count_menu_action) # Display count here

        # Help Menu
        self.help_menu.addAction(self.help_locations_action)

    def init_context_menu(self):
        """Initialize the right-click context menu for the bookmark list."""
        self.bookmark_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.context_menu = QMenu(self)

        # Action to copy SQL (same as double-click/Enter)
        copy_sql_icon = QIcon.fromTheme("edit-copy", QIcon())
        self.copy_sql_action_context = QAction(copy_sql_icon, "Copy SQL", self)
        self.copy_sql_action_context.triggered.connect(self.handle_item_action_from_context)

        # Action to copy just the file URL
        copy_url_icon = QIcon.fromTheme("text-html", QIcon()) # Using HTML icon as proxy for URL
        self.copy_url_action_context = QAction(copy_url_icon, "Copy File URL", self)
        self.copy_url_action_context.triggered.connect(self.copy_bookmark_url_from_context)

        self.context_menu.addAction(self.copy_sql_action_context)
        self.context_menu.addAction(self.copy_url_action_context)

    @Slot(QListWidgetItem, QListWidgetItem)
    def on_current_item_changed(self, current: QListWidgetItem, previous: QListWidgetItem):
        """Update the preview pane when the current item changes (e.g., via keyboard navigation)."""
        if current:
            logging.debug(f"Current item changed to: {current.data(Qt.ItemDataRole.UserRole).get('title', 'N/A')}")
            self.update_preview_pane(current)
        else:
            # Clear preview if selection is lost
            self.preview_pane.clear()
            self.preview_pane.setPlaceholderText("Select a bookmark to preview its SQL content...")
            logging.debug("List selection cleared.")

    def connect_signals(self):
        """Connect UI signals to their corresponding slots."""
        # Search
        self.search_box.textChanged.connect(self.filter_bookmarks)
        # Re-filter when search scope changes
        self.search_title_radio.toggled.connect(self.filter_bookmarks)
        self.search_syntax_radio.toggled.connect(self.filter_bookmarks)
        # self.search_both_radio has no separate signal needed, it's handled by the others

        # List interactions
        self.bookmark_list.itemDoubleClicked.connect(self.handle_item_action) # Double-click copies SQL
        self.bookmark_list.itemClicked.connect(self.update_preview_pane) # Single click updates preview
        self.bookmark_list.currentItemChanged.connect(self.on_current_item_changed) # Arrow key navigation updates preview
        self.bookmark_list.customContextMenuRequested.connect(self.show_context_menu) # Right-click

        # Buttons
        # self.up_button.clicked.connect(self.move_up) # Disabled for now
        # self.down_button.clicked.connect(self.move_down) # Disabled for now
        self.hide_button.clicked.connect(self.hide_to_tray) # Hide window button
        self.custom_close_button.clicked.connect(self.quit_application) # Close button quits app

        # Font size
        self.font_size_combo.currentTextChanged.connect(self.update_font_size)

        # Menu actions (File)
        self.open_file_action.triggered.connect(self.select_and_load_file)
        self.choose_sql_location_action.triggered.connect(self.choose_sql_location)
        self.choose_datagrip_action.triggered.connect(self.choose_datagrip_executable)
        self.exit_action.triggered.connect(self.quit_application)

        # Menu actions (View)
        self.choose_favicon_action.triggered.connect(self.choose_main_icon)
        self.choose_tray_icon_action.triggered.connect(self.choose_tray_icon)
        self.clear_counts_action.triggered.connect(self.clear_usage_counts)
        self.transparency_slider_action.triggered.connect(self.show_transparency_dialog)

        # Menu actions (Help)
        self.help_locations_action.triggered.connect(self.show_help_locations)

        # NOTE: Tray Actions connected later in connect_tray_signals when tray icon is confirmed available

    # --- Icon/Path Choosing Logic ---
    def choose_icon(self, target_icon_path, title="Choose Icon File"):
        """Generic helper to choose an .ico file and copy it to the target path."""
        # Start file dialog in the directory of the current icon, or the app's icon dir
        start_dir = os.path.dirname(target_icon_path) if os.path.exists(os.path.dirname(target_icon_path)) else ICON_DIR
        if not os.path.isdir(start_dir): start_dir = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.PicturesLocation) # Absolute fallback

        fpath, _ = QFileDialog.getOpenFileName(self, title, start_dir, "Icon Files (*.ico);;All Files (*)")
        updated = False
        if fpath and os.path.exists(fpath):
            try:
                # Ensure target directory exists
                os.makedirs(os.path.dirname(target_icon_path), exist_ok=True)
                # Copy the chosen file to the application's data directory
                shutil.copy2(fpath, target_icon_path)
                logging.info(f"Successfully copied selected icon '{fpath}' to '{target_icon_path}'")
                updated = True
            except Exception as e:
                logging.error(f"Failed to copy icon file from '{fpath}' to '{target_icon_path}': {e}", exc_info=True)
                QMessageBox.critical(self, "Icon Copy Error", f"Failed to copy the selected icon:\n{e}")
        elif fpath: # User selected a path but it doesn't exist (shouldn't happen with QFileDialog?)
            QMessageBox.warning(self, "File Not Found", f"The selected icon file could not be found:\n{fpath}")

        return updated, target_icon_path # Return success status and the target path

    def choose_main_icon(self):
        """Handles choosing and applying the main application icon."""
        updated, path = self.choose_icon(MAIN_ICON_FILE, "Choose Main Application Icon (.ico)")
        if updated:
            icon = QIcon(path)
            if not icon.isNull():
                QApplication.instance().setWindowIcon(icon) # Set for the whole app
                self.setWindowIcon(icon) # Set for this window instance
                logging.info("Main application icon updated successfully.")
                QMessageBox.information(self, "Icon Updated", "The main application icon has been updated.")
            else:
                logging.error(f"Failed to load the new main application icon from {path} even after copying.")
                QMessageBox.warning(self, "Icon Load Error", f"The new icon was copied, but could not be loaded.\nPlease ensure it's a valid .ico file.")

    def choose_tray_icon(self):
        """Handles choosing and applying the system tray icon."""
        updated, path = self.choose_icon(TRAY_ICON_FILE, "Choose System Tray Icon (.ico)")
        if updated and self.tray_icon: # Only proceed if update was successful and tray icon exists
            icon = QIcon(path)
            if not icon.isNull():
                self.tray_icon.setIcon(icon)
                logging.info("System tray icon updated successfully.")
                QMessageBox.information(self, "Icon Updated", "The system tray icon has been updated.")
            else:
                logging.error(f"Failed to load the new tray icon from {path} even after copying.")
                QMessageBox.warning(self, "Icon Load Error", f"The new tray icon was copied, but could not be loaded.\nPlease ensure it's a valid .ico file.")
        elif updated and not self.tray_icon:
             logging.warning("Tray icon file updated, but no tray icon object exists (system tray unavailable?).")
             QMessageBox.information(self, "Icon Copied", "Tray icon file updated, but the tray icon is not currently active.")

    def choose_sql_location(self):
        """Allows user to select the root directory for their SQL files."""
        current_dir = self.settings.get('sql_root_directory', '')
        # Start in the current directory or user's home if not set
        start_dir = current_dir if os.path.isdir(current_dir) else os.path.expanduser("~")

        directory = QFileDialog.getExistingDirectory(self, "Choose SQL Root Directory", start_dir)

        if directory: # User selected a directory
            try:
                self.settings.set('sql_root_directory', directory)
                self.settings.save_settings()
                logging.info(f"SQL root directory set to: {directory}")
                QMessageBox.information(self, "Directory Set", f"SQL root directory successfully set to:\n{directory}")
                # Re-resolve and update preview if an item is selected
                current_item = self.bookmark_list.currentItem()
                if current_item:
                    self.update_preview_pane(current_item)
            except Exception as e:
                logging.error(f"Error saving SQL root directory setting: {e}", exc_info=True)
                QMessageBox.critical(self, "Error Saving Setting", f"Failed to save the selected SQL directory:\n{e}")

    def choose_datagrip_executable(self):
        """Allows user to select the DataGrip executable (placeholder)."""
        # This function is currently disabled but structure is here
        current_path = self.settings.get('datagrip_path', DEFAULT_DATAGRIP_PATH)
        init_dir = os.path.dirname(current_path) if current_path else os.path.expanduser("~")

        fpath, _ = QFileDialog.getOpenFileName(self, "Choose DataGrip Executable", init_dir, "Executables (*.exe);;All Files (*)")

        if fpath and os.path.exists(fpath):
            try:
                self.settings.set('datagrip_path', fpath)
                self.settings.save_settings()
                logging.info(f"DataGrip executable path set to: {fpath}")
                QMessageBox.information(self, "Path Set", f"DataGrip executable path successfully set:\n{fpath}")
            except Exception as e:
                logging.error(f"Error saving DataGrip executable path: {e}", exc_info=True)
                QMessageBox.critical(self, "Error Saving Setting", f"Failed to save the selected DataGrip path:\n{e}")
        elif fpath:
            QMessageBox.warning(self, "File Not Found", f"The selected executable file could not be found:\n{fpath}")

    # --- Styling and Font Update ---
    def apply_styles(self):
        """Apply custom styles using stylesheets for a dark theme."""
        # Using more specific selectors where possible (e.g., QPushButton#hide_button)
        # Font properties removed from here for QListWidget/QTextEdit as they are set programmatically
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b; /* Dark background */
                color: #dcdcdc; /* Light grey text */
                font-family: 'Segoe UI', Arial, sans-serif; /* Default font */
            }
            QWidget#centralWidget { /* Style the central widget specifically */
                background-color: #2b2b2b;
            }
            QMenuBar {
                background-color: #3c3c3c; /* Slightly lighter dark */
                color: #dcdcdc;
                border-bottom: 1px solid #555555; /* Subtle separator */
            }
            QMenuBar::item {
                padding: 4px 10px;
                background: transparent;
            }
            QMenuBar::item:selected { /* Hover/selection */
                background-color: #0078d7; /* Standard blue highlight */
                color: #ffffff;
            }
            QMenuBar::item:pressed {
                background-color: #005ba1; /* Darker blue when pressed */
            }
            QMenu {
                background-color: #3c3c3c;
                color: #dcdcdc;
                border: 1px solid #555555; /* Border around menu */
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 20px; /* More padding for readability */
                /* Font size managed by system or explicitly if needed */
            }
            QMenu::item:selected {
                background-color: #0078d7;
                color: #ffffff;
            }
            QMenu::separator {
                height: 1px;
                background-color: #555555;
                margin: 4px 0px;
            }
            QLineEdit {
                background-color: #3c3c3c;
                color: #dcdcdc;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
                /* font-size: 12px; */ /* Let system handle or set via QFont */
            }
            QLineEdit:focus {
                border: 1px solid #0078d7; /* Blue border on focus */
                background-color: #4a4a4a; /* Slightly lighter background on focus */
            }
            QListWidget { /* Base style for the list */
                background-color: #3c3c3c;
                color: #dcdcdc;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 0px; /* Delegate handles padding */
                outline: none; /* Remove focus outline */
                /* Font set programmatically */
                /* Item selection colors handled by delegate/palette */
            }
            QTextEdit { /* Style for the preview pane */
                background-color: #252526; /* Even darker background for code */
                color: #dcdcdc;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
                /* Font set programmatically */
            }
            QPushButton {
                background-color: #4a4a4a; /* Default button grey */
                color: #dcdcdc;
                border: 1px solid #666666;
                padding: 6px 12px;
                /* font-size: 12px; */
                border-radius: 3px;
                min-width: 80px; /* Minimum button width */
                margin: 0px 2px; /* Small horizontal margin */
            }
            QPushButton:hover {
                background-color: #5a5a5a; /* Lighter grey on hover */
                border-color: #777777;
            }
            QPushButton:pressed {
                background-color: #3a3a3a; /* Darker grey when pressed */
            }
            QPushButton#hide_button { /* Specific style for Hide button */
                background-color: #34577d; /* Blue */
                color: white;
                border-color: #456a9c;
            }
            QPushButton#hide_button:hover {
                background-color: #456a9c;
                border-color: #5a85b0;
            }
            QPushButton#hide_button:pressed {
                background-color: #284363;
            }
            QPushButton#custom_close_button { /* Specific style for Close button */
                background-color: #7d3434; /* Red */
                color: white;
                border-color: #9c4545;
            }
            QPushButton#custom_close_button:hover {
                background-color: #9c4545;
                border-color: #b05a5a;
            }
            QPushButton#custom_close_button:pressed {
                background-color: #632828;
            }
            QLabel { /* Default label style */
                color: #dcdcdc;
                /* font-size: 12px; */
                background-color: transparent; /* Ensure no background unless specified */
                border: none;
                padding: 0px;
            }
            QLabel#bookmark_count_label { /* Specific padding for count label */
                padding: 5px;
            }
            QRadioButton {
                color: #dcdcdc;
                /* font-size: 12px; */
                background-color: transparent;
            }
            QRadioButton::indicator { /* Style the radio button circle */
                width: 13px;
                height: 13px;
            }
            QComboBox {
                background-color: #3c3c3c;
                color: #dcdcdc;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px; /* Padding inside the combo box text area */
                /* font-size: 12px; */
                min-width: 40px; /* Min width for font size dropdown */
            }
            QComboBox:focus {
                border: 1px solid #0078d7; /* Highlight on focus */
            }
            QComboBox::drop-down { /* Style the dropdown arrow area */
                border: none;
                background-color: #4a4a4a; /* Button-like background */
                width: 15px;
                border-top-right-radius: 3px; /* Match main radius */
                border-bottom-right-radius: 3px;
            }
            QComboBox::down-arrow {
                 /* Use a theme icon or custom image if desired, removing default arrow */
                 /* image: url(:/icons/down_arrow.png); */
                 /* For now, let the system draw it or leave it blank */
                 image: url(no_arrow.png); /* Explicitly remove default arrow if needed */
            }
            QComboBox QAbstractItemView { /* Style the dropdown list */
                background-color: #3c3c3c;
                color: #dcdcdc;
                selection-background-color: #0078d7; /* Highlight color for selected item */
                border: 1px solid #555555; /* Border around the dropdown */
                outline: none;
            }
            QSplitter::handle {
                background-color: #555555; /* Color of the splitter handle */
            }
            QSplitter::handle:horizontal {
                width: 5px; /* Width of vertical splitter */
            }
            QSplitter::handle:vertical {
                height: 5px; /* Height of horizontal splitter */
            }
            QSplitter::handle:pressed {
                background-color: #0078d7; /* Color when dragging */
            }
            QDialog { /* Base style for dialogs */
                background-color: #2b2b2b;
                color: #dcdcdc;
                /* font-size: 12px; */
            }
            QDialog QLabel { /* Ensure dialog labels are transparent */
                background-color: transparent;
            }
            QDialog QPushButton { /* Smaller min width for dialog buttons */
                min-width: 60px;
            }
        """)
        # Set object names used in the stylesheet selectors
        self.central_widget.setObjectName("centralWidget")
        self.hide_button.setObjectName("hide_button")
        self.custom_close_button.setObjectName("custom_close_button")
        self.bookmark_count_label.setObjectName("bookmark_count_label")

    @Slot(str)
    def update_font_size(self, size_text):
        """Update font size for list and preview, preserving selection."""
        try:
            font_size = int(size_text)
            if font_size < 6 or font_size > 72: # Basic sanity check
                logging.warning(f"Font size {font_size} out of reasonable range, ignoring.")
                return

            font = QFont("Segoe UI", font_size) # Use Segoe UI or fallback like Arial
            logging.info(f"Attempting to set font size to {font_size}pt")

            # --- Preserve Selection ---
            selected_item_id = None
            current_item = self.bookmark_list.currentItem()
            if current_item:
                data = current_item.data(Qt.ItemDataRole.UserRole)
                if isinstance(data, dict):
                    selected_item_id = data.get('id')
            # --- End Preserve Selection ---

            # Apply font to relevant widgets
            self.bookmark_list.setFont(font)
            self.preview_pane.setFont(font)
            # Font size combo itself might need font update if system doesn't handle it
            # self.font_size_combo.setFont(font) # Usually not needed

            # Re-apply delegate to ensure sizeHint uses new font (might not be strictly necessary)
            # self.bookmark_list.setItemDelegate(BookmarkDelegate(self.bookmark_list))

            # Force layout update for the list widget items
            self.bookmark_list.updateGeometries()
            # Trigger a repaint might help sometimes, but updateGeometries should handle it
            # self.bookmark_list.viewport().update()

            # --- Restore Selection ---
            if selected_item_id:
                restored = False
                for i in range(self.bookmark_list.count()):
                    item = self.bookmark_list.item(i)
                    data = item.data(Qt.ItemDataRole.UserRole)
                    if isinstance(data, dict) and data.get('id') == selected_item_id:
                        self.bookmark_list.setCurrentItem(item, QItemSelectionModel.SelectionFlag.SelectCurrent)
                        self.bookmark_list.scrollToItem(item, QAbstractItemView.ScrollHint.EnsureVisible)
                        restored = True
                        break
                if restored:
                    logging.debug(f"Restored selection to item ID {selected_item_id}")
                else:
                    logging.warning(f"Could not restore selection for item ID {selected_item_id} after font change.")
            # --- End Restore Selection ---

            # Update preview explicitly if an item is selected
            if self.bookmark_list.currentItem():
                self.update_preview_pane(self.bookmark_list.currentItem())
            else:
                # Clear preview if nothing is selected after font change
                self.preview_pane.clear()
                self.preview_pane.setPlaceholderText("Select a bookmark to preview its SQL content...")


            logging.info(f"Font size successfully set to {font_size}pt")
            self.settings.set('font_size', size_text) # Save the new setting
            # No need to call save_settings here, happens on close/quit

        except ValueError:
            logging.error(f"Invalid font size text received: '{size_text}'")
        except Exception as e:
            logging.error(f"Unexpected error updating font size: {e}", exc_info=True)


    # --- Tray Icon and Related Actions ---
    def init_tray_icon(self):
        """Initialize system tray icon and its context menu."""
        if not QSystemTrayIcon.isSystemTrayAvailable():
            logging.warning("System tray is not available on this system. Tray icon disabled.")
            self.tray_icon = None
            # Disable tray-related UI elements
            self.hide_button.setEnabled(False)
            self.hide_button.setToolTip("System tray not available")
            return

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setToolTip(APP_NAME)

        # --- Load Tray Icon ---
        icon = QIcon() # Start with an empty icon
        loaded = False
        # 1. Try loading from the specified TRAY_ICON_FILE
        if os.path.exists(TRAY_ICON_FILE):
            temp_icon = QIcon(TRAY_ICON_FILE)
            if not temp_icon.isNull():
                icon = temp_icon
                loaded = True
                logging.info(f"Loaded tray icon from file: {TRAY_ICON_FILE}")
            else:
                logging.warning(f"Tray icon file exists ({TRAY_ICON_FILE}) but failed to load (invalid format?).")

        # 2. Fallback: Try loading from theme icon (edit-copy or similar)
        if not loaded:
            theme_icon = QIcon.fromTheme("accessories-text-editor", QIcon.fromTheme("edit-copy")) # Example fallback themes
            if not theme_icon.isNull():
                icon = theme_icon
                loaded = True
                logging.info("Loaded tray icon from system theme.")
            else:
                 logging.warning("Could not load tray icon from specified file or system theme.")

        # 3. Absolute Fallback: Main application icon (if available)
        if not loaded:
             app_icon = QApplication.instance().windowIcon()
             if not app_icon.isNull():
                  icon = app_icon
                  loaded = True
                  logging.info("Using main application icon as fallback tray icon.")

        # --- End Load Tray Icon ---

        if not loaded or icon.isNull():
            logging.error("Failed to find or load any valid icon for the system tray. Tray functionality disabled.")
            self.tray_icon = None # Ensure tray_icon is None if no icon could be set
            self.hide_button.setEnabled(False)
            self.hide_button.setToolTip("Failed to load tray icon")
            return # Cannot proceed without an icon

        self.tray_icon.setIcon(icon)

        # Create the context menu for the tray icon
        self.tray_menu = QMenu(self)
        self.tray_menu.addAction(self.show_window_action) # Show main window
        self.tray_menu.addAction(self.hide_window_action) # Hide main window (redundant if shown from tray)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(self.tray_copy_sql_action) # Open copy dialog
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(self.exit_action) # Exit application

        self.tray_icon.setContextMenu(self.tray_menu)

        # Connect signals for tray icon interaction
        self.connect_tray_signals()

        self.tray_icon.show()

        # Check if it actually became visible
        if self.tray_icon.isVisible():
            logging.info("System tray icon successfully initialized and shown.")
        else:
            # This can happen due to OS limitations or errors
            logging.error("System tray icon was initialized but failed to show. Tray functionality may be limited.")
            # Optionally disable UI elements again if showing failed
            # self.hide_button.setEnabled(False)
            # self.hide_button.setToolTip("Tray icon failed to show")

    def connect_tray_signals(self):
        """Connect signals related to the tray icon."""
        if not self.tray_icon:
            logging.debug("Skipping tray signal connection: tray icon not available.")
            return
        try:
            # Handle clicks on the tray icon
            self.tray_icon.activated.connect(self.handle_tray_icon_activation)
            # Connect actions used in the tray menu
            self.show_window_action.triggered.connect(self.show_window)
            self.hide_window_action.triggered.connect(self.hide_to_tray)
            self.tray_copy_sql_action.triggered.connect(self.show_tray_copy_dialog)
            # exit_action is already connected in connect_signals

            logging.debug("Tray icon signals connected.")
        except Exception as e:
            # Catch potential errors during signal connection (less common)
            logging.error(f"Error connecting tray icon signals: {e}", exc_info=True)

    def handle_tray_icon_activation(self, reason):
        """Handle tray icon activation (clicks)."""
        # Show window on double-click or middle-click (Trigger)
        if reason in (QSystemTrayIcon.ActivationReason.DoubleClick,
                      QSystemTrayIcon.ActivationReason.Trigger): # Trigger is often middle-click
            logging.debug(f"Tray icon activated (reason: {reason}). Showing window.")
            self.show_window()
        # Context menu (right-click) is handled automatically by setContextMenu

    def show_tray_copy_dialog(self):
        """Shows the dialog for selecting a bookmark to copy from the tray."""
        if not self.bookmarks:
            QMessageBox.information(self, "No Bookmarks", "No bookmarks are currently loaded.")
            return

        # Use the cached sorted list if available, otherwise sort the main list
        # This ensures the dialog shows the same order as the main window (including sort by count)
        sorted_list = self.sorted_bookmarks_cache if self.sorted_bookmarks_cache else self.apply_sort(self.bookmarks)

        if not sorted_list:
             QMessageBox.information(self, "No Bookmarks", "No bookmarks available to copy (list empty or filtered out).")
             return

        current_font = self.bookmark_list.font() # Use the same font as the main list
        dialog = TrayCopyDialog(sorted_list, current_font, self) # Parent is the main window
        # Connect the dialog's signal to our handler slot
        dialog.bookmark_selected_for_copy.connect(self.handle_tray_dialog_copy)
        # Show the dialog modally
        dialog.exec()

    @Slot(dict)
    def handle_tray_dialog_copy(self, bookmark_data):
        """Handles the signal from TrayCopyDialog when a bookmark is selected for copying."""
        if bookmark_data and isinstance(bookmark_data, dict):
            title = bookmark_data.get('title', 'N/A')
            bookmark_id = bookmark_data.get('id')
            logging.info(f"Copy action triggered from tray dialog for: '{title}' (ID: {bookmark_id})")

            if bookmark_id:
                # Increment usage count for the selected bookmark
                self.usage_counts.increment_count(bookmark_id)
                self.usage_counts.save_counts() # Save counts immediately after increment

                # Update the count in the main self.bookmarks list and the cache
                new_count = self.usage_counts.get_count(bookmark_id)
                updated_main = False
                for bm in self.bookmarks:
                    if isinstance(bm, dict) and bm.get('id') == bookmark_id:
                        bm['count'] = new_count
                        updated_main = True
                        break
                # Update cache as well, assuming it contains dict references
                for bm_cache in self.sorted_bookmarks_cache:
                     if isinstance(bm_cache, dict) and bm_cache.get('id') == bookmark_id:
                          bm_cache['count'] = new_count
                          break

                # Refresh the main list view if the count was updated
                # (Sort order might change)
                if updated_main:
                    self.update_bookmark_list()
                    logging.info(f"Incremented count for '{title}' via tray dialog. Main list updated.")
                else:
                    logging.warning(f"Could not find bookmark with ID {bookmark_id} in main list to update count.")

            # Retrieve the SQL content using the refined function
            sql_content = self.get_sql_content(bookmark_data)

            if sql_content is not None:
                # Copy the SQL content to the clipboard
                clipboard = QGuiApplication.clipboard()
                clipboard.setText(sql_content.strip())
                logging.info(f"Successfully copied SQL to clipboard for: '{title}'")

                # Show a tray notification bubble
                if self.tray_icon and self.tray_icon.isVisible():
                    self.tray_icon.showMessage(
                        APP_NAME,
                        f"Copied: {title}",
                        QSystemTrayIcon.MessageIcon.Information, # Icon type
                        1500 # Duration in milliseconds
                    )
                # Log the action
                self.log_bookmark_action(bookmark_data, "Copied SQL via Tray Dialog")
            else:
                # Handle case where SQL content couldn't be retrieved
                 logging.warning(f"Failed to retrieve SQL content for tray copy action: '{title}'")
                 if self.tray_icon and self.tray_icon.isVisible():
                     self.tray_icon.showMessage(
                         APP_NAME,
                         f"Error copying: {title}",
                         QSystemTrayIcon.MessageIcon.Warning,
                         2000
                     )
                 # Also show a standard message box as tray messages can be missed
                 QMessageBox.warning(self, "Copy Failed", f"Could not retrieve the SQL content for the selected bookmark:\n'{title}'.\n\nPlease check file paths and logs.")
        else:
            logging.warning("handle_tray_dialog_copy received invalid or empty bookmark data.")

    # --- Window Show/Hide ---
    def show_window(self):
        """Ensure the main window is visible, restored, and activated."""
        if self.isHidden():
             logging.debug("Window was hidden, calling show().")
             self.show() # Makes the window visible if hidden
        if self.isMinimized():
             logging.debug("Window was minimized, calling showNormal().")
             self.showNormal() # Restores from minimized state

        self.activateWindow() # Bring window to the front
        self.raise_() # Raise window above others in the application stack
        logging.info("Main window shown/activated.")

    def hide_to_tray(self):
         """Hide the main window to the system tray if available."""
         # Check if tray icon exists *and* is currently visible
         if self.tray_icon and self.tray_icon.isVisible():
              self.hide() # Hide the main window
              logging.info("Window hidden to system tray.")
              # Optional: Show a notification that it's still running
              self.tray_icon.showMessage(
                  APP_NAME,
                  "Application is running in the system tray.",
                  QSystemTrayIcon.MessageIcon.Information,
                  1000
              )
         else:
              # Fallback behavior if tray icon isn't available or visible
              logging.warning("Hide to tray requested, but no visible tray icon. Minimizing window instead.")
              self.showMinimized() # Minimize the window as a fallback

    # --- Bookmark Loading and Handling ---
    def select_and_load_file(self):
        """Open file dialog, copy the selected XML locally, and load bookmarks."""
        # Determine starting directory for the file dialog
        docs_location = QStandardPaths.standardLocations(QStandardPaths.StandardLocation.DocumentsLocation)
        # Use directory of last opened file, fallback to Documents, fallback to home
        initial_dir = os.path.dirname(self.settings.get('last_file_path', ''))
        if not initial_dir or not os.path.isdir(initial_dir):
            initial_dir = docs_location[0] if docs_location else os.path.expanduser("~")

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open DataGrip Bookmarks XML File",
            initial_dir,
            "XML Files (*.xml);;All Files (*)" # Filter for XML files
        )

        if file_path and os.path.exists(file_path):
            # User selected a valid file
            try:
                # Ensure the target directory for the copy exists
                os.makedirs(os.path.dirname(LAST_BOOKMARKS_COPY), exist_ok=True)
                # Copy the selected file to our app data dir for stable access
                shutil.copy2(file_path, LAST_BOOKMARKS_COPY)
                logging.info(f"Copied selected bookmarks file '{file_path}' to local storage '{LAST_BOOKMARKS_COPY}'")

                # Save the path of the *original* file the user selected
                self.settings.set('last_file_path', file_path)
                # Save the path of the *copy* we are actually loading from
                self.settings.set('loaded_copy_path', LAST_BOOKMARKS_COPY)
                # No need to save settings immediately, happens on quit/close

                # Load bookmarks from the *copied* file
                self.load_bookmarks(LAST_BOOKMARKS_COPY)

            except Exception as e:
                logging.error(f"Error copying or loading selected bookmarks file '{file_path}': {e}", exc_info=True)
                QMessageBox.critical(self, "File Load Error", f"Could not copy or load the selected bookmarks file:\n{e}")
        elif file_path:
            # User entered a path but it doesn't exist
             QMessageBox.warning(self, "File Not Found", f"The specified file could not be found:\n{file_path}")

    def update_bookmark_count(self):
        """Update the bookmark count display label and menu item."""
        count = len(self.bookmarks) if hasattr(self, 'bookmarks') else 0
        display_text = f"Bookmarks: {count}"
        self.bookmark_count_label.setText(display_text)
        self.bookmark_count_menu_action.setText(display_text)
        logging.debug(f"Bookmark count updated to: {count}")

    def load_bookmarks(self, file_path):
        """Load bookmarks from the specified XML file path."""
        logging.info(f"Attempting to load bookmarks from: {file_path}")
        self.loaded_file_path = file_path # Store the path we're loading from

        if not os.path.exists(file_path):
            logging.warning(f"Bookmarks file to load does not exist: {file_path}")
            self.bookmarks = [] # Clear bookmarks if file is gone
            # Use APP_NAME as title if file doesn't exist
            self.setWindowTitle(f"{APP_NAME} - No File Loaded")
            QMessageBox.warning(self, "File Not Found", f"The bookmarks file could not be found at:\n{file_path}\n\nPlease use 'File > Open' to select a valid file.")
        else:
            # Parse the XML file
            loaded_bookmarks = parse_bookmarks_xml(file_path)

            # Update counts for each loaded bookmark from usage data
            for bm in loaded_bookmarks:
                 if isinstance(bm, dict) and 'id' in bm:
                      bm['count'] = self.usage_counts.get_count(bm['id'])

            self.bookmarks = loaded_bookmarks # Store the loaded and count-updated bookmarks
            # Display the original file name in the title bar for user context
            original_file = self.settings.get('last_file_path', file_path) # Prefer original path for title
            self.setWindowTitle(f"{APP_NAME} - {os.path.basename(original_file)}")
            logging.info(f"Finished loading {len(self.bookmarks)} bookmarks.")

        # Update the list widget and count display after loading/clearing
        self.update_bookmark_list()
        self.update_bookmark_count()

        # Save settings (specifically the loaded_copy_path might have changed)
        # Moved saving to close/quit to avoid frequent writes
        # self.settings.save_settings()

    def update_bookmark_list(self):
        """Update list widget: filters, sorts, displays items, and caches the sorted list."""
        logging.debug("Updating bookmark list...")

        # --- Preserve Selection ---
        selected_item_id = None
        current_item = self.bookmark_list.currentItem()
        if current_item:
            data = current_item.data(Qt.ItemDataRole.UserRole)
            if isinstance(data, dict):
                selected_item_id = data.get('id')
        # --- End Preserve Selection ---

        # Disable updates for performance during clear/populate
        self.bookmark_list.setUpdatesEnabled(False)
        self.bookmark_list.clear() # Remove all existing items

        search_term = self.search_box.text()
        # 1. Filter the raw bookmarks based on search
        filtered_bookmarks = self.apply_filter(search_term)
        # 2. Sort the filtered bookmarks
        self.sorted_bookmarks_cache = self.apply_sort(filtered_bookmarks) # Update the cache
        logging.debug(f"Filtered: {len(filtered_bookmarks)}, Sorted/Cached: {len(self.sorted_bookmarks_cache)}")

        restored_idx = -1 # Index of the previously selected item in the new sorted list

        if self.sorted_bookmarks_cache:
            # Hide the "No bookmarks" label and show the list
            self.no_bookmarks_label.hide()
            self.bookmark_list.show()

            # Populate the list widget with sorted items
            for index, bm in enumerate(self.sorted_bookmarks_cache):
                 if not isinstance(bm, dict):
                     logging.warning(f"Skipping non-dict item during list update: {bm}")
                     continue
                 item = QListWidgetItem()
                 item.setData(Qt.ItemDataRole.UserRole, bm) # Store data in the item
                 # Delegate takes care of display based on data
                 self.bookmark_list.addItem(item)
                 # Check if this is the item that was previously selected
                 if selected_item_id and bm.get('id') == selected_item_id:
                     restored_idx = index
        else:
            # Show the "No bookmarks" label and hide the list
            self.bookmark_list.hide()
            if search_term:
                self.no_bookmarks_label.setText("No bookmarks match your search.")
            elif not self.bookmarks:
                 self.no_bookmarks_label.setText("No bookmarks loaded. Use File > Open.")
            else:
                self.no_bookmarks_label.setText("No bookmarks available (check XML file?).")
            self.no_bookmarks_label.show()

        # Re-enable updates now that the list is populated
        self.bookmark_list.setUpdatesEnabled(True)

        # --- Restore Selection ---
        restored_item = None
        if restored_idx != -1:
            try:
                restored_item = self.bookmark_list.item(restored_idx)
                if restored_item:
                    self.bookmark_list.setCurrentItem(restored_item, QItemSelectionModel.SelectionFlag.SelectCurrent)
                    # Scroll smoothly to ensure it's visible
                    self.bookmark_list.scrollToItem(restored_item, QAbstractItemView.ScrollHint.EnsureVisible)
                    logging.debug(f"Restored selection to index {restored_idx} (ID: {selected_item_id})")
                else:
                     logging.warning(f"Found index {restored_idx} for selection restore, but item was null.")
            except Exception as e:
                 logging.error(f"Error restoring selection to index {restored_idx}: {e}", exc_info=True)
        # --- End Restore Selection ---

        # Update highlighting in the preview pane based on the current search term
        # This should happen *after* selection is potentially restored
        self.highlight_search_results()
        logging.debug("Bookmark list update complete.")


    def apply_filter(self, text):
        """Filter bookmarks based on search text and selected search option."""
        if not hasattr(self, 'bookmarks') or not isinstance(self.bookmarks, list):
            logging.error("Apply_filter called but self.bookmarks is not a valid list.")
            return []

        if not text: # No search term, return all bookmarks
            return list(self.bookmarks) # Return a copy

        # Determine search scope (title, syntax, or both)
        search_option = "both" # Default
        # Check radio buttons safely (they might not exist during early init)
        r_title = getattr(self, 'search_title_radio', None)
        r_syntax = getattr(self, 'search_syntax_radio', None)
        if r_title and r_title.isChecked():
            search_option = "title"
        elif r_syntax and r_syntax.isChecked():
            search_option = "syntax"

        filtered_list = []
        text_lower = text.lower() # Case-insensitive search
        logging.debug(f"Filtering with term '{text}' (option: {search_option})")

        # Pre-fetch potentially needed values outside the loop
        sql_root = self.settings.get('sql_root_directory')
        user_home = os.path.expanduser("~") # Get user home once

        for bm in self.bookmarks:
            if not isinstance(bm, dict) or 'id' not in bm: continue # Skip invalid entries

            title_match = False
            syntax_match = False
            title = bm.get('title', '') # Safely get title

            # 1. Check Title Match (if applicable)
            if search_option in ("title", "both"):
                if text_lower in title.lower():
                    title_match = True

            # 2. Check Syntax Match (if applicable and title didn't already match if searching both)
            if search_option in ("syntax", "both") and not (search_option == "both" and title_match):
                # Avoid redundant SQL reading if title already matched in "both" mode
                sql_content = self.get_sql_content(bm, sql_root=sql_root, user_home=user_home)
                if sql_content and text_lower in sql_content.lower():
                    syntax_match = True

            # Add to filtered list if either matches according to the search option
            if title_match or syntax_match:
                filtered_list.append(bm)

        logging.debug(f"Filtering complete. Found {len(filtered_list)} matches.")
        return filtered_list


    def apply_sort(self, bookmark_list):
        """Sort bookmarks primarily by usage count (desc), then title (asc)."""
        if not isinstance(bookmark_list, list):
            logging.error("Apply_sort received non-list input.")
            return [] # Return empty list if input is invalid

        try:
            # Sort key: Tuple where first element is negative count (for descending),
            # second element is lowercase title (for ascending secondary sort).
            # Handle potential non-dict items gracefully in key function.
            return sorted(
                bookmark_list,
                key=lambda b: (
                    -b.get('count', 0) if isinstance(b, dict) else 0, # Negate count for descending
                    b.get('title', '').lower() if isinstance(b, dict) else '' # Lowercase title for ascending
                )
            )
        except Exception as e:
            logging.error(f"Error during bookmark sorting: {e}", exc_info=True)
            return bookmark_list # Return the original list unsorted on error

    @Slot() # Can be triggered by textChanged or radio toggled
    def filter_bookmarks(self):
        """Slot to trigger list update when search text or options change."""
        logging.debug("Filter trigger activated (text or radio change).")
        self.update_bookmark_list()
        # Highlighting is handled within update_bookmark_list now


    def resolve_file_path(self, url, sql_root=None, user_home=None):
        """Resolve file path from URL, considering protocol, placeholders, and SQL root."""
        if not url:
            logging.debug("Resolve_file_path called with empty URL.")
            return None

        # Get SQL root and user home if not provided
        if sql_root is None:
            sql_root = self.settings.get('sql_root_directory', '')
        if user_home is None:
            user_home = os.path.expanduser("~")

        f_path = url
        prefix_found = False

        # 1. Handle file:// prefix (common in DataGrip URLs)
        # Handle both file:/// (correct) and file:// (sometimes seen)
        for prefix in ["file:///", "file://"]:
            if f_path.startswith(prefix):
                f_path = f_path[len(prefix):]
                prefix_found = True
                # On Windows, a path starting with /C:/... might result, remove leading /
                if sys.platform == 'win32' and f_path.startswith('/') and len(f_path) > 2 and f_path[2] == ':':
                     f_path = f_path[1:]
                logging.debug(f"Removed '{prefix}' prefix, path now: {f_path}")
                break

        # 2. Replace $USER_HOME$ placeholder
        if "$USER_HOME$" in f_path:
             f_path = f_path.replace("$USER_HOME$", user_home)
             logging.debug(f"Replaced $USER_HOME$, path now: {f_path}")

        # 3. Normalize path separators
        f_path = os.path.normpath(f_path)
        logging.debug(f"Normalized path: {f_path}")

        # 4. Check if the path (potentially relative or absolute) exists directly
        if os.path.exists(f_path):
            logging.debug(f"Resolved path exists directly: {f_path}")
            return f_path

        # 5. If not found directly, check relative to SQL Root Directory (if set)
        if sql_root and os.path.isdir(sql_root):
            # Assume the f_path might be just the filename or relative within the project
            # Construct path relative to SQL root
            potential_path_in_sql_root = os.path.join(sql_root, os.path.basename(f_path)) # Simplest: just filename
            # More complex: try joining relative path? Might be fragile.
            # potential_path_in_sql_root_rel = os.path.join(sql_root, f_path)

            logging.debug(f"Checking relative to SQL Root '{sql_root}': '{potential_path_in_sql_root}'")
            if os.path.exists(potential_path_in_sql_root):
                logging.debug(f"Found path relative to SQL Root: {potential_path_in_sql_root}")
                return potential_path_in_sql_root
            # elif os.path.exists(potential_path_in_sql_root_rel):
            #     logging.debug(f"Found path relative to SQL Root (using full relative): {potential_path_in_sql_root_rel}")
            #     return potential_path_in_sql_root_rel

        # 6. Add specific OneDrive check (example - adjust pattern if needed)
        # This is heuristic and might need refinement based on actual paths
        onedrive_marker = "OneDrive - Aledade, Inc" # Adapt this marker
        if onedrive_marker in url and user_home: # Check original URL for marker
             try:
                 # Try to reconstruct path relative to user's actual OneDrive folder
                 parts = url.replace("\\", "/").split("/")
                 idx = parts.index(onedrive_marker)
                 # Join user home with parts *after* the marker
                 potential_onedrive_path = os.path.normpath(os.path.join(user_home, *parts[idx:]))
                 logging.debug(f"Checking potential OneDrive path: {potential_onedrive_path}")
                 if os.path.exists(potential_onedrive_path):
                      logging.debug(f"Resolved path via OneDrive heuristic: {potential_onedrive_path}")
                      return potential_onedrive_path
             except (ValueError, IndexError) as e:
                 logging.debug(f"OneDrive marker found but path reconstruction failed: {e}")
             except Exception as e: # Catch other potential errors during reconstruction
                 logging.warning(f"Error during OneDrive path resolution: {e}", exc_info=True)


        # If all checks fail, the path could not be resolved
        logging.warning(f"Could not resolve URL '{url}' to an existing file path. Checked direct path '{f_path}' and relative to SQL root '{sql_root}'.")
        return None


    # --- Get SQL Content (Refined Slicing) ---
    def get_sql_content(self, bookmark_data, sql_root=None, user_home=None, pre_resolved_path=None):
        """Gets the SQL content for a given bookmark, respecting the next bookmark in the same file. Returns None on error."""
        if not bookmark_data or not isinstance(bookmark_data, dict):
            logging.warning("get_sql_content called with invalid bookmark_data.")
            return None

        url = bookmark_data.get('url', '')
        start_line_str = bookmark_data.get('line', '')

        if not url or not start_line_str:
            logging.warning(f"Missing URL or Line in bookmark data: {bookmark_data.get('title', 'N/A')}")
            return None

        try:
            # Line numbers in bookmarks are 1-based
            start_line_1based = int(start_line_str) + 1 # DataGrip line seems 0-based, file lines are 1-based
            if start_line_1based <= 0:
                 logging.warning(f"Invalid start line number {start_line_1based} (from '{start_line_str}') for bookmark '{bookmark_data.get('title', 'N/A')}'")
                 return None
        except ValueError:
            logging.warning(f"Non-integer line number '{start_line_str}' for bookmark '{bookmark_data.get('title', 'N/A')}'")
            return None

        # --- Resolve File Path ---
        file_path = pre_resolved_path # Use pre-resolved path if provided (e.g., during syntax search)
        if not file_path:
             file_path = self.resolve_file_path(url, sql_root=sql_root, user_home=user_home)

        if not file_path or not os.path.exists(file_path):
            logging.warning(f"Could not find file for SQL content. URL: '{url}', Resolved Path: '{file_path}'")
            # Optionally return an error message string?
            # return f"-- Error: Could not find file at resolved path: {file_path}\n-- Original URL: {url}"
            return None
        # --- End Resolve File Path ---

        # --- Determine End Line ---
        end_line_1based = None # Represents the line *number* (1-based) *of* the next bookmark (exclusive end)
        try:
            # Find all bookmark lines for the *same file*
            file_bookmark_lines = []
            for bm in self.bookmarks:
                if isinstance(bm, dict) and bm.get('url') == url:
                    line_str = bm.get('line', '')
                    try:
                        # Convert to 1-based line number for comparison
                        line_1based = int(line_str) + 1
                        if line_1based > 0:
                            file_bookmark_lines.append(line_1based)
                    except ValueError:
                        continue # Ignore bookmarks with invalid line numbers in the same file

            if start_line_1based in file_bookmark_lines:
                file_bookmark_lines.sort() # Sort lines numerically
                try:
                    current_idx = file_bookmark_lines.index(start_line_1based)
                    # If there's a bookmark after this one in the sorted list...
                    if current_idx + 1 < len(file_bookmark_lines):
                        end_line_1based = file_bookmark_lines[current_idx + 1] # Line number of the next bookmark
                        logging.debug(f"Found next bookmark in same file at line {end_line_1based}. Current ends before that.")
                except ValueError:
                    # Should not happen if start_line_1based was in the list
                     logging.error(f"Logic error: start_line {start_line_1based} reported in list but index not found.")
            # else: Bookmark start line not found among bookmarks for this file? Might be an issue, read to EOF.

        except Exception as e:
            # Log error during end line calculation but proceed (read to end)
            logging.error(f"Error determining end line for bookmark '{bookmark_data.get('title', 'N/A')}': {e}", exc_info=True)
        # --- End Determine End Line ---

        # --- Read File Content ---
        try:
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.readlines()

            total_lines = len(lines)
            # Adjust start line to 0-based index for slicing
            start_index_0based = start_line_1based - 1

            if start_index_0based >= total_lines:
                logging.warning(f"Start line ({start_line_1based}) is beyond the end of file ({total_lines} lines) for '{file_path}'.")
                return None # Cannot start reading past the end

            # Determine the end index for slicing (exclusive)
            if end_line_1based is not None:
                 # end_line_1based is the 1-based line number of the *next* bookmark.
                 # We want to slice up to, but *not including*, that line.
                 # So, the 0-based end index is end_line_1based - 1.
                 end_index_0based = end_line_1based - 1
                 # Clamp the end index to the bounds of the file
                 end_index_0based = min(end_index_0based, total_lines)
            else:
                # No next bookmark found, read to the end of the file
                end_index_0based = total_lines

            # Ensure end index is not before start index
            if end_index_0based <= start_index_0based:
                 logging.warning(f"Calculated end index ({end_index_0based}) is not after start index ({start_index_0based}) for '{bookmark_data.get('title', 'N/A')}'. Reading only start line.")
                 # Read just the single start line in this edge case
                 sql_lines = lines[start_index_0based : start_index_0based + 1]
            else:
                 # Perform the slice
                 sql_lines = lines[start_index_0based : end_index_0based]

            logging.debug(f"Read lines {start_line_1based} to {end_index_0based} (exclusive) from '{file_path}'")
            return ''.join(sql_lines) # Join the lines into a single string

        except FileNotFoundError:
             logging.error(f"File not found error during read operation: '{file_path}' (Should have been caught earlier?)")
             return None
        except IOError as e:
             logging.error(f"IOError reading file '{file_path}': {e}", exc_info=True)
             return None
        except Exception as e:
             logging.error(f"Unexpected error reading file content from '{file_path}': {e}", exc_info=True)
             return None
        # --- End Read File Content ---

    # --- Preview Pane and Highlighting ---
    def update_preview_pane(self, item: QListWidgetItem):
        """Update preview pane with SQL content and highlight search terms."""
        sql_content = None
        header_info = "-- Select a bookmark to preview --"
        bookmark_data = None

        if item:
            bookmark_data = item.data(Qt.ItemDataRole.UserRole)

        if bookmark_data and isinstance(bookmark_data, dict) and bookmark_data.get('id'):
            title = bookmark_data.get('title', 'N/A')
            logging.debug(f"Updating preview for: '{title}'")
            # Get the SQL content for this bookmark
            sql_content = self.get_sql_content(bookmark_data)

            if sql_content is not None:
                # Prepare header information
                file_url = bookmark_data.get('url', 'Unknown URL')
                file_path = self.resolve_file_path(file_url) # Try to resolve for display
                display_path = os.path.basename(file_path) if file_path else os.path.basename(file_url.replace("file://", ""))
                line_num = bookmark_data.get('line', '?') # Original 0-based line
                try:
                    display_line = int(line_num) + 1 # Show 1-based line to user
                except ValueError:
                    display_line = line_num # Keep as is if not integer

                header_info = (
                    f"-- Bookmark: {title}\n"
                    f"-- File: {display_path}\n"
                    f"-- Line: {display_line} (Starts on 0-based line {line_num} in XML)\n"
                    f"{'-'*60}\n"
                )
                # Combine header and SQL content
                full_preview_text = header_info + sql_content.strip()
            else:
                # SQL content failed to load, show error in preview
                logging.warning(f"Failed to get SQL content for preview: '{title}'")
                header_info = (
                     f"-- Error loading preview for:\n"
                     f"-- {title}\n"
                     f"-- URL: {bookmark_data.get('url', 'N/A')}\n"
                     f"-- Line: {bookmark_data.get('line', 'N/A')}\n"
                     f"{'-'*60}\n"
                     f"-- Check logs for details (e.g., file not found, parse errors)."
                )
                full_preview_text = header_info
        else:
             # No valid item selected or data missing
             logging.debug("Update preview called with invalid item or data.")
             full_preview_text = header_info # Show default placeholder

        # Set the text in the preview pane
        self.preview_pane.setPlainText(full_preview_text) # Use setPlainText for simpler text

        # Always apply/clear highlighting *after* text is set
        self.highlight_search_results()

    def highlight_search_results(self):
        """Highlights occurrences of the search term in the preview pane if syntax search is active."""
        if not hasattr(self, 'preview_pane'): return # Should not happen

        # Clear existing highlights first
        cursor = self.preview_pane.textCursor()
        cursor.select(QTextCursor.SelectionType.Document)
        default_format = QTextCharFormat() # Get default format from the widget?
        # preview_format = self.preview_pane.currentCharFormat() # Or use current default
        cursor.setCharFormat(default_format)
        # Reset cursor position to avoid selection staying active
        cursor.clearSelection()
        self.preview_pane.setTextCursor(cursor)
        logging.debug("Cleared previous highlights.")

        search_term = self.search_box.text()
        # Only highlight if there's a search term AND syntax/both search is enabled
        search_syntax = getattr(self, 'search_syntax_radio', None)
        search_both = getattr(self, 'search_both_radio', None)
        should_highlight = bool(search_term and (
            (search_syntax and search_syntax.isChecked()) or
            (search_both and search_both.isChecked())
        ))

        if not should_highlight:
             logging.debug("Highlighting skipped (no search term or syntax search not active).")
             return

        document = self.preview_pane.document()
        if not document:
             logging.warning("Preview pane document is null, cannot highlight.")
             return

        # Define the format for highlighted text
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QColor("yellow"))
        highlight_format.setForeground(QColor("black")) # Ensure text is visible on yellow

        # Use QTextDocument.find to iterate through matches
        find_cursor = QTextCursor(document) # Start cursor at the beginning
        first_match_cursor = None # To store the cursor of the first match

        logging.debug(f"Highlighting term: '{search_term}'")
        count = 0
        while True:
            # Find next occurrence (case-insensitive)
            find_cursor = document.find(search_term, find_cursor, QTextDocument.FindFlag.FindCaseSensitively) # Use FindCaseSensitively? Or FindFlag() for case-insensitive? Let's try sensitive.

            if find_cursor.isNull():
                # No more matches found
                break
            else:
                # Apply the highlight format to the found selection
                find_cursor.mergeCharFormat(highlight_format)
                count += 1
                # Store the first match cursor to potentially scroll to it
                if first_match_cursor is None:
                    first_match_cursor = QTextCursor(find_cursor)
                    first_match_cursor.clearSelection() # Don't keep it selected, just position

        logging.debug(f"Found and highlighted {count} occurrences.")

        # Scroll to the first match if found
        if first_match_cursor:
            self.preview_pane.setTextCursor(first_match_cursor)
            self.preview_pane.ensureCursorVisible()
            logging.debug("Scrolled to the first highlight match.")


    # --- Item Actions ---
    def handle_item_action(self, item: QListWidgetItem):
        """Handle double-click/Enter: increment count, copy SQL, log action."""
        if not item:
            logging.warning("handle_item_action called with null item.")
            return

        bookmark_data = item.data(Qt.ItemDataRole.UserRole)
        if bookmark_data and isinstance(bookmark_data, dict) and 'id' in bookmark_data:
            bookmark_id = bookmark_data['id']
            title = bookmark_data.get('title', 'N/A')
            logging.info(f"Action triggered for bookmark: '{title}' (ID: {bookmark_id})")

            # 1. Increment Usage Count
            self.usage_counts.increment_count(bookmark_id)
            self.usage_counts.save_counts() # Persist count change immediately

            # Update count in the main data list and cache for immediate reflection
            new_count = self.usage_counts.get_count(bookmark_id)
            updated_main = False
            for bm in self.bookmarks:
                 if isinstance(bm, dict) and bm.get('id') == bookmark_id:
                      bm['count'] = new_count
                      updated_main = True
                      break
            for bm_cache in self.sorted_bookmarks_cache:
                 if isinstance(bm_cache, dict) and bm_cache.get('id') == bookmark_id:
                      bm_cache['count'] = new_count
                      break

            # 2. Get SQL Content
            sql_content = self.get_sql_content(bookmark_data)

            if sql_content is not None:
                 # 3. Copy to Clipboard
                 clipboard = QGuiApplication.clipboard()
                 clipboard.setText(sql_content.strip())
                 logging.info(f"Successfully copied SQL to clipboard for: '{title}'")
                 # Optional: Show brief confirmation message (can be annoying)
                 # QMessageBox.information(self, "Copied", f"SQL for '{title}' copied to clipboard.")

                 # 4. Log the Action
                 self.log_bookmark_action(bookmark_data, "Copied SQL via Main List Action")
            else:
                 # Handle failure to get SQL content
                 logging.warning(f"Failed to retrieve SQL content for action: '{title}'")
                 QMessageBox.warning(self, "Copy Failed", f"Could not retrieve the SQL content for the bookmark:\n'{title}'.\n\nPlease check file paths and logs.")

            # 5. Refresh List View (because counts/sort order might change)
            # Only refresh if the count was actually updated in the main list
            if updated_main:
                logging.debug("Refreshing list view after item action due to count change.")
                self.update_bookmark_list()
            else:
                 logging.warning(f"Could not find bookmark ID {bookmark_id} in main list to update count for action.")

        else:
            logging.warning("handle_item_action called with invalid item data.")


    def log_bookmark_action(self, bookmark_data, action_type="Unknown Action"):
        """Logs details about a bookmark action (e.g., copy) to a separate file."""
        if not bookmark_data or not isinstance(bookmark_data, dict):
            logging.warning("Attempted to log action with invalid bookmark data.")
            return

        try:
            # Get current timestamp in a standard format
            timestamp = logging.Formatter().formatTime(
                logging.LogRecord(None, None, '', 0, '', (), None, None),
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            bookmark_id = bookmark_data.get('id', '')
            count_after = self.usage_counts.get_count(bookmark_id) if bookmark_id else 'N/A'

            entry = (
                f"\n{'='*50}\n"
                f"Timestamp: {timestamp}\n"
                f"Action: {action_type}\n"
                f"Bookmark Title: {bookmark_data.get('title', 'N/A')}\n"
                f"Bookmark Details: {bookmark_data.get('full_text', 'N/A')}\n"
                f"Bookmark ID: {bookmark_id}\n"
                f"Usage Count After Action: {count_after}\n"
                f"{'='*50}\n"
            )

            # Append the entry to the dedicated actions log file
            with open(BOOKMARK_ACTIONS_LOG, 'a', encoding='utf-8') as f:
                f.write(entry)
            logging.info(f"Logged action '{action_type}' for bookmark '{bookmark_data.get('title', 'N/A')}' to {BOOKMARK_ACTIONS_LOG}")

        except Exception as e:
            # Log error to the main log file if action logging fails
            logging.error(f"Failed to write to bookmark actions log ({BOOKMARK_ACTIONS_LOG}): {e}", exc_info=True)

    @Slot()
    def handle_item_action_from_context(self):
        """Triggered by the 'Copy SQL' context menu action."""
        if self.context_item:
            logging.debug("Context menu 'Copy SQL' triggered.")
            self.handle_item_action(self.context_item) # Reuse the main action handler
        else:
            logging.warning("Context menu 'Copy SQL' triggered but context_item is None.")

    @Slot()
    def copy_bookmark_url_from_context(self):
        """Triggered by the 'Copy File URL' context menu action."""
        if self.context_item:
            data = self.context_item.data(Qt.ItemDataRole.UserRole)
            if isinstance(data, dict):
                url = data.get('url', '')
                if url:
                    clipboard = QGuiApplication.clipboard()
                    clipboard.setText(url)
                    logging.info(f"Copied URL via context menu: {url}")
                    # Optional confirmation:
                    # self.statusBar().showMessage(f"URL copied: {url}", 2000)
                else:
                    QMessageBox.warning(self, "No URL", "The selected bookmark does not have a URL associated with it.")
            else:
                 logging.warning("Context item data is not a dictionary in copy_bookmark_url.")
        else:
            logging.warning("Context menu 'Copy URL' triggered but context_item is None.")

    def move_item(self, direction):
         """Moves the selected item up or down in the list (UI only for now)."""
         # NOTE: This currently only moves the item in the QListWidget.
         # It does NOT reorder the underlying self.bookmarks data or save the order.
         # Proper implementation requires modifying self.bookmarks and likely saving
         # the order, which is complex if combined with count-based sorting.
         # DISABLING the buttons calling this for now.
         logging.warning("Move item functionality is currently UI-only and disabled.")
         return # Exit early

         current_row = self.bookmark_list.currentRow()
         count = self.bookmark_list.count()
         target_row = -1

         if direction == "up":
             if current_row > 0:
                 target_row = current_row - 1
             else:
                  logging.debug("Cannot move item up: Already at the top.")
                  return # Already at top
         elif direction == "down":
             if current_row < count - 1 and current_row != -1:
                 target_row = current_row + 1
             else:
                  logging.debug("Cannot move item down: Already at the bottom or no selection.")
                  return # Already at bottom or nothing selected

         if target_row != -1:
             logging.debug(f"Moving item from row {current_row} to {target_row}")
             # Take the item out
             item_to_move = self.bookmark_list.takeItem(current_row)
             # Insert it at the target row
             self.bookmark_list.insertItem(target_row, item_to_move)
             # Reselect the moved item
             self.bookmark_list.setCurrentRow(target_row)

             # --- Attempt to move in underlying data (Needs refinement) ---
             # This part is tricky because self.bookmarks might not match the
             # current visually sorted/filtered list (self.sorted_bookmarks_cache).
             # A robust implementation would need to map the visual row back to
             # the correct index in self.bookmarks or handle reordering differently.
             # item_data = item_to_move.data(Qt.ItemDataRole.UserRole)
             # try:
             #     item_id = item_data.get('id') if isinstance(item_data, dict) else None
             #     if item_id:
             #         # Find index in the *main* list
             #         current_data_idx = -1
             #         for i, bm in enumerate(self.bookmarks):
             #              if isinstance(bm, dict) and bm.get('id') == item_id:
             #                   current_data_idx = i
             #                   break
             #
             #         if current_data_idx != -1:
             #             target_data_idx = max(0, min(len(self.bookmarks)-1, current_data_idx + (1 if direction=="down" else -1)))
             #             # Find the ID of the item currently at the target_data_idx
             #             # Swap positions in self.bookmarks (this is a simple approach, might need adjustment)
             #             moved_data = self.bookmarks.pop(current_data_idx)
             #             self.bookmarks.insert(target_data_idx, moved_data)
             #             logging.info(f"Moved bookmark '{item_data.get('title')}' {direction} in underlying data (simple swap).")
             #             # Mark data as dirty if saving order is intended
             #             # self.data_dirty = True
             #             # Re-filter and re-sort cache? This might undo the manual move if sorting is active.
             #             # self.sorted_bookmarks_cache = self.apply_sort(self.apply_filter(self.search_box.text()))
             #         else:
             #             logging.warning("Moved item's data not found in self.bookmarks.")
             #     else:
             #          logging.warning("Moved item has no ID, cannot reorder underlying data.")
             # except Exception as e:
             #      logging.error(f"Error reordering underlying bookmark data: {e}", exc_info=True)
             # --- End Attempt to move data ---

    @Slot()
    def move_up(self):
         self.move_item("up")

    @Slot()
    def move_down(self):
         self.move_item("down")

    @Slot(QPoint)
    def show_context_menu(self, pos: QPoint):
        """Show the right-click context menu at the given position."""
        item = self.bookmark_list.itemAt(pos)
        self.context_item = item # Store the item that was right-clicked
        if item:
             # Enable/disable actions based on the item
             data = item.data(Qt.ItemDataRole.UserRole)
             has_url = bool(data and isinstance(data, dict) and data.get('url'))
             self.copy_url_action_context.setEnabled(has_url)
             self.copy_sql_action_context.setEnabled(True) # Always enabled if item exists

             # Map the local position to global screen coordinates and show the menu
             global_pos = self.bookmark_list.mapToGlobal(pos)
             self.context_menu.popup(global_pos)
        else:
             # Right-clicked on empty area? Optionally show a limited menu or nothing.
             logging.debug("Right-click on empty list area, no context menu shown.")
             self.context_item = None

    @Slot()
    def clear_usage_counts(self):
        """Asks for confirmation and clears all bookmark usage counts."""
        reply = QMessageBox.question(
            self,
            "Confirm Clear Counts",
            "Are you sure you want to reset all bookmark usage counts to zero?\nThis cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, # Buttons
            QMessageBox.StandardButton.No # Default button
        )

        if reply == QMessageBox.StandardButton.Yes:
            logging.info("User confirmed clearing usage counts.")
            self.usage_counts.clear_counts()
            self.usage_counts.save_counts() # Persist the cleared counts immediately

            # Reset counts in the currently loaded bookmarks list
            if hasattr(self, 'bookmarks'):
                for bm in self.bookmarks:
                     if isinstance(bm, dict):
                          bm['count'] = 0 # Reset count in memory
            else:
                 logging.warning("Bookmarks list not found while resetting counts in memory.")

            # Update the list view to reflect the cleared counts (sort order will change)
            self.update_bookmark_list()
            QMessageBox.information(self, "Counts Cleared", "All bookmark usage counts have been reset.")
        else:
            logging.info("User cancelled clearing usage counts.")

    @Slot()
    def show_transparency_dialog(self):
        """Shows a dialog to adjust the main window's transparency."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Set Window Opacity")
        dialog.setMinimumWidth(300)

        layout = QVBoxLayout(dialog)
        label = QLabel("Window Opacity (10% = More Transparent, 100% = Opaque):")
        layout.addWidget(label)

        slider_layout = QHBoxLayout()
        slider = QSlider(Qt.Orientation.Horizontal)
        slider_layout.addWidget(slider)

        value_label = QLabel() # To display the current percentage
        value_label.setMinimumWidth(40) # Ensure space for "100%"
        value_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        slider_layout.addWidget(value_label)
        layout.addLayout(slider_layout)

        min_opacity_percent = 10
        max_opacity_percent = 100
        slider.setRange(min_opacity_percent, max_opacity_percent)
        # Get current opacity (0.0-1.0), convert to percent, clamp, and set slider
        current_opacity = max(0.1, min(1.0, self.windowOpacity())) # Ensure current is within bounds
        current_percent = int(current_opacity * 100)
        slider.setValue(current_percent)

        original_opacity = self.windowOpacity() # Store original value for cancel

        # Function to update label and window opacity as slider moves
        def update_opacity(value):
            value_label.setText(f"{value}%")
            self.setWindowOpacity(value / 100.0)

        slider.valueChanged.connect(update_opacity)
        update_opacity(current_percent) # Set initial label text

        # Standard OK/Cancel buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        # Execute the dialog
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Save the new opacity setting
            final_opacity = slider.value() / 100.0
            self.settings.set('transparency', final_opacity)
            # Saving happens on close/quit
            logging.info(f"Window transparency set to {final_opacity*100:.0f}% and saved.")
        else:
            # Restore original opacity on cancel
            self.setWindowOpacity(original_opacity)
            logging.info("Transparency adjustment cancelled, opacity restored.")


    @Slot()
    def show_help_locations(self):
        """Displays a dialog showing key file/folder locations."""
        help_text = generate_help_locations_text() # Regenerate text each time

        dialog = QDialog(self)
        dialog.setWindowTitle(f"{APP_NAME} - File Locations")
        dialog.setMinimumSize(550, 400) # Make dialog reasonably sized

        layout = QVBoxLayout(dialog)

        text_edit = QTextEdit()
        text_edit.setPlainText(help_text)
        text_edit.setReadOnly(True)
        # Use a monospace font for better path alignment if available
        mono_font = QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont)
        text_edit.setFont(mono_font)
        layout.addWidget(text_edit)

        button_layout = QHBoxLayout()
        # Buttons to open relevant directories
        log_dir_button = QPushButton("Open Log Directory")
        log_dir_button.setToolTip(f"Open: {LOG_DIR}")
        config_dir_button = QPushButton("Open Config Directory")
        config_dir_button.setToolTip(f"Open: {CONFIG_DIR}")
        # Add more buttons if needed (e.g., for ICON_DIR)

        ok_button = QPushButton("OK")
        ok_button.setDefault(True) # Default button

        button_layout.addWidget(log_dir_button)
        button_layout.addWidget(config_dir_button)
        button_layout.addStretch()
        button_layout.addWidget(ok_button)
        layout.addLayout(button_layout)

        # Connect buttons
        ok_button.clicked.connect(dialog.accept)
        # Use lambda to pass the specific directory path to the open function
        log_dir_button.clicked.connect(lambda: self.open_directory(LOG_DIR))
        config_dir_button.clicked.connect(lambda: self.open_directory(CONFIG_DIR))

        dialog.exec()

    def open_directory(self, path):
        """Opens the specified directory path in the system's file explorer."""
        if os.path.isdir(path):
            try:
                # Use QDesktopServices for better cross-platform compatibility
                from PySide6.QtGui import QDesktopServices
                from PySide6.QtCore import QUrl

                # Convert local path to a file URL
                url = QUrl.fromLocalFile(os.path.normpath(path))
                if QDesktopServices.openUrl(url):
                    logging.info(f"Successfully requested to open directory: {path}")
                else:
                    logging.error(f"QDesktopServices failed to open directory URL: {url.toString()}")
                    QMessageBox.warning(self, "Open Directory Failed", f"Could not open the directory:\n{path}")
            except ImportError:
                 logging.error("QDesktopServices not available. Cannot open directory.")
                 QMessageBox.critical(self, "Error", "Could not open directory (Component missing).")
            except Exception as e:
                logging.error(f"Failed to open directory '{path}': {e}", exc_info=True)
                QMessageBox.warning(self, "Open Directory Error", f"An error occurred while trying to open the directory:\n{path}\n\n{e}")
        else:
            logging.warning(f"Cannot open directory: Path does not exist or is not a directory: {path}")
            QMessageBox.warning(self, "Directory Not Found", f"The directory could not be found:\n{path}")

    # --- Application Close/Quit Logic ---
    def quit_application(self):
        """Saves state and cleanly exits the application."""
        logging.info("Quit action triggered. Saving state and exiting application.")
        self.save_state() # Ensure state is saved before quitting
        if self.tray_icon:
            self.tray_icon.hide() # Hide tray icon before quitting
            logging.debug("Tray icon hidden.")
        QApplication.instance().quit() # Tell the Qt application event loop to exit

    def closeEvent(self, event):
        """Handles the window close event (e.g., clicking the 'X' button)."""
        logging.info("Window close event triggered.")

        # Check if the tray icon exists and is currently visible
        if self.tray_icon and self.tray_icon.isVisible():
            # If tray is active, hide the window instead of closing the app
            logging.info("Tray icon is visible. Hiding window to tray instead of closing.")
            self.hide() # Hide the main window
            event.ignore() # Tell Qt to ignore the close event, preventing app exit
            # Optional: Show notification
            # self.tray_icon.showMessage(APP_NAME,"Minimized to tray.", QSystemTrayIcon.MessageIcon.Information, 1000)
        else:
            # If no visible tray icon, proceed with quitting the application
            logging.info("No visible tray icon. Proceeding with application quit.")
            self.save_state() # Save state before accepting the close
            event.accept() # Allow the window to close and the application to exit

    def save_state(self):
        """Save window geometry, splitter state, settings, and usage counts."""
        logging.info("Saving application state...")
        try:
            # Only save geometry if the window is in a 'normal' state (not minimized/hidden)
            if not self.isMinimized() and not self.isHidden():
                 # Save window geometry as hex string
                 geometry_hex = self.saveGeometry().toHex().data().decode('ascii')
                 self.settings.set('window_geometry', geometry_hex)
                 logging.debug(f"Saved window geometry: {geometry_hex[:30]}...") # Log truncated hex

                 # Save splitter state if the splitter exists
                 if hasattr(self, 'splitter') and self.splitter:
                      splitter_state_hex = self.splitter.saveState().toHex().data().decode('ascii')
                      self.settings.set('splitter_state', splitter_state_hex)
                      logging.debug(f"Saved splitter state: {splitter_state_hex[:30]}...") # Log truncated hex
                 else:
                      logging.debug("Splitter widget not found, skipping state save.")
            else:
                 # Don't save geometry if minimized/hidden to avoid restoring in that state
                 logging.info("Window is minimized or hidden, skipping geometry/splitter state save.")

        except Exception as e:
            logging.warning(f"Failed to get or save geometry/splitter state: {e}", exc_info=True)

        # Always save settings and usage counts regardless of window state
        try:
             self.settings.save_settings()
        except Exception as e:
             logging.error(f"Error during settings save: {e}", exc_info=True)

        try:
             self.usage_counts.save_counts()
        except Exception as e:
             logging.error(f"Error during usage counts save: {e}", exc_info=True)

        logging.info("Application state saving process completed.")

# --- Entry Point ---
if __name__ == '__main__':
    # Logging is already set up globally now

    exit_code = 0 # Default exit code
    try:
        # **FIX:** Remove deprecated High DPI attributes for Qt >= 6
        # High DPI scaling is generally handled automatically or via environment variables (e.g., QT_ENABLE_HIGHDPI_SCALING=1)
        # QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
        # QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

        # Create the application instance
        app = QApplication(sys.argv)
        app.setQuitOnLastWindowClosed(False) # Prevent app exit when window hides to tray

        # --- Load Main Application Icon ---
        main_icon = QIcon() # Start empty
        icon_loaded = False
        if os.path.exists(MAIN_ICON_FILE):
            temp_icon = QIcon(MAIN_ICON_FILE)
            if not temp_icon.isNull():
                main_icon = temp_icon
                icon_loaded = True
                logging.info(f"Loaded main application icon from: {MAIN_ICON_FILE}")
            else:
                logging.warning(f"Main icon file exists ({MAIN_ICON_FILE}) but failed to load.")
        if not icon_loaded:
            # Fallback to a generic theme icon
            theme_icon = QIcon.fromTheme("application-x-executable", QIcon.fromTheme("text-editor"))
            if not theme_icon.isNull():
                 main_icon = theme_icon
                 icon_loaded = True
                 logging.info("Using system theme icon as main application icon.")
            else:
                 logging.warning("Could not load main icon from file or theme.")

        if icon_loaded:
             app.setWindowIcon(main_icon)
        # --- End Load Main Icon ---

        # Initialize settings and usage counts handlers
        app_settings = AppSettings()
        usage_counts_data = UsageCounts()

        # Create and show the main window
        window = FloatingBookmarksWindow(app_settings, usage_counts_data)
        # Window calls self.show() in its __init__

        # Start the Qt event loop
        logging.info("Starting Qt application event loop...")
        exit_code = app.exec()
        logging.info(f"Qt application event loop finished. Exit code: {exit_code}.")

    except Exception as e:
        # Catch any unhandled exceptions at the top level
        logging.critical(f"Unhandled top-level exception occurred: {e}", exc_info=True)
        # Try to show a message box, but this might fail if Qt isn't running
        try:
            QMessageBox.critical(
                None, # No parent window
                "Fatal Application Error",
                f"An unhandled error occurred and the application must close:\n\n{e}\n\nPlease check the log file for details:\n{LOG_FILE}"
            )
        except Exception as msg_e:
            # Fallback to printing if message box fails
            print(f"FATAL APPLICATION ERROR: {e}", file=sys.stderr)
            print(f"Log File: {LOG_FILE}", file=sys.stderr)
            print(f"(Error showing message box: {msg_e})", file=sys.stderr)
        exit_code = 1 # Indicate an error exit

    finally:
        logging.info("="*20 + f" {APP_NAME} End " + "="*20 + f" (Exit Code: {exit_code})")
        sys.exit(exit_code) # Ensure the script exits with the correct code