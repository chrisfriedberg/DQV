# -*- coding: utf-8 -*-
import sys
import os
import json
import shutil
from xml.etree import ElementTree as ET
import logging
import subprocess
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QListWidget, QListWidgetItem,
    QLabel, QDialog, QPushButton, QFileDialog, QMenu, QMessageBox, QTextEdit, QSplitter,
    QAbstractItemView, QMenuBar, QSlider, QMainWindow, QSystemTrayIcon, QStyledItemDelegate, QStyle,
    QRadioButton, QComboBox, QButtonGroup, QDialogButtonBox, QAction, QCheckBox
)
from PyQt5.QtCore import (
    Qt, QSize, QPoint, QSettings, QStandardPaths, QRect, pyqtSignal as Signal, pyqtSlot as Slot,
    QItemSelectionModel 
)
from PyQt5.QtGui import (
    QColor, QFont, QGuiApplication, QIcon, QPainter, QTextDocument, QFontMetrics,
    QKeyEvent, QCursor, QTextCharFormat, QTextCursor, QKeySequence
)
# QScintilla imports
from PyQt5.Qsci import QsciScintilla, QsciLexerSQL, QsciAPIs
from PyQt5.QtGui import QFontDatabase
# --- Custom SQL Syntax Highlighter ---
from PyQt5.QtGui import QSyntaxHighlighter
class SQLSyntaxHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighting_rules = []
        
        # SQL Keywords
        keyword_format = QTextCharFormat()
        keyword_format.setForeground(QColor("#569CD6"))  # Blue for keywords
        keyword_format.setFontWeight(QFont.Weight.Bold)
        
        keywords = [
            "\\bSELECT\\b", "\\bFROM\\b", "\\bWHERE\\b", "\\bJOIN\\b", "\\bLEFT\\b", "\\bRIGHT\\b", 
            "\\bINNER\\b", "\\bOUTER\\b", "\\bGROUP\\s+BY\\b", "\\bORDER\\s+BY\\b", "\\bHAVING\\b", 
            "\\bLIMIT\\b", "\\bOFFSET\\b", "\\bUNION\\b", "\\bINSERT\\b", "\\bUPDATE\\b", "\\bDELETE\\b", 
            "\\bCREATE\\b", "\\bALTER\\b", "\\bDROP\\b", "\\bTABLE\\b", "\\bINDEX\\b", "\\bVIEW\\b", 
            "\\bAS\\b", "\\bON\\b", "\\bAND\\b", "\\bOR\\b", "\\bNOT\\b", "\\bIN\\b", "\\bBETWEEN\\b", 
            "\\bLIKE\\b", "\\bIS\\s+NULL\\b", "\\bIS\\s+NOT\\s+NULL\\b", "\\bASC\\b", "\\bDESC\\b", 
            "\\bDISTINCT\\b", "\\bCOUNT\\b", "\\bSUM\\b", "\\bAVG\\b", "\\bMIN\\b", "\\bMAX\\b"
        ]
        
        # Add keyword patterns
        for pattern in keywords:
            self.highlighting_rules.append((pattern, keyword_format))
        
        # String formats
        string_format = QTextCharFormat()
        string_format.setForeground(QColor("#CE9178"))  # Brown for strings
        self.highlighting_rules.append(("'[^']*'", string_format))
        self.highlighting_rules.append(("\"[^\"]*\"", string_format))
        
        # Number format
        number_format = QTextCharFormat()
        number_format.setForeground(QColor("#B5CEA8"))  # Green for numbers
        self.highlighting_rules.append(("\\b\\d+\\b", number_format))
        
        # Comment format (single line)
        comment_format = QTextCharFormat()
        comment_format.setForeground(QColor("#6A9955"))  # Green for comments
        self.highlighting_rules.append(("--[^\n]*", comment_format))
        
        # Operator format
        operator_format = QTextCharFormat()
        operator_format.setForeground(QColor("#D7BA7D"))  # Gold for operators
        operators = ["=", "<", ">", "<=", ">=", "<>", "!=", "\\+", "-", "\\*", "/", "%"]
        for op in operators:
            self.highlighting_rules.append((op, operator_format))
        
        # Multi-line comment format (stored separately as it needs special handling)
        self.multi_line_comment_format = QTextCharFormat()
        self.multi_line_comment_format.setForeground(QColor("#6A9955"))
        self.comment_start_expression = r"/\*"
        self.comment_end_expression = r"\*/"
        
    def highlightBlock(self, text):
        # Apply single-line rules
        for pattern, format in self.highlighting_rules:
            import re
            matches = re.finditer(pattern, text, re.IGNORECASE)
            for match in matches:
                start = match.start()
                length = match.end() - match.start()
                self.setFormat(start, length, format)
        
        # Handle multi-line comments
        self.setCurrentBlockState(0)
        
        start_index = 0
        if self.previousBlockState() != 1:
            import re
            match = re.search(self.comment_start_expression, text)
            if match:
                start_index = match.start()
            else:
                start_index = -1
        
        while start_index >= 0:
            import re
            end_match = re.search(self.comment_end_expression, text[start_index:])
            if end_match:
                end_index = start_index + end_match.end()
                comment_length = end_index - start_index
                self.setFormat(start_index, comment_length, self.multi_line_comment_format)
                start_index = text.find("/*", end_index)
            else:
                # Comment continues to next block
                self.setCurrentBlockState(1)
                comment_length = len(text) - start_index
                self.setFormat(start_index, comment_length, self.multi_line_comment_format)
                break

# --- Application Metadata and Data Folder Setup ---
APP_NAME = "DQV"  # Changed from "DGBookmarksViewer" to "DQV"
ORG_NAME = "MyLocalScripts"
# QApplication organization and app name will be set after creating the application instance
APP_DATA_DIR = os.path.join(QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation), APP_NAME)
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
INTERNAL_VAULT_DIR = os.path.join(APP_DATA_DIR, "query_vault")  # Directory for storing internal queries

# Create directories if they don't exist
for directory in [APP_DATA_DIR, LOG_DIR, CONFIG_DIR, BOOKMARKS_COPY_DIR, ICON_DIR, HELP_DIR, INTERNAL_VAULT_DIR]:
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
INTERNAL_VAULT_FILE = os.path.join(INTERNAL_VAULT_DIR, "query_vault.json")  # File for storing internal queries

# Define default DataGrip path (adjust if necessary)
DEFAULT_DATAGRIP_PATH = r"C:\Users\cfriedberg\AppData\Local\JetBrains\DataGrip 2024.1.4\bin\datagrip64.exe"

# Constants for data source modes
SOURCE_DATAGRIP = "datagrip"
SOURCE_INTERNAL = "internal"

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
        if 'data_source' not in self.settings:
            self.settings['data_source'] = SOURCE_DATAGRIP # Default to DataGrip source for backward compatibility

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

# --- Query Vault for Internal Storage ---
class QueryVault:
    def __init__(self):
        self.vault_path = INTERNAL_VAULT_FILE
        self.queries = []
        self.load_vault()
        
    def load_vault(self):
        """Load internal queries from the vault file."""
        if os.path.exists(self.vault_path):
            try:
                with open(self.vault_path, 'r', encoding='utf-8') as f:
                    self.queries = json.load(f)
                logging.info(f"Internal query vault loaded from {self.vault_path}, {len(self.queries)} queries found")
            except json.JSONDecodeError as e:
                logging.error(f"Error decoding vault JSON from {self.vault_path}: {e}", exc_info=True)
                self.queries = [] # Reset to empty on decode error
            except Exception as e:
                logging.error(f"Error loading vault from {self.vault_path}: {e}", exc_info=True)
                self.queries = []
        else:
            logging.info(f"Query vault file not found at {self.vault_path}. Starting with empty vault.")
            self.queries = []
            
    def save_vault(self):
        """Save internal queries to the vault file."""
        try:
            os.makedirs(os.path.dirname(self.vault_path), exist_ok=True)
            with open(self.vault_path, 'w', encoding='utf-8') as f:
                json.dump(self.queries, f, indent=4, ensure_ascii=False)
            logging.info(f"Query vault saved to {self.vault_path}, {len(self.queries)} queries")
        except Exception as e:
            logging.error(f"Error saving query vault to {self.vault_path}: {e}", exc_info=True)
            
    def get_queries(self):
        """Return a copy of all queries."""
        return list(self.queries)
    
    def add_query(self, query_data):
        """Add a new query to the vault."""
        if not isinstance(query_data, dict):
            logging.error(f"Cannot add query: data is not a dictionary")
            return False
            
        # Ensure the query has a unique ID
        if 'id' not in query_data:
            # Generate a unique ID if none exists
            import uuid
            query_data['id'] = str(uuid.uuid4())
            
        # Add timestamp if not present
        if 'created_at' not in query_data:
            from datetime import datetime
            query_data['created_at'] = datetime.now().isoformat()
            
        # Add to vault
        self.queries.append(query_data)
        logging.info(f"Added query '{query_data.get('title', 'Untitled')}' to vault with ID {query_data['id']}")
        return True
        
    def update_query(self, query_id, updated_data):
        """Update an existing query by ID."""
        for i, query in enumerate(self.queries):
            if query.get('id') == query_id:
                # Update fields while preserving ID and created_at
                query_id = query['id']  # Preserve ID
                created_at = query.get('created_at')  # Preserve creation time
                
                # Update with new data
                self.queries[i] = updated_data
                
                # Restore preserved fields
                self.queries[i]['id'] = query_id
                if created_at:
                    self.queries[i]['created_at'] = created_at
                
                # Add modified timestamp
                from datetime import datetime
                self.queries[i]['modified_at'] = datetime.now().isoformat()
                
                logging.info(f"Updated query with ID {query_id}")
                return True
                
        logging.warning(f"Failed to update query: No query found with ID {query_id}")
        return False
        
    def delete_query(self, query_id):
        """Delete a query by ID."""
        initial_count = len(self.queries)
        self.queries = [q for q in self.queries if q.get('id') != query_id]
        
        if len(self.queries) < initial_count:
            logging.info(f"Deleted query with ID {query_id}")
            return True
        else:
            logging.warning(f"Failed to delete query: No query found with ID {query_id}")
            return False
            
    def get_query_by_id(self, query_id):
        """Get a query by ID."""
        for query in self.queries:
            if query.get('id') == query_id:
                return query
        return None
        
    def get_all_labels(self):
        """Get a list of all unique labels used across all queries."""
        labels = set()
        for query in self.queries:
            query_labels = query.get('labels', [])
            if isinstance(query_labels, list):
                labels.update(query_labels)
        return sorted(list(labels))
        
    def add_label_to_query(self, query_id, label):
        """Add a label to a query."""
        for query in self.queries:
            if query.get('id') == query_id:
                if 'labels' not in query:
                    query['labels'] = []
                if label not in query['labels']:
                    query['labels'].append(label)
                    from datetime import datetime
                    query['modified_at'] = datetime.now().isoformat()
                    logging.info(f"Added label '{label}' to query {query_id}")
                    return True
                return True  # Label already exists, still success
        return False
        
    def remove_label_from_query(self, query_id, label):
        """Remove a label from a query."""
        for query in self.queries:
            if query.get('id') == query_id:
                if 'labels' in query and label in query['labels']:
                    query['labels'].remove(label)
                    from datetime import datetime
                    query['modified_at'] = datetime.now().isoformat()
                    logging.info(f"Removed label '{label}' from query {query_id}")
                    return True
        return False
        
    def get_queries_by_label(self, label):
        """Get all queries with a specific label."""
        return [query for query in self.queries if 'labels' in query and label in query['labels']]

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
    LABEL_MARGIN = 4  # Margin between labels
    LABEL_PADDING = 4  # Padding inside each label
    LABEL_COLORS = {
        "Critical": QColor(255, 0, 0, 80),  # Red with transparency
        "Draft": QColor(255, 165, 0, 80),   # Orange with transparency
        "Final": QColor(0, 128, 0, 80),     # Green with transparency
        # Default color for other labels will be handled dynamically
    }
    
    def calculate_fixed_height(self, font):
        # Calculate a consistent height based on font metrics
        font_height = QFontMetrics(font).height()
        return int(font_height * self.FIXED_ITEM_HEIGHT_FACTOR)
        
    def paint(self, painter: QPainter, option, index):
        # Get bookmark data from the item's UserRole
        bookmark_data = index.data(Qt.ItemDataRole.UserRole)
        if not bookmark_data or not isinstance(bookmark_data, dict):
            # If no valid bookmark data, fall back to standard delegate painting
            super().paint(painter, option, index)
            return
            
        # Get title
        title = bookmark_data.get('title', 'Untitled')
        
        # Get usage count
        count = bookmark_data.get('count', 0)
        
        # Get labels (if any)
        labels = bookmark_data.get('labels', [])
        
        # Calculate item rect with padding
        item_rect = option.rect.adjusted(self.ITEM_PADDING, self.ITEM_PADDING, 
                                       -self.ITEM_PADDING, -self.ITEM_PADDING)
                                       
        # Setup painter based on item state (selected, hover, etc.)
        painter.save()
        
        # Draw background rect based on selection state
        if option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, option.palette.highlight())
            painter.setPen(option.palette.highlightedText().color())
        else:
            painter.setPen(option.palette.text().color())
            
        # Draw the usage count box on the left
        count_box_rect = QRect(item_rect.left(), item_rect.top(), 
                             self.COUNT_BOX_WIDTH, item_rect.height())
        # Use a lighter background for count                  
        count_bg_color = option.palette.alternateBase().color()
        painter.fillRect(count_box_rect, count_bg_color)
        painter.drawRect(count_box_rect) # Draw border
        
        # Draw count text
        painter.drawText(count_box_rect, Qt.AlignmentFlag.AlignCenter, str(count))
        
        # Set up text rect to right of count box
        text_rect = QRect(count_box_rect.right() + self.TEXT_SPACING, 
                        item_rect.top(), 
                        item_rect.width() - self.COUNT_BOX_WIDTH - self.TEXT_SPACING,
                        item_rect.height())
                        
        # Draw title text, elide if necessary
        font = option.font
        font_metrics = QFontMetrics(font)
        elided_title = font_metrics.elidedText(
            title, 
            Qt.TextElideMode.ElideRight, 
            text_rect.width()
        )
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, elided_title)
        
        # Draw labels if present
        if labels and isinstance(labels, list):
            # Start from the right side
            label_x = option.rect.right() - self.ITEM_PADDING
            label_y = option.rect.top() + self.ITEM_PADDING
            
            # Smaller font for labels
            label_font = QFont(option.font)
            label_font.setPointSize(max(8, option.font.pointSize() - 2))
            painter.setFont(label_font)
            
            # Draw each label from right to left
            for label in labels:
                label_text = str(label)
                label_width = font_metrics.horizontalAdvance(label_text) + 2 * self.LABEL_PADDING
                
                # Determine label color
                if label_text in self.LABEL_COLORS:
                    label_color = self.LABEL_COLORS[label_text]
                else:
                    # Generate a color based on the label text
                    # Use a hash function to get a consistent color for the same label
                    hash_val = sum(ord(c) for c in label_text)
                    hue = hash_val % 360
                    # Use HSV to generate a consistent pastel color
                    from PyQt5.QtGui import QColor
                    label_color = QColor()
                    label_color.setHsv(hue, 180, 240, 80)  # Light pastel with transparency
                
                # Draw label background
                label_rect = QRect(label_x - label_width, label_y, 
                                 label_width, label_font.pointSize() + 2 * self.LABEL_PADDING)
                painter.fillRect(label_rect, label_color)
                painter.drawRect(label_rect)
                
                # Draw label text
                painter.drawText(label_rect, Qt.AlignmentFlag.AlignCenter, label_text)
                
                # Move to the next label position (to the left)
                label_x -= (label_width + self.LABEL_MARGIN)
        
        painter.restore()
    
    def sizeHint(self, option, index):
        # Return the calculated fixed height and base width
        fixed_height = self.calculate_fixed_height(option.font)
        # Standard width calculation - let the view determine this
        return QSize(0, fixed_height)

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
    """Main application window for displaying and interacting with bookmarks/queries."""
    
    def create_actions(self):
        """Creates all actions for menus, toolbars, and tray icon."""
        from PyQt5.QtWidgets import QAction
        from PyQt5.QtGui import QKeySequence

        # Create window actions
        self.show_window_action = QAction("Show Window", self)
        self.show_window_action.triggered.connect(self.show_window)

        self.hide_window_action = QAction("Hide to Tray", self)
        self.hide_window_action.triggered.connect(self.hide_to_tray)

        self.tray_copy_sql_action = QAction("Copy SQL to Clipboard", self)
        self.tray_copy_sql_action.triggered.connect(self.copy_current_query_to_clipboard)

        self.exit_action = QAction("Exit DQV", self)
        self.exit_action.triggered.connect(self.close_app)
        
        # Add bookmark count menu action
        self.bookmark_count_menu_action = QAction("Bookmarks: 0", self)
        self.bookmark_count_menu_action.setEnabled(False)  # This is just an indicator, not clickable

        # Add dark mode toggle action
        self.toggle_theme_action = QAction("Toggle Dark Mode", self)
        self.toggle_theme_action.setCheckable(True)
        self.toggle_theme_action.setChecked(True)  # Default to dark mode
        self.toggle_theme_action.triggered.connect(self.toggle_theme)

        # Create context menu actions
        self.copy_sql_action_context = QAction("Copy SQL", self)
        self.copy_sql_action_context.triggered.connect(self.handle_item_action_from_context)
        
        self.copy_url_action_context = QAction("Copy File URL", self)
        self.copy_url_action_context.triggered.connect(self.copy_bookmark_url_from_context)
        
        self.edit_query_action_context = QAction("Edit Query", self)
        self.edit_query_action_context.triggered.connect(self.edit_query)
        
        self.delete_query_action_context = QAction("Delete Query", self)
        self.delete_query_action_context.triggered.connect(self.delete_query)
        
        self.manage_labels_action_context = QAction("Manage Labels", self)
        self.manage_labels_action_context.triggered.connect(self.manage_labels)
        
        self.import_to_vault_action_context = QAction("Import to Vault", self)
        self.import_to_vault_action_context.triggered.connect(self.import_to_vault)

        # --- File-menu specific actions ---
        self.open_file_action = QAction("&Open DataGrip XML...", self)
        self.open_file_action.setShortcut(QKeySequence.Open)
        self.open_file_action.triggered.connect(self.select_and_load_file)

        self.set_sql_root_action = QAction("Set SQL Root Directory...", self)
        self.set_sql_root_action.triggered.connect(self.set_sql_root_directory)

        self.clear_counts_action = QAction("Clear Usage Counts", self)
        self.clear_counts_action.triggered.connect(self.clear_usage_counts)

        # --- View-menu extras ---
        self.transparency_action = QAction("Window Opacity...", self)
        self.transparency_action.triggered.connect(self.show_transparency_dialog)

        # Preview edit toggle
        self.toggle_edit_action = QAction("Editable Preview", self)
        self.toggle_edit_action.setCheckable(True)
        self.toggle_edit_action.setChecked(False)  # Default read-only
        self.toggle_edit_action.toggled.connect(self.toggle_preview_editable)

        # --- Help-menu actions ---
        self.help_locations_action = QAction("File Locations...", self)
        self.help_locations_action.triggered.connect(self.show_help_locations)

        self.about_action = QAction("About", self)
        self.about_action.triggered.connect(self.show_about_dialog)

        logging.info("create_actions() successfully created and connected.")

    def add_actions_to_menus(self):
        """Populate the File / View / Help menus and rebuild the tray context menu."""
        # ----- File Menu -----
        self.file_menu.addAction(self.open_file_action)
        self.file_menu.addAction(self.set_sql_root_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.clear_counts_action)
        self.file_menu.addSeparator()
        self.file_menu.addAction(self.exit_action)

        # ----- View Menu -----
        self.view_menu.addAction(self.toggle_theme_action)
        self.view_menu.addAction(self.transparency_action)
        self.view_menu.addAction(self.toggle_edit_action)
        self.view_menu.addSeparator()

        # ----- Help Menu -----
        self.help_menu.addAction(self.help_locations_action)
        self.help_menu.addSeparator()
        self.help_menu.addAction(self.about_action)

        # ----- System-Tray Menu -----
        self.tray_menu = QMenu("Tray Menu", self)
        self.tray_menu.addAction(self.show_window_action)
        self.tray_menu.addAction(self.hide_window_action)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(self.tray_copy_sql_action)
        self.tray_menu.addSeparator()
        self.tray_menu.addAction(self.exit_action)

        if getattr(self, 'tray_icon', None):
            self.tray_icon.setContextMenu(self.tray_menu)

        logging.info("add_actions_to_menus: Menus populated.")

    def __init__(self, settings: AppSettings, usage_counts: UsageCounts):
        super().__init__()
        
        # Store settings and usage counting instances
        self.settings = settings
        self.usage_counts = usage_counts
        
        # Initialize query vault for internal storage
        self.query_vault = QueryVault()
        
        # Initialize bookmarks list to empty
        self.bookmarks = []
        # Cache for sorted bookmarks after filtering
        self.sorted_bookmarks_cache = []
        # File source tracking
        self.loaded_file_path = None
        # Current data source (DataGrip XML or Internal Vault)
        self.current_data_source = self.settings.get('data_source', SOURCE_DATAGRIP)
        
        # Temporary variables for context operations
        self.context_menu_item = None  # For tracking item under context menu
        
        # Window title setup 
        self.setWindowTitle(f"{APP_NAME}")
        
        # Apply saved window geometry if available
        self._load_and_apply_geometry()
        
        # Apply window transparency setting
        self._apply_transparency()
        
        # Initialize UI components
        self.setup_ui()
        
        # Create all actions (must be done before menus)
        self.create_actions()
        
        # Add actions to menus (must be done after create_actions)
        self.add_actions_to_menus()
        
        # Initialize context menu (must be done after create_actions)
        self.init_context_menu()
        
        # Load and apply theme preference
        is_dark = self.settings.get('dark_theme', True)  # Default to dark theme
        self.toggle_theme_action.setChecked(is_dark)
        self.apply_styles(is_dark=is_dark)
        
        # Initialize tray icon (must be done after create_actions)
        self.tray_icon = None
        self.init_tray_icon()
        
        # Connect signals (must be done after UI and actions are created)
        self.connect_signals()
        
        # Load initial data based on source
        if self.current_data_source == SOURCE_DATAGRIP:
            last_file = self.settings.get('loaded_copy_path')
            if last_file and os.path.isfile(last_file):
                self.load_bookmarks(last_file)
            else:
                self.update_bookmark_list()  # Show empty state
        else:
            self.load_queries_from_vault()
        
        # Show the window, init complete
        self.show()
        logging.info(f"{APP_NAME} window initialized and shown.")

    def connect_signals(self):
        """Connect all signals to their slots."""
        # Search and filter signals
        self.search_box.textChanged.connect(self.filter_bookmarks)
        self.search_title_radio.toggled.connect(self.filter_bookmarks)
        self.search_syntax_radio.toggled.connect(self.filter_bookmarks)
        self.search_both_radio.toggled.connect(self.filter_bookmarks)
        self.label_filter_combo.currentIndexChanged.connect(self.filter_bookmarks)
        
        # Font size signal
        self.font_size_combo.currentTextChanged.connect(self.update_font_size)
        
        # Data source signal
        self.source_combo.currentIndexChanged.connect(self.on_source_changed)
        
        # Button signals
        self.hide_button.clicked.connect(self.hide_to_tray)
        self.custom_close_button.clicked.connect(self.close_app)
        
        # List signals
        self.bookmark_list.itemDoubleClicked.connect(self.handle_item_action)
        
        logging.info("Signals connected successfully.")

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
        """Setup the UI components."""
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # --- Menu Bar ---
        self.menu_bar = self.menuBar()
        self.file_menu = self.menu_bar.addMenu("&File")
        self.view_menu = self.menu_bar.addMenu("&View")
        self.help_menu = self.menu_bar.addMenu("&Help")

        # --- Source Toggle UI ---
        source_layout = QHBoxLayout()
        source_layout.setSpacing(10)
        source_label = QLabel("Data Source:")
        source_label.setObjectName("source_label")
        self.source_combo = QComboBox()
        self.source_combo.addItem("DataGrip Export (XML)", SOURCE_DATAGRIP)
        self.source_combo.addItem("Internal Query Vault", SOURCE_INTERNAL)
        
        # Set current index based on settings
        current_source = self.settings.get('data_source', SOURCE_DATAGRIP)
        index = 0 if current_source == SOURCE_DATAGRIP else 1
        self.source_combo.setCurrentIndex(index)
        
        source_layout.addWidget(source_label)
        source_layout.addWidget(self.source_combo)
        source_layout.addStretch(1)  # Push widgets to the left
        
        main_layout.addLayout(source_layout)

        # --- Top Layout (Search, Font, Count) ---
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # --- Search Box ---
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search Queries...")
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
        
        # --- Label Filter ---
        label_filter_layout = QHBoxLayout()
        label_filter_layout.addWidget(QLabel("Label:"))
        self.label_filter_combo = QComboBox()
        self.label_filter_combo.addItem("All Labels", None)  # Default option
        self.update_label_filter_dropdown()  # Populate with available labels
        label_filter_layout.addWidget(self.label_filter_combo)
        top_layout.addLayout(label_filter_layout)
        
        top_layout.addStretch(1) # Add some space before font settings

        # --- Font Size Combo Box ---
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

        # --- Bookmark Count ---
        self.bookmark_count_label = QLabel("Queries: 0")
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
        self.no_bookmarks_label = QLabel("No queries loaded.")
        self.no_bookmarks_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        list_lay.addWidget(self.no_bookmarks_label)
        self.no_bookmarks_label.hide() # Initially hidden

        self.splitter.addWidget(list_cont)

        # Right side: Preview Pane
        # Replace existing QTextEdit preview pane with QScintilla
        self.preview_pane = QsciScintilla()
        self.preview_pane.setReadOnly(True)

        # Set a monospace font for code display
        code_font = QFont('Consolas, Courier New, monospace', 10)
        code_font.setFixedPitch(True)
        self.preview_pane.setFont(code_font)

        # Configure SQL syntax highlighting
        self.sql_lexer = QsciLexerSQL()
        self.sql_lexer.setDefaultFont(code_font)
        self.preview_pane.setLexer(self.sql_lexer)

        # Configure display options
        self.preview_pane.setMarginWidth(0, '00000')  # Line numbers width
        self.preview_pane.setMarginLineNumbers(0, True)
        self.preview_pane.setMarginsForegroundColor(QColor('#CCCCCC'))
        self.preview_pane.setMarginsBackgroundColor(QColor('#252526'))
        self.preview_pane.setUtf8(True)

        # Set dark theme colors for the editor
        self.preview_pane.setPaper(QColor('#252526'))  # Background
        self.preview_pane.setColor(QColor('#dcdcdc'))  # Text color

        # Configure line wrapping
        self.preview_pane.setWrapMode(QsciScintilla.WrapNone)

        # Configure selection colors
        self.preview_pane.setSelectionBackgroundColor(QColor('#264F78'))
        self.preview_pane.setSelectionForegroundColor(QColor('#ffffff'))
        
        # --- Build preview container with footer controls ---
        preview_cont = QWidget()
        preview_layout = QVBoxLayout(preview_cont)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(0)

        # Add main editor
        preview_layout.addWidget(self.preview_pane, 1)  # Stretch factor 1

        # Footer layout with Editable toggle & Add button
        preview_footer_layout = QHBoxLayout()
        preview_footer_layout.setContentsMargins(4, 4, 4, 4)

        self.edit_toggle_checkbox = QCheckBox("Editable")
        self.edit_toggle_checkbox.setChecked(False)
        self.edit_toggle_checkbox.toggled.connect(self.toggle_preview_editable)

        self.add_button = QPushButton("Add")
        self.add_button.clicked.connect(self.add_new_query)

        preview_footer_layout.addWidget(self.edit_toggle_checkbox)
        preview_footer_layout.addWidget(self.add_button)
        preview_footer_layout.addStretch()

        preview_layout.addLayout(preview_footer_layout)

        # Add the container to the splitter (instead of raw preview_pane)
        self.splitter.addWidget(preview_cont)

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
        # Add / Move Up / Move Down buttons
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
        
        # Set object names used in the stylesheet selectors
        self.centralWidget().setObjectName("centralWidget")
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
            
            # Set font for QScintilla preview pane
            self.preview_pane.setFont(font)
            
            # Update font for the SQL lexer too
            if hasattr(self, 'sql_lexer') and self.sql_lexer:
                lexer_font = QFont(font)
                self.sql_lexer.setFont(lexer_font)
                
            # Update the margin width for line numbers when font changes
            if hasattr(self, 'preview_pane'):
                fontmetrics = QFontMetrics(font)
                self.preview_pane.setMarginWidth(0, fontmetrics.horizontalAdvance("00000") + 5)
                
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
                self.preview_pane.setText("-- Select a bookmark to preview its SQL content...")


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
        self.tray_icon.activated.connect(self.handle_tray_icon_activation)

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
        logging.debug("Updating query list...")

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
                self.no_bookmarks_label.setText("No queries match your search.")
            elif not self.bookmarks:
                if self.current_data_source == SOURCE_DATAGRIP:
                    self.no_bookmarks_label.setText("No queries loaded. Use File > Open to load XML.")
                else:
                    self.no_bookmarks_label.setText("No queries in vault. Add new queries to get started.")
            else:
                if self.current_data_source == SOURCE_DATAGRIP:
                    self.no_bookmarks_label.setText("No queries available (check XML file?).")
                else:
                    self.no_bookmarks_label.setText("No queries available (vault may be corrupted).")
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
        
        # Update window title based on data source
        if self.current_data_source == SOURCE_DATAGRIP:
            original_file = self.settings.get('last_file_path', self.loaded_file_path) 
            if original_file and os.path.exists(original_file):
                self.setWindowTitle(f"{APP_NAME} - DataGrip Mode - {os.path.basename(original_file)}")
            else:
                self.setWindowTitle(f"{APP_NAME} - DataGrip Mode (No File Loaded)")
        else:
            self.setWindowTitle(f"{APP_NAME} - Internal Query Vault")
        
        logging.debug("Query list update complete.")

    def filter_bookmarks(self):
        """Filter bookmarks based on search term and update the list."""
        logging.debug("Filtering bookmarks...")
        self.update_bookmark_list()
        
    def apply_filter(self, search_term):
        """Filter bookmarks based on search criteria and return filtered list."""
        logging.debug(f"Applying filter with search term: '{search_term}'")
        if not self.bookmarks:
            return []
            
        # First apply label filter
        filtered = self.apply_label_filter(self.bookmarks)
        
        # If no search term, just return the label-filtered list
        if not search_term:
            return filtered
            
        # Apply search criteria
        search_term = search_term.lower()
        results = []
        
        for bm in filtered:
            title = bm.get('name', '').lower()
            
            # Determine search scope based on radio button selection
            search_title = self.search_title_radio.isChecked() or self.search_both_radio.isChecked()
            search_syntax = self.search_syntax_radio.isChecked() or self.search_both_radio.isChecked()
            
            # Match title if searching titles
            if search_title and search_term in title:
                results.append(bm)
                continue
                
            # Match SQL content if searching syntax
            if search_syntax:
                # Get SQL content
                sql_content = self.get_sql_content(bm)
                if sql_content and search_term in sql_content.lower():
                    results.append(bm)
                    
        logging.debug(f"Search filter results: {len(results)}/{len(filtered)} queries matching '{search_term}'")
        return results
    
    def apply_sort(self, bookmarks):
        """Sort the filtered bookmarks by name (alphabetically)."""
        if not bookmarks:
            return []
            
        # Simple alphabetical sort by name
        return sorted(bookmarks, key=lambda bm: bm.get('name', '').lower())

    # --- Preview Pane and Highlighting ---
    def update_preview_pane(self, item: QListWidgetItem):
        """Update the preview pane with the SQL content of the selected bookmark/query."""
        if not item:
            self.preview_pane.clear()
            self.preview_pane.setText("-- Select a query to preview its SQL content...")
            return
            
        data = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(data, dict):
            self.preview_pane.setText("-- Error: Invalid query data.")
            return
            
        # Get SQL content using our helper that handles both DataGrip and internal storage
        sql_content = self.get_sql_content(data)
        
        if not sql_content:
            # No content available
            if self.current_data_source == SOURCE_INTERNAL:
                self.preview_pane.setText("-- This query has no content.")
            else:
                # For DataGrip mode, show more helpful error
                self.preview_pane.setText(
                    "-- SQL content could not be loaded.\n"
                    "-- Possible reasons:\n"
                    "-- 1. SQL file not found (check SQL Root Directory in File menu)\n"
                    "-- 2. Read permission denied\n"
                    "-- 3. File has been moved or deleted"
                )
            return
            
        # Update the preview pane with the content (QScintilla will handle the syntax highlighting)
        self.preview_pane.setText(sql_content)
        
        # Apply search term highlighting if there's a current search
        self.highlight_search_results()

    def highlight_sql_syntax(self):
        """SQL syntax highlighting is now handled by QScintilla's lexer."""
        # This method is kept as a placeholder to avoid breaking any existing calls
        # The actual highlighting is done automatically by the QsciLexerSQL
        pass

    def highlight_search_results(self):
        """Highlight search terms in the preview pane using QScintilla search."""
        # Get the current search text
        search_text = self.search_box.text()
        if not search_text:
            # Clear any existing indicators if search is empty
            self.preview_pane.clearIndicatorRange(0, 0, self.preview_pane.lines(), 0, 0)
            return  # No search term to highlight
            
        # Use QScintilla's search and indicator mechanism
        # First, define an indicator for search results
        SEARCH_INDICATOR = 0  # Use indicator number 0
        self.preview_pane.indicatorDefine(QsciScintilla.SquigglePixmapIndicator, SEARCH_INDICATOR)
        self.preview_pane.setIndicatorForegroundColor(QColor("#4a4a0a"), SEARCH_INDICATOR)  # Yellow-ish background
        
        # Clear any existing indicators
        self.preview_pane.clearIndicatorRange(0, 0, self.preview_pane.lines(), 0, 0)
        
        # Get the text to search within
        text = self.preview_pane.text()
        if not text:
            return
            
        # Find all occurrences and highlight them
        pos = 0
        while True:
            # Find the next occurrence
            pos = text.lower().find(search_text.lower(), pos)
            if pos == -1:
                break
                
            # Convert position to line and index
            line, index = self._position_to_line_index(text, pos)
            
            # Set indicator for the found text
            self.preview_pane.fillIndicatorRange(line, index, line, index + len(search_text), SEARCH_INDICATOR)
            
            # Move to position after this match
            pos += len(search_text)
            
    def _position_to_line_index(self, text, position):
        """Helper function to convert a flat position to line and index within that line."""
        # Split text into lines
        lines = text.split('\n')
        
        # Track position
        current_pos = 0
        for line_num, line in enumerate(lines):
            line_length = len(line) + 1  # +1 for newline character
            if current_pos + line_length > position:
                # Found the line, calculate index within the line
                index = position - current_pos
                return line_num, index
            current_pos += line_length
            
        # Default fallback (shouldn't reach here with valid input)
        return 0, 0

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
        if self.context_menu_item:
            logging.debug("Context menu 'Copy SQL' triggered.")
            self.handle_item_action(self.context_menu_item) # Reuse the main action handler
        else:
            logging.warning("Context menu 'Copy SQL' triggered but context_item is None.")

    @Slot()
    def copy_bookmark_url_from_context(self):
        """Triggered by the 'Copy File URL' context menu action."""
        if self.context_menu_item:
            data = self.context_menu_item.data(Qt.ItemDataRole.UserRole)
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
        self.context_menu_item = item # Store the item that was right-clicked
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
             self.context_menu_item = None

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
                from PyQt5.QtGui import QDesktopServices
                from PyQt5.QtCore import QUrl

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

    def load_queries_from_vault(self):
        """Load queries from the internal query vault."""
        self.query_vault.load_vault()
        self.bookmarks = self.query_vault.get_queries()
        self.update_bookmark_list()
        logging.info("Queries loaded from internal query vault.")
        self.setWindowTitle(f"{APP_NAME} - Internal Query Vault")

    def update_label_filter_dropdown(self):
        """Update the label filter dropdown with all available labels."""
        current_label = None
        
        # Store currently selected label if any
        if hasattr(self, 'label_filter_combo') and self.label_filter_combo.currentIndex() > 0:
            current_label = self.label_filter_combo.currentData()
            
        # Clear and add the default "All Labels" option
        if hasattr(self, 'label_filter_combo'):
            self.label_filter_combo.clear()
            self.label_filter_combo.addItem("All Labels", None)
        
        # Get all unique labels from the appropriate source
        all_labels = []
        if self.current_data_source == SOURCE_INTERNAL:
            all_labels = self.query_vault.get_all_labels()
        else:
            # For DataGrip mode, extract labels from loaded bookmarks
            labels = set()
            for bm in self.bookmarks:
                if isinstance(bm, dict) and 'labels' in bm and isinstance(bm['labels'], list):
                    labels.update(bm['labels'])
            all_labels = sorted(list(labels))
        
        # Add all labels to the dropdown
        if hasattr(self, 'label_filter_combo'):
            for label in all_labels:
                self.label_filter_combo.addItem(label, label)
                
            # Restore previously selected label if it still exists
            if current_label:
                index = self.label_filter_combo.findData(current_label)
                if index >= 0:
                    self.label_filter_combo.setCurrentIndex(index)
                    
            logging.debug(f"Updated label filter dropdown with {len(all_labels)} labels")

    def apply_label_filter(self, bookmarks):
        """Filter bookmarks by selected label."""
        if not hasattr(self, 'label_filter_combo'):
            return bookmarks
            
        selected_label = self.label_filter_combo.currentData()
        
        # If no label filter active, return all bookmarks
        if selected_label is None:
            return bookmarks
            
        # Filter by the selected label
        filtered = []
        for bm in bookmarks:
            if isinstance(bm, dict) and 'labels' in bm and isinstance(bm['labels'], list):
                if selected_label in bm['labels']:
                    filtered.append(bm)
        
        logging.debug(f"Label filter '{selected_label}' active: {len(filtered)}/{len(bookmarks)} queries matching")
        return filtered

    @Slot()
    def on_source_changed(self):
        """Handle data source change."""
        if not hasattr(self, 'source_combo'):
            return
            
        new_source = self.source_combo.currentData()
        if new_source == self.current_data_source:
            return  # No change
            
        logging.info(f"Data source changing from {self.current_data_source} to {new_source}")
        
        # Save the new source setting
        self.settings.set('data_source', new_source)
        self.current_data_source = new_source
        
        # Load data from the appropriate source
        if new_source == SOURCE_DATAGRIP:
            # Load from DataGrip XML export
            last_file = self.settings.get('loaded_copy_path')
            if last_file and os.path.isfile(last_file):
                self.load_bookmarks(last_file)
            else:
                # No file loaded yet
                self.bookmarks = []
                self.update_bookmark_list()
                self.setWindowTitle(f"{APP_NAME} - DataGrip Mode (No File Loaded)")
        else:
            # Load from internal vault
            self.load_queries_from_vault()
            
        # Update UI
        self.update_label_filter_dropdown()
        
    def add_query_to_vault(self, title, sql_content, labels=None):
        """Add a new query to the internal vault."""
        if self.current_data_source != SOURCE_INTERNAL:
            logging.warning("Cannot add query to vault: current source is not Internal Query Vault")
            return False
            
        # Create query data structure
        query_data = {
            'title': title,
            'sql_content': sql_content,
            'labels': labels or [],
            'count': 0
        }
        
        # Add to vault
        if self.query_vault.add_query(query_data):
            # Save vault
            self.query_vault.save_vault()
            # Reload queries to refresh the list
            self.load_queries_from_vault()
            logging.info(f"Added new query '{title}' to vault")
            return True
        return False
        
    def edit_query_in_vault(self, query_id, title=None, sql_content=None, labels=None):
        """Edit an existing query in the internal vault."""
        if self.current_data_source != SOURCE_INTERNAL:
            logging.warning("Cannot edit query in vault: current source is not Internal Query Vault")
            return False
            
        # Get the current query data
        current_query = self.query_vault.get_query_by_id(query_id)
        if not current_query:
            logging.warning(f"Cannot edit query: query with ID {query_id} not found")
            return False
            
        # Update only the provided fields
        if title is not None:
            current_query['title'] = title
        if sql_content is not None:
            current_query['sql_content'] = sql_content
        if labels is not None:
            current_query['labels'] = labels
            
        # Update in vault
        if self.query_vault.update_query(query_id, current_query):
            # Save vault
            self.query_vault.save_vault()
            # Reload queries to refresh the list
            self.load_queries_from_vault()
            logging.info(f"Updated query with ID {query_id} in vault")
            return True
        return False
        
    def add_label_to_query_in_vault(self, query_id, label):
        """Add a label to a query in the internal vault."""
        if self.current_data_source != SOURCE_INTERNAL:
            logging.warning("Cannot add label: current source is not Internal Query Vault")
            return False
            
        if self.query_vault.add_label_to_query(query_id, label):
            # Save vault
            self.query_vault.save_vault()
            # Reload queries to refresh the list
            self.load_queries_from_vault()
            logging.info(f"Added label '{label}' to query {query_id}")
            return True
        return False
        
    def remove_label_from_query_in_vault(self, query_id, label):
        """Remove a label from a query in the internal vault."""
        if self.current_data_source != SOURCE_INTERNAL:
            logging.warning("Cannot remove label: current source is not Internal Query Vault")
            return False
            
        if self.query_vault.remove_label_from_query(query_id, label):
            # Save vault
            self.query_vault.save_vault()
            # Reload queries to refresh the list
            self.load_queries_from_vault()
            logging.info(f"Removed label '{label}' from query {query_id}")
            return True
        return False
        
    def delete_query_from_vault(self, query_id):
        """Delete a query from the internal vault."""
        if self.current_data_source != SOURCE_INTERNAL:
            logging.warning("Cannot delete query: current source is not Internal Query Vault")
            return False
            
        if self.query_vault.delete_query(query_id):
            # Save vault
            self.query_vault.save_vault()
            # Reload queries to refresh the list
            self.load_queries_from_vault()
            logging.info(f"Deleted query with ID {query_id} from vault")
            return True
        return False
        
    def import_bookmark_to_vault(self, bookmark_data):
        """Import a DataGrip bookmark into the internal vault."""
        if not isinstance(bookmark_data, dict):
            logging.warning("Cannot import bookmark: invalid data format")
            return False
            
        # Get SQL content
        sql_content = self.get_sql_content(bookmark_data)
        if not sql_content:
            logging.warning(f"Cannot import bookmark '{bookmark_data.get('title', 'Untitled')}': failed to retrieve SQL content")
            return False
            
        # Create query data
        query_data = {
            'title': bookmark_data.get('title', 'Imported Query'),
            'sql_content': sql_content,
            'labels': ['Imported'],
            'count': bookmark_data.get('count', 0)
        }
        
        # Add to vault
        if self.query_vault.add_query(query_data):
            # Save vault
            self.query_vault.save_vault()
            logging.info(f"Imported bookmark '{query_data['title']}' to vault")
            return True
        return False

    def get_sql_content(self, bookmark_data, sql_root=None, user_home=None, pre_resolved_path=None):
        """Get SQL content from a bookmark data dict, handling both DataGrip bookmarks and internal queries."""
        if not bookmark_data or not isinstance(bookmark_data, dict):
            return ""
            
        # First check if this is an internal vault query with direct SQL content
        if 'sql_content' in bookmark_data and bookmark_data['sql_content']:
            return bookmark_data['sql_content']
            
        # If not, attempt to load from file (DataGrip bookmark)
        url = bookmark_data.get('url', '')
        if not url:
            logging.warning(f"Bookmark '{bookmark_data.get('title', 'Untitled')}' has no URL.")
            return ""
            
        # Use provided SQL root directory or get from settings
        sql_root_dir = sql_root if sql_root else self.settings.get('sql_root_directory')
        # Use provided user home or get from os.path
        home_dir = user_home if user_home else os.path.expanduser("~")
        
        # If path already resolved (to avoid redundant work in multiple calls)
        if pre_resolved_path and os.path.isfile(pre_resolved_path):
            file_path = pre_resolved_path
        else:
            # Resolve the URL to a file path
            file_path = self.resolve_file_path(url, sql_root=sql_root_dir, user_home=home_dir)
            
        # Read and return file content if resolved successfully
        if file_path and os.path.isfile(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                return content
            except Exception as e:
                logging.error(f"Failed to read SQL file '{file_path}': {e}", exc_info=True)
                return f"// Error reading file: {e}"
        else:
            logging.warning(f"SQL file not found for bookmark '{bookmark_data.get('title', 'Untitled')}': {file_path}")
            return "// File not found (check SQL root directory setting)"

    def copy_current_query_to_clipboard(self):
        """Copy the currently selected query to clipboard."""
        current_item = self.bookmark_list.currentItem()
        if current_item:
            self.handle_item_action(current_item)
        else:
            logging.warning("No query selected to copy to clipboard")

    def close_app(self):
        """Close the application."""
        self.save_state()
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.instance().quit()

    def init_context_menu(self):
        """Initialize the context menu for the bookmark list."""
        self.context_menu = QMenu(self)
        
        # Create actions for the context menu
        self.copy_sql_action_context = QAction("Copy SQL", self)
        self.copy_sql_action_context.triggered.connect(self.handle_item_action_from_context)
        
        self.copy_url_action_context = QAction("Copy File URL", self)
        self.copy_url_action_context.triggered.connect(self.copy_bookmark_url_from_context)
        
        # Add basic copy actions first
        self.context_menu.addAction(self.copy_sql_action_context)
        self.context_menu.addAction(self.copy_url_action_context)
        
        # Additional actions
        self.context_menu.addSeparator()
        self.context_menu.addAction(self.edit_query_action_context)
        self.context_menu.addAction(self.delete_query_action_context)
        self.context_menu.addAction(self.manage_labels_action_context)
        self.context_menu.addAction(self.import_to_vault_action_context)
        
        # Connect the context menu to the list widget
        self.bookmark_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.bookmark_list.customContextMenuRequested.connect(self.show_context_menu)
        
        logging.info("Context menu initialized and connected.")

    def apply_styles(self, is_dark=True):
        """Apply stylesheet and visual styling to the application."""
        if is_dark:
            # Dark theme colors
            colors = {
                'bg': '#2d2d2d',
                'text': '#ffffff',
                'input_bg': '#3d3d3d',
                'border': '#555555',
                'selection': '#264f78',
                'hover': '#404040',
                'button': '#0078d4',
                'button_hover': '#1084d8',
                'button_pressed': '#006cbd',
                'disabled': '#555555',
                'disabled_text': '#888888',
                'preview_bg': '#1e1e1e',
                'preview_text': '#d4d4d4'
            }
        else:
            # Light theme colors
            colors = {
                'bg': '#ffffff',
                'text': '#000000',
                'input_bg': '#f0f0f0',
                'border': '#cccccc',
                'selection': '#cce8ff',
                'hover': '#e5e5e5',
                'button': '#0078d4',
                'button_hover': '#1084d8',
                'button_pressed': '#006cbd',
                'disabled': '#e0e0e0',
                'disabled_text': '#666666',
                'preview_bg': '#ffffff',
                'preview_text': '#000000'
            }

        style_sheet = """
            QMainWindow, QWidget {
                background-color: %(bg)s;
                color: %(text)s;
            }
            QLineEdit {
                background-color: %(input_bg)s;
                color: %(text)s;
                border: 1px solid %(border)s;
                border-radius: 3px;
                padding: 2px;
            }
            QListWidget {
                background-color: %(bg)s;
                color: %(text)s;
                border: 1px solid %(border)s;
                border-radius: 3px;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: %(selection)s;
                color: %(text)s;
            }
            QListWidget::item:hover {
                background-color: %(hover)s;
            }
            QPushButton {
                background-color: %(button)s;
                color: #ffffff;
                border: none;
                border-radius: 3px;
                padding: 5px 15px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: %(button_hover)s;
            }
            QPushButton:pressed {
                background-color: %(button_pressed)s;
            }
            QPushButton:disabled {
                background-color: %(disabled)s;
                color: %(disabled_text)s;
            }
            QLabel {
                color: %(text)s;
            }
            QMenuBar {
                background-color: %(bg)s;
                color: %(text)s;
            }
            QMenuBar::item {
                background-color: transparent;
                padding: 4px 10px;
            }
            QMenuBar::item:selected {
                background-color: %(hover)s;
            }
            QMenu {
                background-color: %(bg)s;
                color: %(text)s;
                border: 1px solid %(border)s;
            }
            QMenu::item {
                padding: 5px 20px;
            }
            QMenu::item:selected {
                background-color: %(hover)s;
            }
            QSplitter::handle {
                background-color: %(border)s;
            }
            QComboBox {
                background-color: %(input_bg)s;
                color: %(text)s;
                border: 1px solid %(border)s;
                border-radius: 3px;
                padding: 2px;
            }
            QComboBox QAbstractItemView {
                background-color: %(bg)s;
                color: %(text)s;
                selection-background-color: %(selection)s;
            }
            QRadioButton {
                color: %(text)s;
                spacing: 5px;
            }
            QRadioButton::indicator {
                width: 13px;
                height: 13px;
            }
            QRadioButton::indicator:unchecked {
                background-color: %(input_bg)s;
                border: 2px solid %(border)s;
                border-radius: 7px;
            }
            QRadioButton::indicator:checked {
                background-color: %(button)s;
                border: 2px solid %(button)s;
                border-radius: 7px;
            }
            #bookmark_count_label {
                color: %(disabled_text)s;
                padding-right: 10px;
            }
            #hide_button {
                background-color: %(button)s;
            }
            #hide_button:hover {
                background-color: %(button_hover)s;
            }
            #custom_close_button {
                background-color: #c42b1c;
            }
            #custom_close_button:hover {
                background-color: #d13438;
            }
        """ % colors
        
        self.setStyleSheet(style_sheet)
        
        # Apply theme to QScintilla preview pane
        if hasattr(self, 'preview_pane'):
            bg_col = QColor(colors['preview_bg'])
            fg_col = QColor(colors['preview_text'])
            self.preview_pane.setColor(fg_col)  # Default text color
            self.preview_pane.setPaper(bg_col)  # Background color
            self.preview_pane.setMarginsBackgroundColor(bg_col)
            self.preview_pane.setMarginsForegroundColor(fg_col)

            # Ensure QScintilla lexer adopts the same palette
            if hasattr(self, 'sql_lexer') and self.sql_lexer:
                self.sql_lexer.setDefaultColor(fg_col)
                self.sql_lexer.setDefaultPaper(bg_col)
                for style_idx in range(128):
                    if self.sql_lexer.description(style_idx):
                        self.sql_lexer.setColor(fg_col, style_idx)
                        self.sql_lexer.setPaper(bg_col, style_idx)

        logging.info(f"Applied {'dark' if is_dark else 'light'} theme styles.")

    def toggle_theme(self):
        """Toggle between light and dark theme."""
        # Check current theme state from the action
        is_dark = self.toggle_theme_action.isChecked()
        # Apply the appropriate theme
        self.apply_styles(is_dark=is_dark)
        # Save the theme preference in settings
        self.settings.set('dark_theme', is_dark)
        logging.info(f"Theme toggled to {'dark' if is_dark else 'light'} mode")

    @Slot()
    def edit_query(self):
        """Placeholder for editing the selected query/bookmark."""
        if self.context_menu_item:
            data = self.context_menu_item.data(Qt.ItemDataRole.UserRole)
            title = data.get('title', 'Untitled') if isinstance(data, dict) else 'Untitled'
            logging.info(f"Edit requested for item: {title}")
            QMessageBox.information(self, "Edit Query", "The edit feature is not implemented yet.")
        else:
            logging.warning("edit_query triggered but no context_menu_item is set.")
            QMessageBox.information(self, "Edit Query", "Please right-click a query and choose Edit to modify it.")

    @Slot()
    def delete_query(self):
        """Delete the selected query from its source after confirmation."""
        if not self.context_menu_item:
            logging.warning("delete_query triggered but no context_menu_item is set.")
            return

        data = self.context_menu_item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(data, dict):
            logging.warning("delete_query: item data is invalid or not a dict.")
            return

        title = data.get('title', 'Untitled')
        query_id = data.get('id')

        reply = QMessageBox.question(
            self,
            "Delete Query",
            f"Are you sure you want to delete '{title}'? This action cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # Delete depending on current data source
        if self.current_data_source == SOURCE_INTERNAL:
            if query_id and self.delete_query_from_vault(query_id):
                QMessageBox.information(self, "Delete", f"'{title}' has been deleted.")
            else:
                QMessageBox.warning(self, "Delete Failed", "Could not delete the selected query.")
        else:
            QMessageBox.warning(self, "Delete Not Supported", "Deleting is only supported for Internal Query Vault items.")

    @Slot()
    def manage_labels(self):
        """Placeholder to manage labels for the selected query."""
        QMessageBox.information(self, "Manage Labels", "Label management is not implemented yet.")
        logging.info("manage_labels called but not yet implemented.")

    @Slot()
    def import_to_vault(self):
        """Import a DataGrip bookmark into the internal vault."""
        if not self.context_menu_item:
            logging.warning("import_to_vault triggered but no context_menu_item is set.")
            return

        data = self.context_menu_item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(data, dict):
            logging.warning("import_to_vault: item data is invalid or not a dict.")
            return

        title = data.get('title', 'Untitled')
        if self.current_data_source == SOURCE_DATAGRIP:
            if self.import_bookmark_to_vault(data):
                QMessageBox.information(self, "Import Successful", f"'{title}' has been imported to the Internal Query Vault.")
            else:
                QMessageBox.warning(self, "Import Failed", "Failed to import the selected bookmark to the vault. Check logs for details.")
        else:
            QMessageBox.information(self, "Already in Vault", "The selected query is already in the Internal Query Vault.")

    def resolve_file_path(self, url_string, sql_root=None, user_home=None):
        """Resolve JetBrains-style URL placeholders and return a valid file path if it exists.

        Parameters
        ----------
        url_string : str
            The raw URL string stored in a DataGrip bookmark (may start with 'file://', or contain
            placeholders like $PROJECT_DIR$).
        sql_root : str, optional
            Override for the SQL root directory. Falls back to the setting 'sql_root_directory'.
        user_home : str, optional
            Override for the user's home directory. Defaults to os.path.expanduser("~").

        Returns
        -------
        str | None
            The resolved absolute path if the file exists on disk, otherwise None.
        """
        if not url_string:
            return None

        from urllib.parse import urlparse, unquote

        # Determine directories to substitute
        sql_root_dir = sql_root if sql_root is not None else self.settings.get('sql_root_directory')
        home_dir = user_home if user_home is not None else os.path.expanduser("~")

        # Start with the raw string
        path_str = str(url_string)

        # Strip URI scheme if present
        if path_str.startswith("file://"):
            path_str = unquote(urlparse(path_str).path)
            # On Windows, remove leading slash before drive letter (e.g. /C:/.. -> C:/..)
            if os.name == 'nt' and path_str.startswith('/') and len(path_str) > 2 and path_str[2] == ':':
                path_str = path_str[1:]

        # Replace JetBrains placeholders
        if sql_root_dir:
            path_str = path_str.replace("$PROJECT_DIR$", sql_root_dir)
        path_str = path_str.replace("$USER_HOME$", home_dir)

        # If still relative, join with sql_root_dir if available
        if not os.path.isabs(path_str):
            if sql_root_dir:
                candidate = os.path.normpath(os.path.join(sql_root_dir, path_str))
            else:
                candidate = os.path.abspath(path_str)
        else:
            candidate = os.path.normpath(path_str)

        return candidate if os.path.isfile(candidate) else None

    @Slot()
    def set_sql_root_directory(self):
        """Prompt user to select the root directory that replaces $PROJECT_DIR$ placeholders."""
        current_root = self.settings.get('sql_root_directory', os.path.expanduser("~"))
        dir_path = QFileDialog.getExistingDirectory(self, "Select SQL Root Directory", current_root)
        if dir_path:
            self.settings.set('sql_root_directory', dir_path)
            QMessageBox.information(self, "SQL Root Directory Set", f"Queries will resolve $PROJECT_DIR$ to:\n{dir_path}")
            logging.info(f"SQL root directory updated to: {dir_path}")

    @Slot()
    def show_about_dialog(self):
        """Display an About dialog with basic app information."""
        QMessageBox.about(
            self,
            f"About {APP_NAME}",
            (
                f"<b>{APP_NAME} – Dynamic Query Vault</b><br><br>"
                "A tool to browse DataGrip export bookmarks and your own internal SQL vault, "
                "copy queries quickly, and manage labels.<br><br>"
                "Version: 0.1.0 (alpha)"
            ),
        )

    @Slot(bool)
    def toggle_preview_editable(self, checked=False):
        """Sync preview pane editability with checkbox/menu."""
        # `checked` may be omitted when the slot is invoked without parameters.
        checked_bool = bool(checked)

        self.preview_pane.setReadOnly(not checked_bool)

        # Sync counterparts without recursion
        if hasattr(self, 'toggle_edit_action') and self.toggle_edit_action.isChecked() != checked_bool:
            self.toggle_edit_action.blockSignals(True)
            self.toggle_edit_action.setChecked(checked_bool)
            self.toggle_edit_action.blockSignals(False)

        if hasattr(self, 'edit_toggle_checkbox') and self.edit_toggle_checkbox.isChecked() != checked_bool:
            self.edit_toggle_checkbox.blockSignals(True)
            self.edit_toggle_checkbox.setChecked(checked_bool)
            self.edit_toggle_checkbox.blockSignals(False)

        logging.info(f"Preview pane edit mode {'ON' if checked_bool else 'OFF'}.")

    @Slot()
    def add_new_query(self):
        # Implement the logic for adding a new query
        logging.info("Add new query functionality not implemented yet.")
        QMessageBox.information(self, "Add New Query", "This functionality is not implemented yet.")

if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication

    app = QApplication(sys.argv)
    viewer = FloatingBookmarksWindow(AppSettings(), UsageCounts())
    viewer.show()
    sys.exit(app.exec_())