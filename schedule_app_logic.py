import sys
import os
import pandas as pd
import re
import json
import urllib.request
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QCheckBox, QScrollArea, QFrame,
    QMessageBox, QSizePolicy, QMenu, QDialog, QColorDialog, QTabWidget, QGroupBox,
    QDialogButtonBox, QProgressBar, QListWidget, QDateEdit, QCalendarWidget, QSlider
)
from PyQt6.QtCore import Qt, pyqtSignal, QEvent, QSize, QDate
from PyQt6.QtGui import QPalette, QColor, QIcon

# --- Configuration ---
def get_app_path():
    """Returns the path to the folder containing the exe (if frozen) or script."""
    if getattr(sys, 'frozen', False):
        # We are running as an exe (Launcher)
        return os.path.dirname(sys.executable)
    else:
        # We are running as a script
        return os.path.dirname(os.path.abspath(__file__))

# Define paths relative to the executable location
BASE_DIR = get_app_path()
EXCEL_FOLDER_NAME = "excel_files"
EXCEL_FOLDER = os.path.join(BASE_DIR, EXCEL_FOLDER_NAME)
RECENTLY_DELETED_FILE = "recently_deleted_employees.xlsx" # For undo functionality
EMPLOYEE_FILE = "Employees_Full_List.xlsx" # Master list
SETTINGS_FILE = "app_settings.json"
DRAFT_FILE = "draft_schedule.json"
DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
INITIAL_DISPLAY_ROWS_SCHEDULE = 20 # Initial rows for final schedule view
INITIAL_ROWS_UNASSIGNED = 20 # Initial rows for unassigned view

# --- Dummy File Creation ---
if not os.path.exists(EXCEL_FOLDER):
    os.makedirs(EXCEL_FOLDER)
    print(f"Created directory: {EXCEL_FOLDER}")
    dummy_master_path = os.path.join(EXCEL_FOLDER, EMPLOYEE_FILE)
    if not os.path.exists(dummy_master_path):
        try:
            dummy_df = pd.DataFrame({
                'Name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve', 'Frank'],
                'Sun': ['Yes', 'No', 'Yes', 'Yes', 'No', 'Yes'],
                'Mon': ['Yes', 'Yes', 'No', 'Yes', 'Yes', 'No'],
                'Tue': ['No', 'Yes', 'Yes', 'No', 'Yes', 'Yes'],
                'Wed': ['Yes', 'No', 'No', 'Yes', 'No', 'No'],
                'Thu': ['Yes', 'Yes', 'Yes', 'Yes', 'Yes', 'Yes'],
                'Fri': ['No', 'Yes', 'No', 'No', 'Yes', 'No'],
                'Sat': ['Yes', 'No', 'Yes', 'Yes', 'No', 'Yes'],
                'Set Schedule': ['Yes', 'No', 'No', 'No', 'No', 'No'],
                'Max Per Week': [4, 5, 3, 5, 5, 2],
                'Time Off Dates': ['07/04/2024', '', '12/20/2024-12/28/2024', '01/01/2025', '', '']
            })
            dummy_df.to_excel(dummy_master_path, index=False)
            print(f"Created dummy master file: {dummy_master_path}")
        except Exception as e:
            print(f"Error creating dummy master file: {e}")

master_employee_file_path = os.path.join(EXCEL_FOLDER, EMPLOYEE_FILE)
try:
    xlsx_files = [f for f in os.listdir(EXCEL_FOLDER) if f.endswith('.xlsx') and f != EMPLOYEE_FILE]
except FileNotFoundError:
    xlsx_files = []
    print(f"Error: Directory not found: {EXCEL_FOLDER}")

# --- Data Classes (Unchanged from original) ---
class Employee:
    def __init__(self, name, availability):
        self.name = name
        self.availability = availability
    def __repr__(self):
        return f"Employee({self.name})"

class Schedule:
    def __init__(self):
        self.scheduled = {day: [] for day in DAYS}
    def clear_schedule(self):
        self.scheduled = {day: [] for day in DAYS}
    def assign_employee(self, emp_to_assign: Employee, day: str):
        if not isinstance(emp_to_assign, Employee): return False
        if any(e.name == emp_to_assign.name for e in self.scheduled[day]): return False
        self.scheduled[day].append(emp_to_assign)
        self.scheduled[day].sort(key=lambda e: e.name)
        return True
    def remove_employee(self, emp_to_remove: Employee, day: str):
        if not isinstance(emp_to_remove, Employee): return False
        initial_len = len(self.scheduled[day])
        self.scheduled[day] = [e for e in self.scheduled[day] if e.name != emp_to_remove.name]
        return len(self.scheduled[day]) < initial_len
    def is_scheduled_on_any_day(self, emp_to_check: Employee):
        if not isinstance(emp_to_check, Employee): return False
        return any(emp_to_check.name == e.name for day_list in self.scheduled.values() for e in day_list)

# --- Week Selector Dialog ---
class WeekSelectorDialog(QDialog):
    """A popup dialog containing just a calendar for selecting a week."""
    def __init__(self, parent=None, current_date=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Popup)

        self.selected_date = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(1, 3, 0, 0)

        self.calendar = QCalendarWidget()
        if current_date:
            self.calendar.setSelectedDate(current_date)

        self.calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.ISOWeekNumbers)

        self.week_header_label = QLabel("Week", self)
        self.week_header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        font = self.calendar.font()
        font.setBold(True)
        self.week_header_label.setFont(font)

        # Find the navigation bar to dynamically style our label later
        self.nav_bar = self.calendar.findChild(QWidget, "qt_calendar_navigationbar")

        self.calendar.clicked.connect(self.on_date_selected)
        layout.addWidget(self.calendar)

    def showEvent(self, event):
        """Called just before the dialog is shown. Used for final positioning and styling."""
        super().showEvent(event)
        self.reposition_and_style_header()

    def reposition_and_style_header(self):
        """Calculates position and copies style from the calendar's nav bar."""
        nav_bar_height = 30 # Default height
        bg_color = "transparent"

        if self.nav_bar:
            nav_bar_height = self.nav_bar.height()
            # Extract background-color from the navigation bar's stylesheet
            match = re.search(r"background-color:\s*([^;}]+)", self.nav_bar.styleSheet())
            if match:
                bg_color = match.group(1)

        # Apply the extracted background color to our label for a seamless look
        self.week_header_label.setStyleSheet(f"background-color: {bg_color};")

        # Fine-tuned position
        x_offset = 1
        y_offset = nav_bar_height + 1
        header_height = 25
        column_width = 35

        self.week_header_label.setGeometry(x_offset, y_offset, column_width, header_height)
        self.week_header_label.raise_()

    def on_date_selected(self, date):
        self.selected_date = date
        self.accept()

# --- Settings and Target Employees Window ---
class SettingsWindow(QDialog):
    """
    A dialog window for managing application settings, including target
    employee counts and theme colors.
    """
    def __init__(self, main_window):
        super().__init__(main_window)
        self.main_window = main_window
        self.setWindowTitle("⚙️ Settings")
        self.setMinimumSize(600, 450)

        # Temporary storage for settings until 'Apply' is clicked
        self.temp_theme_config = self.main_window.theme_config.copy()
        self.temp_max_per_day = self.main_window.max_per_day.copy()

        # Main layout
        layout = QVBoxLayout(self)

        # Tabs for different settings categories
        tab_widget = QTabWidget()
        tab_widget.addTab(self._create_general_tab(), "General")
        tab_widget.addTab(self._create_theme_tab('light'), "Light Theme")
        tab_widget.addTab(self._create_theme_tab('dark'), "Dark Theme")
        layout.addWidget(tab_widget)

        # Standard OK/Apply/Cancel buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Apply |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept) # OK
        button_box.rejected.connect(self.reject) # Cancel
        button_box.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.apply_changes)
        layout.addWidget(button_box)

    def _create_general_tab(self):
        """Creates the 'General' settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # --- Target Employees Settings ---
        targets_group = QGroupBox("Target Employees Per Day")
        targets_layout = QGridLayout(targets_group)

        self.max_entries_edits = {}
        for i, day in enumerate(DAYS):
            day_label = QLabel(day)
            entry = QLineEdit(str(self.temp_max_per_day.get(day, 35)))
            entry.setFixedWidth(50)
            entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.max_entries_edits[day] = entry

            col_layout = QVBoxLayout()
            col_layout.addWidget(day_label)
            col_layout.addWidget(entry)
            targets_layout.addLayout(col_layout, 0, i)

        layout.addWidget(targets_group)

        # --- Font Size Settings (with Sliders) ---
        font_group = QGroupBox("Font Sizes")
        font_layout = QGridLayout(font_group)

        # --- Header Font Slider ---
        font_layout.addWidget(QLabel("Headers and Buttons:"), 0, 0)

        self.header_font_slider = QSlider(Qt.Orientation.Horizontal)
        self.header_font_slider.setRange(8, 14)
        initial_header_size = self.temp_theme_config.get('header_font_size', 10)
        self.header_font_slider.setValue(initial_header_size)

        # Label to show the slider's current value
        self.header_font_value_label = QLabel(f"{initial_header_size} pt")
        self.header_font_value_label.setFixedWidth(40) # Ensures layout stability

        # Connect slider movement to update the value label
        self.header_font_slider.valueChanged.connect(
            lambda value: self.header_font_value_label.setText(f"{value} pt")
        )

        font_layout.addWidget(self.header_font_slider, 0, 1)
        font_layout.addWidget(self.header_font_value_label, 0, 2)

        # --- Body Font Slider ---
        font_layout.addWidget(QLabel("Body Text:"), 1, 0)

        self.body_font_slider = QSlider(Qt.Orientation.Horizontal)
        self.body_font_slider.setRange(8, 14)
        initial_body_size = self.temp_theme_config.get('body_font_size', 9)
        self.body_font_slider.setValue(initial_body_size)

        # Label to show the slider's current value
        self.body_font_value_label = QLabel(f"{initial_body_size} pt")
        self.body_font_value_label.setFixedWidth(40)

        # Connect slider movement to update the value label
        self.body_font_slider.valueChanged.connect(
            lambda value: self.body_font_value_label.setText(f"{value} pt")
        )

        font_layout.addWidget(self.body_font_slider, 1, 1)
        font_layout.addWidget(self.body_font_value_label, 1, 2)

        font_layout.setColumnStretch(1, 1) # Make the slider column stretch
        layout.addWidget(font_group)

        # --- Application Update Settings ---
        update_group = QGroupBox("Application Update")
        update_layout = QVBoxLayout(update_group)

        # URL Input
        url_layout = QHBoxLayout()
        self.update_url_edit = QLineEdit()
        # IMPORTANT: Replace this URL with the raw URL of your Python file on GitHub
        self.update_url_edit.setText("https://raw.githubusercontent.com/benberryhill/ScheduleApp-PyQt6/refs/heads/working_branch/dist/ScheduleApp/schedule_app_logic.py")
        self.update_url_edit.setPlaceholderText("Enter raw GitHub file URL...")
        url_layout.addWidget(QLabel("Update URL:"))
        url_layout.addWidget(self.update_url_edit)

        # Update Button
        self.update_button = QPushButton("Check for Updates")
        self.update_button.clicked.connect(self._check_for_updates)

        update_layout.addLayout(url_layout)
        update_layout.addWidget(self.update_button)
        layout.addWidget(update_group)

        layout.addStretch()
        return widget

    def _create_theme_tab(self, theme_mode):
        """Creates a tab for managing theme-specific colors ('light' or 'dark')."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        group = QGroupBox("Final Schedule Header Colors")
        grid_layout = QGridLayout(group)
        grid_layout.setColumnStretch(1, 1)

        header_colors = self.temp_theme_config[f'{theme_mode}_header_colors']

        # Define nice labels for the UI
        status_labels = {
            "empty": "Empty (No one scheduled)",
            "warning": "Warning (Has people, but below target)",
            "semi_full": "Full (Target met)",
            "full": "Over Capacity (More than target)",
        }

        for i, (status, color_hex) in enumerate(header_colors.items()):
            label = QLabel(status_labels.get(status, status.title()) + ":")

            # The callback updates the temporary config when a new color is picked
            callback = lambda color, s=status: self.temp_theme_config[f'{theme_mode}_header_colors'].update({s: color.name()})

            color_button = self._create_color_picker_button(color_hex, callback)

            grid_layout.addWidget(label, i, 0)
            grid_layout.addWidget(color_button, i, 1)

        # --- Highlight Color Group ---
        highlight_group = QGroupBox("Selection Highlight Color")
        highlight_layout = QHBoxLayout(highlight_group)
        highlight_layout.addWidget(QLabel("Highlight Color:"))

        # The callback now updates the nested dictionary for the specific theme
        callback = lambda color, mode=theme_mode: self.temp_theme_config['highlight_colors'].update({mode: color.name()})

        # Get the initial color from the correct theme
        initial_color = self.temp_theme_config['highlight_colors'][theme_mode]

        color_button = self._create_color_picker_button(initial_color, callback)
        highlight_layout.addWidget(color_button)
        highlight_layout.addStretch()

        layout.addWidget(group)
        layout.addWidget(highlight_group)
        layout.addStretch()
        return widget

    def _create_color_picker_button(self, initial_color, on_color_changed):
        """Helper to create a button that opens a color dialog."""
        button = QPushButton()
        button.setFixedSize(120, 25)

        def pick_color():
            current_color = QColor(button.property("color_hex"))
            new_color = QColorDialog.getColor(current_color, self, "Select Color")
            if new_color.isValid():
                self._update_button_color(button, new_color)
                on_color_changed(new_color)

        button.clicked.connect(pick_color)
        self._update_button_color(button, QColor(initial_color))
        return button

    def _update_button_color(self, button, color):
        """Updates a button's appearance to show the selected color."""
        button.setProperty("color_hex", color.name())
        # Set text color to black or white for contrast
        text_color = "white" if color.lightness() < 128 else "black"
        button.setStyleSheet(f"background-color: {color.name()}; color: {text_color};")
        button.setText(color.name())

    def _check_for_updates(self):
        """Downloads and validates a new script version from a URL."""
        url = self.update_url_edit.text().strip()
        if not url:
            QMessageBox.warning(self, "No URL", "Please provide a valid URL for the update.")
            return

        # Determine target file to overwrite
        if getattr(sys, 'frozen', False):
            # If frozen, we overwrite the external logic file next to the exe
            current_script_path = os.path.join(os.path.dirname(sys.executable), "schedule_app_logic.py")
        else:
            # If testing, we overwrite this file itself
            current_script_path = os.path.abspath(__file__)

        # Provide immediate feedback to the user
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Checking for Updates")
        msg_box.setText("Downloading and checking for a new version...")
        msg_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
        msg_box.show()
        # Force the UI to draw the message box immediately
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.processEvents()

        try:
            # Download the new script content from the URL
            with urllib.request.urlopen(url) as response:
                if response.getcode() != 200:
                    raise urllib.error.URLError(f"Server returned status code {response.getcode()}")
                new_content = response.read()

            # Read the content of the current script to compare
            with open(current_script_path, 'rb') as f:
                current_content = f.read()

            msg_box.hide()
            msg_box.close() # Done checking, close the message box
            if app_instance:
                app_instance.processEvents()

            if new_content == current_content:
                QMessageBox.information(self, "Up to Date", "You are already using the latest version of the application.")
            else:
                reply = QMessageBox.question(self, "Update Found",
                                             "A new version is available.\n\n"
                                             "The application will restart automatically to apply changes. Continue?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                             QMessageBox.StandardButton.No)

                if reply == QMessageBox.StandardButton.Yes:
                    try:
                        # Attempt to overwrite the old script file
                        with open(current_script_path, 'wb') as f:
                            f.write(new_content)

                        QMessageBox.information(self, "Update Complete", "Update successful. Restarting...")
                        # Launch a new instance of the executable
                        if getattr(sys, 'frozen', False):
                            subprocess.Popen([sys.executable])
                        else:
                            # If running from python script (dev mode)
                            subprocess.Popen([sys.executable, sys.argv[0]])

                        self.main_window.close() # Close this instance

                    except IOError as e:
                        QMessageBox.critical(self, "Update Error", f"Could not write to the application file. Please check permissions.\nError: {e}")

        except (urllib.error.URLError, urllib.error.HTTPError) as e:
            msg_box.close()
            QMessageBox.critical(self, "Update Failed", f"Could not download the update. Please check your internet connection and the URL.\nError: {e}")
        except Exception as e:
            msg_box.close()
            QMessageBox.critical(self, "An Error Occurred", f"An unexpected error occurred during the update process:\n{e}")

    def apply_changes(self):
        """Applies the settings to the main window without closing the dialog."""
        # Apply target counts
        for day, entry in self.max_entries_edits.items():
            if entry.text().isdigit():
                self.main_window.max_per_day[day] = int(entry.text())

        # Update temp_theme_config from all UI controls before applying
        self.temp_theme_config['header_font_size'] = self.header_font_slider.value()
        self.temp_theme_config['body_font_size'] = self.body_font_slider.value()

        # Apply theme config
        self.main_window.theme_config = self.temp_theme_config.copy()

        # Tell the main window to refresh its styles and UI
        self.main_window.apply_settings_and_refresh()

        # Save the new settings to the file
        self.main_window.save_settings()

    def accept(self):
        """Applies changes and closes the dialog (OK button)."""
        self.apply_changes()
        super().accept()

# --- Employee Editor Window ---
class EmployeeEditorWindow(QDialog):
    """A dialog for adding, editing, and deleting employees from the master list."""
    def __init__(self, main_window, selected_employee_name=None):
        super().__init__(main_window)
        self.main_window = main_window
        self.setWindowTitle("Manage Employees")
        self.setMinimumSize(550, 400)

        # --- UI Setup ---
        main_layout = QVBoxLayout(self)

        # Mode selection
        mode_layout = QHBoxLayout()
        self.employee_selector_combo = QComboBox()
        self.add_new_mode_check = QCheckBox("Add New Employee")
        mode_layout.addWidget(self.employee_selector_combo)
        mode_layout.addWidget(self.add_new_mode_check)
        main_layout.addLayout(mode_layout)

        # Editor fields group
        editor_group = QGroupBox("Employee Details")
        grid = QGridLayout(editor_group)

        # Row 0: Name
        grid.addWidget(QLabel("Name:"), 0, 0)
        self.edit_name = QLineEdit()
        grid.addWidget(self.edit_name, 0, 1)

        # Row 1: Availability
        grid.addWidget(QLabel("Availability:"), 1, 0, alignment=Qt.AlignmentFlag.AlignTop)
        avail_layout = QHBoxLayout()
        self.availability_boxes = {}
        for day in DAYS:
            chk = QCheckBox(day)
            self.availability_boxes[day] = chk
            avail_layout.addWidget(chk)
        grid.addLayout(avail_layout, 1, 1)

        # Row 2: Other settings
        grid.addWidget(QLabel("Settings:"), 2, 0)
        settings_layout = QHBoxLayout()
        self.set_schedule_check = QCheckBox("Set Schedule")
        self.set_schedule_check.setToolTip("If checked, this employee is automatically scheduled on all their available days.")
        self.max_days_edit = QLineEdit()
        self.max_days_edit.setPlaceholderText("e.g., 5")
        self.max_days_edit.setFixedWidth(60)
        settings_layout.addWidget(self.set_schedule_check)
        settings_layout.addWidget(QLabel("Max Days/Week:"))
        settings_layout.addWidget(self.max_days_edit)
        settings_layout.addStretch()
        grid.addLayout(settings_layout, 2, 1)

        main_layout.addWidget(editor_group)
        main_layout.addStretch()

        # Row 3: Time Off
        time_off_group = QGroupBox("Time Off Management")
        time_off_layout = QHBoxLayout(time_off_group)

        self.time_off_list = QListWidget()
        time_off_layout.addWidget(self.time_off_list)

        time_off_buttons_layout = QVBoxLayout()
        self.add_time_off_btn = QPushButton("Add")
        self.edit_time_off_btn = QPushButton("Edit")
        self.delete_time_off_btn = QPushButton("Delete")
        time_off_buttons_layout.addWidget(self.add_time_off_btn)
        time_off_buttons_layout.addWidget(self.edit_time_off_btn)
        time_off_buttons_layout.addWidget(self.delete_time_off_btn)
        time_off_buttons_layout.addStretch()

        time_off_layout.addLayout(time_off_buttons_layout)
        main_layout.addWidget(time_off_group)

        # Action Buttons Layout
        action_buttons_layout = QHBoxLayout()

        self.delete_button = QPushButton("Delete Employee")
        self.delete_button.setStyleSheet("""
            QPushButton {
                background-color: #cd5c5c;
                color: white;
                font-weight: bold;
                border: none;
                padding: 5px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #b02a2a;
            }
            QPushButton:pressed {
                background-color: #902a2a;
            }
        """)
        self.undo_button = QPushButton("Undo Last Delete")

        # Dialog Buttons (Apply/Close)
        self.button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Apply | QDialogButtonBox.StandardButton.Close
        )

        action_buttons_layout.addWidget(self.delete_button)
        action_buttons_layout.addWidget(self.undo_button)
        action_buttons_layout.addStretch()
        action_buttons_layout.addWidget(self.button_box)

        main_layout.addLayout(action_buttons_layout)

        # --- Connections ---
        self.add_new_mode_check.toggled.connect(self.toggle_mode)
        self.employee_selector_combo.currentIndexChanged.connect(self.populate_fields_from_selection)

        self.button_box.button(QDialogButtonBox.StandardButton.Apply).clicked.connect(self.save_changes)
        self.button_box.button(QDialogButtonBox.StandardButton.Close).clicked.connect(self.reject)

        self.add_time_off_btn.clicked.connect(self._add_time_off)
        self.edit_time_off_btn.clicked.connect(self._edit_time_off)
        self.delete_time_off_btn.clicked.connect(self._delete_time_off)

        self.delete_button.clicked.connect(self.delete_employee)
        self.undo_button.clicked.connect(self.undo_last_delete)

        # --- Initial State ---
        self._refresh_employee_list(select_name=selected_employee_name)
        self.toggle_mode(self.add_new_mode_check.isChecked())

    def toggle_mode(self, is_add_mode):
        """Switches the UI between 'Add New' and 'Edit Existing' modes."""
        self.employee_selector_combo.setDisabled(is_add_mode)
        self.edit_name.setReadOnly(not is_add_mode)
        self.delete_button.setDisabled(is_add_mode or self.employee_selector_combo.currentIndex() == 0)

        if is_add_mode:
            self.setWindowTitle("Add New Employee")
            self.clear_fields()
            self.edit_name.setFocus()
        else:
            self.setWindowTitle("Edit Employee")
            self.populate_fields_from_selection()

    def clear_fields(self):
        """Resets all input fields to their default state."""
        self.edit_name.clear()
        for chk in self.availability_boxes.values():
            chk.setChecked(False)
        self.set_schedule_check.setChecked(False)
        self.max_days_edit.setText("4") # Default to 4 days
        self.time_off_list.clear()

    def populate_fields_from_selection(self):
        """Loads the selected employee's data from the master file into the form."""
        if self.add_new_mode_check.isChecked():
            return

        selected_name = self.employee_selector_combo.currentText()
        self.delete_button.setDisabled(self.employee_selector_combo.currentIndex() == 0)

        if self.employee_selector_combo.currentIndex() == 0:
            self.clear_fields()
            return

        try:
            df = pd.read_excel(master_employee_file_path)
            emp_data = df[df['Name'] == selected_name].iloc[0]
            self.edit_name.setText(emp_data['Name'])

            for day, chk in self.availability_boxes.items():
                chk.setChecked(str(emp_data.get(day, 'No')).lower() == 'yes')

            self.set_schedule_check.setChecked(str(emp_data.get('Set Schedule', 'No')).lower() == 'yes')
            self.max_days_edit.setText(str(emp_data.get('Max Per Week', 7)))

            self.time_off_list.clear()
            time_off_str = emp_data.get('Time Off Dates', '')
            if pd.notna(time_off_str) and time_off_str:
                # Split by semicolon and remove any empty strings from extra separators
                entries = [entry.strip() for entry in time_off_str.split(';') if entry.strip()]
                self.time_off_list.addItems(entries)

        except (IndexError, FileNotFoundError) as e:
            QMessageBox.critical(self, "Error", f"Could not find or load data for '{selected_name}'.\nError: {e}")
            self.clear_fields()

    def save_changes(self):
        """Validates input, saves to Excel, and performs a non-destructive UI update."""
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Employee name cannot be empty.")
            return

        try:
            max_days = int(self.max_days_edit.text())
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Input Error", "Max Days/Week must be a number between 1 and 7.")
            return

        try:
            df = pd.read_excel(master_employee_file_path)
        except FileNotFoundError:
            df = pd.DataFrame(columns=['Name'] + DAYS + ['Set Schedule', 'Max Per Week', 'Time Off Dates'])

        original_name = self.employee_selector_combo.currentText()
        is_add_mode = self.add_new_mode_check.isChecked()

        if (is_add_mode or (not is_add_mode and name != original_name)) and name in df['Name'].values:
            QMessageBox.warning(self, "Input Error", f"An employee named '{name}' already exists.")
            return

        # --- START OF CHANGE ---
        # Get the state of time-off BEFORE changes are saved
        old_time_off_string = ""
        if not is_add_mode:
            emp_data_row = df[df['Name'] == original_name]
            if not emp_data_row.empty:
                # Ensure we handle potential missing column or NaN values gracefully
                old_time_off_val = emp_data_row.iloc[0].get('Time Off Dates')
                if pd.notna(old_time_off_val):
                    old_time_off_string = str(old_time_off_val)

        # Get the new time-off state from the UI
        new_time_off_items = [self.time_off_list.item(i).text() for i in range(self.time_off_list.count())]
        new_time_off_string = "; ".join(new_time_off_items)

        # Parse both old and new dates to find what was added and removed
        old_dates = self._parse_time_off_entries_to_dates(old_time_off_string.split(';'))
        new_dates = self._parse_time_off_entries_to_dates(new_time_off_items)

        added_dates = new_dates - old_dates
        removed_dates = old_dates - new_dates
        # --- END OF CHANGE ---

        row_data = {
            'Name': name,
            'Set Schedule': 'Yes' if self.set_schedule_check.isChecked() else 'No',
            'Max Per Week': max_days,
            'Time Off Dates': new_time_off_string,
            **{day: ('Yes' if chk.isChecked() else 'No') for day, chk in self.availability_boxes.items()}
        }

        if is_add_mode:
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)
        else:
            idx = df.index[df['Name'] == original_name].tolist()[0]
            for col, value in row_data.items():
                df.loc[idx, col] = value

        df.to_excel(master_employee_file_path, index=False)
        QMessageBox.information(self, "Success", f"Changes for '{name}' have been saved.")

        # Perform a non-destructive update
        main_app = self.main_window
        main_app.time_off_manager.load_time_off_data(master_employee_file_path)

        emp_obj = None
        if is_add_mode:
            emp_obj = Employee(name, {d: row_data[d] == 'Yes' for d in DAYS})
            main_app.employees.append(emp_obj)
            main_app.employees.sort(key=lambda e: e.name)
        else:
            emp_obj = next((e for e in main_app.employees if e.name == original_name), None)

        if emp_obj:
            emp_obj.name = name
            emp_obj.availability = {day: row_data[day] == 'Yes' for day in DAYS}

            # --- START OF CHANGE ---
            # Now, update the live schedule based on added/removed dates
            start_of_week = main_app.current_week_start_date
            end_of_week = start_of_week.addDays(6)

            # 1. Handle newly ADDED time-off (remove from schedule)
            for date in added_dates:
                if start_of_week <= date <= end_of_week:
                    day_index = start_of_week.daysTo(date)
                    day_key = DAYS[day_index]
                    main_app.schedule.remove_employee(emp_obj, day_key)

            # 2. Handle REMOVED time-off (add back to schedule if "Set Schedule")
            if self.set_schedule_check.isChecked():
                for date in removed_dates:
                    if start_of_week <= date <= end_of_week:
                        day_index = start_of_week.daysTo(date)
                        day_key = DAYS[day_index]
                        # Add back if they are available and not already there
                        if emp_obj.availability.get(day_key, False):
                            main_app.schedule.assign_employee(emp_obj, day_key)
            # --- END OF CHANGE ---

        main_app.update_all_views()
        self._refresh_employee_list(select_name=name)

    def delete_employee(self):
        """Moves the selected employee to the recently_deleted file."""
        name_to_delete = self.employee_selector_combo.currentText()
        if not name_to_delete or self.employee_selector_combo.currentIndex() == 0:
            return

        reply = QMessageBox.question(self, "Confirm Deletion",
                                     f"Are you sure you want to delete '{name_to_delete}'?\nThis can be undone with the 'Undo Last Delete' button.",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.No:
            return

        try:
            # 1. Read master list
            df_master = pd.read_excel(master_employee_file_path)
            row_to_delete = df_master[df_master['Name'] == name_to_delete]

            if row_to_delete.empty:
                QMessageBox.critical(self, "Error", f"Could not find '{name_to_delete}' in the master file.")
                return

            # 2. Add to recently_deleted file
            deleted_file_path = os.path.join(EXCEL_FOLDER, RECENTLY_DELETED_FILE)
            try:
                df_deleted = pd.read_excel(deleted_file_path)
            except FileNotFoundError:
                df_deleted = pd.DataFrame(columns=df_master.columns)

            df_deleted = pd.concat([df_deleted, row_to_delete], ignore_index=True)
            df_deleted.to_excel(deleted_file_path, index=False)

            # 3. Remove from master list
            df_master = df_master[df_master['Name'] != name_to_delete]
            df_master.to_excel(master_employee_file_path, index=False)

            QMessageBox.information(self, "Success", f"Employee '{name_to_delete}' has been deleted.")

            # --- Refresh everything but keep window open ---
            self.main_window.reload_from_master_file()
            self._refresh_employee_list() # Resets to "Select Employee..."

        except Exception as e:
            QMessageBox.critical(self, "Delete Error", f"Could not delete employee:\n{e}")

    def _refresh_employee_list(self, select_name=None):
        """Reloads the employee names from the main window into the combo box."""
        current_selection = select_name or self.employee_selector_combo.currentText()

        self.employee_selector_combo.blockSignals(True)
        self.employee_selector_combo.clear()

        all_names = sorted([e.name for e in self.main_window.employees])
        self.employee_selector_combo.addItems(["Select Employee..."] + all_names)

        if current_selection in all_names:
            self.employee_selector_combo.setCurrentText(current_selection)
        else:
            self.employee_selector_combo.setCurrentIndex(0)

        self.employee_selector_combo.blockSignals(False)
        self.populate_fields_from_selection()

    def undo_last_delete(self):
        """Restores the most recently deleted employee from the backup file."""
        deleted_file_path = os.path.join(EXCEL_FOLDER, RECENTLY_DELETED_FILE)
        if not os.path.exists(deleted_file_path):
            QMessageBox.information(self, "Undo", "No recently deleted employees to restore.")
            return

        try:
            df_deleted = pd.read_excel(deleted_file_path)
            if df_deleted.empty:
                QMessageBox.information(self, "Undo", "The deleted employees list is empty.")
                return

            last_deleted_row = df_deleted.iloc[-1:].copy()
            restored_name = last_deleted_row.iloc[0]['Name']

            df_master = pd.read_excel(master_employee_file_path)
            if restored_name in df_master['Name'].values:
                QMessageBox.critical(self, "Restore Error", f"Cannot restore '{restored_name}'. An employee with this name already exists.")
                return

            # Add the restored employee back to the master list
            df_master = pd.concat([df_master, last_deleted_row], ignore_index=True)
            df_master.to_excel(master_employee_file_path, index=False)

            # Remove the entry from the deleted list
            df_deleted = df_deleted.iloc[:-1]
            df_deleted.to_excel(deleted_file_path, index=False)

            QMessageBox.information(self, "Success", f"Restored '{restored_name}' to the master list.")
            self.main_window.reload_from_master_file()
            self._refresh_employee_list(select_name=restored_name)

        except Exception as e:
            QMessageBox.critical(self, "Restore Error", f"Failed to restore employee:\n{e}")

    def _parse_time_off_entries_to_dates(self, entries):
        """Parses a list of string entries into a set of unique QDate objects."""
        all_dates = set()
        for entry in entries:
            entry = entry.strip()
            if not entry:
                continue
            try:
                if '-' in entry:
                    start_str, end_str = entry.split('-')
                    start_date = QDate.fromString(start_str.strip(), "MM/dd/yyyy")
                    end_date = QDate.fromString(end_str.strip(), "MM/dd/yyyy")
                    if start_date.isValid() and end_date.isValid():
                        current_date = start_date
                        while current_date <= end_date:
                            all_dates.add(current_date)
                            current_date = current_date.addDays(1)
                else:
                    single_date = QDate.fromString(entry.strip(), "MM/dd/yyyy")
                    if single_date.isValid():
                        all_dates.add(single_date)
            except Exception as e:
                print(f"Could not parse date entry '{entry}' for comparison: {e}")
        return all_dates

    def _add_time_off(self):
        """Opens a dialog to add a new time off entry."""
        dialog = TimeOffDialog(self)
        if dialog.exec():
            date_string = dialog.get_date_string()
            if date_string:
                self.time_off_list.addItem(date_string)

    def _edit_time_off(self):
        """Opens a dialog to edit the selected time off entry."""
        selected_item = self.time_off_list.currentItem()
        if not selected_item:
            QMessageBox.information(self, "No Selection", "Please select a time off entry to edit.")
            return

        dialog = TimeOffDialog(self, date_string=selected_item.text())
        if dialog.exec():
            date_string = dialog.get_date_string()
            if date_string:
                selected_item.setText(date_string)

    def _delete_time_off(self):
        """Deletes the selected time off entry from the list."""
        selected_item = self.time_off_list.currentItem()
        if not selected_item:
            QMessageBox.information(self, "No Selection", "Please select a time off entry to delete.")
            return

        reply = QMessageBox.question(self, "Confirm Delete", "Are you sure you want to delete this time off entry?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.time_off_list.takeItem(self.time_off_list.row(selected_item))

# --- Collapsible Frame Widget ---
class CollapsibleFrame(QFrame):
    """A custom frame that can be collapsed or expanded, correctly resizing within a layout."""
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(0)
        main_layout.setAlignment(Qt.AlignmentFlag.AlignTop) # Important for collapsing

        self.toggle_button = QPushButton(f"▶ {title}")
        self.toggle_button.setStyleSheet("text-align: left; font-weight: bold; border: none; padding: 5px;")
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(True) # Start opened

        self.content_frame = QFrame()
        self.content_frame.setFrameShape(QFrame.Shape.NoFrame)
        self.content_frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.content_frame.setVisible(False)

        self.toggle_button.toggled.connect(self._on_toggle)

        main_layout.addWidget(self.toggle_button)
        main_layout.addWidget(self.content_frame)

        # --- KEY CHANGE 1: Set the initial size policy for the collapsed state ---
        # This tells the layout that this widget's preferred vertical size is based
        # on its contents, and it should not be stretched vertically.
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

    def _on_toggle(self, checked):
        """Handles the show/hide logic and dynamically changes the widget's size policy."""
        if checked:
            self.content_frame.setVisible(True)
            self.toggle_button.setText(f"▼ {self.toggle_button.text()[2:]}")
            # --- KEY CHANGE 2: When EXPANDED, tell the layout it can grow vertically.
            self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        else:
            self.content_frame.setVisible(False)
            self.toggle_button.setText(f"▶ {self.toggle_button.text()[2:]}")
            # --- KEY CHANGE 3: When COLLAPSED, tell the layout it should be its minimum possible size.
            self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        # --- KEY CHANGE 4: Tell the parent widget's layout to reactivate and recalculate. ---
        # This is more forceful and reliable than updateGeometry() for this specific case.
        if self.parentWidget() and self.parentWidget().layout():
            self.parentWidget().layout().activate()

    def setContentLayout(self, layout):
        """Sets the layout for the collapsible content area."""
        old_layout = self.content_frame.layout()
        if old_layout is not None:
            # Properly dispose of the old layout
            QWidget().setLayout(old_layout)
        self.content_frame.setLayout(layout)

# --- Custom Qt Widget for Click/Drag Events ---
class ClickableLabel(QLabel):
    """ A custom QLabel that can hold employee/day data and handle mouse clicks. """
    def __init__(self, main_window, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.main_window = main_window
        self.employee_obj = None
        self.day_key = None  # This must hold the day (e.g., 'Monday') for context menu actions
        self.is_draggable = False
        self.is_clickable_remove = False # Set to True for labels in the final schedule

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.MouseButton.LeftButton:
                # If this label is in the final schedule, trigger the highlight function.
                if self.is_clickable_remove and self.employee_obj:
                    self.main_window.on_schedule_label_highlight_click(self.employee_obj)
                    return # Exclusive action for this type of label

                # Handle left-click drag events for other labels
                if self.is_draggable and self.employee_obj:
                    self.main_window.on_drag_start(self)

            elif event.button() == Qt.MouseButton.RightButton:
                # Handle right-click events to show a context menu
                if self.employee_obj:
                    self.show_context_menu(event.globalPosition().toPoint())
            else:
                # For other mouse buttons
                super().mousePressEvent(event)
        except Exception as e:
            print(f"Error during mousePressEvent: {e}")

    def show_context_menu(self, global_position):
        """ Create and show a context menu with actions relevant to the label's location. """
        try:
            # A menu is only relevant if we have an employee and a day context.
            if not self.employee_obj or not self.day_key:
                return

            menu = QMenu(self)
            action_added = False

            # Context 1: Label is in the "Final Schedule" list.
            # `is_clickable_remove` is True only for these labels.
            if self.is_clickable_remove:
                remove_action = menu.addAction(f"Remove '{self.employee_obj.name}' from schedule")
                remove_action.triggered.connect(self.on_remove_from_schedule)
                action_added = True

            # Context 2: Label is in the "Unassigned Employees" list.
            # These labels are draggable, have a day_key, but are not `is_clickable_remove`.
            elif self.is_draggable and not self.is_clickable_remove:
                add_action = menu.addAction(f"Add '{self.employee_obj.name}' to schedule")
                add_action.triggered.connect(self.on_add_to_schedule)
                action_added = True

            # If no relevant action was found, do not show a menu.
            if not action_added:
                return

            # Apply dynamic style based on dark mode
            if self.main_window.is_dark_mode:
                menu.setStyleSheet("""
                    QMenu {
                        background-color: #2d2d2d;  /* Dark background */
                        color: #ffffff;             /* White text */
                        border: 1px solid #444444;  /* Border with dark gray */
                    }
                    QMenu::item {
                        padding: 5px 20px;
                        background-color: transparent;
                    }
                    QMenu::item:selected {
                        background-color: #444444; /* Lighter gray on hover */
                        color: #e0e0e0;             /* Slightly lighter text */
                    }
                """)
            else:
                menu.setStyleSheet("""
                    QMenu {
                        background-color: white;
                        color: black;
                        border: 1px solid gray;
                    }
                    QMenu::item {
                        padding: 5px 20px;
                        background-color: transparent;
                    }
                    QMenu::item:selected {
                        background-color: #0078d7; /* Blue highlight */
                        color: white;
                    }
                """)

            menu.exec(global_position)
        except Exception as e:
            print(f"Error during context menu creation: {e}")

    def on_add_to_schedule(self):
        """ Handle the 'Add to Final Schedule' context menu action. """
        try:
            if not self.employee_obj or not self.day_key:
                return

            # Call the scheduling logic
            success = self.main_window.schedule.assign_employee(self.employee_obj, self.day_key)
            if success:
                self.main_window.update_all_views()
            else:
                print(f"Failed to add {self.employee_obj.name} to the schedule on {self.day_key}.")
        except Exception as e:
            print(f"Error during 'Add to Final Schedule' action: {e}")

    def on_remove_from_schedule(self):
        """ Handle the 'Remove from Schedule' context menu action. """
        try:
            if not self.employee_obj or not self.day_key:
                return

            # Call the removal logic from the main schedule object
            success = self.main_window.schedule.remove_employee(self.employee_obj, self.day_key)
            if success:
                self.main_window.update_all_views()
                # The call to select_employee_by_object is now removed.
            else:
                # This case is unlikely if the UI is correct, but good to have
                print(f"UI state issue: Could not find {self.employee_obj.name} to remove from {self.day_key}")
        except Exception as e:
            print(f"Error during 'Remove from Schedule' action: {e}")

# --- Time Off Handling ---
class TimeOffDialog(QDialog):
    """A dialog for entering or editing a single/ranged time off period."""
    def __init__(self, parent=None, date_string=None):
        super().__init__(parent)
        self.setWindowTitle("Add/Edit Time Off")

        # UI Setup
        layout = QVBoxLayout(self)
        grid = QGridLayout()

        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("MM/dd/yyyy")
        self.start_date_edit.setDate(QDate.currentDate())

        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("MM/dd/yyyy")
        self.end_date_edit.setDate(QDate.currentDate())

        self.single_day_check = QCheckBox("Single Day Event")
        self.single_day_check.toggled.connect(self.end_date_edit.setDisabled)

        grid.addWidget(QLabel("Start Date:"), 0, 0)
        grid.addWidget(self.start_date_edit, 0, 1)
        grid.addWidget(QLabel("End Date:"), 1, 0)
        grid.addWidget(self.end_date_edit, 1, 1)
        grid.addWidget(self.single_day_check, 2, 1)

        layout.addLayout(grid)

        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # Populate fields if editing an existing entry
        if date_string:
            self.parse_date_string(date_string)

    def parse_date_string(self, date_string):
        """Parses a date string (e.g., 'MM/DD/YYYY' or 'MM/DD/YYYY-MM/DD/YYYY') to populate the fields."""
        try:
            if '-' in date_string:
                start_str, end_str = date_string.split('-')
                self.start_date_edit.setDate(QDate.fromString(start_str.strip(), "MM/dd/yyyy"))
                self.end_date_edit.setDate(QDate.fromString(end_str.strip(), "MM/dd/yyyy"))
                self.single_day_check.setChecked(False)
            else:
                self.start_date_edit.setDate(QDate.fromString(date_string.strip(), "MM/dd/yyyy"))
                self.end_date_edit.setDate(self.start_date_edit.date())
                self.single_day_check.setChecked(True)
        except Exception as e:
            print(f"Error parsing date string '{date_string}': {e}")

    def get_date_string(self):
        """Returns the formatted date string based on the dialog's inputs."""
        start_date = self.start_date_edit.date().toString("MM/dd/yyyy")
        if self.single_day_check.isChecked():
            return start_date
        else:
            end_date = self.end_date_edit.date().toString("MM/dd/yyyy")
            # Ensure start date is not after end date
            if self.start_date_edit.date() > self.end_date_edit.date():
                QMessageBox.warning(self, "Date Error", "The end date cannot be before the start date.")
                return None
            return f"{start_date}-{end_date}"

# --- Time Off Manager ---
class TimeOffManager:
    """Handles parsing, storing, and checking employee time off dates."""
    def __init__(self):
        self.time_off_data = {}

    def load_time_off_data(self, file_path):
        """Reads the master Excel file and parses all time off dates."""
        self.time_off_data.clear()
        try:
            df = pd.read_excel(file_path)
            if 'Time Off Dates' not in df.columns:
                return

            for _, row in df.iterrows():
                name = row['Name']
                date_str = row['Time Off Dates']
                if pd.notna(date_str) and name:
                    self.time_off_data[name] = self._parse_date_string_list(date_str)
        except Exception as e:
            print(f"Error loading time off data: {e}")

    def _parse_date_string_list(self, date_string_list):
        """Parses a semicolon-separated string of dates/ranges into QDate tuples."""
        ranges = []
        entries = [entry.strip() for entry in date_string_list.split(';') if entry.strip()]
        for entry in entries:
            try:
                if '-' in entry:
                    start_str, end_str = entry.split('-')
                    start_date = QDate.fromString(start_str.strip(), "MM/dd/yyyy")
                    end_date = QDate.fromString(end_str.strip(), "MM/dd/yyyy")
                    if start_date.isValid() and end_date.isValid():
                        ranges.append((start_date, end_date))
                else:
                    single_date = QDate.fromString(entry.strip(), "MM/dd/yyyy")
                    if single_date.isValid():
                        ranges.append((single_date, single_date))
            except Exception as e:
                print(f"Could not parse date entry '{entry}': {e}")
        return ranges

    def is_employee_off(self, employee_name, check_date):
        """Checks if an employee has a specific date off."""
        if employee_name not in self.time_off_data or not isinstance(check_date, QDate):
            return False

        for start_date, end_date in self.time_off_data[employee_name]:
            if start_date <= check_date <= end_date:
                return True
        return False

# --- Main Application ---
class SchedulerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Semi-Auto Scheduler")
        self.setGeometry(100, 100, 1800, 950)

        self.employees = []
        self.schedule = Schedule()
        self.time_off_manager = TimeOffManager()
        self.current_week_start_date = QDate()
        self.max_per_day = {day: 35 for day in DAYS}

        self.selected_employee_obj = None
        self.selected_row_widgets = None
        self.drag_data = {'label_widget': None, 'employee': None}
        self.highlighted_employee_obj = None # To track the employee for grid highlighting

        self.grid_header_widgets = []
        self.day_column_header_widgets = []
        self.table_column_header_widgets = []
        self.employee_progress_styles = {}
        self.no_set_days_rows = []

        # For mapping widgets to employee objects, similar to the original's approach
        self.widget_to_employee = {}

        self.theme_config = {
            'highlight_colors': {
                'light': QColor(0, 120, 215).name(),  # Default system blue for light mode
                'dark': QColor(0, 90, 158).name()     # A darker blue for dark mode
            },
            'header_font_size': 10,
            'body_font_size': 9,
            'light_header_colors': {
                "empty": "#f0f0f0",
                "warning": "indianred",
                "semi_full": "royalblue",
                "full": "lightseagreen",
            },
            'dark_header_colors': {
                "empty": "#3c3c3c",
                "warning": "darkred",
                "semi_full": "darkblue",
                "full": "teal",
            }
        }
        # Now, load settings from file, which will overwrite the defaults
        self.load_settings()

        # Apply the loaded palette settings immediately so they are used for initial style generation.
        # This ensures user-defined highlight colors are active on startup.
        palette = self.palette()
        # On startup, we are always in light mode.
        initial_theme_mode = 'light'
        highlight_color_hex = self.theme_config['highlight_colors'][initial_theme_mode]
        highlight_color = QColor(highlight_color_hex)
        palette.setColor(QPalette.ColorRole.Highlight, highlight_color)
        
        # Set the text color for the highlight based on the highlight color's brightness
        text_color = Qt.GlobalColor.white if highlight_color.lightness() < 128 else Qt.GlobalColor.black
        palette.setColor(QPalette.ColorRole.HighlightedText, text_color)
        
        # Apply the new palette to the entire application
        self.setPalette(palette)

        # This dictionary will hold the generated stylesheets
        self.header_styles = {}

        # Theming
        self.is_dark_mode = False
        self._setup_styles()
        app_instance = QApplication.instance()
        if app_instance:
            app_instance.setStyleSheet(self.light_stylesheet) # Apply the light theme on startup
        self._generate_header_styles() # Generate styles from the new config
        self._generate_progressbar_styles()
        self._generate_employee_progressbar_styles()
        self._update_highlight_style() # Initialize the highlight style based on the current palette
        # Define styles for the schedule day headers
        self.header_base_style_template = "QLabel {{ font-weight: bold; padding: 4px; border-radius: 4px; color: {text_color}; background-color: {bg_color}; }}"
        
        self.dark_header_styles = {
            "empty": self.header_base_style_template.format(bg_color="#3c3c3c", text_color="white"), # Darker gray for empty
            "warning": self.header_base_style_template.format(bg_color="darkred", text_color="white"), # Darker red for warning
            "semi_full": self.header_base_style_template.format(bg_color="#0068F0", text_color="white"), # Darker blue for semi-full
            "full": self.header_base_style_template.format(bg_color="00C3C3", text_color="white"), # Teal for over capacity
        }
        self.light_header_styles = {
            "empty": self.header_base_style_template.format(bg_color="#f0f0f0", text_color="black"), # Lighter gray for empty
            "warning": self.header_base_style_template.format(bg_color="indianred", text_color="black"),
            "semi_full": self.header_base_style_template.format(bg_color="royalblue", text_color="white"), # White text for Royal Blue
            "full": self.header_base_style_template.format(bg_color="lightseagreen", text_color="black"),
        }

        # Load initial data to determine UI size
        initial_employee_df, initial_employee_count = self._load_and_count_master_file_data()
        self.max_display_rows_per_list = max(30, initial_employee_count)
        print(self.max_display_rows_per_list)

        # Create UI
        self._setup_ui()

        # Initialize the week selector
        self._on_week_changed(QDate.currentDate())

        # Apply initial font styles after UI is built
        self._apply_font_styles()

        # Load data into UI
        self.load_master_employees(initial_employee_df)

        self.check_and_load_draft()

    def _setup_styles(self):
        # self.light_stylesheet = ""
        self.light_stylesheet = """
            QWidget {
                background-color: transparent;
                color: #000000;
                border: none;
            }
            QMainWindow, QDialog {
                background-color: #f0f0f0;
            }
            QFrame, QScrollArea, QListWidget {
                background-color: #ffffff;
            }
            QFrame {
                border-radius: 5px;
            }
            QLabel {
                background-color: transparent;
            }
            QPushButton {
                background-color: #ffffff;
                color: #000000;
                padding: 5px;
                border-radius: 5px;
                border: 1px solid #adadad;
            }
            QPushButton:hover {
                background-color: #e5f1fb;
                border-color: #0078d7;
            }
            QPushButton:pressed {
                background-color: #cce4f7;
            }
            QLineEdit, QComboBox, QDateEdit {
                padding: 5px;
                border: 1px solid #abadb3;
                border-radius: 5px;
                background-color: #ffffff;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left: 1px solid #dcdcdc;
            }
            QComboBox::drop-down {
                border: none;
            }
            QCheckBox::indicator {
                width: 13px;
                height: 13px;
            }
            QGroupBox {
                border: 1px solid #d1d1d1;
                border-radius: 5px;
                margin-top: 15px; /* Create space for the title */
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #000000;
                background-color: #f0f0f0; /* Match window background */
            }
            QTabBar::tab {
                background-color: #e1e1e1;
                padding: 8px 15px;
                border: 1px solid #d1d1d1;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:hover {
                background-color: #e5f1fb;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
            }
            QTabWidget::pane {
                border: 1px solid #d1d1d1;
                border-top: none;
            }
            QCalendarWidget QWidget {
                alternate-background-color: #e1e1e1;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: #000000;
                selection-background-color: #0078d7;
                selection-color: white;
            }
            #qt_calendar_navigationbar {
                background-color: #e1e1e1;
                color: black;
            }
            #qt_calendar_prevmonth, #qt_calendar_nextmonth {
                color: black;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 10px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #c1c1c1;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a8a8a8;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }
            QScrollBar:horizontal {
                border: none;
                background: #f0f0f0;
                height: 10px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:horizontal {
                background: #c1c1c1;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #a8a8a8;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                height: 0px;
            }
            QCheckBox::indicator {
                width: 13px;
                height: 13px;
            }
        """
        self.dark_stylesheet = """
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                border: none;
            }
            QMainWindow, QDialog {
                background-color: #2b2b2b;
            }
            QFrame, QScrollArea, QListWidget {
                background-color: #3c3c3c;
                border-radius: 3px;
            }
            QLabel {
                background-color: transparent;
            }
            QPushButton {
                background-color: #555555;
                color: #ffffff;
                padding: 5px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #6a6a6a;
            }
            QPushButton:pressed {
                background-color: #4a4a4a;
            }
            QLineEdit, QComboBox, QDateEdit {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 5px;
                background-color: #3c3c3c;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left: 1px solid #555555;
            }
            QComboBox::drop-down {
                border: none;
            }
            QCheckBox::indicator {
                width: 13px;
                height: 13px;
            }
            QGroupBox {
                border: 1px solid #555555;
                border-radius: 5px;
                margin-top: 15px; /* Create space for the title */
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #dddddd;
                background-color: #2b2b2b; /* Match window background */
            }
            QTabBar::tab {
                background-color: #3c3c3c;
                padding: 8px 15px;
                border: 1px solid #555555;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:hover {
                background-color: #4c4c4c;
            }
            QTabBar::tab:selected {
                background-color: #2b2b2b;
            }
            QTabWidget::pane {
                border: 1px solid #555555;
                border-top: none;
            }
            QCalendarWidget QWidget {
                alternate-background-color: #4a4a4a;
            }
            QCalendarWidget QAbstractItemView:enabled {
                color: #e0e0e0;
                selection-background-color: #0078d7;
                selection-color: white;
            }
            #qt_calendar_navigationbar {
                background-color: #555555;
                color: white;
            }
            #qt_calendar_prevmonth, #qt_calendar_nextmonth {
                color: white;
            }
            QScrollBar:vertical {
                border: none;
                background: #3c3c3c;
                width: 8px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #555555;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background: #6a6a6a;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }
            QScrollBar:horizontal {
                border: none;
                background: #3c3c3c;
                height: 10px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:horizontal {
                background: #555555;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #6a6a6a;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                border: none;
                background: none;
                height: 0px;
            }
            QCheckBox::indicator {
                width: 13px;
                height: 13px;
            }
        """

    def _apply_font_styles(self):
        """Applies font sizes from theme_config to relevant widgets."""
        try:
            header_size = self.theme_config.get('header_font_size', 10)
            body_size = self.theme_config.get('body_font_size', 9)

            # --- Header Font ---
            header_font = self.font() # Get base font
            header_font.setPointSize(header_size)
            header_font.setBold(True)

            # Apply header font to Top Bar buttons
            top_bar_buttons = [
                self.settings_button, self.manage_employees_button,
                self.auto_schedule_button, self.week_selector_button, self.export_button,
                self.toggle_theme_button
            ]
            for button in top_bar_buttons:
                button.setFont(header_font)

            # Collapsible frame titles
            self.unassigned_frame.toggle_button.setFont(header_font)
            self.all_employees_frame.toggle_button.setFont(header_font)

            # Final schedule day name labels (e.g., "Sun", "Mon")
            for day_widgets in self.final_schedule_day_headers.values():
                day_widgets['label'].setFont(header_font)

            # All other grid headers ("Name", "Availability", etc.)
            for widget in self.grid_header_widgets:
                widget.setFont(header_font)

            # --- Body Font ---
            body_font = self.font() # Get base font
            body_font.setPointSize(body_size)
            # Make the day headers bold but using the body font size
            body_font_bold = self.font()
            body_font_bold.setPointSize(body_size)
            body_font_bold.setBold(True)

            # All table-like grid headers ("Name", "Availability", etc.)
            for widget in self.table_column_header_widgets:
                widget.setFont(body_font_bold)

            # Final schedule day name labels (e.g., "Sun", "Mon")
            for day_widgets in self.final_schedule_day_headers.values():
                day_widgets['label'].setFont(body_font_bold) # Changed from header_font

            # Unassigned grid day name labels
            for widget in self.day_column_header_widgets:
                widget.setFont(body_font_bold)

            # Final schedule employee name labels
            for day_list in self.schedule_labels.values():
                for label in day_list:
                    label.setFont(body_font)

            # All Employees grid body labels
            for row_data in self.all_employees_rows:
                for widget in row_data['conceptual_row_widgets']:
                    widget.setFont(body_font)

            # Unassigned Employees grid body labels
            for day_list in self.unassigned_labels.values():
                for label in day_list:
                    label.setFont(body_font)

            # "No Set Days" grid body labels
            for row_data in self.no_set_days_rows:
                row_data['name_lbl'].setFont(body_font)

        except Exception as e:
            print(f"Error applying font styles: {e}")

    def _generate_header_styles(self):
        """
        Generates header stylesheets from the theme_config dictionary.
        This makes the theme dynamically configurable.
        """
        base_template = "QLabel {{ font-weight: bold; padding: 4px; border-radius: 4px; color: {text_color}; background-color: {bg_color}; }}"

        # --- Generate Light Theme Styles ---
        light_styles = {}
        for status, bg_hex in self.theme_config['light_header_colors'].items():
            bg_color = QColor(bg_hex)
            text_color = "white" if bg_color.lightness() < 128 else "black"
            light_styles[status] = base_template.format(bg_color=bg_hex, text_color=text_color)

        # --- Generate Dark Theme Styles ---
        dark_styles = {}
        for status, bg_hex in self.theme_config['dark_header_colors'].items():
            bg_color = QColor(bg_hex)
            text_color = "white" if bg_color.lightness() < 128 else "black"
            dark_styles[status] = base_template.format(bg_color=bg_hex, text_color=text_color)

        self.header_styles = {'light': light_styles, 'dark': dark_styles}

    def _generate_progressbar_styles(self):
        """Generates QProgressBar stylesheets from the theme_config."""
        self.progressbar_styles = {'light': {}, 'dark': {}}
        base_template = """
            QProgressBar {{
                border: 1px solid {border_color};
                border-radius: 5px;
                text-align: center;
                background-color: {bar_bg};
                color: {text_color};
            }}
            QProgressBar::chunk {{
                background-color: {chunk_color};
                border-radius: 4px;
            }}
        """

        # --- Generate Light Theme Styles ---
        for status, bg_hex in self.theme_config['light_header_colors'].items():
            chunk_color = QColor(bg_hex)
            text_color = "white" if chunk_color.lightness() < 128 else "black"
            style = base_template.format(
                border_color='grey',
                bar_bg='#e0e0e0',
                text_color=text_color,
                chunk_color=bg_hex
            )
            self.progressbar_styles['light'][status] = style

        # --- Generate Dark Theme Styles ---
        for status, bg_hex in self.theme_config['dark_header_colors'].items():
            chunk_color = QColor(bg_hex)
            text_color = "white" if chunk_color.lightness() < 128 else "black"
            style = base_template.format(
                border_color='#555555',
                bar_bg='#3c3c3c',
                text_color=text_color,
                chunk_color=bg_hex
            )
            self.progressbar_styles['dark'][status] = style

    def apply_settings_and_refresh(self):
        """
        Applies all settings from the config and refreshes the entire UI.
        This is the central point for applying changes from the SettingsWindow.
        """
        # 1. Update max day values in the main UI (if they exist, for compatibility)
        # The primary source of truth is now self.max_per_day
        if hasattr(self, 'max_entries'):
            for day, value in self.max_per_day.items():
                if day in self.max_entries:
                    self.max_entries[day].setText(str(value))

        # 2. Apply palette for highlight color
        palette = self.palette()
        theme_mode = 'dark' if self.is_dark_mode else 'light'
        highlight_color_hex = self.theme_config['highlight_colors'][theme_mode]
        highlight_color = QColor(highlight_color_hex)
        palette.setColor(QPalette.ColorRole.Highlight, highlight_color)
        text_color = Qt.GlobalColor.white if highlight_color.lightness() < 128 else Qt.GlobalColor.black
        palette.setColor(QPalette.ColorRole.HighlightedText, text_color)
        self.setPalette(palette)

        # 3. Regenerate all dynamic styles
        self._generate_header_styles()
        self._generate_progressbar_styles()
        self._update_highlight_style()

        # 4. Refresh all views to show changes
        self.update_all_views()

        # 5. Apply the new font styles
        self._apply_font_styles()

    def _open_settings_window(self):
        """Opens the main settings dialog."""
        # The dialog is modal, so we don't need to check if it's already visible.
        # A new instance is created each time.
        dialog = SettingsWindow(self)
        dialog.exec() # exec() shows the dialog and blocks until closed

    def _load_and_count_master_file_data(self):
        try:
            if not os.path.exists(master_employee_file_path): return None, 0
            df = pd.read_excel(master_employee_file_path)
            valid_names = df['Name'].dropna().astype(str).str.strip().unique()
            return df, len(valid_names)
        except Exception as e:
            print(f"Error reading master file for count and data: {e}")
            return None, 0

    def _setup_ui(self):
        # Central Widget and Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top Bar
        top_bar = self._create_top_bar()
        main_layout.addLayout(top_bar)

        # Middle Content Area (Split View)
        middle_layout = QHBoxLayout()
        main_layout.addLayout(middle_layout)

        left_column = self._create_left_column()
        right_column = self._create_right_column()

        middle_layout.addLayout(left_column, 1) # Weight 1
        middle_layout.addLayout(right_column, 1) # Weight 1

    def _create_top_bar(self):
        layout = QHBoxLayout()

        self.settings_button = QPushButton("⚙️ Settings")
        self.settings_button.clicked.connect(self._open_settings_window)
        font = self.settings_button.font()
        font.setPointSize(11)
        self.settings_button.setFont(font)

        self.manage_employees_button = QPushButton("🧑‍💼 Manage Employees")
        self.manage_employees_button.clicked.connect(self._open_employee_editor)

        self.auto_schedule_button = QPushButton("Auto Schedule Available Employees")
        self.auto_schedule_button.clicked.connect(self.auto_schedule_available_employees)

        self.week_selector_button = QPushButton("Select a week...")
        self.week_selector_button.clicked.connect(self._open_week_selector)

        self.export_button = QPushButton("Export Schedule")
        self.export_button.clicked.connect(self.export_schedule)

        self.toggle_theme_button = QCheckBox("Dark Mode")
        self.toggle_theme_button.toggled.connect(self.toggle_theme)

        layout.addWidget(self.settings_button)
        layout.addWidget(self.manage_employees_button)
        layout.addWidget(self.auto_schedule_button)
        layout.addWidget(self.week_selector_button)
        layout.addStretch(1)
        layout.addWidget(self.toggle_theme_button)
        layout.addWidget(self.export_button)
        return layout

    def _create_left_column(self):
        layout = QVBoxLayout()

        # Final Schedule Preview with Scroll
        self.final_schedule_scroll_area = QScrollArea()
        self.final_schedule_scroll_area.setWidgetResizable(True)
        self.final_schedule_frame_container = QFrame()
        self.final_schedule_frame_container.setFrameShape(QFrame.Shape.StyledPanel)

        # Add the final schedule frame container to the scroll area
        self.final_schedule_scroll_area.setWidget(self.final_schedule_frame_container)
        layout.addWidget(self.final_schedule_scroll_area, 1)  # Stretchable

        # Build the schedule preview inside the container
        self._build_schedule_preview(self.final_schedule_frame_container)
        return layout

    def _create_right_column(self):
        layout = QVBoxLayout()

        # Unassigned Employees (Now a CollapsibleFrame)
        self.unassigned_frame = CollapsibleFrame("Unassigned Employees (Available Days)")
        unassigned_content_layout = QVBoxLayout()
        self.unassigned_scroll_area = QScrollArea()
        self.unassigned_scroll_area.setWidgetResizable(True)
        unassigned_content_layout.addWidget(self.unassigned_scroll_area)
        self.unassigned_frame.setContentLayout(unassigned_content_layout)
        self._build_unassigned_grid()

        # All Employees (Unchanged)
        self.all_employees_frame = CollapsibleFrame("All Employees (Master List)")
        all_employees_content_layout = QVBoxLayout()
        self.all_employees_scroll_area = QScrollArea()
        self.all_employees_scroll_area.setWidgetResizable(True)
        all_employees_content_layout.addWidget(self.all_employees_scroll_area)
        self.all_employees_frame.setContentLayout(all_employees_content_layout)
        self._build_all_employees_grid()

        # Add frames to the layout with stretch factors to share space
        layout.addWidget(self.unassigned_frame, 1)
        layout.addWidget(self.all_employees_frame, 1)
        return layout

    def _build_schedule_preview(self, parent_container):
        self.schedule_layout = QGridLayout(parent_container)
        self.schedule_labels = {day: [] for day in DAYS}
        self.final_schedule_day_headers = {}
        self.schedule_col_layouts = {}

        for i, day in enumerate(DAYS):
            # Create a container widget for each header cell
            header_container = QWidget()
            header_layout = QVBoxLayout(header_container)
            header_layout.setContentsMargins(2, 2, 2, 2)
            header_layout.setSpacing(3)

            # 1. The Day Label
            day_label = QLabel(f"<b>{day}</b>")
            day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

            # 2. The Progress Bar
            progress_bar = QProgressBar()
            progress_bar.setFixedHeight(20)
            progress_bar.setTextVisible(True)
            progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)

            header_layout.addWidget(day_label)
            header_layout.addWidget(progress_bar)

            # Store references to both widgets for easy updates
            self.final_schedule_day_headers[day] = {
                'label': day_label,
                'progress': progress_bar
            }
            self.schedule_layout.addWidget(header_container, 0, i)

            # Use a widget with a layout for the column of labels
            col_widget = QWidget()
            col_layout = QVBoxLayout(col_widget)
            col_layout.setContentsMargins(0, 0, 0, 0)
            col_layout.setSpacing(1)

            # Store the layout
            self.schedule_col_layouts[day] = col_layout

            # Create initial batch of rows
            for r in range(INITIAL_DISPLAY_ROWS_SCHEDULE):
                lbl = ClickableLabel(self, "")
                lbl.setFixedHeight(20)
                lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                lbl.is_clickable_remove = True
                lbl.day_key = day  # Assign the day for this label
                lbl.employee_obj = None  # Will be updated when employees are added
                col_layout.addWidget(lbl)
                self.schedule_labels[day].append(lbl)

            col_layout.addStretch(1)
            self.schedule_layout.addWidget(col_widget, 1, i)
            self.schedule_layout.setColumnStretch(i, 1)

    def _build_all_employees_grid(self):
        container = QWidget()
        self.all_employees_scroll_area.setWidget(container)
        layout = QGridLayout(container)

        # Headers
        headers_text = ["<b>Name</b>", "<b>Availability</b>", "<b>Time Off Dates</b>", "<b>Shifts This Week</b>"]
        for i, text in enumerate(headers_text):
            header_label = QLabel(text)
            layout.addWidget(header_label, 0, i)
            self.table_column_header_widgets.append(header_label)
        layout.setColumnStretch(0, 2)
        layout.setColumnStretch(1, 2)
        layout.setColumnStretch(2, 2)
        layout.setColumnStretch(3, 3)

        self.all_employees_rows = []
        for r in range(self.max_display_rows_per_list):
            row_idx = r + 1
            name_lbl = ClickableLabel(self, "") # Still a ClickableLabel for drag-and-drop
            name_lbl.is_draggable = True
            avail_lbl = QLabel("")
            time_off_lbl = QLabel("")
            progress_bar = QProgressBar()

            row_widgets = [name_lbl, avail_lbl, time_off_lbl, progress_bar]
            alt_row_widgets = [name_lbl, avail_lbl, time_off_lbl]
            for i, widget in enumerate(row_widgets):
                widget.setFixedHeight(22)
                layout.addWidget(widget, row_idx, i)

            self.all_employees_rows.append({
                'name_lbl': name_lbl,
                'avail_lbl': avail_lbl,
                'time_off_lbl': time_off_lbl,
                'progress_bar': progress_bar,
                'conceptual_row_widgets': row_widgets,
                'alt_row_widgets': alt_row_widgets,
                'employee': None
            })
        layout.setRowStretch(self.max_display_rows_per_list + 1, 1)

    def _build_unassigned_grid(self):
        container = QWidget()
        self.unassigned_scroll_area.setWidget(container)
        self.unassigned_grid_layout = QGridLayout(container)

        self.unassigned_labels = {day: [] for day in DAYS}
        for c, day in enumerate(DAYS):
            header_label = QLabel(f"<b>{day}</b>")
            self.unassigned_grid_layout.addWidget(header_label, 0, c, alignment=Qt.AlignmentFlag.AlignCenter)
            self.day_column_header_widgets.append(header_label)

            for r in range(INITIAL_ROWS_UNASSIGNED):
                lbl = ClickableLabel(self, "")
                lbl.setFixedHeight(20)
                lbl.is_draggable = True
                self.unassigned_grid_layout.addWidget(lbl, r + 1, c)
                self.unassigned_labels[day].append(lbl)
            self.unassigned_grid_layout.setColumnStretch(c, 1)

        # Set an initial stretch
        self.unassigned_grid_layout.setRowStretch(INITIAL_ROWS_UNASSIGNED + 1, 1)

    def _open_employee_editor(self):
        """Opens the employee editor dialog."""
        # Pass the name of the currently highlighted employee, if any
        selected_name = self.highlighted_employee_obj.name if self.highlighted_employee_obj else None
        dialog = EmployeeEditorWindow(self, selected_name)
        dialog.exec()

    def _open_week_selector(self):
        """Opens a popup calendar to select a week."""
        dialog = WeekSelectorDialog(self, current_date=self.current_week_start_date)

        # Position the popup below the button
        button_pos = self.week_selector_button.mapToGlobal(self.week_selector_button.rect().bottomLeft())
        dialog.move(button_pos)

        if dialog.exec():
            if dialog.selected_date:
                self._on_week_changed(dialog.selected_date)

    def _generate_employee_progressbar_styles(self):
        """Generates QProgressBar styles for employee weekly shift counts."""
        self.employee_progress_styles = {'light': {}, 'dark': {}}

        # We'll map our new statuses to the existing theme colors for consistency
        status_color_map = {
            "empty": "empty",          # 0 shifts
            "low": "warning",          # 1-2 shifts
            "approaching_max": "semi_full", # max - 1 shifts
            "at_max": "full"           # max shifts
        }

        for theme in ['light', 'dark']:
            for status, color_key in status_color_map.items():
                bg_hex = self.theme_config[f'{theme}_header_colors'][color_key]
                chunk_color = QColor(bg_hex)
                text_color = "white" if chunk_color.lightness() < 128 else "black"

                # Use a simplified template since we don't need a background bar
                style = f"""
                    QProgressBar {{
                        border: 1px solid grey;
                        border-radius: 5px;
                        text-align: center;
                        background-color: transparent;
                        color: {"white" if theme == 'dark' else "black"};
                    }}
                    QProgressBar::chunk {{
                        background-color: {bg_hex};
                        border-radius: 4px;
                        margin: 1px;
                    }}
                """
                self.employee_progress_styles[theme][status] = style

    def reload_from_master_file(self):
        """Reloads all employee data and time off data, then refreshes the entire UI."""
        self.schedule.clear_schedule()
        self.time_off_manager.load_time_off_data(master_employee_file_path) # <-- ADD THIS
        df, count = self._load_and_count_master_file_data()
        self.load_master_employees(df)
        print("UI reloaded from master file.")

    def _on_week_changed(self, new_date):
        """Snaps the selected date to the start of the week (Sunday) and refreshes views."""
        # Day of the week: Monday=1, Sunday=7. We want to find the previous Sunday.
        day_of_week = new_date.dayOfWeek()

        # If it's already Sunday, we don't need to change it. Otherwise, subtract days.
        days_to_subtract = 0 if day_of_week == 7 else day_of_week

        start_of_week = new_date.addDays(-days_to_subtract)

        # Update the state and UI only if the week has actually changed
        if self.current_week_start_date != start_of_week:
            self.current_week_start_date = start_of_week

            # Update the button text to show the full week range
            end_of_week = self.current_week_start_date.addDays(6)
            date_format = "M/d/yyyy"
            self.week_selector_button.setText(
                f"{self.current_week_start_date.toString(date_format)} - {end_of_week.toString(date_format)}"
            )

            print(f"Week changed. New start date: {self.current_week_start_date.toString('yyyy-MM-dd')}")
            self.reload_from_master_file() # Reload and refresh everything

# --- Core Logic Methods ---

    def load_master_employees(self, df_master):
        self.employees.clear()
        self.schedule.clear_schedule()
        if df_master is None:
            QMessageBox.critical(self, "Error", "Master employee data could not be loaded.")
            return

        try:
            loaded_names = set()
            for _, row in df_master.iterrows():
                name = str(row.get('Name', '')).strip()
                if not name or name in loaded_names: continue
                loaded_names.add(name)
                
                availability = {day: str(row.get(day, '')).strip().lower() == 'yes' for day in DAYS}
                emp = Employee(name, availability)
                self.employees.append(emp)

                # Check if the employee is a "set schedule" employee
                set_schedule_status = str(row.get('Set Schedule', '')).strip().lower()
                if set_schedule_status == 'yes':
                    # Automatically assign the employee to the schedule for available days
                    for day in DAYS:
                        day_index = DAYS.index(day)
                        actual_date = self.current_week_start_date.addDays(day_index)
                        is_off = self.time_off_manager.is_employee_off(emp.name, actual_date)
                        if emp.availability.get(day, False) and not is_off: # Ensure employee is available on this day
                            self.schedule.assign_employee(emp, day)

            self.employees.sort(key=lambda e: e.name)
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Error processing master employee data:\n{e}")
            self.employees.clear()
        
        self.update_all_views()
        if not self.employees: self.clear_editor_fields()

    def update_all_views(self):
        self.update_all_employees_grid()
        self.update_unassigned_grid()
        self.update_final_schedule_display()

    def get_alternating_row_style(self, index):
        if not self.is_dark_mode:
            colors = ("#ffffff", "#f0f0f0")
        else:
            colors = ("#3c3c3c", "#313131")
        return f"background-color: {colors[index % 2]};"

    def update_all_employees_grid(self):
        try:
            # Read the file once to get all data
            df_master = pd.read_excel(master_employee_file_path)
            # Create a lookup map for the time off dates
            time_off_map = dict(zip(df_master['Name'], df_master.get('Time Off Dates', '')))
        except Exception as e:
            print(f"Could not read master employee file for time off:\n{e}")
            time_off_map = {}

        # Create a lookup map for the progress bar
        df_master = pd.read_excel(master_employee_file_path)
        max_per_week_map = dict(zip(df_master['Name'], df_master.get('Max Per Week', [7]*len(df_master))))

        current_schedule_count = {emp.name: 0 for emp in self.employees}
        for day in DAYS:
            for emp in self.schedule.scheduled.get(day, []):
                current_schedule_count[emp.name] += 1

        theme_mode = 'dark' if self.is_dark_mode else 'light'
        style_map = self.employee_progress_styles[theme_mode]

        for i, row_data in enumerate(self.all_employees_rows):
            if i < len(self.employees):
                emp = self.employees[i]
                current_count = current_schedule_count.get(emp.name, 0)
                max_val = int(max_per_week_map.get(emp.name, 7))

                row_data['employee'] = emp
                row_data['name_lbl'].setText(emp.name)
                row_data['name_lbl'].employee_obj = emp # Attach obj for drag/click
                row_data['avail_lbl'].setText(', '.join([d for d, v in emp.availability.items() if v]))

                raw_time_off_str = time_off_map.get(emp.name, '')
                formatted_time_off = self._format_time_off_for_display(raw_time_off_str)
                row_data['time_off_lbl'].setText(formatted_time_off)

                pb = row_data['progress_bar']
                pb.setRange(0, max_val)
                pb.setValue(current_count)
                pb.setFormat(f"{current_count} / {max_val}")

                status = "empty"
                if current_count >= max_val: status = "at_max"
                elif current_count == max_val - 1: status = "approaching_max"
                elif 0 < current_count: status = "low"

                pb.setStyleSheet(style_map.get(status, ""))
                pb.show()

                for lbl in row_data['alt_row_widgets']:
                    lbl.setStyleSheet(self.get_alternating_row_style(i))
                    lbl.show()
            else: # Hide unused rows
                row_data['employee'] = None
                for lbl in row_data['conceptual_row_widgets']:
                    lbl.setText("")
                    lbl.hide()

    def update_unassigned_grid(self):
        """
        Updates the grid displaying employees available but not yet scheduled for each day,
        and who are still under their max per week limit. Dynamically expands rows if needed.
        """
        try:
            df = pd.read_excel(master_employee_file_path)
            # Create a map for Max Per Week (default to 7 if missing)
            max_per_week_map = dict(zip(df['Name'], df.get('Max Per Week', [7]*len(df))))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read master employee file:\n{e}")
            return

        # 1. Count how many times each employee is already scheduled this week
        current_schedule_count = {emp.name: 0 for emp in self.employees}
        for day in DAYS:
            for emp in self.schedule.scheduled.get(day, []):
                current_schedule_count[emp.name] += 1

        # 2. Calculate the lists of available employees for each day first
        day_availability_lists = {}
        for day_index, day in enumerate(DAYS):
            actual_date = self.current_week_start_date.addDays(day_index)
            available_list = []

            for emp in self.employees:
                is_off = self.time_off_manager.is_employee_off(emp.name, actual_date)

                # Logic: Available on day AND Not off AND Not scheduled today AND Under weekly max
                if (
                        emp.availability.get(day, False) and
                        not is_off and
                        all(e.name != emp.name for e in self.schedule.scheduled[day]) and
                        current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                ):
                    available_list.append(emp)

            # Sort by name for consistent display
            available_list.sort(key=lambda e: e.name)
            day_availability_lists[day] = available_list

        # 3. Dynamic Expansion: Check if we have enough widgets, create more if needed
        for c, day in enumerate(DAYS):
            needed = len(day_availability_lists[day])
            current_widgets = self.unassigned_labels[day]

            while len(current_widgets) < needed:
                # Row index is length of current widgets + 1 (because row 0 is headers)
                row_idx = len(current_widgets) + 1

                lbl = ClickableLabel(self, "")
                lbl.setFixedHeight(20)
                lbl.is_draggable = True

                # Apply current body font settings to new labels
                body_font = self.font()
                body_font.setPointSize(self.theme_config.get('body_font_size', 9))
                lbl.setFont(body_font)

                # Add to the layout (ensure self.unassigned_grid_layout was created in _build)
                self.unassigned_grid_layout.addWidget(lbl, row_idx, c)

                # Add to our local list tracking
                current_widgets.append(lbl)

        # 4. Reset visibility of all existing labels
        # This clears old data and handles cases where the list shrinks
        for day in DAYS:
            for lbl in self.unassigned_labels[day]:
                lbl.setText("")
                lbl.employee_obj = None
                lbl.day_key = day
                lbl.hide()

        # 5. Populate the grid with the calculated data
        for day in DAYS:
            available_list = day_availability_lists[day]
            for r_idx, emp in enumerate(available_list):
                # We guaranteed enough widgets in Step 3, so this is safe
                lbl = self.unassigned_labels[day][r_idx]

                lbl.setText(emp.name)
                lbl.employee_obj = emp

                # Handle Highlighting
                if self.highlighted_employee_obj and self.highlighted_employee_obj.name == emp.name:
                    lbl.setStyleSheet(self.highlight_style)
                else:
                    lbl.setStyleSheet(self.get_alternating_row_style(r_idx))

                lbl.show()

    def update_final_schedule_display(self):
        current_max = self.max_per_day

        for day, labels in self.schedule_labels.items():
            emps_on_day = self.schedule.scheduled[day]
            count = len(emps_on_day)
            max_val = current_max.get(day, 35)

            # If we have more employees than labels, create new labels on the fly
            while len(labels) < count:
                new_lbl = ClickableLabel(self, "")
                new_lbl.setFixedHeight(20)
                new_lbl.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                new_lbl.is_clickable_remove = True
                new_lbl.day_key = day
                new_lbl.employee_obj = None

                # Apply current font settings to new label
                body_font = self.font()
                body_font.setPointSize(self.theme_config.get('body_font_size', 9))
                new_lbl.setFont(body_font)

                # Insert before the stretch item (which is usually the last item)
                layout = self.schedule_col_layouts[day]
                cnt = layout.count()
                # Insert at count-1 to stay above the stretch
                layout.insertWidget(cnt - 1, new_lbl)

                # Add to our local list
                labels.append(new_lbl)

            # Get the header widgets (label and progress bar)
            header_widgets = self.final_schedule_day_headers[day]
            progress_bar = header_widgets['progress']

            # Configure the progress bar's range, value, and text format
            progress_bar.setRange(0, max_val)
            progress_bar.setValue(count)
            progress_bar.setFormat(f"{count} / {max_val}") # Custom text format
            if count > max_val:
                progress_bar.setRange(0, count)
                progress_bar.setValue(count)
            else:
                # Otherwise, use the standard range and value.
                progress_bar.setRange(0, max_val)
                progress_bar.setValue(count)

            # Determine the visual status
            theme_mode = 'dark' if self.is_dark_mode else 'light'
            status = "empty"
            if count > 0:
                status = "warning"
            if count >= max_val:
                status = "semi_full"
            if count > max_val:
                status = "full"

            # Apply the corresponding style to the progress bar
            style = self.progressbar_styles[theme_mode].get(status, "")
            progress_bar.setStyleSheet(style)

            # The original day label no longer needs its style updated, just its background
            header_widgets['label'].setStyleSheet("background-color: transparent;")

            # Update the text labels for the day's schedule entries (these are the inner labels, not the headers)
            for i, lbl in enumerate(labels):
                if i < len(emps_on_day):
                    emp = emps_on_day[i]
                    lbl.setText(emp.name)
                    lbl.employee_obj = emp
                    lbl.day_key = day
                    if self.highlighted_employee_obj and self.highlighted_employee_obj.name == emp.name:
                        lbl.setStyleSheet(self.highlight_style)
                    else:
                        lbl.setStyleSheet(self.get_alternating_row_style(i))
                    lbl.show()
                else:
                    lbl.setText("")
                    lbl.employee_obj = None
                    lbl.day_key = None
                    lbl.hide()

    def auto_schedule_available_employees(self):
        """
        Automatically schedules employees day-by-day. It prioritizes filling harder days first.
        For weekdays, it prioritizes senior employees. For weekends, it prioritizes newer employees
        and tries to avoid scheduling them on both Saturday and Sunday if possible.
        """
        try:
            df = pd.read_excel(master_employee_file_path)
            employee_map = {e.name: e for e in self.employees}
            max_per_week_map = dict(zip(df['Name'], df.get('Max Per Week', [7]*len(df))))
            seniority_ordered_employees = [employee_map[name] for name in df['Name'].dropna() if name in employee_map]
            reverse_seniority_employees = list(reversed(seniority_ordered_employees))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read master employee file for auto-scheduling:\n{e}")
            return

        max_per_day = self.max_per_day
        day_availability_counts = {day: sum(1 for e in seniority_ordered_employees if e.availability.get(day, False)) for day in DAYS}
        sorted_days = sorted(DAYS, key=lambda d: day_availability_counts[d])

        current_schedule_count = {emp.name: 0 for emp in self.employees}
        for day in DAYS:
            for emp in self.schedule.scheduled.get(day, []):
                current_schedule_count[emp.name] += 1

        employees_added_count = 0

        for day in sorted_days:
            day_index = DAYS.index(day)
            actual_date = self.current_week_start_date.addDays(day_index)
            needed = max_per_day.get(day, 0) - len(self.schedule.scheduled.get(day, []))
            if needed <= 0:
                continue

            # --- NEW: Logic splits for Weekends vs Weekdays ---
            if day in ["Sat", "Sun"]:
                # --- WEEKEND LOGIC (Two-Pass System) ---
                employee_pool = reverse_seniority_employees
                other_weekend_day = "Sun" if day == "Sat" else "Sat"

                # --- Pass 1: Prioritize employees NOT working the other weekend day ---
                eligible_pass_1 = [
                    emp for emp in employee_pool
                    if emp.availability.get(day, False)
                       and not self.time_off_manager.is_employee_off(emp.name, actual_date)
                       and current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                       and emp.name not in [e.name for e in self.schedule.scheduled[day]]
                       # This is the key new condition for Pass 1
                       and emp.name not in [e.name for e in self.schedule.scheduled[other_weekend_day]]
                ]
                eligible_pass_1.sort(key=lambda e: current_schedule_count[e.name])

                # Assign from the ideal pool first
                for emp in eligible_pass_1:
                    if needed <= 0: break
                    self.schedule.assign_employee(emp, day)
                    current_schedule_count[emp.name] += 1
                    employees_added_count += 1
                    needed -= 1

                # If we've filled the schedule, move to the next day
                if needed <= 0:
                    continue

                # --- Pass 2: If still needed, fill with remaining staff (who might work both days) ---
                eligible_pass_2 = [
                    emp for emp in employee_pool
                    if emp.availability.get(day, False)
                       and not self.time_off_manager.is_employee_off(emp.name, actual_date)
                       and current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                       and emp.name not in [e.name for e in self.schedule.scheduled[day]]
                ]
                eligible_pass_2.sort(key=lambda e: current_schedule_count[e.name])

                # Assign from the less-ideal pool to fill the gaps
                for emp in eligible_pass_2:
                    if needed <= 0: break
                    self.schedule.assign_employee(emp, day)
                    current_schedule_count[emp.name] += 1
                    employees_added_count += 1
                    needed -= 1

            else:
                # --- WEEKDAY LOGIC (Original logic) ---
                employee_pool = seniority_ordered_employees
                eligible_employees = [
                    emp for emp in employee_pool
                    if emp.availability.get(day, False)
                       and not self.time_off_manager.is_employee_off(emp.name, actual_date)
                       and current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                       and emp.name not in [e.name for e in self.schedule.scheduled.get(day, [])]
                ]
                eligible_employees.sort(key=lambda e: current_schedule_count[e.name])

                for emp in eligible_employees:
                    if needed <= 0: break
                    self.schedule.assign_employee(emp, day)
                    current_schedule_count[emp.name] += 1
                    employees_added_count += 1
                    needed -= 1

        # --- Step 5: Done ---
        self.update_all_views()
        QMessageBox.information(self, "Auto-Schedule Complete",
                                f"Finished auto-scheduling. {employees_added_count} new assignments were made.")

    def _format_time_off_for_display(self, time_off_str):
        """Converts a raw time off string into a more readable, shorthand format."""
        if not time_off_str or pd.isna(time_off_str):
            return ""

        current_year = QDate.currentDate().year()

        def format_date(q_date):
            """Formats a single QDate object."""
            if q_date.year() == current_year:
                return q_date.toString("M/d")
            else:
                return q_date.toString("M/d/yy")

        entries = [entry.strip() for entry in time_off_str.split(';') if entry.strip()]
        formatted_entries = []

        for entry in entries:
            try:
                if '-' in entry:
                    start_str, end_str = entry.split('-')
                    start_date = QDate.fromString(start_str.strip(), "MM/dd/yyyy")
                    end_date = QDate.fromString(end_str.strip(), "MM/dd/yyyy")
                    if start_date.isValid() and end_date.isValid():
                        formatted_entries.append(f"{format_date(start_date)}-{format_date(end_date)}")
                else:
                    single_date = QDate.fromString(entry.strip(), "MM/dd/yyyy")
                    if single_date.isValid():
                        formatted_entries.append(format_date(single_date))
            except Exception:
                # If parsing fails, just use the original entry
                formatted_entries.append(entry)

        return "; ".join(formatted_entries)

    def on_schedule_label_highlight_click(self, employee_obj):
        """ Handles left-clicks on the final schedule to highlight the employee everywhere. """
        if self.highlighted_employee_obj == employee_obj:
            # If clicking the same employee, toggle the highlight off.
            self.highlighted_employee_obj = None
        else:
            # Otherwise, set the new employee to be highlighted.
            self.highlighted_employee_obj = employee_obj

        # Redraw the two grids that can show the highlight.
        self.update_final_schedule_display()
        self.update_unassigned_grid()

    def export_schedule(self):
        if not any(self.schedule.scheduled.values()):
            QMessageBox.information(self, "Export", "Schedule is empty.")
            return
        
        try:
            filename = f"Generated_Schedule_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            path = os.path.join(EXCEL_FOLDER, filename)
            
            max_len = max(len(v) for v in self.schedule.scheduled.values())
            data = {day: [e.name for e in emps] + [''] * (max_len - len(emps)) 
                    for day, emps in self.schedule.scheduled.items()}
            
            df = pd.DataFrame(data)
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Export Successful", f"Schedule saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export schedule:\n{e}")

    def load_settings(self):
        """Loads settings from the JSON file on startup."""
        settings_path = os.path.join(EXCEL_FOLDER, SETTINGS_FILE)
        try:
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    loaded_settings = json.load(f)

                # Get the theme config part of the loaded settings
            if 'theme_config' in loaded_settings:
                loaded_theme_conf = loaded_settings['theme_config']

                # Check for the old 'highlight_bg' key and migrate it
                if 'highlight_bg' in loaded_theme_conf and 'highlight_colors' not in loaded_theme_conf:
                    print("Migrating old 'highlight_bg' setting to new format.")
                    old_color = loaded_theme_conf.pop('highlight_bg')
                    loaded_theme_conf['highlight_colors'] = {
                        'light': old_color,
                        'dark': self.theme_config['highlight_colors']['dark'] # Use default for dark
                    }

                # Update the application's config with the loaded one.
                # This is a simple "deep update" for one level of nesting.
                for key, value in loaded_theme_conf.items():
                    if isinstance(value, dict) and key in self.theme_config:
                        self.theme_config[key].update(value)
                    else:
                        self.theme_config[key] = value

            if 'max_per_day' in loaded_settings:
                self.max_per_day.update(loaded_settings.get('max_per_day', {}))

            print("Settings loaded successfully.")
        except (json.JSONDecodeError, TypeError) as e:
            print(f"Error loading settings file, using defaults. Error: {e}")
        except Exception as e:
            print(f"An unexpected error occurred while loading settings: {e}")

    def save_settings(self):
        """Saves current settings to the JSON file."""
        settings_path = os.path.join(EXCEL_FOLDER, SETTINGS_FILE)
        try:
            settings_to_save = {
                'theme_config': self.theme_config,
                'max_per_day': self.max_per_day
            }
            with open(settings_path, 'w') as f:
                json.dump(settings_to_save, f, indent=4)
            print("Settings saved successfully.")
        except Exception as e:
            print(f"Error saving settings: {e}")

    def save_draft(self):
        """Saves the current schedule to a JSON file."""
        try:
            # Convert objects to a dictionary of lists of names
            # Structure: {'Mon': ['Alice', 'Bob'], 'Tue': ['Charlie']}
            draft_data = {
                day: [e.name for e in self.schedule.scheduled[day]]
                for day in DAYS
            }

            draft_path = os.path.join(EXCEL_FOLDER, DRAFT_FILE)
            with open(draft_path, 'w') as f:
                json.dump(draft_data, f, indent=4)
            print(f"Draft saved to {draft_path}")
        except Exception as e:
            print(f"Error saving draft: {e}")

    def check_and_load_draft(self):
        """Checks if a draft exists and asks user to load it."""
        draft_path = os.path.join(EXCEL_FOLDER, DRAFT_FILE)

        if not os.path.exists(draft_path):
            return

        reply = QMessageBox.question(
            self,
            "Load Draft?",
            "A saved draft from a previous session was found.\n\nDo you want to load it?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                with open(draft_path, 'r') as f:
                    draft_data = json.load(f)

                # Create a lookup map: Name -> Employee Object
                # This ensures we are using the live employee objects with current availability info
                employee_map = {e.name: e for e in self.employees}

                self.schedule.clear_schedule()

                # Reconstruct schedule
                for day, names in draft_data.items():
                    if day in DAYS:
                        for name in names:
                            if name in employee_map:
                                self.schedule.assign_employee(employee_map[name], day)

                self.update_all_views()
                QMessageBox.information(self, "Draft Loaded", "Previous schedule has been restored.")

            except Exception as e:
                QMessageBox.warning(self, "Load Error", f"Could not load draft file:\n{e}")

# --- Drag/Drop and Click Handling ---
    def on_drag_start(self, label_widget):
        if label_widget and label_widget.employee_obj:
            self.drag_data = {'label_widget': label_widget, 'employee': label_widget.employee_obj}
            # Visual feedback for drag start
            label_widget.setStyleSheet(f"background-color: {self.palette().color(QPalette.ColorRole.Highlight).name()};")

    def mouseReleaseEvent(self, event):
        # Immediately restore the original label's style if a drag was in progress.
        # This removes the highlight instantly, before any other logic runs.
        if self.drag_data.get('label_widget'):
            original_widget = self.drag_data['label_widget']
            original_style = self.drag_data.get('original_style', '')
            if original_widget:
                original_widget.setStyleSheet(original_style)

        if not self.drag_data.get('employee'):
            return

        employee_to_drop = self.drag_data['employee']
        target_day = None

        # Check if dropped over the final schedule area
        fs_widget = self.final_schedule_frame_container
        if fs_widget.rect().contains(fs_widget.mapFromGlobal(event.globalPosition().toPoint())):
            pos_in_widget = fs_widget.mapFromGlobal(event.globalPosition().toPoint())

            # Find column (day)
            col_width = fs_widget.width() / len(DAYS)
            col_index = int(pos_in_widget.x() // col_width)
            if 0 <= col_index < len(DAYS):
                target_day = DAYS[col_index]

        if target_day:
            day_index = DAYS.index(target_day)
            actual_date = self.current_week_start_date.addDays(day_index)
            if self.time_off_manager.is_employee_off(employee_to_drop.name, actual_date):
                reply = QMessageBox.question(self, "Time Off Warning",
                                             f"'{employee_to_drop.name}' has requested this day off.\n\nSchedule them anyway?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                             QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No:
                    # Cleanup and refresh without assigning
                    self.drag_data = {'label_widget': None, 'employee': None}
                    self.update_all_views()
                    return

            if not employee_to_drop.availability.get(target_day, False):
                reply = QMessageBox.question(self, "Availability Warning",
                                             f"'{employee_to_drop.name}' is not normally available on {target_day}.\nSchedule anyway?",
                                             QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if reply == QMessageBox.StandardButton.No:
                    target_day = None # Cancel the drop

            if target_day and not self.schedule.assign_employee(employee_to_drop, target_day):
                QMessageBox.information(self, "Already Scheduled", f"'{employee_to_drop.name}' is already scheduled on {target_day}.")

        # Cleanup and refresh
        self.drag_data = {'label_widget': None, 'employee': None}
        self.update_all_views()
        # Highlight the employee that was just dragged
        self.on_schedule_label_highlight_click(employee_to_drop)

    def _update_highlight_style(self):
        """
        Generates the stylesheet for highlighting based on the app's current palette.
        This ensures the highlight color matches the native selection color (e.g., blue).
        """
        highlight_bg = self.palette().color(QPalette.ColorRole.Highlight).name()
        highlight_text = self.palette().color(QPalette.ColorRole.HighlightedText).name()
        self.highlight_style = f"background-color: {highlight_bg}; color: {highlight_text}; border-radius: 3px;"

    def toggle_theme(self, checked):
        self.is_dark_mode = checked
        app_instance = QApplication.instance()
        if app_instance:
            if checked:
                app_instance.setStyleSheet(self.dark_stylesheet)
            else:
                app_instance.setStyleSheet(self.light_stylesheet)

        # Re-apply the palette to get the correct highlight color for the new theme
        palette = self.palette()
        theme_mode = 'dark' if self.is_dark_mode else 'light'
        highlight_color_hex = self.theme_config['highlight_colors'][theme_mode]
        highlight_color = QColor(highlight_color_hex)
        palette.setColor(QPalette.ColorRole.Highlight, highlight_color)
        text_color = Qt.GlobalColor.white if highlight_color.lightness() < 128 else Qt.GlobalColor.black
        palette.setColor(QPalette.ColorRole.HighlightedText, text_color)
        self.setPalette(palette)

        self._update_highlight_style()
        self.update_all_views()

    def closeEvent(self, event):
        """Called automatically when the main window is closed."""

        # Check if schedule has any data
        has_data = any(len(self.schedule.scheduled[d]) > 0 for d in DAYS)

        if has_data:
            reply = QMessageBox.question(
                self,
                "Save Changes?",
                "Do you want to save the current schedule as a draft before exiting?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Yes
            )

            if reply == QMessageBox.StandardButton.Cancel:
                event.ignore() # Stop the window from closing
                return

            if reply == QMessageBox.StandardButton.Yes:
                self.save_draft()

        # Save settings (Theme, etc.) - Existing Logic
        self.save_settings()
        event.accept() # Close the window

def main():
    # Only create QApplication if one doesn't exist (safety check)
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    # app.setStyle("Breeze")
    window = SchedulerApp()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

# package into exe with all needed files and imports
# pyinstaller --noconfirm -D --windowed --icon="app_icon.ico" --add-data="excel_files;excel_files" schedule_app_pyqt6.py