import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QCheckBox, QScrollArea, QFrame,
    QMessageBox, QSizePolicy, QMenu, QDialog, QColorDialog, QTabWidget, QGroupBox,
    QDialogButtonBox, QProgressBar, QListWidget, QDateEdit
)
from PyQt6.QtCore import Qt, pyqtSignal, QEvent, QSize, QDate
from PyQt6.QtGui import QPalette, QColor, QIcon

# --- Configuration ---
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
EXCEL_FOLDER_NAME = "excel_files"
EXCEL_FOLDER = resource_path(EXCEL_FOLDER_NAME)
RECENTLY_DELETED_FILE = "recently_deleted.xlsx" # For undo functionality
EMPLOYEE_FILE = "Employees_Full_List.xlsx" # Master list
DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
MAX_DISPLAY_ROWS_SCHEDULE = 60 # Max rows per day in Final Schedule
MAX_ROWS_UNASSIGNED_PER_DAY_COLUMN = 60 # Max rows per day in the Unassigned Employees grid

# --- Dummy File Creation (Same as original) ---
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
                'Notes': ['Team Lead', '', 'Part-time', 'New Hire', '', 'Floater'],
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
    def __init__(self, name, availability, notes):
        self.name = name
        self.availability = availability
        self.notes = notes
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

        # --- Highlight Color Settings ---
        highlight_group = QGroupBox("Selection Highlight Color")
        highlight_layout = QHBoxLayout(highlight_group)
        highlight_layout.addWidget(QLabel("Highlight Color:"))

        self.highlight_color_button = self._create_color_picker_button(
            self.temp_theme_config['highlight_bg'],
            lambda color: self.temp_theme_config.update({'highlight_bg': color.name()})
        )
        highlight_layout.addWidget(self.highlight_color_button)
        highlight_layout.addStretch()
        layout.addWidget(highlight_group)

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

        layout.addWidget(group)
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

    def apply_changes(self):
        """Applies the settings to the main window without closing the dialog."""
        # Apply target counts
        for day, entry in self.max_entries_edits.items():
            if entry.text().isdigit():
                self.main_window.max_per_day[day] = int(entry.text())

        # Apply theme config
        self.main_window.theme_config = self.temp_theme_config.copy()

        # Tell the main window to refresh its styles and UI
        self.main_window.apply_settings_and_refresh()

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

        # Row 1: Notes
        grid.addWidget(QLabel("Notes:"), 1, 0)
        self.edit_notes = QLineEdit()
        grid.addWidget(self.edit_notes, 1, 1)

        # Row 2: Availability
        grid.addWidget(QLabel("Availability:"), 2, 0, alignment=Qt.AlignmentFlag.AlignTop)
        avail_layout = QHBoxLayout()
        self.availability_boxes = {}
        for day in DAYS:
            chk = QCheckBox(day)
            self.availability_boxes[day] = chk
            avail_layout.addWidget(chk)
        grid.addLayout(avail_layout, 2, 1)

        # Row 3: Other settings
        grid.addWidget(QLabel("Settings:"), 3, 0)
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
        grid.addLayout(settings_layout, 3, 1)

        main_layout.addWidget(editor_group)
        main_layout.addStretch()

        # Row 4: Time Off
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
        self.edit_notes.clear()
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

            notes_value = emp_data.get('Notes', '')
            display_notes = str(notes_value) if pd.notna(notes_value) else ''
            self.edit_notes.setText(display_notes)

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

    def save_changes(self): # Renamed from save_and_close
        """Validates input and saves changes (add or update) to the master Excel file."""
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Employee name cannot be empty.")
            return

        try:
            max_days = int(self.max_days_edit.text())
            if not (1 <= max_days <= 7): raise ValueError
        except (ValueError, TypeError):
            QMessageBox.warning(self, "Input Error", "Max Days/Week must be a number between 1 and 7.")
            return

        # Define all columns the app manages and set their expected type to string
        all_managed_cols = ['Name', 'Notes', 'Set Schedule', 'Max Per Week', 'Time Off Dates'] + DAYS
        dtype_map = {col: str for col in all_managed_cols}

        try:
            # Read the Excel file, forcing all our columns to be treated as strings
            df = pd.read_excel(master_employee_file_path, dtype=dtype_map)
        except (FileNotFoundError, ValueError):
            # If a file doesn't exist or a column is missing, create a new blank DataFrame
            df = pd.DataFrame(columns=all_managed_cols)
            # Ensure the new blank DataFrame also has the correct types
            df = df.astype(dtype_map)

        # As a safeguard, ensure all managed columns actually exist in the DataFrame
        for col in all_managed_cols:
            if col not in df.columns:
                df[col] = ''

        # Final cast to ensure any newly added columns have the correct string type
        df = df.astype(dtype_map)

        if 'Time Off Dates' not in df.columns:
            df['Time Off Dates'] = ''
        # Explicitly cast the column to 'object' to prevent dtype warnings.
        # .astype(str) would work but 'object' is more idiomatic for mixed/string data.
        df['Time Off Dates'] = df['Time Off Dates'].astype('object')

        original_name = self.employee_selector_combo.currentText()
        is_add_mode = self.add_new_mode_check.isChecked()

        # Check for name collision on add OR rename
        if (is_add_mode or (not is_add_mode and name != original_name)) and name in df['Name'].values:
            QMessageBox.warning(self, "Input Error", f"An employee named '{name}' already exists.")
            return

        # Get time off data from the list widget
        time_off_items = [self.time_off_list.item(i).text() for i in range(self.time_off_list.count())]
        time_off_string = "; ".join(time_off_items)

        # Prepare data row
        row_data = {
            'Name': name,
            'Notes': self.edit_notes.text().strip(),
            'Set Schedule': 'Yes' if self.set_schedule_check.isChecked() else 'No',
            'Max Per Week': max_days,
            'Time Off Dates': time_off_string,
            **{day: ('Yes' if chk.isChecked() else 'No') for day, chk in self.availability_boxes.items()}
        }

        if is_add_mode:
            new_row_df = pd.DataFrame([row_data])
            df = pd.concat([df, new_row_df], ignore_index=True)
        else:
            idx = df.index[df['Name'] == original_name].tolist()
            if not idx:
                QMessageBox.critical(self, "Save Error", f"Could not find '{original_name}' to update.")
                return

            # Use .loc to update all columns from the dictionary
            for col, value in row_data.items():
                df.loc[idx[0], col] = value

        df.to_excel(master_employee_file_path, index=False)
        QMessageBox.information(self, "Success", f"Changes for '{name}' have been saved.")

        # --- Refresh everything but keep window open ---
        self.main_window.reload_from_master_file()
        self._refresh_employee_list(select_name=name)
        # Note: self.accept() is REMOVED

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

# --- Main Application ---
class SchedulerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Semi-Auto Scheduler")
        self.setGeometry(100, 100, 1800, 950)

        self.employees = []
        self.schedule = Schedule()
        self.max_per_day = {day: 35 for day in DAYS}

        # self.max_entries = {}
        # for i, day in enumerate(DAYS):
        #     entry = QLineEdit(str(self.max_per_day[day]))
        #     self.max_entries[day] = entry

        self.selected_employee_obj = None
        self.selected_row_widgets = None
        self.drag_data = {'label_widget': None, 'employee': None}
        self.highlighted_employee_obj = None # To track the employee for grid highlighting

        # For mapping widgets to employee objects, similar to the original's approach
        self.widget_to_employee = {}

        self.theme_config = {
            'highlight_bg': QColor(0, 120, 215).name(), # Default system blue
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
        # This dictionary will hold the generated stylesheets
        self.header_styles = {}

        # Theming
        self.is_dark_mode = False
        self._setup_styles()
        self._generate_header_styles() # Generate styles from the new config
        self._generate_progressbar_styles()
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

        # Load data into UI
        self.load_master_employees(initial_employee_df)

    def _setup_styles(self):
        self.light_stylesheet = "" # Default Qt style
        self.dark_stylesheet = """
            QWidget {
                background-color: #2b2b2b;
                color: #ffffff;
                border: none;
            }
            QMainWindow {
                background-color: #2b2b2b;
            }
            QFrame {
                background-color: #3c3c3c;
                border-radius: 5px;
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
            QLineEdit, QComboBox {
                padding: 5px;
                border: 1px solid #555555;
                border-radius: 5px;
                background-color: #3c3c3c;
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
            QScrollArea {
                background-color: #3c3c3c;
                border-radius: 5px;
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
        highlight_color = QColor(self.theme_config['highlight_bg'])
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

        self.export_button = QPushButton("Export Schedule")
        self.export_button.clicked.connect(self.export_schedule)

        self.toggle_theme_button = QCheckBox("Dark Mode")
        self.toggle_theme_button.toggled.connect(self.toggle_theme)

        layout.addWidget(self.settings_button)
        layout.addWidget(self.manage_employees_button)
        layout.addWidget(self.auto_schedule_button)
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

        # Unassigned Employees
        unassigned_frame = QFrame()
        unassigned_frame.setFrameShape(QFrame.Shape.StyledPanel)
        unassigned_layout = QVBoxLayout(unassigned_frame)
        unassigned_layout.addWidget(QLabel("<h3>Unassigned Employees (Available Days)</h3>"))
        self.unassigned_scroll_area = QScrollArea()
        self.unassigned_scroll_area.setWidgetResizable(True)
        unassigned_layout.addWidget(self.unassigned_scroll_area)
        self._build_unassigned_grid()
        
        # All Employees
        all_employees_frame = QFrame()
        all_employees_frame.setFrameShape(QFrame.Shape.StyledPanel)
        all_employees_layout = QVBoxLayout(all_employees_frame)
        all_employees_layout.addWidget(QLabel("<h3>All Employees (Master List)</h3>"))
        self.all_employees_scroll_area = QScrollArea()
        self.all_employees_scroll_area.setWidgetResizable(True)
        all_employees_layout.addWidget(self.all_employees_scroll_area)
        self._build_all_employees_grid()

        layout.addWidget(unassigned_frame, 1) # Stretch
        layout.addWidget(all_employees_frame, 1) # Stretch
        return layout

    def _build_schedule_preview(self, parent_container):
        self.schedule_layout = QGridLayout(parent_container)
        self.schedule_labels = {day: [] for day in DAYS}
        self.final_schedule_day_headers = {}

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

            for r in range(MAX_DISPLAY_ROWS_SCHEDULE):
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
        layout.addWidget(QLabel("<b>Name</b>"), 0, 0)
        layout.addWidget(QLabel("<b>Availability</b>"), 0, 1)
        layout.addWidget(QLabel("<b>Notes</b>"), 0, 2)
        layout.addWidget(QLabel("<b>Time Off Dates</b>"), 0, 3)
        layout.setColumnStretch(0, 2)
        layout.setColumnStretch(1, 3)
        layout.setColumnStretch(2, 2)
        layout.setColumnStretch(3, 3)

        self.all_employees_rows = []
        for r in range(self.max_display_rows_per_list):
            row_idx = r + 1
            name_lbl = ClickableLabel(self, "") # Still a ClickableLabel for drag-and-drop
            name_lbl.is_draggable = True
            avail_lbl = QLabel("")
            notes_lbl = QLabel("")
            time_off_lbl = QLabel("")
            
            row_widgets = [name_lbl, avail_lbl, notes_lbl, time_off_lbl]
            for i, widget in enumerate(row_widgets):
                widget.setFixedHeight(22)
                layout.addWidget(widget, row_idx, i)
            
            self.all_employees_rows.append({
                'name_lbl': name_lbl, 'avail_lbl': avail_lbl, 'notes_lbl': notes_lbl,
                'time_off_lbl': time_off_lbl,
                'conceptual_row_widgets': row_widgets, 'employee': None
            })
        layout.setRowStretch(self.max_display_rows_per_list + 1, 1)

    def _build_unassigned_grid(self):
        container = QWidget()
        self.unassigned_scroll_area.setWidget(container)
        layout = QGridLayout(container)

        self.unassigned_labels = {day: [] for day in DAYS}
        for c, day in enumerate(DAYS):
            layout.addWidget(QLabel(f"<b>{day}</b>"), 0, c, alignment=Qt.AlignmentFlag.AlignCenter)
            for r in range(MAX_ROWS_UNASSIGNED_PER_DAY_COLUMN):
                lbl = ClickableLabel(self, "")
                lbl.setFixedHeight(20)
                lbl.is_draggable = True # These labels can be dragged
                layout.addWidget(lbl, r + 1, c)
                self.unassigned_labels[day].append(lbl)
            layout.setColumnStretch(c, 1)
        layout.setRowStretch(MAX_ROWS_UNASSIGNED_PER_DAY_COLUMN + 1, 1)

    def _open_employee_editor(self):
        """Opens the employee editor dialog."""
        # Pass the name of the currently highlighted employee, if any
        selected_name = self.highlighted_employee_obj.name if self.highlighted_employee_obj else None
        dialog = EmployeeEditorWindow(self, selected_name)
        dialog.exec()

    def reload_from_master_file(self):
        """Reloads all employee data from the master file and refreshes the entire UI."""
        self.schedule.clear_schedule()
        df, count = self._load_and_count_master_file_data()
        self.load_master_employees(df)
        print("UI reloaded from master file.")

# --- Core Logic Methods (largely adapted from original) ---

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
                notes = str(row.get('Notes', '')).strip()
                emp = Employee(name, availability, notes)
                self.employees.append(emp)

                # Check if the employee is a "set schedule" employee
                set_schedule_status = str(row.get('Set Schedule', '')).strip().lower()
                if set_schedule_status == 'yes':
                    # Automatically assign the employee to the schedule for available days
                    for day in DAYS:
                        if emp.availability.get(day, False): # Ensure employee is available on this day
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
            colors = ("transparent", "#f0f0f0")
        else:
            colors = ("transparent", "#313131")
        return f"background-color: {colors[index % 2]};"

    def update_all_employees_grid(self):
        try:
            # Read the file once to get all data
            df_master = pd.read_excel(master_employee_file_path)
            # Create a lookup map for the time off dates
            time_off_map = dict(zip(df_master['Name'], df_master.get('Time Off Dates', '')))
        except Exception as e:
            print(f"Could not read master file for time off data: {e}")
            time_off_map = {}

        for i, row_data in enumerate(self.all_employees_rows):
            if i < len(self.employees):
                emp = self.employees[i]
                row_data['employee'] = emp
                row_data['name_lbl'].setText(emp.name)
                row_data['name_lbl'].employee_obj = emp # Attach obj for drag/click
                row_data['avail_lbl'].setText(', '.join([d for d, v in emp.availability.items() if v]))
                row_data['notes_lbl'].setText(emp.notes)
                
                notes_display = emp.notes if emp.notes and emp.notes.lower() != 'nan' else ""
                row_data['notes_lbl'].setText(notes_display)

                time_off_str = time_off_map.get(emp.name, '')
                row_data['time_off_lbl'].setText(str(time_off_str) if pd.notna(time_off_str) else "")
                
                for lbl in row_data['conceptual_row_widgets']:
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
        and who are still under their max per week limit.
        """
        try:
            df = pd.read_excel(master_employee_file_path)
            max_per_week_map = dict(zip(df['Name'], df.get('Max Per Week', [7]*len(df))))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read master employee file:\n{e}")
            return

        # Count how many times each employee is already scheduled this week
        current_schedule_count = {emp.name: 0 for emp in self.employees}
        for day in DAYS:
            for emp in self.schedule.scheduled.get(day, []):
                current_schedule_count[emp.name] += 1

        # Reset all labels in the unassigned grid
        for day in DAYS:
            for lbl in self.unassigned_labels[day]:
                lbl.setText("")
                lbl.employee_obj = None
                lbl.day_key = day
                lbl.hide()

        # Populate grid day-by-day
        for day_index, day in enumerate(DAYS):
            available_and_unscheduled_for_day = []
            for emp in self.employees:
                if (
                    emp.availability.get(day, False) and
                    all(e.name != emp.name for e in self.schedule.scheduled[day]) and
                    current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                ):
                    available_and_unscheduled_for_day.append(emp)

            # Sort by name for consistent display
            available_and_unscheduled_for_day.sort(key=lambda e: e.name)

            # Show them in the unassigned label grid
            for r_idx, emp in enumerate(available_and_unscheduled_for_day):
                if r_idx < len(self.unassigned_labels[day]):
                    lbl = self.unassigned_labels[day][r_idx]
                    lbl.setText(emp.name)
                    lbl.employee_obj = emp
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

            # Get the header widgets (label and progress bar)
            header_widgets = self.final_schedule_day_headers[day]
            progress_bar = header_widgets['progress']

            # Configure the progress bar's range, value, and text format
            progress_bar.setRange(0, max_val)
            progress_bar.setValue(count)
            progress_bar.setFormat(f"{count} / {max_val}") # Custom text format

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

# --- Drag/Drop and Click Handling ---
    def on_drag_start(self, label_widget):
        if label_widget and label_widget.employee_obj:
            self.drag_data = {'label_widget': label_widget, 'employee': label_widget.employee_obj}
            # Visual feedback for drag start
            label_widget.setStyleSheet(f"background-color: {self.palette().color(QPalette.ColorRole.Highlight).name()};")

    def mouseReleaseEvent(self, event):
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
        if checked:
            app.setStyleSheet(self.dark_stylesheet)
        else:
            app.setStyleSheet(self.light_stylesheet)
        self._update_highlight_style()
        self.update_all_views()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Breeze")
    window = SchedulerApp()
    window.show()
    sys.exit(app.exec())

# package into exe with all needed files and imports
# pyinstaller --noconfirm -D --windowed --icon="app_icon.ico" --add-data="excel_files;excel_files" schedule_app_pyqt6.py