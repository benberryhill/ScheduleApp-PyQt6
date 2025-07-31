import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QCheckBox, QScrollArea, QFrame,
    QMessageBox, QSizePolicy, QMenu, QDialog, QColorDialog, QTabWidget, QGroupBox,
    QDialogButtonBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QEvent, QSize
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
                'Notes': ['Team Lead', '', 'Part-time', 'New Hire', '', 'Floater']
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
                # Re-select the employee to maintain context in the editor
                self.main_window.select_employee_by_object(self.employee_obj)
            else:
                # This case is unlikely if the UI is correct, but good to have
                print(f"UI state issue: Could not find {self.employee_obj.name} to remove from {self.day_key}")
        except Exception as e:
            print(f"Error during 'Remove from Schedule' action: {e}")

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
        # Make the icon a bit larger for visibility
        font = self.settings_button.font()
        font.setPointSize(11)
        self.settings_button.setFont(font)

        self.auto_schedule_button = QPushButton("Auto Schedule Available Employees")
        self.auto_schedule_button.clicked.connect(self.auto_schedule_available_employees)

        self.export_button = QPushButton("Export Schedule")
        self.export_button.clicked.connect(self.export_schedule)

        self.toggle_theme_button = QCheckBox("Dark Mode")
        self.toggle_theme_button.toggled.connect(self.toggle_theme)

        layout.addWidget(self.settings_button)
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

        # Editor
        edit_info_frame = self._create_editor_frame()

        layout.addWidget(unassigned_frame, 1) # Stretch
        layout.addWidget(all_employees_frame, 1) # Stretch
        layout.addWidget(edit_info_frame)
        return layout

    def _create_editor_frame(self):
        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.StyledPanel)
        main_layout = QVBoxLayout(frame)
        
        # Name and Notes
        name_notes_layout = QHBoxLayout()
        self.edit_name = QLineEdit()
        self.edit_name.setPlaceholderText("Name")
        self.edit_notes = QLineEdit()
        self.edit_notes.setPlaceholderText("Notes")
        name_notes_layout.addWidget(self.edit_name)
        name_notes_layout.addWidget(self.edit_notes)
        
        # Availability Checkboxes
        availability_layout = QHBoxLayout()
        self.availability_boxes = {}
        for day in DAYS:
            chk = QCheckBox(day)
            checkbox_style = """
                QCheckBox {
                    background-color: transparent; /* Key: Make the entire QCheckBox widget background transparent */
                    border: none;                   /* Remove any border on the QCheckBox widget itself */
                    margin-left: 4px;               /* Push the QCheckBox widget slightly right */
                    padding: 0;                     /* Remove default internal padding of the QCheckBox */
                    
                    /* Ensure label color inherits properly */
                    color: inherit; 
                }
                QCheckBox::indicator {
                    /* Style the actual checkable box (the indicator) */
                    /* This is what ensures the box itself is visible and has a consistent look */
                    border: 1px solid gray;        /* Add a consistent border to the indicator */
                    border-radius: 3px;            /* Slightly rounded corners */
                    width: 14px;                   /* Fixed size for consistency */
                    height: 14px;                  /* Fixed size for consistency */
                    background-color: transparent; /* Ensure the indicator's own background is transparent */
                    
                    /* When checked, Qt will draw the checkmark inside this styled indicator */
                }
                QCheckBox::indicator:checked {
                    /* Style for the indicator when it is checked */
                    background-color: #4CAF50; /* Green for checked, a common UI color */
                    border-color: #388E3C;     /* Darker green border for checked */
                    /* If you want the checkmark to be white: */
                    /* This is tricky and platform dependent. Often, the system draws it. */
                    /* If you must force a color, it often involves custom painter, which is complex. */
                    /* For now, let's rely on the default checkmark color or ensure good contrast. */
                }

                /* Hover effects are good for usability */
                QCheckBox:hover {
                    background-color: rgba(128, 128, 128, 30); /* Subtle background tint on hover */
                }
                QCheckBox::indicator:hover {
                    border-color: #555; /* Slightly darker border on hover */
                }
                QCheckBox::indicator:checked:hover {
                    background-color: #66BB6A; /* Lighter green on hover when checked */
                    border-color: #4CAF50;
                }
            """
            chk.setStyleSheet(checkbox_style)
            self.availability_boxes[day] = chk
            availability_layout.addWidget(chk)
            availability_layout.addStretch()

        # Buttons
        button_layout = QHBoxLayout()
        self.add_emp_button = QPushButton("Add New Employee")
        self.add_emp_button.clicked.connect(self.add_new_employee)
        self.update_emp_button = QPushButton("Update Selected Employee")
        self.update_emp_button.clicked.connect(self.update_employee)
        self.del_curr_emp_button = QPushButton("Delete Selected Employee")
        self.del_curr_emp_button.clicked.connect(self.del_curr_employee)
        self.undo_del_button = QPushButton("Undo Delete")
        self.undo_del_button.clicked.connect(self.undo_del_employee)
        button_layout.addStretch()
        button_layout.addWidget(self.add_emp_button)
        button_layout.addWidget(self.update_emp_button)
        button_layout.addWidget(self.del_curr_emp_button)
        button_layout.addWidget(self.undo_del_button)
        button_layout.addStretch()

        main_layout.addLayout(name_notes_layout)
        main_layout.addLayout(availability_layout)
        main_layout.addLayout(button_layout)
        return frame

    def _build_schedule_preview(self, parent_container):
        self.schedule_layout = QGridLayout(parent_container)
        self.schedule_labels = {day: [] for day in DAYS}
        self.final_schedule_day_headers = {}

        for i, day in enumerate(DAYS):
            header = QLabel(f"<b>{day}</b>")
            header.setAlignment(Qt.AlignmentFlag.AlignCenter)
            # header.setAutoFillBackground(True) # Needed for setting background color
            self.final_schedule_day_headers[day] = header
            self.schedule_layout.addWidget(header, 0, i)

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
        layout.setColumnStretch(0, 2)
        layout.setColumnStretch(1, 3)
        layout.setColumnStretch(2, 4)

        self.all_employees_rows = []
        for r in range(self.max_display_rows_per_list):
            row_idx = r + 1
            name_lbl = ClickableLabel(self, "")
            # This is where we enable click-to-select functionality
            name_lbl.mousePressEvent = lambda event, emp_row=r: self.on_employee_label_click(emp_row)
            name_lbl.is_draggable = True

            avail_lbl = QLabel("")
            notes_lbl = QLabel("")
            
            row_widgets = [name_lbl, avail_lbl, notes_lbl]
            for i, widget in enumerate(row_widgets):
                widget.setFixedHeight(22)
                layout.addWidget(widget, row_idx, i)
            
            self.all_employees_rows.append({
                'name_lbl': name_lbl, 'avail_lbl': avail_lbl, 'notes_lbl': notes_lbl,
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
        # Store selected employee to re-select after update
        current_selected_name = self.selected_employee_obj.name if self.selected_employee_obj else None
        
        self.update_all_employees_grid()
        self.update_unassigned_grid()
        self.update_final_schedule_display()
        
        self.deselect_employee() # Clears visual selection
        if current_selected_name:
            emp_to_reselect = next((e for e in self.employees if e.name == current_selected_name), None)
            if emp_to_reselect:
                self.select_employee_by_object(emp_to_reselect)
            else:
                self.clear_editor_fields()
        else:
            self.clear_editor_fields()

    def get_alternating_row_style(self, index):
        if not self.is_dark_mode:
            colors = ("transparent", "#f0f0f0")
        else:
            colors = ("transparent", "#313131")
        return f"background-color: {colors[index % 2]};"

    def update_all_employees_grid(self):
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
                
                for lbl in row_data['conceptual_row_widgets']:
                    lbl.setStyleSheet(self.get_alternating_row_style(i))
                    lbl.show()
            else: # Hide unused rows
                row_data['employee'] = None
                for lbl in row_data['conceptual_row_widgets']:
                    lbl.setText("")
                    lbl.hide()
        self.highlight_selected_row()

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
            max_val = current_max.get(day, 999)
            header = self.final_schedule_day_headers[day]

            # Select the correct map of styles based on the current theme
            theme_mode = 'dark' if self.is_dark_mode else 'light'
            style_map = self.header_styles[theme_mode]

            # Determine the status to select the correct style
            status = "empty"
            if count > 0:
                status = "warning"
            if count >= max_val:
                status = "semi_full"
            if count > max_val: # If more employees are scheduled than the max
                status = "full"

            # Apply the determined stylesheet to the header label
            header.setStyleSheet(style_map.get(status, ""))

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
        Automatically schedules employees day-by-day, prioritizing days with the fewest total available workers.
        For each day, assigns the best available employees who haven't reached their max weekly limit.
        """
        try:
            df = pd.read_excel(master_employee_file_path)
            employee_map = {e.name: e for e in self.employees}
            max_per_week_map = dict(zip(df['Name'], df.get('Max Per Week', [7]*len(df))))
            seniority_ordered_employees = [employee_map[name] for name in df['Name'].dropna() if name in employee_map]
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Could not read master employee file for auto-scheduling:\n{e}")
            return

        # --- Step 1: Get max needed per day from UI ---
        max_per_day = self.max_per_day

        # --- Step 2: Count how many people are available per day (day difficulty) ---
        day_availability_counts = {day: 0 for day in DAYS}
        for emp in seniority_ordered_employees:
            for day in DAYS:
                if emp.availability.get(day, False):
                    day_availability_counts[day] += 1

        # Order days from hardest to fill to easiest
        sorted_days = sorted(DAYS, key=lambda d: day_availability_counts[d])

        # --- Step 3: Track how many times each employee is scheduled ---
        current_schedule_count = {emp.name: 0 for emp in self.employees}
        for day in DAYS:
            for emp in self.schedule.scheduled.get(day, []):
                current_schedule_count[emp.name] += 1

        # --- Step 4: Assign employees to each day in difficulty order ---
        employees_added_count = 0

        for day in sorted_days:
            needed = max_per_day.get(day, 0) - len(self.schedule.scheduled.get(day, []))
            if needed <= 0:
                continue  # Day already full

            # Filter eligible employees: available that day and under their weekly limit
            eligible_employees = [
                emp for emp in seniority_ordered_employees
                if emp.availability.get(day, False)
                and current_schedule_count.get(emp.name, 0) < max_per_week_map.get(emp.name, 7)
                and emp.name not in [e.name for e in self.schedule.scheduled.get(day, [])]
            ]

            # Sort by fewest total current assignments to balance load
            eligible_employees.sort(key=lambda e: current_schedule_count[e.name])

            # Assign up to needed
            for emp in eligible_employees[:needed]:
                self.schedule.assign_employee(emp, day)
                current_schedule_count[emp.name] += 1
                employees_added_count += 1

        # --- Step 5: Done ---
        self.update_all_views()
        QMessageBox.information(self, "Auto-Schedule Complete",
                                f"Finished auto-scheduling. {employees_added_count} new assignments were made.")

    def on_employee_label_click(self, row_index):
        if row_index < len(self.all_employees_rows):
            employee = self.all_employees_rows[row_index].get('employee')
            if employee:
                if self.selected_employee_obj == employee:
                    self.deselect_employee()
                    self.clear_editor_fields()
                else:
                    self.select_employee_by_object(employee)

    def select_employee_by_object(self, employee):
        self.deselect_employee()
        self.selected_employee_obj = employee
        
        self.edit_name.setText(employee.name)
        self.edit_name.setEnabled(False) # Prevent editing name
        self.edit_notes.setText(employee.notes)
        notes_display = employee.notes if employee.notes and employee.notes.lower() != 'nan' else ""
        self.edit_notes.setText(notes_display) 
        for day, chk in self.availability_boxes.items():
            chk.setChecked(employee.availability.get(day, False))

        self.highlight_selected_row()

    def deselect_employee(self):
        if self.selected_row_widgets:
            # Find index to restore correct alternating color
            try:
                idx = self.employees.index(self.selected_employee_obj)
                style = self.get_alternating_row_style(idx)
                for w in self.selected_row_widgets:
                    w.setStyleSheet(style)
            except (ValueError, AttributeError):
                pass # Employee might have been removed

        self.selected_employee_obj = None
        self.selected_row_widgets = None
        if hasattr(self, 'edit_name'):
            self.edit_name.setEnabled(True)

    def highlight_selected_row(self):
        if self.selected_employee_obj:
            try:
                idx = self.employees.index(self.selected_employee_obj)
                if idx < len(self.all_employees_rows):
                    row_data = self.all_employees_rows[idx]
                    for w in row_data['conceptual_row_widgets']:
                        w.setStyleSheet(self.highlight_style)
                    self.selected_row_widgets = row_data['conceptual_row_widgets']
            except ValueError:
                self.selected_row_widgets = None

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

    def clear_editor_fields(self):
        self.edit_name.setEnabled(True)
        self.edit_name.clear()
        self.edit_notes.clear()
        for chk in self.availability_boxes.values():
            chk.setChecked(False)

    def add_new_employee(self):
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.critical(self, "Input Error", "Employee name cannot be empty.")
            return
        if any(e.name.lower() == name.lower() for e in self.employees):
            QMessageBox.critical(self, "Input Error", f"Employee '{name}' already exists.")
            return

        notes = self.edit_notes.text().strip()
        availability = {day: chk.isChecked() for day, chk in self.availability_boxes.items()}
        new_emp = Employee(name, availability, notes)
        
        self.employees.append(new_emp)
        self.employees.sort(key=lambda e: e.name)
        self.save_master_list()
        self.update_all_views()
        self.clear_editor_fields()
        QMessageBox.information(self, "Success", f"Employee '{name}' added.")

    def update_employee(self):
        if not self.selected_employee_obj:
            QMessageBox.warning(self, "Update Error", "No employee selected.")
            return

        # Update the in-memory employee object first
        emp = self.selected_employee_obj
        emp.notes = self.edit_notes.text().strip()
        emp.availability = {day: chk.isChecked() for day, chk in self.availability_boxes.items()}

        # Directly update the Excel file to preserve order and other columns ---
        try:
            # Read the master file into a DataFrame
            df = pd.read_excel(master_employee_file_path)

            # Find the index of the employee to update by name
            idx_list = df.index[df['Name'] == emp.name].tolist()
            if not idx_list:
                QMessageBox.critical(self, "Save Error", f"Could not find '{emp.name}' in the master file to update. It may have been renamed or deleted externally.")
                return

            idx = idx_list[0] # Get the integer index

            # Update the DataFrame at the specific row for Notes and availability days
            df.loc[idx, 'Notes'] = emp.notes
            for day, available in emp.availability.items():
                if day in df.columns: # Only update columns that exist in the file
                    df.loc[idx, day] = 'Yes' if available else 'No'

            # Save the entire modified DataFrame back to the Excel file
            # This preserves row order, other columns (like 'Set Schedule'), and formatting
            df.to_excel(master_employee_file_path, index=False)

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to update master list file:\n{e}")
            return

        # Refresh the UI to reflect the changes
        self.update_all_views()
        QMessageBox.information(self, "Success", f"Employee '{emp.name}' updated in the master file.")

    def del_curr_employee(self):
        if not self.selected_employee_obj:
            QMessageBox.warning(self, "Delete Error", "No employee selected to delete.")
            return

        emp_to_delete = self.selected_employee_obj
        
        reply = QMessageBox.question(self, "Confirm Deletion",
            f"Are you sure you want to delete '{emp_to_delete.name}'?\nThis can be undone with the 'Undo Delete' button.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.No:
            return

        # --- Proceed with deletion ---
        # 1. Prepare deleted employee data
        deleted_emp_data = {
            'Name': emp_to_delete.name,
            'Notes': emp_to_delete.notes,
            **{d: ('Yes' if emp_to_delete.availability.get(d) else 'No') for d in DAYS}
        }
        
        # 2. Load or create the recently_deleted file
        deleted_file_path = os.path.join(EXCEL_FOLDER, RECENTLY_DELETED_FILE)
        try:
            if os.path.exists(deleted_file_path):
                df_deleted = pd.read_excel(deleted_file_path)
            else:
                df_deleted = pd.DataFrame(columns=['Name'] + DAYS + ['Notes'])
            
            # 3. Append and save
            new_row_df = pd.DataFrame([deleted_emp_data])
            df_deleted = pd.concat([df_deleted, new_row_df], ignore_index=True)
            df_deleted.to_excel(deleted_file_path, index=False)

        except Exception as e:
            QMessageBox.critical(self, "File Error", f"Could not update the deleted employees list:\n{e}")
            return

        # 4. Remove from main employee list
        self.employees.remove(emp_to_delete)
        
        # 5. Remove from any scheduled day
        for day in DAYS:
            self.schedule.remove_employee(emp_to_delete, day)

        # 6. Save the master list
        self.save_master_list()
        
        # 7. Update UI
        self.deselect_employee()
        self.clear_editor_fields()
        self.update_all_views()
        
        QMessageBox.information(self, "Success", f"Employee '{emp_to_delete.name}' has been deleted.")

    def undo_del_employee(self):
        deleted_file_path = os.path.join(EXCEL_FOLDER, RECENTLY_DELETED_FILE)

        # 1. Check if the file exists and is not empty
        if not os.path.exists(deleted_file_path):
            QMessageBox.information(self, "Undo Delete", "No recently deleted employees to restore.")
            return
        
        try:
            df_deleted = pd.read_excel(deleted_file_path)
            if df_deleted.empty:
                QMessageBox.information(self, "Undo Delete", "The deleted employees list is empty.")
                return

            # 2. Get the last deleted employee (last row)
            last_deleted_series = df_deleted.iloc[-1]
            
            # 3. Check for name collision
            restored_name = last_deleted_series['Name']
            if any(e.name.lower() == restored_name.lower() for e in self.employees):
                QMessageBox.critical(self, "Restore Error", f"An employee named '{restored_name}' already exists in the master list.")
                return

            # 4. Convert row back to Employee object
            availability = {day: str(last_deleted_series.get(day, '')).strip().lower() == 'yes' for day in DAYS}
            notes = str(last_deleted_series.get('Notes', '')).strip()
            restored_emp = Employee(restored_name, availability, notes)
            
            # 5. Add back to main list and sort
            self.employees.append(restored_emp)
            self.employees.sort(key=lambda e: e.name)
            
            # 6. Remove from deleted list and save
            df_deleted = df_deleted.iloc[:-1] # Drop the last row
            df_deleted.to_excel(deleted_file_path, index=False)

            # 7. Save the master list
            self.save_master_list()
            
            # 8. Update UI
            self.update_all_views()
            
            QMessageBox.information(self, "Success", f"Restored '{restored_name}' to the master list.")

        except Exception as e:
            QMessageBox.critical(self, "Restore Error", f"Failed to restore employee:\n{e}")

    def save_master_list(self):
        """
        Saves the current state of self.employees to the master Excel file.
        This version preserves column order and extra columns (e.g., 'Set Schedule')
        by reading the original file first. It's used for Add/Delete/Undo operations.
        """
        try:
            original_data = {}
            all_columns = ['Name'] + DAYS + ['Notes']

            # 1. Try to read the original file to get its structure and extra data
            try:
                df_orig = pd.read_excel(master_employee_file_path)
                all_columns = df_orig.columns.tolist() # Preserve original column order
                app_managed_cols = ['Name'] + DAYS + ['Notes']
                extra_cols = [c for c in all_columns if c not in app_managed_cols]

                if extra_cols:
                    # Create a lookup dictionary for the extra data of each employee
                    for _, row in df_orig.iterrows():
                        if 'Name' in row and pd.notna(row['Name']):
                            original_data[row['Name']] = {col: row[col] for col in extra_cols}
            except Exception:
                # File might not exist or is malformed, proceed with the default structure
                pass

            # 2. Build the list of data rows from the current in-memory employee list
            data_to_save = []
            for e in self.employees:
                row_data = {
                    'Name': e.name,
                    'Notes': e.notes,
                    **{d: ('Yes' if e.availability.get(d, False) else 'No') for d in DAYS}
                }
                # If we have stored extra data for this employee, add it back in
                if e.name in original_data:
                    row_data.update(original_data[e.name])
                data_to_save.append(row_data)

            # 3. Create the final DataFrame
            df_final = pd.DataFrame(data_to_save)

            # 4. Ensure all original columns are present and in the correct order
            # This prevents columns from being dropped or reordered
            df_final = df_final.reindex(columns=all_columns, fill_value='')

            # 5. Save to the master file, overwriting it with the preserved structure
            df_final.to_excel(master_employee_file_path, index=False)

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save master list:\n{e}")

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
        # Re-select the employee that was just dragged
        self.select_employee_by_object(employee_to_drop)

    def on_final_schedule_click(self, label_widget):
        employee = label_widget.employee_obj
        day = label_widget.day_key
        if employee and day:
            reply = QMessageBox.question(self, "Remove Employee",
                f"Remove '{employee.name}' from {day}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                if self.schedule.remove_employee(employee, day):
                    self.update_all_views()
                    self.select_employee_by_object(employee) # Reselect after removal

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
    window = SchedulerApp()
    window.show()
    sys.exit(app.exec())

# highlight names. settings menu. changed to global styles and max per day settings