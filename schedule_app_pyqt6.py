import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QLineEdit, QCheckBox, QScrollArea, QFrame,
    QMessageBox, QSizePolicy, QMenu
)
from PyQt6.QtCore import Qt, pyqtSignal, QEvent
from PyQt6.QtGui import QPalette, QColor

# --- Configuration ---
EXCEL_FOLDER = "excel_files"
RECENTLY_DELETED_FILE = "recently_deleted.xlsx" # For undo functionality
EMPLOYEE_FILE = "Employees_Full_List.xlsx" # Master list
DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
MAX_DISPLAY_ROWS_SCHEDULE = 50 # Max rows per day in Final Schedule
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

# --- Custom Qt Widget for Click/Drag Events ---
class ClickableLabel(QLabel):
    """ A custom QLabel that can hold employee/day data and handle mouse clicks. """
    def __init__(self, main_window, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.main_window = main_window
        self.employee_obj = None
        self.day_key = None  # This must hold the day (e.g., 'Monday') for context menu actions
        self.is_draggable = False
        self.is_clickable_remove = False

    def mousePressEvent(self, event):
        try:
            if event.button() == Qt.MouseButton.LeftButton:
                # Handle left-click events
                if self.is_draggable and self.employee_obj:
                    self.main_window.on_drag_start(self)
                elif self.is_clickable_remove and self.employee_obj and self.day_key:
                    self.main_window.on_final_schedule_click(self)
            elif event.button() == Qt.MouseButton.RightButton:
                # Handle right-click events
                if self.employee_obj:
                    self.show_context_menu(event.globalPosition().toPoint())
            else:
                # For other mouse buttons
                super().mousePressEvent(event)
        except Exception as e:
            print(f"Error during mousePressEvent: {e}")

    def show_context_menu(self, global_position):
        """ Create and show a context menu on right-click. """
        try:
            menu = QMenu(self)

            # Add actions to the context menu
            add_to_schedule_action = menu.addAction("Add to Final Schedule")
            add_to_schedule_action.triggered.connect(self.on_add_to_schedule)

            # Safely display the menu at the correct global position
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
        """ Handle the 'Add to Final Schedule' option. """
        try:
            if not self.employee_obj:
                print("Error: No employee object found for this label.")
                return
            if not self.day_key:
                print("Error: No day key assigned to this label.")
                return

            # Call the scheduling logic
            success = self.main_window.schedule.assign_employee(self.employee_obj, self.day_key)
            if success:
                self.main_window.update_all_views()
                print(f"Added {self.employee_obj.name} to the final schedule on {self.day_key}.")
            else:
                print(f"Failed to add {self.employee_obj.name} to the schedule on {self.day_key}.")
        except Exception as e:
            print(f"Error during 'Add to Final Schedule' action: {e}")

# --- Main Application ---
class SchedulerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Semi-Auto Scheduler")
        self.setGeometry(100, 100, 1700, 950)

        self.employees = []
        self.schedule = Schedule()
        self.max_per_day = {day: 20 for day in DAYS}
        self.selected_employee_obj = None
        self.selected_row_widgets = None
        self.drag_data = {'label_widget': None, 'employee': None}

        # For mapping widgets to employee objects, similar to the original's approach
        self.widget_to_employee = {}

        # Theming
        self.is_dark_mode = False
        self._setup_styles()
        
        # Define styles for the schedule day headers
        self.header_base_style_template = "QLabel {{ font-weight: bold; padding: 4px; border-radius: 4px; color: {text_color}; background-color: {bg_color}; }}"
        
        self.dark_header_styles = {
            "empty": self.header_base_style_template.format(bg_color="#3c3c3c", text_color="white"), # Darker gray for empty
            "warning": self.header_base_style_template.format(bg_color="darkred", text_color="white"), # Darker red for warning
            "semi_full": self.header_base_style_template.format(bg_color="darkblue", text_color="white"), # Darker blue for semi-full
            "full": self.header_base_style_template.format(bg_color="teal", text_color="white"), # Teal for over capacity
        }
        self.light_header_styles = {
            "empty": self.header_base_style_template.format(bg_color="#f0f0f0", text_color="black"), # Lighter gray for empty
            "warning": self.header_base_style_template.format(bg_color="indianred", text_color="black"),
            "semi_full": self.header_base_style_template.format(bg_color="royalblue", text_color="white"), # White text for Royal Blue
            "full": self.header_base_style_template.format(bg_color="lightseagreen", text_color="black"),
        }

        # Load initial data to determine UI size
        initial_employee_df, initial_employee_count = self._load_and_count_master_file_data()
        self.max_display_rows_per_list = max(30, initial_employee_count + 10)

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
                width: 10px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #555555;
                min-height: 20px;
                border-radius: 5px;
            }
            QCheckBox::indicator {
                width: 13px;
                height: 13px;
            }
        """

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
        self.schedule_select = QComboBox()
        self.schedule_select.addItems(["Select a schedule..."] + xlsx_files)
        self.schedule_select.setMinimumWidth(300)
        self.schedule_select.currentTextChanged.connect(self.load_schedule_from_selection)
        
        self.export_button = QPushButton("Export Schedule")
        self.export_button.clicked.connect(self.export_schedule)

        self.toggle_theme_button = QCheckBox("Dark Mode")
        self.toggle_theme_button.toggled.connect(self.toggle_theme)

        layout.addWidget(self.schedule_select)
        layout.addStretch(1)
        layout.addWidget(self.toggle_theme_button)
        layout.addWidget(self.export_button)
        return layout

    def _create_left_column(self):
        layout = QVBoxLayout()

        # Final Schedule Preview with Scroll
        final_schedule_scroll_area = QScrollArea()
        final_schedule_scroll_area.setWidgetResizable(True)
        self.final_schedule_frame_container = QFrame()
        self.final_schedule_frame_container.setFrameShape(QFrame.Shape.StyledPanel)

        # Add the final schedule frame container to the scroll area
        final_schedule_scroll_area.setWidget(self.final_schedule_frame_container)
        layout.addWidget(QLabel("<h2>Final Schedule</h2>"))
        layout.addWidget(final_schedule_scroll_area, 1)  # Stretchable

        # Build the schedule preview inside the container
        self._build_schedule_preview(self.final_schedule_frame_container)

        # Max per Day Settings
        max_settings_frame = self._create_max_settings_frame()
        layout.addWidget(max_settings_frame)
        return layout

    def _create_max_settings_frame(self):
        frame = QFrame()
        frame.setFrameShape(QFrame.Shape.StyledPanel)
        layout = QHBoxLayout(frame)
        
        self.day_max_grid = QGridLayout()
        self.max_entries = {}
        for i, day in enumerate(DAYS):
            day_label = QLabel(day)
            day_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            entry = QLineEdit(str(self.max_per_day[day]))
            entry.setFixedWidth(45)
            entry.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.max_entries[day] = entry
            
            v_box = QVBoxLayout()
            v_box.addWidget(day_label)
            v_box.addWidget(entry)
            self.day_max_grid.addLayout(v_box, 0, i)

        self.update_max_button = QPushButton("Update Max")
        self.update_max_button.clicked.connect(self.update_max_values_and_refresh)

        layout.addLayout(self.day_max_grid)
        layout.addStretch(1)
        layout.addWidget(self.update_max_button)
        return frame

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

    def load_schedule_from_selection(self, selected_filename):
        if selected_filename == "Select a schedule...":
            self.schedule.clear_schedule()
            self.update_all_views()
            return
        
        schedule_file_path = os.path.join(EXCEL_FOLDER, selected_filename)
        try:
            df = pd.read_excel(schedule_file_path)
            self.schedule.clear_schedule()
            master_map = {emp.name: emp for emp in self.employees}
            
            for day in DAYS:
                if day in df.columns:
                    for name in df[day].dropna():
                        name = str(name).strip()
                        if name in master_map:
                            self.schedule.assign_employee(master_map[name], day)

        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Error loading schedule from {selected_filename}:\n{e}")
            self.schedule.clear_schedule()
        finally:
            self.update_all_views()

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
        Updates the grid displaying employees available but not yet scheduled for each day.
        """
        # Reset all labels in the unassigned grid first
        for day in DAYS:
            for lbl in self.unassigned_labels[day]:
                lbl.setText("")
                lbl.employee_obj = None
                lbl.day_key = day
                lbl.hide()

        # Populate the grid day by day
        for day_index, day in enumerate(DAYS):
            # Find employees available on this specific day AND not scheduled on this specific day
            available_and_unscheduled_for_day = []
            for emp in self.employees:
                # Check if employee is available on this specific day
                if emp.availability.get(day, False):
                    # Check if employee is NOT already scheduled on this specific day
                    is_scheduled_today = any(e.name == emp.name for e in self.schedule.scheduled[day])
                    if not is_scheduled_today:
                        available_and_unscheduled_for_day.append(emp)

            # Sort them by name for consistent display
            available_and_unscheduled_for_day.sort(key=lambda e: e.name)

            # Populate the labels for this day's column
            for r_idx, emp in enumerate(available_and_unscheduled_for_day):
                if r_idx < len(self.unassigned_labels[day]):
                    lbl = self.unassigned_labels[day][r_idx]
                    lbl.setText(emp.name)
                    lbl.employee_obj = emp # Attach employee object for drag
                    lbl.setStyleSheet(self.get_alternating_row_style(r_idx)) # Apply alternating colors
                    lbl.show()

    def update_final_schedule_display(self):
        current_max = {day: int(self.max_entries[day].text()) if self.max_entries[day].text().isdigit() else 999 for day in DAYS}

        for day, labels in self.schedule_labels.items():
            emps_on_day = self.schedule.scheduled[day]
            count = len(emps_on_day)
            max_val = current_max[day]
            header = self.final_schedule_day_headers[day]

            # Select the correct map of styles based on the current theme
            style_map = self.dark_header_styles if self.is_dark_mode else self.light_header_styles

            # Determine the status to select the correct style
            status = "empty"
            if count > 0:
                status = "warning"
            if count >= max_val:
                status = "semi_full"
            if count > max_val: # If more employees are scheduled than the max
                status = "full"

            # Apply the determined stylesheet to the header label
            header.setStyleSheet(style_map[status])

            # Update the text labels for the day's schedule entries (these are the inner labels, not the headers)
            for i, lbl in enumerate(labels):
                if i < len(emps_on_day):
                    emp = emps_on_day[i]
                    lbl.setText(emp.name)
                    lbl.employee_obj = emp
                    lbl.day_key = day
                else:
                    lbl.setText("")
                    lbl.employee_obj = None
                    lbl.day_key = None
                # Apply alternating row style to these inner labels
                lbl.setStyleSheet(self.get_alternating_row_style(i))

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
                    highlight_color = self.palette().color(QPalette.ColorRole.Highlight)
                    style = f"background-color: {highlight_color.name()};"
                    for w in row_data['conceptual_row_widgets']:
                        w.setStyleSheet(style)
                    self.selected_row_widgets = row_data['conceptual_row_widgets']
            except ValueError:
                self.selected_row_widgets = None

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
        
        emp = self.selected_employee_obj
        emp.notes = self.edit_notes.text().strip()
        emp.availability = {day: chk.isChecked() for day, chk in self.availability_boxes.items()}
        
        self.save_master_list()
        self.update_all_views()
        QMessageBox.information(self, "Success", f"Employee '{emp.name}' updated.")

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
        try:
            data = [{'Name': e.name, 'Notes': e.notes,
                     **{d: ('Yes' if e.availability.get(d) else 'No') for d in DAYS}}
                    for e in self.employees]
            df = pd.DataFrame(data)
            df = df[['Name'] + DAYS + ['Notes']]
            df.to_excel(master_employee_file_path, index=False)
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save master list:\n{e}")

    def update_max_values_and_refresh(self):
        self.update_final_schedule_display()

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

    def toggle_theme(self, checked):
        self.is_dark_mode = checked
        if checked:
            app.setStyleSheet(self.dark_stylesheet)
        else:
            app.setStyleSheet(self.light_stylesheet)
        self.update_all_views()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SchedulerApp()
    window.show()
    sys.exit(app.exec())