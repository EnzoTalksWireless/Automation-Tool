import sys
import logging
import json
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QCalendarWidget,
    QListWidget, QTabWidget, QSpinBox, QCheckBox, QMessageBox,
    QFileDialog, QScrollArea, QFrame, QDialog, QListWidgetItem,
    QProgressDialog, QGroupBox, QTextEdit, QStyle, QStyleFactory,
    QGridLayout


)
from PyQt6.QtCore import Qt, QPoint, QSize, pyqtSignal, QThread
from PyQt6.QtGui import QIcon, QDragEnterEvent, QDropEvent, QPalette, QColor, QFont
import os
import time
import keyboard

from automation_steps import StepType, STEP_DIALOGS
from recorder import ActionRecorder


from executor import WorkflowExecutor

# Directory constants
WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGES_DIR = os.path.join(WORKSPACE_DIR, "images")
AUTOMATIONS_DIR = os.path.join(WORKSPACE_DIR, "Saved Automations")

# Ensure directories exist
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(AUTOMATIONS_DIR, exist_ok=True)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automation.log'),
        logging.StreamHandler()
    ]
)

# Modern UI Style Sheet
STYLE_SHEET = """
QMainWindow, QWidget {
    background-color: #1e1e1e;
    color: #e0e0e0;
    font-family: system;
    font-size: 10pt;
}

QPushButton {
    background-color: #2d2d2d;
    color: #e0e0e0;
    border: 1px solid #3d3d3d;
    border-radius: 4px;
    padding: 4px 8px;
    min-width: 70px;
    min-height: 24px;
}

QPushButton:hover {
    background-color: #3d3d3d;
    border: 1px solid #4d4d4d;
}

QPushButton:pressed {
    background-color: #4d4d4d;
}

/* Style for emoji characters in buttons */
QPushButton[text*="⏺️"], QPushButton[text*="▶️"], QPushButton[text*="⏸️"], 
QPushButton[text*="⏯️"], QPushButton[text*="⏹️"], QPushButton[text*="✏️"], 
QPushButton[text*="🗑️"], QPushButton[text*="➕"], QPushButton[text*="🔍"], 
QPushButton[text*="📋"], QPushButton[text*="🧹"], QPushButton[text*="📁"],
QPushButton[text*="✅"], QPushButton[text*="❌"] {
    font-size: 10pt;
    padding-top: 4px;
    padding-bottom: 4px;
}

/* Style specific to automation steps in the list */
QListWidget::item {
    min-height: 35px;  /* Reduced height for each step in the list */
    font-size: 9pt;    /* Smaller font size */
    padding: 2px;      /* Reduced padding */
}

QLineEdit, QSpinBox, QComboBox, QTextEdit {
    background-color: #2d2d2d;
    color: #e0e0e0;
    border: 1px solid #3d3d3d;
    border-radius: 4px;
    padding: 2px 4px;
}

QLineEdit:focus, QSpinBox:focus, QComboBox:focus, QTextEdit:focus {
    border: 1px solid #0078D4;
}

QListWidget {
    background-color: #2d2d2d;
    color: #e0e0e0;
    border: 1px solid #3d3d3d;
    border-radius: 4px;
}

QListWidget::item:selected {
    background-color: #0078D4;
    color: white;
}

QListWidget::item:hover {
    background-color: #3d3d3d;
}

QTabWidget::pane {
    border: 1px solid #3d3d3d;
    background-color: #1e1e1e;
}

QTabWidget::tab-bar {
    left: 5px;
}

QTabBar::tab {
    background-color: #2d2d2d;
    color: #e0e0e0;
    border: 1px solid #3d3d3d;
    border-bottom: none;
    padding: 5px 10px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
}

QTabBar::tab:selected {
    background-color: #3d3d3d;
    border-bottom: none;
}

QTabBar::tab:hover {
    background-color: #4d4d4d;
}

QGroupBox {
    margin-top: 12px;
    padding-top: 16px;
    border: 1px solid #3d3d3d;
    border-radius: 4px;
}

QGroupBox::title {
    color: #e0e0e0;
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px;
}

QTextEdit {
    font-family: 'Consolas', 'Courier New', monospace;
    background-color: #1e1e1e;
    color: #e0e0e0;
}

QScrollBar:vertical {
    border: none;
    background-color: #2d2d2d;
    width: 10px;
    margin: 0;
}

QScrollBar::handle:vertical {
    background-color: #4d4d4d;
    min-height: 20px;
    border-radius: 5px;
}

QScrollBar::handle:vertical:hover {
    background-color: #5d5d5d;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    border: none;
    background-color: #2d2d2d;
    height: 10px;
    margin: 0;
}

QScrollBar::handle:horizontal {
    background-color: #4d4d4d;
    min-width: 20px;
    border-radius: 5px;
}

QScrollBar::handle:horizontal:hover {
    background-color: #5d5d5d;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* Removing checkbox styling to use default */
/* QCheckBox {
    color: #e0e0e0;
    spacing: 5px;
}

QCheckBox::indicator {
    width: 15px;
    height: 15px;
    border: 1px solid #3d3d3d;
    border-radius: 3px;
    background-color: #2d2d2d;
}

QCheckBox::indicator:checked {
    background-color: #0078D4;
}

QCheckBox::indicator:hover {
    border: 1px solid #4d4d4d;
} */

QMenuBar {
    background-color: #1e1e1e;
    color: #e0e0e0;
    border-bottom: 1px solid #3d3d3d;
}

QMenuBar::item {
    background-color: transparent;
    padding: 4px 8px;
}

QMenuBar::item:selected {
    background-color: #3d3d3d;
}

QMenu {
    background-color: #2d2d2d;
    color: #e0e0e0;
    border: 1px solid #3d3d3d;
}

QMenu::item {
    padding: 4px 20px;
}

QMenu::item:selected {
    background-color: #3d3d3d;
}

QStatusBar {
    background-color: #1e1e1e;
    color: #e0e0e0;
    border-top: 1px solid #3d3d3d;
}

/* Make the window more compact */
QMainWindow {
    min-width: 800px;
    min-height: 600px;
}

/* Adjust layout spacing */
QVBoxLayout, QHBoxLayout {
    spacing: 4px;
    margin: 2px;
}
"""

# Step Colors
STEP_COLORS = {
    StepType.MOUSE_CLICK: "#4da6ff",     # Blue
    StepType.KEYBOARD_TYPE: "#66ff66",   # Green
    StepType.KEYBOARD_SPECIAL: "#00cc99", # Teal
    StepType.WAIT: "#ffcc00",            # Orange
    StepType.LOOP_START: "#cc66ff",      # Purple
    StepType.LOOP_END: "#cc66ff"         # Purple
}

class AutomationStep(QFrame):
    """Represents a single automation step in the workflow"""
    stepMoved = pyqtSignal(int, int)  # Signal for drag and drop reordering
    paramsChanged = pyqtSignal(dict)  # Signal for parameter updates

    def __init__(self, step_type, params=None, parent=None):
        super().__init__(parent)
        self.step_type = step_type
        self.params = params or {}
        self.setup_ui()
        self.paramsChanged.connect(self._update_ui)

    def setup_ui(self):
        # Clear any existing layout
        if self.layout():
            QWidget().setLayout(self.layout())
        
        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        
        # Set minimum height for the step - reduced for smaller text
        self.setMinimumHeight(40)
        
        # Get color for this step type
        step_color = STEP_COLORS.get(self.step_type, "#e0e0e0")
        
        # Set styling for the step frame with colored border and left accent
        self.setStyleSheet(f"""
            QFrame {{
                background-color: #2d2d2d;
                border: 1px solid #3d3d3d;
                border-left: 5px solid {step_color};
                border-radius: 4px;
                padding: 4px;
            }}
            QLabel {{
                color: #e0e0e0;
                font-size: 9pt;
            }}
            QPushButton {{
                background-color: #424242;
                color: #e0e0e0;
                border: 1px solid #555555;
                font-size: 9pt;
                padding: 3px 6px;
                border-radius: 3px;
                min-width: 0px;
            }}
            QPushButton:hover {{
                background-color: #4d4d4d;
            }}
        """)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(6)
        
        # Reorder handle
        handle_label = QLabel("⋮⋮")
        handle_label.setStyleSheet("color: #666666; font-size: 10pt; font-weight: bold;")
        handle_label.setCursor(Qt.CursorShape.SizeVerCursor)
        layout.addWidget(handle_label)
        
        # Step type icon/emoji
        icon_map = {
            StepType.MOUSE_CLICK: "🖱️",
            StepType.KEYBOARD_TYPE: "⌨️",
            StepType.KEYBOARD_SPECIAL: "🔣",
            StepType.WAIT: "⏱️",
            StepType.LOOP_START: "🔄",
            StepType.LOOP_END: "↩️"
        }
        icon = icon_map.get(self.step_type, "❓")
        icon_label = QLabel(icon)
        layout.addWidget(icon_label)
        
        # Step name/description with colored type
        info_layout = QVBoxLayout()
        info_layout.setSpacing(1)
        
        self.name_label = QLabel(self.params.get("name", ""))
        self.name_label.setStyleSheet("font-weight: bold;")
        
        # Colored type label
        self.type_label = QLabel(self.step_type)
        self.type_label.setStyleSheet(f"color: {step_color}; font-size: 8pt;")
        
        info_layout.addWidget(self.name_label)
        info_layout.addWidget(self.type_label)
        layout.addLayout(info_layout, stretch=1)
        
        # Detail labels based on type
        if self.step_type == StepType.MOUSE_CLICK:
            coords = f"({self.params.get('x', '?')}, {self.params.get('y', '?')})"
            self.coord_label = QLabel(coords)
            self.coord_label.setStyleSheet("color: #aaaaaa; font-size: 8pt;")
            layout.addWidget(self.coord_label)

        elif self.step_type == StepType.WAIT:
            duration = f"{self.params.get('duration', 1)}s"
            duration_label = QLabel(duration)
            duration_label.setStyleSheet(f"color: {step_color}; font-weight: bold;")
            layout.addWidget(duration_label)
            
        elif self.step_type == StepType.LOOP_START:
            iters = f"x{self.params.get('iterations', 1)}"
            iter_label = QLabel(iters)
            iter_label.setStyleSheet(f"color: {step_color}; font-weight: bold;")
            layout.addWidget(iter_label)
            
        # Add spacer to push buttons to the right
        layout.addStretch(1)

        # Edit button with enhanced emoji
        edit_btn = QPushButton("✏️ Edit")
        edit_btn.clicked.connect(self.edit_step)
        edit_btn.setStyleSheet("""
            QPushButton {
                color: #e0e0e0;
                background-color: #3a3a3a;
                border: 1px solid #4d4d4d;
                font-size: 9pt;
            }
            QPushButton:hover {
                background-color: #4d4d4d;
            }
        """)
        layout.addWidget(edit_btn)

        # Delete button with enhanced emoji
        delete_btn = QPushButton("🗑️ Delete")
        delete_btn.clicked.connect(self.delete_step)
        delete_btn.setStyleSheet("""
            QPushButton {
                color: #e0e0e0;
                background-color: #3a3a3a;
                border: 1px solid #4d4d4d;
                font-size: 9pt;
            }
            QPushButton:hover {
                background-color: #4d4d4d;
                color: #ff8888;
            }
        """)
        layout.addWidget(delete_btn)

    def _update_ui(self):
        """Update UI elements with current parameters"""
        # Ensure we're on the main thread
        if not self.thread() == QApplication.instance().thread():
            return
            
        self.name_label.setText(self.params.get("name", ""))
        if self.step_type == StepType.MOUSE_CLICK and hasattr(self, 'coord_label'):
            coords = f"({self.params.get('x', '?')}, {self.params.get('y', '?')})"
            self.coord_label.setText(coords)

    def update_params(self, new_params):
        """Thread-safe method to update parameters"""
        self.params.update(new_params)
        # Emit signal to trigger UI update in main thread
        self.paramsChanged.emit(self.params.copy())

    def edit_step(self):
        if self.step_type in STEP_DIALOGS:
            # Pass the current parameters to the dialog
            dialog = STEP_DIALOGS[self.step_type](params=self.params.copy())
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.update_params(dialog.get_params())

    def delete_step(self):
        """Delete this step from the list"""
        # Find the QListWidgetItem that contains this widget
        list_widget = self.parent().parent()  # Get the QListWidget
        for i in range(list_widget.count()):
            item = list_widget.item(i)
            if list_widget.itemWidget(item) == self:
                # Remove both the widget and the item
                list_widget.takeItem(i)
                self.deleteLater()
                break

    def get_data(self):
        """Get step data for saving"""
        return {
            "type": self.step_type,
            "params": self.params
        }

class AddStepDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add Automation Step")
        self.setModal(True)
        self.selected_type = None
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Set dialog styling for dark theme
        self.setStyleSheet("""
            QDialog {
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QLabel {
                color: #e0e0e0;
                font-size: 10pt;
            }
            QLabel[text^="🔍"] {
                font-size: 11pt;
                color: #88CCFF;
            }
            QPushButton {
                background-color: #424242;
                color: #e0e0e0;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #4e4e4e;
            }
            QPushButton:pressed {
                background-color: #353535;
            }
            QPushButton[text^="✅"] {
                background-color: #2a5d3c;
            }
            QPushButton[text^="✅"]:hover {
                background-color: #336d49;
            }
            QPushButton[text^="❌"] {
                background-color: #5d2a2a;
            }
            QPushButton[text^="❌"]:hover {
                background-color: #6d3333;
            }
            QComboBox {
                background-color: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 2px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 15px;
                border-left-width: 1px;
                border-left-color: #555555;
                border-left-style: solid;
            }
            QComboBox QAbstractItemView {
                background-color: #3d3d3d;
                color: #e0e0e0;
                selection-background-color: #4e4e4e;
            }
        """)

        # Step type selection
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            StepType.MOUSE_CLICK,
            StepType.KEYBOARD_TYPE,
            StepType.KEYBOARD_SPECIAL,
            StepType.WAIT,
            StepType.LOOP_START,
            StepType.LOOP_END
        ])
        
        # Create label with enhanced emoji
        step_type_label = QLabel("🔍 Step Type:")
        step_type_label.setStyleSheet("font-size: 11pt; font-weight: bold; color: #88CCFF;")
        
        layout.addWidget(step_type_label)
        layout.addWidget(self.type_combo)

        # Buttons
        button_layout = QHBoxLayout()
        ok_btn = QPushButton("✅ OK")
        cancel_btn = QPushButton("❌ Cancel")
        ok_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(ok_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

    def get_selected_type(self):
        return self.type_combo.currentText()

class ExecutorThread(QThread):
    """Thread for executing automation workflows"""
    def __init__(self, executor, steps):
        super().__init__()
        self.executor = executor
        self.steps = steps
        self.loop_count = 1  # Default to 1 loop

    def start(self, loop_count=1):
        """Start the thread with specified loop count"""
        self.loop_count = loop_count
        super().start()

    def run(self):
        """Execute the workflow in a separate thread"""
        self.executor.execute_workflow(self.steps, self.loop_count)

class AutomationToolGUI(QMainWindow):
    emergency_stop_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automation Tool")
        self.setMinimumSize(800, 600)  # Reduced window size
        self.setWindowState(Qt.WindowState.WindowMaximized)  # Start maximized
        
        # Set application icon
        icon_path = os.path.join(WORKSPACE_DIR, "icon.ico")
        fallback_icon_path = os.path.join(WORKSPACE_DIR, "icon.png")
        
        app_icon = QIcon()
        if os.path.exists(icon_path):
            app_icon.addFile(icon_path)
        elif os.path.exists(fallback_icon_path):
            app_icon.addFile(fallback_icon_path)
        
        if not app_icon.isNull():
            self.setWindowIcon(app_icon)
            # Set taskbar icon
            if hasattr(sys, 'getwindowsversion'):  # Windows only
                import ctypes
                myappid = 'mycompany.erpautomation.1.0'  # Arbitrary string
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
        # Set modern style
        self.setStyleSheet(STYLE_SHEET)
        
        # Initialize components
        self.recorder = ActionRecorder()

        self.executor = WorkflowExecutor()
        self.executor_thread = None
        self.current_workflow_path = None
        self.coordinate_recording = False
        self.progress_dialog = None  # Initialize as None
        
        self.setup_ui()
        
        # Connect signals
        self.setup_signals()
        self.setup_menu()
        
        # Connect emergency stop signal using a lambda that calls the method
        # This handles the signal connection properly
        self.emergency_stop_signal.connect(self.emergency_stop)
        
        # Setup fail-safe shortcut
        # Use simple lambda to emit signal from background thread
        keyboard.add_hotkey('ctrl+alt+x', self._on_failsafe_triggered, suppress=True)

    def _on_failsafe_triggered(self):
        """Handle failsafe hotkey press"""
        logging.critical("CRITICAL: Failsafe hotkey (Ctrl+Alt+X) triggered by user!")
        print("CRITICAL: Failsafe hotkey triggered!")
        # Use invokeMethod to ensure thread safety if called from keyboard thread
        self.emergency_stop_signal.emit()

    def setup_signals(self):
        # Recorder signals
        self.recorder.action_recorded.connect(self.on_action_recorded)
        self.recorder.recording_stopped.connect(self.on_recording_stopped)
        self.recorder.coordinate_recorded.connect(self.on_coordinate_recorded)
        self.recorder.recording_armed.connect(self.on_recording_armed_changed)
        
        # Executor signals
        self.executor.step_started.connect(self.on_step_started)
        self.executor.step_completed.connect(self.on_step_completed)
        self.executor.step_error.connect(self.on_step_error)
        self.executor.workflow_completed.connect(self.on_workflow_completed)
        self.executor.debug_info.connect(self.on_debug_info)  # Connect debug signal

        


    def setup_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(8)  # Reduced spacing
        main_layout.setContentsMargins(8, 8, 8, 8)  # Reduced margins

        # Create left panel (Steps List)
        left_panel = self.create_left_panel()
        
        # Create right panel (Configuration)
        right_panel = self.create_right_panel()
        
        # Add panels to main layout
        main_layout.addWidget(left_panel, 2)  # Increased ratio for left panel
        main_layout.addWidget(right_panel, 1)  # Decreased ratio for right panel

    def create_left_panel(self):
        """Create the left panel with workflow steps"""
        left_panel = QWidget()
        layout = QVBoxLayout(left_panel)

        # Settings group
        settings_group = QGroupBox("⚙️ Settings")
        settings_layout = QVBoxLayout()

        # Debug mode checkbox with icon
        self.debug_mode = QCheckBox("🔍 Debug Mode")
        self.debug_mode.setChecked(True)
        self.debug_mode.toggled.connect(self.toggle_debug_mode)
        settings_layout.addWidget(self.debug_mode)

        # Add settings to group
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        # Steps list
        steps_group = QGroupBox("📋 Workflow Steps")
        steps_layout = QVBoxLayout()
        
        self.steps_list = QListWidget()
        self.steps_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.steps_list.model().rowsMoved.connect(self._on_steps_reordered)
        steps_layout.addWidget(self.steps_list)
        
        steps_group.setLayout(steps_layout)
        layout.addWidget(steps_group)

        return left_panel

    def create_right_panel(self):
        """Create the right panel with debug output"""
        right_panel = QWidget()
        layout = QVBoxLayout(right_panel)

        # Create tabs
        tabs = QTabWidget()
        
        # Debug tab
        debug_tab = QWidget()
        debug_layout = QVBoxLayout(debug_tab)
        
        # Debug controls
        debug_controls = QHBoxLayout()
        
        # Clear log button
        clear_btn = QPushButton("🧹 Clear Log")
        clear_btn.clicked.connect(self.clear_debug_log)
        debug_controls.addWidget(clear_btn)
        
        # Open debug folder button
        open_folder_btn = QPushButton("📁 Open Debug Folder")
        open_folder_btn.clicked.connect(self.open_debug_directory)
        debug_controls.addWidget(open_folder_btn)
        
        debug_layout.addLayout(debug_controls)
        
        # Debug output
        self.debug_text = QTextEdit()
        self.debug_text.setReadOnly(True)
        self.debug_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #e0e0e0;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 10pt;
            }
        """)
        debug_layout.addWidget(self.debug_text)
        
        tabs.addTab(debug_tab, "🔍 Debug")
        

        
        layout.addWidget(tabs)
        
        # Workflow controls
        controls_group = QGroupBox("Workflow Controls")
        controls_layout = QVBoxLayout()
        
        # Add step and record buttons
        top_controls = QHBoxLayout()
        
        # Add step button
        add_step_btn = QPushButton("➕ Add Step")
        add_step_btn.clicked.connect(self.add_step)
        top_controls.addWidget(add_step_btn)
        
        # Record button
        record_btn = QPushButton("⏺️ Record")
        record_btn.setCheckable(True)
        record_btn.clicked.connect(self.toggle_coordinate_recording)
        top_controls.addWidget(record_btn)
        self.record_btn = record_btn  # Store as instance variable
        
        controls_layout.addLayout(top_controls)
        
        # Loop count control
        loop_layout = QHBoxLayout()
        loop_layout.addWidget(QLabel("Loop Workflow:"))
        self.loop_count = QSpinBox()
        self.loop_count.setRange(1, 999)
        self.loop_count.setValue(1)
        self.loop_count.setToolTip("Number of times to repeat the workflow")
        loop_layout.addWidget(self.loop_count)
        controls_layout.addLayout(loop_layout)
        
        # Run controls
        run_layout = QHBoxLayout()
        
        # Run button
        self.run_btn = QPushButton("▶️ Run")
        self.run_btn.clicked.connect(self.run_workflow)
        self.run_btn.setShortcut("F5")
        self.run_btn.setToolTip("Run Workflow (F5)")
        run_layout.addWidget(self.run_btn)
        
        controls_layout.addLayout(run_layout)
        
        controls_layout.addLayout(run_layout)
        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)
        
        return right_panel

    def setup_menu(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("📂 File")
        
        new_action = file_menu.addAction("🆕 New")
        new_action.triggered.connect(self.new_automation)
        
        save_action = file_menu.addAction("💾 Save")
        save_action.triggered.connect(self.save_automation)
        
        load_action = file_menu.addAction("📂 Load")
        load_action.triggered.connect(self.load_automation)
        
        # Add separator before Quit
        file_menu.addSeparator()
        
        # Add Quit action
        quit_action = file_menu.addAction("❌ Quit")
        quit_action.triggered.connect(self.close)
        
        # Tools menu
        tools_menu = menubar.addMenu("🔧 Tools")
        
        self.record_action = tools_menu.addAction("⏺️ Start Recording")
        self.record_action.triggered.connect(self.toggle_recording)
        


        
        # Help menu
        help_menu = menubar.addMenu("❓ Help")
        
        general_help_action = help_menu.addAction("📚 User Guide")
        general_help_action.triggered.connect(self.show_general_help)
        
        steps_help_action = help_menu.addAction("🔍 Automation Steps")
        steps_help_action.triggered.connect(self.show_steps_help)
        
        recording_help_action = help_menu.addAction("⏺️ Recording Guide")
        recording_help_action.triggered.connect(self.show_recording_help)
        


        
        help_menu.addSeparator()
        
        about_action = help_menu.addAction("ℹ️ About")
        about_action.triggered.connect(self.show_about)

    def add_step(self):
        dialog = AddStepDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            step_type = dialog.get_selected_type()
            if step_type in STEP_DIALOGS:
                config_dialog = STEP_DIALOGS[step_type]()
                if config_dialog.exec() == QDialog.DialogCode.Accepted:
                    params = config_dialog.get_params()
                    step = AutomationStep(step_type, params)
                    item = QListWidgetItem()
                    item.setSizeHint(step.sizeHint())
                    self.steps_list.addItem(item)
                    self.steps_list.setItemWidget(item, step)

    def new_automation(self):
        self.steps_list.clear()
        self.current_workflow_path = None

    def save_automation(self):
        """Save the current automation workflow"""
        if len(self.steps_list) == 0:
            QMessageBox.warning(self, "No Steps", "No steps to save. Add some automation steps first.")
            return
        
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Save Automation Workflow",
            os.path.join(AUTOMATIONS_DIR, "automation.json"),
            "JSON Files (*.json)"
        )
        
        if file_name:
            # Ensure file has .json extension
            if not file_name.lower().endswith('.json'):
                file_name += '.json'
                
            try:
                # Ensure directory exists
                os.makedirs(os.path.dirname(file_name), exist_ok=True)
                
                # Collect data from steps
                workflow = {
                    "version": "1.0",
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "debug_mode": self.debug_mode.isChecked(),
                    "steps": []
                }
                
                for i in range(self.steps_list.count()):
                    item = self.steps_list.item(i)
                    step_widget = self.steps_list.itemWidget(item)
                    
                    if step_widget:
                        workflow["steps"].append({
                            "type": step_widget.step_type,
                            "params": step_widget.params
                        })
                
                with open(file_name, 'w') as f:
                    json.dump(workflow, f, indent=4)
                
                self.current_workflow_path = file_name
                QMessageBox.information(self, "Success", "Automation saved successfully!")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save automation: {str(e)}")

    def load_automation(self):
        """Load an automation workflow"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Load Automation Workflow",
            AUTOMATIONS_DIR,
            "JSON Files (*.json)"
        )
        
        if file_name:
            try:
                with open(file_name, 'r') as f:
                    workflow = json.load(f)
                
                # Clear current steps
                self.steps_list.clear()
                
                # Load debug mode
                self.debug_mode.setChecked(workflow.get("debug_mode", True))
                
                # Load steps
                for step_data in workflow.get("steps", []):
                    step_type = step_data["type"]
                    params = step_data["params"]
                    
                    # Create step widget
                    step = AutomationStep(step_type, params)
                    
                    # Add to list
                    item = QListWidgetItem()
                    item.setSizeHint(step.sizeHint())
                    self.steps_list.addItem(item)
                    self.steps_list.setItemWidget(item, step)
                
                self.current_workflow_path = file_name
                
                # Get just the filename without path for the message
                file_basename = os.path.basename(file_name)
                QMessageBox.information(self, "Success", f"Automation '{file_basename}' loaded successfully!")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load automation: {str(e)}")

    def toggle_recording(self):
        """Handle recording button in menu"""
        if not self.recorder.is_recording:
            # Show countdown dialog
            countdown_dialog = QDialog(self)
            countdown_dialog.setWindowTitle("Recording will start in...")
            countdown_dialog.setFixedSize(300, 150)
            layout = QVBoxLayout(countdown_dialog)
            
            # Countdown label with large font
            countdown_label = QLabel("10")
            countdown_label.setStyleSheet("font-size: 48pt; color: #FF4444; font-weight: bold;")
            countdown_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(countdown_label)
            
            # Instructions label
            instructions = QLabel("Press ESC anytime to stop recording")
            instructions.setStyleSheet("color: #e0e0e0;")
            instructions.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.addWidget(instructions)
            
            countdown_dialog.show()
            
            # Countdown timer
            for i in range(10, 0, -1):
                countdown_label.setText(str(i))
                QApplication.processEvents()
                # Play notification sound (Windows beep)
                if sys.platform == 'win32':
                    import winsound
                    winsound.Beep(1000, 100)  # 1000Hz for 100ms
                time.sleep(1)
            
            countdown_dialog.close()
            
            # Play start recording sound (longer beep)
            if sys.platform == 'win32':
                winsound.Beep(2000, 500)  # 2000Hz for 500ms
            
            # Start recording and setup stop shortcut
            self.recorder.start_recording()
            self.record_action.setText("⏹️ Stop Recording (ESC)")
            keyboard.add_hotkey('esc', self.stop_recording, suppress=True)
        else:
            self.stop_recording()

    def stop_recording(self):
        """Stop recording and cleanup"""
        if self.recorder.is_recording:
            self.recorder.stop_recording()
            self.record_action.setText("⏺️ Start Recording")
            try:
                keyboard.remove_hotkey('esc')  # Remove the ESC shortcut
            except Exception:
                pass
            # Play stop recording sound
            if sys.platform == 'win32':
                import winsound
                winsound.Beep(500, 500)  # 500Hz for 500ms

    def on_action_recorded(self, step_type, params):
        step = AutomationStep(step_type, params)
        item = QListWidgetItem()
        item.setSizeHint(step.sizeHint())
        self.steps_list.addItem(item)
        self.steps_list.setItemWidget(item, step)

    def on_recording_stopped(self):
        self.record_action.setText("⏺️ Start Recording")




    def run_workflow(self):
        """Run the current workflow"""
        if not self.steps_list.count():
            QMessageBox.warning(self, "No Steps", "Please add some steps to the workflow first.")
            return

        # Create progress dialog
        self.progress_dialog = QProgressDialog("Preparing to run workflow...", "Cancel", 0, self.steps_list.count(), self)
        self.progress_dialog.setWindowTitle("Running Workflow")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setAutoReset(False)
        
        # Get all steps
        steps = []
        for i in range(self.steps_list.count()):
            item = self.steps_list.item(i)
            widget = self.steps_list.itemWidget(item)
            if widget:
                step_data = widget.get_data()
                steps.append(step_data)

        if not steps:
            self.progress_dialog.close()
            self.progress_dialog = None
            return

        # Create and start executor thread
        self.executor = WorkflowExecutor()
        self.executor_thread = ExecutorThread(self.executor, steps)
        
        # Connect signals
        self.executor.step_started.connect(self.on_step_started)
        self.executor.step_completed.connect(self.on_step_completed)
        self.executor.step_error.connect(self.on_step_error)
        self.executor.workflow_completed.connect(self.on_workflow_completed)
        self.executor.debug_info.connect(self.on_debug_info)
        self.executor_thread.finished.connect(self.on_executor_thread_finished)
        
        # Set debug mode
        self.executor.debug_mode = self.debug_mode.isChecked()
        
        # Get loop count
        loop_count = self.loop_count.value()
        
        # Start execution
        self.executor_thread.start(loop_count)
        
        # Update UI state
        # self.pause_btn.setEnabled(True)
        # self.pause_btn.setChecked(False)
        # self.pause_btn.setText("⏸️ Pause")

    def on_step_started(self, step_index, step_type):
        """Handle step started signal"""
        if self.progress_dialog:
            self.progress_dialog.setLabelText(f"Executing step {step_index + 1}: {step_type}")
            self.progress_dialog.setValue(step_index)

    def on_step_completed(self, step_index):
        """Handle step completed signal"""
        if self.progress_dialog:
            self.progress_dialog.setValue(step_index + 1)

    def on_step_error(self, step_index, error):
        """Handle step error signal"""
        QMessageBox.critical(
            self,
            "Error",
            f"Error executing step {step_index + 1}:\n{error}"
        )
        if self.progress_dialog:
            self.progress_dialog.cancel()
            self.progress_dialog = None
        
        # Reset UI state
        # self.pause_btn.setEnabled(False)
        # self.stop_btn.setEnabled(False)
        # self.resume_btn.setEnabled(False)

    def on_workflow_completed(self):
        """Handle workflow completed signal"""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        # self.pause_btn.setEnabled(False)
        # self.stop_btn.setEnabled(False)
        # self.resume_btn.setEnabled(False)

    def on_executor_thread_finished(self):
        """Handle executor thread finished signal"""
        if self.progress_dialog:
            self.progress_dialog.close()
            self.progress_dialog = None
        # self.pause_btn.setEnabled(False)
        # self.stop_btn.setEnabled(False)
        # self.resume_btn.setEnabled(False)



    def toggle_debug_mode(self, state):
        """Toggle debug mode"""
        self.executor.debug_mode = bool(state)
        self.debug_text.append(f"Debug mode {'enabled' if state else 'disabled'}")

    def clear_debug_log(self):
        """Clear the debug log"""
        self.debug_text.clear()

    def open_debug_directory(self):
        """Open the debug directory in file explorer"""
        try:
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__)), "debug")
            if os.path.exists(debug_dir):
                # Use os.startfile for Windows
                if sys.platform == 'win32':
                    os.startfile(debug_dir)
                # Use xdg-open for Linux
                elif sys.platform == 'linux':
                    os.system(f'xdg-open "{debug_dir}"')
                # Use open for macOS
                elif sys.platform == 'darwin':
                    os.system(f'open "{debug_dir}"')
            else:
                QMessageBox.warning(self, "Directory Not Found", 
                    "Debug directory does not exist. It will be created when you run an automation.")
        except Exception as e:
            QMessageBox.warning(self, "Error", 
                f"Could not open debug directory: {str(e)}")



    def stop_workflow(self):
        """Stop the current workflow execution"""
        if self.executor_thread and self.executor_thread.isRunning():
            self.executor.stop()
            
            # Non-blocking wait loop to keep UI responsive
            # We wait up to 2 seconds for the thread to stop gracefully
            for _ in range(20):
                if self.executor_thread.wait(100):
                    break
                QApplication.processEvents()
            
            if self.executor_thread.isRunning():
                self.debug_text.append("Force stopping workflow...")
                self.executor_thread.terminate()
                self.executor_thread.wait()
            
            self.debug_text.append("Workflow stopped")
            
            # Clean up progress dialog if it exists
            if self.progress_dialog:
                self.progress_dialog.close()
                self.progress_dialog = None
            
            # Reset UI state
            # self.pause_btn.setEnabled(False)
            # self.pause_btn.setChecked(False)
            # self.pause_btn.setText("⏸️ Pause")

    def on_debug_info(self, info):
        """Handle debug information from executor"""
        # Format the message with timestamp and appropriate color coding
        timestamp = time.strftime("%H:%M:%S")
        
        # Color coding based on message content
        if "✓" in info:  # Success messages
            formatted_msg = f'<span style="color: #00FF00;">[{timestamp}] {info}</span>'
        elif "❌" in info:  # Error messages
            formatted_msg = f'<span style="color: #FF0000;">[{timestamp}] {info}</span>'
        elif "===" in info:  # Step headers
            formatted_msg = f'<span style="color: #00FFFF;">[{timestamp}] {info}</span>'
        elif "Warning" in info:  # Warnings
            formatted_msg = f'<span style="color: #FFA500;">[{timestamp}] {info}</span>'
        else:  # Regular debug messages
            formatted_msg = f'<span style="color: #FFFFFF;">[{timestamp}] {info}</span>'
        
        # Add the message to the debug text area
        self.debug_text.append(formatted_msg)
        
        # Scroll to the bottom to show the latest message
        self.debug_text.verticalScrollBar().setValue(
            self.debug_text.verticalScrollBar().maximum()
        )
        
        # Update the application to ensure the UI remains responsive
        QApplication.processEvents()

    def closeEvent(self, event):
        """Handle application close event"""
        # Remove fail-safe shortcut
        try:
            keyboard.remove_hotkey('ctrl+alt+x')
        except Exception:
            pass
        self.stop_workflow()
        event.accept()

    def toggle_coordinate_recording(self):
        """Toggle coordinate recording mode"""
        if not self.coordinate_recording:
            # Show instructions
            QMessageBox.information(
                self,
                "Record Coordinates",
                "1. Click 'OK' to start recording\n"
                "2. Press Ctrl+Shift+F8 to capture coordinates\n"
                "3. Press ESC to cancel"
            )
            
            # Start recording
            self.coordinate_recording = True
            self.recorder.start_coordinate_recording()
            self.record_btn.setStyleSheet("background-color: #FFA500;")  # Orange for recording state
            self.statusBar().showMessage("Recording armed - Press Ctrl+Shift+F8 when ready to capture coordinates")
        else:
            # Stop recording
            self.coordinate_recording = False
            self.recorder.stop_coordinate_recording()
            self.record_btn.setStyleSheet("")  # Reset to default style
            self.statusBar().showMessage("Recording stopped")
            self.record_btn.setChecked(False)

    def on_recording_armed_changed(self, armed):
        """Handle recording armed state change"""
        # Find and update the record button
        for widget in self.findChildren(QPushButton):
            if widget.text() == "⏺️ Record":
                if armed:
                    widget.setStyleSheet("background-color: #FFA500;")  # Orange for armed state
                    self.statusBar().showMessage("Recording armed - Press Ctrl+Shift+F8 when ready to capture coordinates")
                else:
                    widget.setStyleSheet("")  # Reset to default style
                    self.statusBar().clearMessage()
                break

    def on_coordinate_recorded(self, x, y):
        """Handle recorded coordinates"""
        if self.coordinate_recording:
            # Update the currently selected step with new coordinates
            current_item = self.steps_list.currentItem()
            if current_item:
                step_widget = self.steps_list.itemWidget(current_item)
                if step_widget:
                    if step_widget.step_type == StepType.MOUSE_CLICK:
                        # Handle mouse click coordinates
                        old_coords = (step_widget.params.get("x"), step_widget.params.get("y"))
                        step_widget.update_params({
                            "click_type": "coordinates",
                            "x": x,
                            "y": y
                        })
                        message = (
                            f"Coordinates recorded for step: {step_widget.params.get('name', 'Unnamed Step')}\n"
                            f"New coordinates: ({x}, {y})\n"
                        )
                        if old_coords[0] is not None:
                            message += f"Previous coordinates: ({old_coords[0]}, {old_coords[1]})"

                    
                    QMessageBox.information(
                        self,
                        "Coordinates Recorded",
                        message
                    )
                    
                    # Remind to save if workflow has changed
                    if self.current_workflow_path:
                        self.statusBar().showMessage("Remember to save your workflow to keep these coordinates!", 5000)
            
            # Reset recording state
            self.coordinate_recording = False

    def _cleanup_invalid_items(self):
        """Clean up any invalid or empty items in the list"""
        items_to_remove = []
        for i in range(self.steps_list.count()):
            item = self.steps_list.item(i)
            if not item or not self.steps_list.itemWidget(item):
                items_to_remove.append(i)
        
        # Remove items in reverse order to maintain correct indices
        for i in reversed(items_to_remove):
            self.steps_list.takeItem(i)

    def _on_steps_reordered(self, parent, start, end, destination, row):
        """Handle steps reordering after drag and drop"""
        # Clean up any invalid items that might have been created during drag and drop
        self._cleanup_invalid_items()
        
        # Ensure all remaining items have valid widgets
        for i in range(self.steps_list.count()):
            item = self.steps_list.item(i)
            if item and not self.steps_list.itemWidget(item):
                # If we find an item without a widget, remove it
                self.steps_list.takeItem(i)

    def show_general_help(self):
        """Show general help dialog"""
        dialog = GeneralHelpDialog(self)
        dialog.exec()
        
    def show_steps_help(self):
        """Show steps help dialog"""
        dialog = StepsHelpDialog(self)
        dialog.exec()
        
    def show_recording_help(self):
        """Show recording help dialog"""
        dialog = RecordingHelpDialog(self)
        dialog.exec()
        


        
    def show_about(self):
        """Show about dialog"""
        dialog = AboutDialog(self)
        dialog.exec()

    def emergency_stop(self):
        """Emergency stop for the automation workflow"""
        try:
            if hasattr(self, 'executor') and self.executor:
                self.executor.stop()
            
            if hasattr(self, 'executor_thread') and self.executor_thread and self.executor_thread.isRunning():
                self.executor_thread.quit()
                # Give it a short time to quit gracefully
                if not self.executor_thread.wait(1000):  # Wait up to 1 second
                    self.executor_thread.terminate()  # Force quit if necessary
            
            # Reset UI state
            self.run_btn.setEnabled(True)
            # self.pause_btn.setEnabled(False)
            # self.stop_btn.setEnabled(False)
            
            # Clear any highlighted steps
            for i in range(self.steps_list.count()):
                item = self.steps_list.item(i)
                widget = self.steps_list.itemWidget(item)
                if widget:
                    widget.setStyleSheet("")
            
            QMessageBox.information(self, "Automation Stopped", 
                "The automation has been stopped.")
            
        except Exception as e:
            QMessageBox.warning(self, "Error", 
                f"Error during emergency stop: {str(e)}")
            # Ensure UI is reset even if there's an error
            self.run_btn.setEnabled(True)
            # self.pause_btn.setEnabled(False)
            # self.stop_btn.setEnabled(False)

# Add the help dialog base class and specialized dialogs
class HelpDialog(QDialog):
    """Base class for help dialogs"""
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumSize(700, 600)
        self.setup_ui()
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Help text
        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        self.help_text.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d;
                color: #e0e0e0;
                font-family: system;
                font-size: 10pt;
                padding: 10px;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
            }
        """)
        layout.addWidget(self.help_text)
        
        # Close button
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)
        
    def set_help_content(self, html_content):
        self.help_text.setHtml(html_content)

class GeneralHelpDialog(HelpDialog):
    """Dialog for general user guide"""
    def __init__(self, parent=None):
        super().__init__("User Guide", parent)
        
        content = """
        <h1>🤖 Automation Tool User Guide</h1>
        
        <h2>Introduction</h2>
        <p>The Automation Tool allows you to create, save, load, and execute automation workflows for repetitive tasks.</p>
        
        <h2>Safety Features</h2>
        <p><b>⚠️ Emergency Stop:</b> Press <span style="color: #FF4444; font-weight: bold;">Ctrl+Alt+X</span> at any time to immediately stop a running automation. This is the global fail-safe that overrides all operations.</p>
        
        <h2>Main Interface Components</h2>
        <ul>
            <li><b>Left Panel (Workflow):</b> Your list of automation steps.</li>
            <li><b>Right Panel (Controls):</b> Debug console and the Run execution control.</li>
        </ul>
        
        <h2>Creating a Workflow</h2>
        <ol>
            <li>Click <b>"➕ Add Step"</b> to choose an action.</li>
            <li>Configure parameters (coordinates, text, duration).</li>
            <li>Use <b>"✏️ Edit"</b> to modify steps.</li>
            <li>Drag and drop to reorder steps.</li>
        </ol>
        
        <h2>Running a Workflow</h2>
        <ol>
            <li>Set the <b>Loop Workflow</b> counter (optional).</li>
            <li>Click <b>"▶️ Run"</b> or press <b>F5</b> to start.</li>
            <li>The automation will execute all steps in order.</li>
            <li>To stop execution, press <b>Ctrl+Alt+X</b>.</li>
        </ol>
        
        <h2>Recording Actions</h2>
        <ol>
            <li>Click <b>"⏺️ Tools > Start Recording"</b>.</li>
            <li>Perform your mouse clicks and typing.</li>
            <li>Press <b>ESC</b> to stop recording.</li>
            <li>The actions will be converted into editable steps.</li>
        </ol>
        
        <h2>Debug Mode</h2>
        <ul>
            <li>Enable <b>Debug Mode</b> to see detailed execution logs.</li>
            <li>Logs (and screenshots if applicable) are saved in the debug folder.</li>
        </ul>
        
        <h2>Tips</h2>
        <ul>
            <li>Always test new workflows with a small loop count first.</li>
            <li>Use the fail-safe (<b>Ctrl+Alt+X</b>) if the mouse moves unexpectedly.</li>
            <li>Save your workflows (`.json`) regularly.</li>
        </ul>
        """
        
        self.set_help_content(content)

class StepsHelpDialog(HelpDialog):
    """Dialog for automation steps documentation"""
    def __init__(self, parent=None):
        super().__init__("Automation Steps Guide", parent)
        
        content = """
        <h1>📋 Automation Steps Documentation</h1>
        
        <h2>Mouse Click Step</h2>
        <p><b>Purpose:</b> Perform a mouse click at specific coordinates or on an image.</p>
        <p><b>Parameters:</b></p>
        <ul>
            <li><b>Click Type:</b> Choose between coordinate-based or image-based clicking</li>
            <li><b>Mouse Button:</b> Select left or right mouse button</li>
            <li><b>Coordinates:</b> Specify X and Y screen coordinates (for coordinate-based clicks)</li>
            <li><b>Image Path:</b> Path to the reference image (for image-based clicks)</li>
            <li><b>Confidence:</b> Matching threshold for image recognition (0.1-1.0)</li>
            <li><b>Duration:</b> How long the mouse movement takes (in seconds)</li>
            <li><b>Text Input After Click:</b> Optional text to type after clicking</li>
            <li><b>Delay Before Typing:</b> Wait time before typing (in seconds)</li>
            <li><b>Special Key:</b> Optional special key to press after typing (Enter, Tab, etc.)</li>
        </ul>
        <p><b>Example:</b> Click on the login button and enter username</p>
        <pre>
        Step Name: "Click Login Button and Enter Username"
        Click Type: Image-based
        Mouse Button: Left Click
        Image Path: "login_button.png"
        Confidence: 0.9
        Text Input After Click: Enabled
        Text to Type: "admin"
        Delay Before Typing: 0.5 seconds
        Special Key: Tab (to move to password field)
        </pre>
        
        <h2>Keyboard Type Step</h2>
        <p><b>Purpose:</b> Type text at the current cursor position with support for multiple text inputs.</p>
        <p><b>Parameters:</b></p>
        <ul>
            <li><b>Input Type:</b> Choose between single text or multiple text inputs</li>
            <li><b>Text Input:</b> The text to type (single input) or list of texts (multiple inputs)</li>
            <li><b>Special Key:</b> Optional special key to press after typing</li>
            <li><b>Delay:</b> Delay between keystrokes (in seconds)</li>
        </ul>
        <p><b>Example 1:</b> Type a single search query and press Enter</p>
        <pre>
        Step Name: "Search for Product"
        Input Type: Single
        Text Input: "wireless headphones"
        Special Key: Enter
        Delay: 0.1 seconds
        </pre>
        <p><b>Example 2:</b> Type multiple product codes in a loop</p>
        <pre>
        Step Name: "Enter Product Codes"
        Input Type: Multiple
        Text List: ["ABC123", "XYZ789", "DEF456"]
        Special Key: Enter
        Delay: 0.1 seconds
        </pre>
        <p>When using multiple text inputs, the step will:</p>
        <ul>
            <li>Cycle through the list of texts in order</li>
            <li>Work seamlessly with workflow loops</li>
            <li>Type each text in sequence without clearing previous text</li>
        </ul>
        
        <h2>Keyboard Special Step</h2>
        <p><b>Purpose:</b> Execute keyboard shortcuts or special key combinations.</p>
        <p><b>Parameters:</b></p>
        <ul>
            <li><b>Key:</b> The main key to press</li>
            <li><b>Modifiers:</b> Combination of Ctrl, Alt, Shift, and/or Windows keys</li>
        </ul>
        <p><b>Example:</b> Press Ctrl+C to copy selected content</p>
        <pre>
        Step Name: "Copy Selection"
        Key: C
        Modifiers: Ctrl
        </pre>
        <p><b>Example:</b> Press Alt+Tab to switch applications</p>
        <pre>
        Step Name: "Switch Applications"
        Key: Tab
        Modifiers: Alt
        </pre>
        

        """
        
        self.set_help_content(content)

class RecordingHelpDialog(HelpDialog):
    """Dialog for recording guide"""
    def __init__(self, parent=None):
        super().__init__("Recording Guide", parent)
        
        content = """
        <h1>⏺️ Recording Guide</h1>
        
        <h2>Action Recording</h2>
        
        <p><b>⚠ Important Note:</b> This feature is currently unstable. Some recorded steps may not work as expected and might require manual adjustment.</p>
        
        <p>Action recording allows you to create automation steps by performing the actions yourself:</p>

        <h3>How to Record Actions</h3>
        <ol>
            <li>Click the "⏺️ Start Recording" option in the Tools menu or click the "⏺️ Record" button</li>
            <li>Perform the actions you want to automate (mouse clicks, keyboard typing, etc.)</li>
            <li>Click "⏹️ Stop Recording" when finished</li>
        </ol>
        
        <h3>How to Record Actions</h3>
        <ol>
            <li>Click the "⏺️ Start Recording" option in the Tools menu or click the "⏺️ Record" button</li>
            <li>Perform the actions you want to automate (mouse clicks, keyboard typing, etc.)</li>
            <li>Click "⏹️ Stop Recording" when finished</li>
        </ol>
        
        <p><b>Recorded Actions Include:</b></p>
        <ul>
            <li>Mouse clicks (with coordinates)</li>
            <li>Keyboard typing (text input)</li>
            <li>Special key presses</li>
            <li>Mouse drag operations</li>
        </ul>
        
        <p><b>Example:</b> Recording a login sequence</p>
        <ol>
            <li>Start recording</li>
            <li>Click on the username field</li>
            <li>Type your username</li>
            <li>Press Tab to move to the password field</li>
            <li>Type your password</li>
            <li>Click the login button</li>
            <li>Stop recording</li>
        </ol>
        
        <h2>Coordinate Recording</h2>
        <p>Coordinate recording allows you to update the coordinates for an existing step:</p>
        
        <h3>How to Record Coordinates</h3>
        <ol>
            <li>Select the step you want to update in the workflow list</li>
            <li>Click the "⏺️ Record" button at the bottom of the steps list</li>
            <li>Position your mouse where you want to click</li>
            <li>Press Ctrl+Shift+F8 to capture the coordinates</li>
        </ol>
        
        <p><b>Example:</b> Updating click coordinates for a button that moved</p>
        <ol>
            <li>Select the "Click Login Button" step in your workflow</li>
            <li>Click "⏺️ Record"</li>
            <li>Position your mouse over the login button's new position</li>
            <li>Press Ctrl+Shift+F8</li>
            <li>The step will be updated with the new coordinates</li>
        </ol>
        
        <h2>Recording Tips</h2>
        <ul>
            <li><b>Move deliberately:</b> Make your mouse movements and clicks deliberate and clear</li>
            <li><b>Take your time:</b> Don't rush through the actions as the recorder needs time to process each action</li>
            <li><b>Review recorded steps:</b> After recording, review each step and edit if necessary</li>
            <li><b>Test the workflow:</b> Run the recorded workflow to ensure it works as expected</li>
            <li><b>Edit parameters manually:</b> You can fine-tune the recorded steps by clicking the "✏️ Edit" button</li>
        </ul>
        
        <h2>Keyboard Recording</h2>
        <p>The recorder captures text as you type it, but there are some special considerations:</p>
        <ul>
            <li>Special keys (Enter, Tab, etc.) are recorded as separate steps</li>
            <li>Key combinations (Ctrl+C, Alt+Tab, etc.) are recorded as Keyboard Special steps</li>
            <li>If you need to type slowly or with specific timing, edit the delay parameter after recording</li>
        </ul>
        """
        
        self.set_help_content(content)




class AboutDialog(QDialog):
    """Dialog showing information about the application"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Automation Tool")
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)  # Increase spacing between elements

        # Title and Version
        title_label = QLabel("Automation Tool")
        title_label.setStyleSheet("font-size: 24pt; font-weight: bold; color: #88CCFF;")
        layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignCenter)

        version_label = QLabel("Version 4.0")
        version_label.setStyleSheet("font-size: 14pt; color: #00FF00;")
        layout.addWidget(version_label, alignment=Qt.AlignmentFlag.AlignCenter)

        # New Features Section
        new_features_group = QGroupBox("What's New in Version 4.0")
        new_features_layout = QVBoxLayout()
        
        features_text = QLabel(
            "• Simplified Controls\n"
            "  - Single 'Run' button for cleaner experience\n"
            "  - Added Color Coding for different types of steps\n\n"
            "• Failsafe System\n"
            "  - New specific hotkey (Ctrl+Alt+X) for emergency stops\n\n"
            "• Visual Enhancements\n"
            "  - Color-coded automation steps\n"
            "  - Improved responsiveness"
        )
        features_text.setStyleSheet("color: #e0e0e0; font-size: 11pt;")
        new_features_layout.addWidget(features_text)
        new_features_group.setLayout(new_features_layout)
        layout.addWidget(new_features_group)

        # Description
        description = QLabel(
            "A powerful automation tool for creating and running custom workflows.\n"
            "Automate repetitive tasks with mouse clicks, keyboard input, and image recognition."
        )
        description.setWordWrap(True)
        description.setStyleSheet("color: #e0e0e0; font-size: 11pt;")
        layout.addWidget(description, alignment=Qt.AlignmentFlag.AlignCenter)

        # Copyright
        copyright_label = QLabel("© 2026 All rights reserved")
        copyright_label.setStyleSheet("color: #888888;")
        layout.addWidget(copyright_label, alignment=Qt.AlignmentFlag.AlignCenter)

        # Close Button
        close_btn = QPushButton("Close")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #424242;
                color: #e0e0e0;
                border: 1px solid #555555;
                border-radius: 3px;
                padding: 5px 20px;
                font-size: 11pt;
            }
            QPushButton:hover {
                background-color: #4e4e4e;
            }
            QPushButton:pressed {
                background-color: #353535;
            }
        """)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignCenter)

        # Set dialog style
        self.setStyleSheet("""
            QDialog {
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QGroupBox {
                border: 2px solid #3d3d3d;
                border-radius: 5px;
                margin-top: 1em;
                padding: 10px;
                color: #88CCFF;
                font-size: 12pt;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
        """)

        # Set fixed size for the dialog
        self.setFixedSize(500, 600)

if __name__ == '__main__':
    # Suppress Qt DPI warnings
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_LOGGING_RULES"] = "qt.qpa.window=false"  # Suppress DPI warning messages
    
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    window = AutomationToolGUI()
    window.show()
    sys.exit(app.exec())
