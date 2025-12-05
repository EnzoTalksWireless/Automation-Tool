import threading
import time
import pyautogui
import keyboard
import mouse
from PyQt6.QtCore import QObject, pyqtSignal
from automation_steps import StepType

class ActionRecorder(QObject):
    """Records user actions for automation"""
    action_recorded = pyqtSignal(str, dict)  # Signal emitted when an action is recorded
    recording_stopped = pyqtSignal()  # Signal emitted when recording stops
    coordinate_recorded = pyqtSignal(int, int)  # Signal emitted when coordinates are recorded
    recording_armed = pyqtSignal(bool)  # Signal emitted when recording is armed/disarmed

    def __init__(self):
        super().__init__()
        self.is_recording = False
        self.is_coordinate_recording = False
        self.is_coordinate_armed = False
        self.record_thread = None
        self.last_mouse_pos = None
        self.last_click_time = 0
        self.key_buffer = []
        self.key_buffer_time = 0

    def start_recording(self):
        """Start recording user actions"""
        if not self.is_recording:
            self.is_recording = True
            self.record_thread = threading.Thread(target=self._record_loop)
            self.record_thread.daemon = True
            self.record_thread.start()

    def stop_recording(self):
        """Stop recording user actions"""
        self.is_recording = False
        if self.record_thread:
            self.record_thread.join()
        self.recording_stopped.emit()

    def start_coordinate_recording(self):
        """Arm the coordinate recording"""
        if not self.is_coordinate_armed:
            self.is_coordinate_armed = True
            keyboard.on_press(lambda e: self._start_actual_recording(e) 
                            if e.name == 'f8' and keyboard.is_pressed('ctrl') and keyboard.is_pressed('shift')
                            else None)
            self.recording_armed.emit(True)

    def stop_coordinate_recording(self):
        """Stop recording coordinates"""
        if self.is_coordinate_armed or self.is_coordinate_recording:
            self.is_coordinate_armed = False
            self.is_coordinate_recording = False
            mouse.unhook_all()
            keyboard.unhook_all()
            self.recording_armed.emit(False)

    def _start_actual_recording(self, event):
        """Start the actual coordinate recording when Ctrl+Shift+F8 is pressed"""
        if self.is_coordinate_armed:
            self.stop_coordinate_recording()
            # Get current mouse position
            x, y = pyautogui.position()
            self.coordinate_recorded.emit(x, y)

    def _record_loop(self):
        """Main recording loop"""
        # Initialize listeners
        keyboard.hook(self._on_keyboard_event)
        mouse.hook(self._on_mouse_event)

        while self.is_recording:
            # Process any buffered keyboard input
            current_time = time.time()
            if self.key_buffer and (current_time - self.key_buffer_time) > 1.0:
                self._flush_key_buffer()
            time.sleep(0.01)

        # Clean up
        keyboard.unhook_all()
        mouse.unhook_all()

    def _on_keyboard_event(self, event):
        """Handle keyboard events"""
        if not self.is_recording or not event.event_type == keyboard.KEY_DOWN:
            return

        # Special keys handling
        special_keys = {
            'enter': 'Enter',
            'tab': 'Tab',
            'space': 'Space',
            'backspace': 'Backspace',
            'delete': 'Delete',
            'esc': 'Escape',
            'up': 'Up',
            'down': 'Down',
            'left': 'Left',
            'right': 'Right'
        }

        if event.name in special_keys:
            self._flush_key_buffer()  # Flush any pending regular text
            self.action_recorded.emit(StepType.KEYBOARD_SPECIAL, {
                "name": f"Press {special_keys[event.name]}",
                "special_key": special_keys[event.name]
            })
        else:
            # Buffer regular keystrokes
            if event.name.isalnum() or event.name in ['.', '/', '-', ' ']:
                self.key_buffer.append(event.name)
                self.key_buffer_time = time.time()

    def _on_mouse_event(self, event):
        """Handle mouse events"""
        if not self.is_recording:
            return

        current_time = time.time()
        
        # Handle clicks
        if hasattr(event, 'button') and event.button in [mouse.LEFT, mouse.RIGHT]:
            # Avoid duplicate events
            if current_time - self.last_click_time < 0.1:
                return
                
            self.last_click_time = current_time
            # Get current mouse position using pyautogui
            x, y = pyautogui.position()
            
            # Flush any pending keyboard input
            self._flush_key_buffer()
            
            self.action_recorded.emit(StepType.MOUSE_CLICK, {
                "name": f"Click at ({x}, {y})",
                "click_type": "coordinates",
                "x": x,
                "y": y
            })

        # Handle moves to update position tracking if needed, or just pass
        elif hasattr(event, 'dx') or hasattr(event, 'dy'):
            pass

    def _flush_key_buffer(self):
        """Flush the keyboard buffer as a typing action"""
        if self.key_buffer:
            text = ''.join(self.key_buffer)
            self.action_recorded.emit(StepType.KEYBOARD_TYPE, {
                "name": f"Type '{text}'",
                "text": text,
                "delay": 10  # Default delay between keystrokes
            })
            self.key_buffer = []

    def take_screenshot(self, x, y, width, height):
        """Take a screenshot of a specific region"""
        try:
            screenshot = pyautogui.screenshot(region=(x, y, width, height))
            return screenshot
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            return None 