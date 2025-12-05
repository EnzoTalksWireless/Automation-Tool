import time

import pyautogui
import keyboard
import cv2
import numpy as np
from PIL import Image
import logging
from PyQt6.QtCore import QObject, pyqtSignal
from automation_steps import StepType
import os
import tempfile
import sys
from PyQt6.QtWidgets import QInputDialog, QApplication
import appdirs

def get_application_path():
    """Get the correct application path whether running as script or frozen exe"""
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (pyinstaller)
        return os.path.dirname(sys.executable)
    else:
        # If the application is run from a Python interpreter
        return os.path.dirname(os.path.abspath(__file__))

# Directory constants
APPLICATION_PATH = get_application_path()
IMAGES_DIR = os.path.join(APPLICATION_PATH, "images")
AUTOMATIONS_DIR = os.path.join(APPLICATION_PATH, "Saved Automations")
DEBUG_DIR = os.path.join(APPLICATION_PATH, "debug")

# Try importing Tesseract, but don't fail if not available
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False

class WorkflowExecutor(QObject):
    """Executes automation workflows"""
    step_started = pyqtSignal(int, str)  # Signal emitted when a step starts
    step_completed = pyqtSignal(int)  # Signal emitted when a step completes
    step_error = pyqtSignal(int, str)  # Signal emitted when a step encounters an error
    workflow_completed = pyqtSignal()  # Signal emitted when the workflow completes
    debug_info = pyqtSignal(str)  # Signal for debug information
    loop_iteration_completed = pyqtSignal(int)  # Signal emitted when a loop iteration completes


    def __init__(self):
        super().__init__()
        self.running = False
        self.paused = False
        self.debug_mode = True  # Enable debug mode by default
        pyautogui.PAUSE = 1.0  # Increase delay for better reliability
        pyautogui.FAILSAFE = True  # Enable fail-safe feature
        
        # Ensure debug directory exists
        os.makedirs(DEBUG_DIR, exist_ok=True)
        
        # Configure logging with error handling
        log_file = os.path.join(DEBUG_DIR, 'automation.log')
        try:
            logging.basicConfig(
                level=logging.DEBUG,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler()
                ]
            )
        except Exception as e:
            # If file logging fails, fall back to console-only logging
            logging.basicConfig(
                level=logging.DEBUG,
                format='%(asctime)s - %(levelname)s - %(message)s',
                handlers=[logging.StreamHandler()]
            )
            print(f"Warning: Could not create log file. Falling back to console logging. Error: {str(e)}")

        if TESSERACT_AVAILABLE:
            self._debug_msg("Tesseract OCR is available")
        else:
            self._debug_msg("Tesseract OCR is not available, will use OpenCV text detection as fallback")
            


    def execute_workflow(self, steps, loop_count=1):
        """Execute a sequence of automation steps with support for nested loops and delays"""
        self.running = True
        self.paused = False
        
        try:
            total_steps = len(steps)
            self._debug_msg(f"\n=== Starting Workflow Execution (Global Loops: {loop_count}) ===")
            
            # Reset text input indices for multiple text inputs
            self._reset_text_indices(steps)
            
            for global_loop in range(loop_count):
                if not self.running:
                    self._debug_msg("\n❌ Workflow execution stopped by user")
                    break
                
                self._debug_msg(f"\n=== Starting Global Loop Iteration {global_loop + 1}/{loop_count} ===")
                
                # Loop state management
                # Stack to keep track of active loops: [{"start_index": int, "iterations": int, "current_iter": int}]
                loop_stack = []
                
                # Instruction pointer
                i = 0
                
                while i < len(steps):
                    if not self.running:
                        break
                        
                    while self.paused:
                        time.sleep(0.1)
                    
                    step = steps[i]
                    step_type = step["type"]
                    params = step["params"]
                    step_name = params.get("name", f"Step {i+1}")
                    
                    # Log step execution (except strictly control flow steps to avoid spam, or log them differently)
                    if step_type not in [StepType.LOOP_END]:
                        self._debug_msg(f"\n=== Step {i+1}/{total_steps}: {step_name} ({step_type}) ===")
                        self.step_started.emit(i, step_type)
                    
                    try:
                        # Handle Loop Start
                        if step_type == StepType.LOOP_START:
                            iterations = params.get("iterations", 1)
                            self._debug_msg(f"Processing Loop Start at index {i}. Iterations requested: {iterations}")
                            
                            # Check if we are already in this loop (top of stack matches this index)
                            if loop_stack and loop_stack[-1]["start_index"] == i:
                                self._debug_msg(f"  -> Re-entering Loop Start (already in stack)")
                            else:
                                # New entry into this loop
                                loop_stack.append({
                                    "start_index": i,
                                    "iterations": iterations,
                                    "current_iter": 0
                                })
                                self._debug_msg(f"  -> New Loop Started (Stack depth: {len(loop_stack)})")
                            
                            self.step_completed.emit(i)
                            i += 1
                            continue

                        # Handle Loop End
                        elif step_type == StepType.LOOP_END:
                            if not loop_stack:
                                self._debug_msg("Warning: Loop End found without matching Start. Ignoring.")
                                i += 1
                                continue
                            
                            current_loop = loop_stack[-1]
                            current_loop["current_iter"] += 1
                            
                            self._debug_msg(f"Loop End reached. Iteration {current_loop['current_iter']} of {current_loop['iterations']}")
                            
                            if current_loop["current_iter"] < current_loop["iterations"]:
                                # Jump back to start + 1 (next step after Loop Start)
                                jump_to = current_loop["start_index"] + 1
                                self._debug_msg(f"  -> Jumping back to step index {jump_to} (Step {jump_to + 1})")
                                i = jump_to
                            else:
                                # Loop finished
                                self._debug_msg("  -> Loop finished. Popping from stack.")
                                loop_stack.pop()
                                i += 1
                            
                            self.step_completed.emit(i - 1 if i > 0 else 0)
                            continue

                        # Execute normal steps
                        
                        # Take debug screenshot before action
                        if self.debug_mode:
                            self._take_debug_screenshot(f"step_{i+1}_before")
                        
                        start_time = time.time()
                        
                        self._execute_step(step_type, params)
                        
                        end_time = time.time()
                        
                        # Take debug screenshot after action
                        if self.debug_mode:
                            time.sleep(0.5)  # Wait for UI to update
                            self._take_debug_screenshot(f"step_{i+1}_after")
                        
                        self.step_completed.emit(i)
                        execution_time = round(end_time - start_time, 2)
                        self._debug_msg(f"✓ Step {i+1} completed successfully (took {execution_time}s)")
                        
                        i += 1

                    except Exception as e:
                        # Convert technical errors to user-friendly messages
                        user_msg = self._get_user_friendly_error(e, step_type)
                        self._debug_msg(f"❌ Error in step {i+1}: {user_msg}")
                        self.step_error.emit(i, user_msg)
                        self._debug_msg(f"Technical details: {str(e)}")
                        
                        # Continue with next step instead of crashing
                        i += 1
                
                # Emit loop iteration completed signal for global loop
                if self.running:
                    self.loop_iteration_completed.emit(global_loop + 1)
                    self._debug_msg(f"=== Global Loop Iteration {global_loop + 1} Completed ===")
            
            self.workflow_completed.emit()
            if self.running:
                self._debug_msg("\n=== Workflow Completed Successfully ===")
            
        finally:
            self.running = False

    def _reset_text_indices(self, steps):
        """Reset the current indices for all steps with multiple inputs"""
        for step in steps:
            if step["type"] == StepType.KEYBOARD_TYPE:
                step["params"]["current_text_index"] = 0
            elif step["type"] == StepType.MOUSE_CLICK and step["params"].get("click_type") == "image" and step["params"].get("input_type") == "multiple":
                step["params"]["current_image_index"] = 0

    def _debug_msg(self, message):
        """Log debug message and emit signal with error handling"""
        try:
            logging.debug(message)
        except Exception as e:
            print(f"Logging failed: {str(e)}")
        self.debug_info.emit(message)

    def _take_debug_screenshot(self, name):
        """Take a debug screenshot with error handling"""
        try:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(DEBUG_DIR, f"{name}_{timestamp}.png")
            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
            self._debug_msg(f"Debug screenshot saved: {filename}")
        except Exception as e:
            self._debug_msg(f"Failed to take debug screenshot: {str(e)}")
            # Continue execution even if screenshot fails
            pass

    def _execute_step(self, step_type, params):
        """Execute a single automation step"""
        if step_type == StepType.MOUSE_CLICK:
            self._execute_mouse_click(params)
        elif step_type == StepType.KEYBOARD_TYPE:
            self._execute_keyboard_type(params)
        elif step_type == StepType.KEYBOARD_SPECIAL:
            self._execute_keyboard_special(params)
        elif step_type == StepType.WAIT:
            self._execute_wait(params)
        elif step_type in [StepType.LOOP_START, StepType.LOOP_END]:
            # Handled in main loop, but just in case
            pass
        else:
            raise ValueError(f"Unknown step type: {step_type}")

    def _execute_wait(self, params):
        """Execute a wait step"""
        duration = params.get("duration", 1)
        self._debug_msg(f"Waiting for {duration} seconds...")
        
        start_time = time.time()
        while time.time() - start_time < duration:
            if not self.running:
                break
            
            while self.paused:
                if not self.running:
                    break
                time.sleep(0.1)
                
            time.sleep(0.1)

    def _execute_mouse_click(self, step_data):
        """Execute a mouse click step"""
        try:
            click_type = step_data.get("click_type", "coordinates")
            button = step_data.get("button", "left")
            duration = step_data.get("duration", 0.5)
            
            # Get click position
            if click_type == "coordinates":
                x = step_data.get("x", 0)
                y = step_data.get("y", 0)
                pyautogui.moveTo(x, y, duration=duration)
            else:  # image-based click
                image_path = step_data.get("image_path")
                confidence = step_data.get("confidence", 0.9)
                
                if not image_path or not os.path.exists(image_path):
                    raise ValueError("Image not found: " + str(image_path))
                
                # Find and click the image
                location = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
                if not location:
                    raise ValueError(f"Could not find image on screen: {image_path}")
                
                x, y = location
                pyautogui.moveTo(x, y, duration=duration)
            
            # Perform the click
            pyautogui.click(button=button)
            
            # Handle text input after click if enabled
            if step_data.get("type_after_click"):
                time.sleep(step_data.get("type_delay", 1))
                
                text_to_type = step_data.get("text_to_type", "")

                
                # Type the text
                if text_to_type:
                    pyautogui.write(text_to_type)
                    
                    # Handle special key after typing
                    special_key = step_data.get("special_key")
                    if special_key:
                        pyautogui.press(special_key.lower())
            
            return True
            
        except Exception as e:
            self._log_error(f"Error in mouse click step: {str(e)}")
            raise



    def _execute_keyboard_type(self, params):
        """Execute a keyboard typing action with support for multiple text inputs"""
        if params.get("input_type") == "multiple":
            # Get the current text from the list
            text_list = params.get("text_list", [])
            if not text_list:
                raise ValueError("No text inputs available in multiple input mode")
            
            current_index = params.get("current_text_index", 0)
            if current_index >= len(text_list):
                current_index = 0
            
            # Get the current text to type
            text = text_list[current_index]
            
            # Update the index for the next iteration
            params["current_text_index"] = (current_index + 1) % len(text_list)
        else:
            text = params["text"]

        delay = params.get("delay", 10) / 1000  # Convert to seconds
        
        # Type the text using pyautogui.write which handles spaces and special characters better
        self._debug_msg(f"Typing text: '{text}' (Length: {len(text)}, Delay: {delay}s)")
        pyautogui.write(text, interval=delay)
        
        # Handle special key if specified
        special_key = params.get("special_key")
        if special_key:
            # Convert special key name to pyautogui format
            key_mapping = {
                "Enter": "enter",
                "Tab": "tab",
                "Space": "space",
                "Backspace": "backspace",
                "Delete": "delete",
                "Escape": "esc",
                "Up": "up",
                "Down": "down",
                "Left": "left",
                "Right": "right"
            }
            
            key_to_press = key_mapping.get(special_key, special_key.lower())
            self._debug_msg(f"Pressing special key: {key_to_press}")
            pyautogui.press(key_to_press)
            time.sleep(0.1)  # Small delay after special key

    def _execute_keyboard_special(self, params):
        """Execute a special keyboard action"""
        try:
            key = params["key"].lower()
            modifiers = params.get("modifiers", [])
            
            if key == "windows":
                key = "win"
            
            if key in ["ctrl", "alt", "shift", "win"]:
                modifiers = [mod for mod in modifiers if mod != key]
            
            key_combo = []
            if "ctrl" in modifiers:
                key_combo.append("ctrl")
            if "alt" in modifiers:
                key_combo.append("alt")
            if "shift" in modifiers:
                key_combo.append("shift")
            if key == "win":
                key_combo.insert(0, "win")
            elif "win" in modifiers:
                key_combo.insert(0, "win")
            
            if key != "win":
                key_combo.append(key)
            
            key_sequence = '+'.join(key_combo)
            self._debug_msg(f"Executing keyboard combination: {key_sequence}")
            keyboard.press_and_release(key_sequence)
            time.sleep(0.1)
            
        except Exception as e:
            self._debug_msg(f"Keyboard special action failed: {str(e)}")
            raise



    def _find_image(self, image_path, confidence=0.9):
        """Find an image on screen and return its center coordinates"""
        try:
            # Convert relative path to absolute path if needed
            if not os.path.isabs(image_path):
                image_path = os.path.join(IMAGES_DIR, os.path.basename(image_path))

            if not os.path.exists(image_path):
                raise FileNotFoundError(
                    f"Image file not found: {os.path.basename(image_path)}\n"
                    f"Please make sure the image exists in the 'images' folder."
                )
            
            # Take screenshot and save for debugging
            screenshot = pyautogui.screenshot()
            debug_screen = os.path.join(DEBUG_DIR, "current_screen.png")
            screenshot.save(debug_screen)
            self._debug_msg(f"Current screen saved to: {debug_screen}")
            
            screenshot_np = np.array(screenshot)
            screenshot_gray = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2GRAY)
            
            # Load and save template image for debugging
            template = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
            if template is None:
                raise RuntimeError(
                    f"Failed to load image: {os.path.basename(image_path)}\n"
                    "Please ensure the image file is a valid image format (PNG, JPG, etc.)"
                )
            
            debug_template = os.path.join(DEBUG_DIR, "template.png")
            cv2.imwrite(debug_template, template)
            self._debug_msg(f"Template image saved to: {debug_template}")
            
            # Perform template matching
            result = cv2.matchTemplate(screenshot_gray, template, cv2.TM_CCOEFF_NORMED)
            min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)
            
            self._debug_msg(f"Best match confidence: {max_val:.4f}")
            
            if max_val >= confidence:
                # Calculate center point
                w, h = template.shape[::-1]
                center_x = max_loc[0] + w//2
                center_y = max_loc[1] + h//2
                
                # Save debug image with match highlighted
                debug_result = screenshot_np.copy()
                cv2.rectangle(
                    debug_result,
                    max_loc,
                    (max_loc[0] + w, max_loc[1] + h),
                    (0, 255, 0),
                    2
                )
                cv2.circle(
                    debug_result,
                    (center_x, center_y),
                    5,
                    (255, 0, 0),
                    -1
                )
                debug_match = os.path.join(DEBUG_DIR, "match_result.png")
                cv2.imwrite(debug_match, debug_result)
                self._debug_msg(f"Match result saved to: {debug_match}")
                
                return (center_x, center_y)
            
            self._debug_msg(
                f"No match found above confidence threshold ({confidence})\n"
                "Try adjusting the confidence level or updating the reference image."
            )
            return None
            
        except Exception as e:
            self._debug_msg(f"Error finding image: {str(e)}")
            return None

    def _find_text(self, text, region=None, confidence=0.7):
        """Find text on screen using OCR"""
        try:
            # Take screenshot of the specified region or full screen
            screenshot = pyautogui.screenshot(region=region)
            screenshot_np = np.array(screenshot)
            
            # Save screenshot for debugging
            debug_screen = os.path.join(DEBUG_DIR, "ocr_screen.png")
            screenshot.save(debug_screen)
            self._debug_msg(f"OCR screen saved to: {debug_screen}")

            # Try Tesseract first if available
            if TESSERACT_AVAILABLE:
                try:
                    self._debug_msg("Attempting to use Tesseract OCR")
                    result = self._find_text_tesseract(screenshot_np, text, confidence)
                    if result:
                        return result
                except Exception as e:
                    self._debug_msg(f"Tesseract OCR failed: {str(e)}, falling back to OpenCV")

            # Fall back to OpenCV text detection
            self._debug_msg("Using OpenCV text detection")
            return self._find_text_opencv(screenshot_np, text, confidence)

        except Exception as e:
            self._debug_msg(f"Error finding text: {str(e)}")
            return None

    def _find_text_tesseract(self, image_np, target_text, confidence):
        """Find text using Tesseract OCR"""
        # Convert to grayscale
        gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
        
        # Apply thresholding to get better OCR results
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Perform OCR
        data = pytesseract.image_to_data(binary, output_type=pytesseract.Output.DICT)
        
        # Search for target text
        for i, text in enumerate(data['text']):
            if target_text.lower() in text.lower():
                confidence_score = float(data['conf'][i]) / 100
                if confidence_score >= confidence:
                    x = data['left'][i]
                    y = data['top'][i]
                    w = data['width'][i]
                    h = data['height'][i]
                    center_x = x + w//2
                    center_y = y + h//2
                    
                    # Save debug image
                    debug_result = image_np.copy()
                    cv2.rectangle(debug_result, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.circle(debug_result, (center_x, center_y), 5, (255, 0, 0), -1)
                    debug_match = os.path.join(DEBUG_DIR, "tesseract_match.png")
                    cv2.imwrite(debug_match, debug_result)
                    
                    self._debug_msg(f"Tesseract found text with confidence: {confidence_score:.4f}")
                    return (center_x, center_y)
        
        return None

    def _find_text_opencv(self, image_np, target_text, confidence):
        """Find text using OpenCV text detection"""
        # Convert to grayscale
        gray = cv2.cvtColor(image_np, cv2.COLOR_RGB2GRAY)
        
        # Apply preprocessing to improve text detection
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        _, binary = cv2.threshold(blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        # Find contours
        contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        
        # Filter contours based on area and aspect ratio
        text_regions = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            aspect_ratio = float(w) / h
            area = cv2.contourArea(contour)
            
            if 0.1 < aspect_ratio < 10 and area > 100:  # Adjust these thresholds as needed
                text_regions.append((x, y, w, h))
        
        # Sort regions by position (left to right, top to bottom)
        text_regions.sort(key=lambda r: (r[1], r[0]))
        
        # Save debug image with detected regions
        debug_result = image_np.copy()
        for x, y, w, h in text_regions:
            cv2.rectangle(debug_result, (x, y), (x + w, y + h), (0, 255, 0), 2)
        debug_match = os.path.join(DEBUG_DIR, "opencv_regions.png")
        cv2.imwrite(debug_match, debug_result)
        
        # For each region, try to match the text pattern
        for x, y, w, h in text_regions:
            roi = gray[y:y+h, x:x+w]
            
            # Apply additional preprocessing for better matching
            _, roi_binary = cv2.threshold(roi, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            # Calculate similarity with target text pattern
            # This is a simplified approach - in practice, you might want to use
            # more sophisticated pattern matching or feature extraction
            similarity = self._calculate_text_similarity(roi_binary, target_text)
            
            if similarity >= confidence:
                center_x = x + w//2
                center_y = y + h//2
                self._debug_msg(f"OpenCV found potential text match with confidence: {similarity:.4f}")
                return (center_x, center_y)
        
        return None

    def _calculate_text_similarity(self, roi, target_text):
        """Calculate similarity between ROI and target text pattern"""
        # This is a simplified implementation
        # You might want to implement more sophisticated text recognition here
        # For now, we'll use basic image statistics as a rough approximation
        
        # Normalize ROI
        roi_norm = cv2.normalize(roi, None, 0, 1, cv2.NORM_MINMAX)
        
        # Calculate basic statistics
        mean = np.mean(roi_norm)
        std = np.std(roi_norm)
        
        # Return a confidence score based on image statistics
        # This is a very basic approach and should be improved based on your needs
        return (mean + std) / 2

    def _get_user_friendly_error(self, error, step_type):
        """Convert technical error messages to user-friendly ones"""
        error_str = str(error).lower()
        
        if "image not found" in error_str:
            return str(error)

    def pause(self):
        """Pause workflow execution"""
        self.paused = True

    def resume(self):
        """Resume workflow execution"""
        self.paused = False

    def stop(self):
        """Stop workflow execution"""
        self.running = False

