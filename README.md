# Python automation Tool

A powerful, visual desktop automation tool built with Python and PyQt6. Create, record, and execute complex workflows to automate repetitive tasks on your computer.

![Automation Tool UI](icon.png)

## 🚀 Key Features

*   **Visual Workflow Editor**: Easy-to-read steps with distinct color coding:
    *   🟦 **Mouse Click**: Automate clicks at specific coordinates.
    *   🟩 **Keyboard Type**: Type text automatically.
    *   teal **Special Keys**: Press combinations like Ctrl+C, Alt+Tab.
    *   🟧 **Wait**: Add delays between actions.
    *   🟪 **Loops**: Repeat a sequence of steps.
*   **Simple Controls**: Minimalist interface with a single **Run** button.
*   **🛡️ Fail-Safe System**: Global **Ctrl+Alt+X** hotkey to immediately kill the automation at any time.
*   **Action Recorder**: Record your mouse clicks and keystrokes to generate steps automatically.
*   **Save & Load**: Store your workflows as JSON files for later use.

## 🛠️ Installation

1.  **Prerequisites**: Ensure you have Python 3.8+ installed.
2.  **Clone/Download**: Get the source code.
3.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```
    *Required libraries: `PyQt6`, `pyautogui`, `keyboard`, `mouse`, `winsound` (Windows), `opencv-python`, `pillow`.*
4.  **Install Tesseract OCR**:
    *   Download the installer from [UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki).
    *   Run the installer and note the installation path (usually `C:\Program Files\Tesseract-OCR`).
    *   Add the Tesseract path to your system's PATH environment variable.

## 📖 Usage Guide

### 1. Launching the App
Run the main script:
```bash
python automate.py
```

### 2. Creating a Workflow
*   **Add Steps**: Click the **"+ Add Step"** button to choose an action.
*   **Edit Parameters**: Click "✏️ Edit" on any step to configure details (e.g., coordinates, text to type, duration).
*   **Reorder**: Drag and drop steps to change their execution order.

### 3. Running Automation
*   **Start**: Click the **▶️ Run** button or press **F5**.
*   **Loop**: Set the "Loop Workflow" counter to repeat the entire sequence multiple times.

### 4. 🛑 Emergency Stop
If you need to stop the automation instantly (e.g., if the mouse is moving uncontrollably), press:
# **Ctrl + Alt + X**
This will immediately terminate the execution and return control to you.

## 🎥 Recording Actions
1.  Click **Tools > ⏺️ Start Recording** (or click the Record button).
2.  Perform your actions (clicks, typing).
3.  Press **ESC** to stop recording.
4.  The recorded actions will be converted into editable steps in your workflow.

> [!NOTE]
> Recording from "Tools" menu is currently in development. Some functions might not work properly. Use at your own risk.

## 📝 Requirements
See `requirements.txt` for the full list of python packages.
