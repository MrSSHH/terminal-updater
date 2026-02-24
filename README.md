# Terminal Automation Suite 🚀

A robust, two-stage automation pipeline designed to filter device inventory and perform bulk firmware updates via computer vision and UI automation. This tool simulates human interaction to handle repetitive tasks in legacy software environments where APIs are unavailable.

---

## 📋 Overview

This suite eliminates the manual labor of updating device versions. It first cleans and segments your data based on specific business rules, then uses **OpenCV** and **PyAutoGUI** to interact with the update software exactly like a human would.

---

## 🏗️ Project Structure

* **`filter_data.py`**: Pre-processes the master Excel sheet, applies filters, and splits data into batch files.
* **`main.py`**: The automation engine that handles UI clicks, image recognition, and version entry.
* **`Assets/`**: Contains the required Excel data and images for UI recognition.
* **`Sounds/`**: Audio cues for process status and alerts.
* **`config.json`**: Stores saved screen coordinates to skip calibration on repeat runs.

---

## 🛠️ Features

### 1. Smart Data Filtering (`filter_data.py`)

* **Exclusion Logic**: Automatically skips "forbidden" customer IDs and specific hardware (e.g., EF500).
* **Version Targeting**: Filters for devices below a specific version threshold (e.g., v808).
* **Batching**: Splits large datasets into chunks to ensure system stability.

### 2. Intelligent Automation (`main.py`)

* **Coordinate Calibration**: Map your UI once; the script remembers coordinates via `config.json`.
* **Computer Vision**: Uses **OpenCV** template matching to verify that the "Update Window" is actually visible before clicking.
* **Safety Kill-Switch**: Press `Ctrl + Esc` at any time to halt the process immediately.
* **Audio Feedback**: Uses distinct beeps (`.mp3`) to notify you of successful steps or required manual intervention.

---

## 🚀 Getting Started

### Prerequisites

* **Windows OS** (Required for `ctypes` alerts and Windows-specific UI focus).
* **Python 3.8+**
* **Display Settings**: Set Windows Display Scaling to **100%** for best image recognition accuracy.

### Installation

1. **Clone the repository:**
```bash
git clone https://github.com/MrSSHH/terminal-updater.git
cd terminal-updater

```


2. **Install dependencies:**
```bash
pip install -r requirements.txt

```



---

## 🖥️ Terminal Usage

### Stage 1: Filter & Prepare Data

Process your master inventory to create the target update list in `Assets\Excels\data.xlsx`.

```bash
$ python filter_data.py
[OK] 1,200 devices scanned.
[OK] 450 devices filtered (v808 threshold).
[OK] Assets\Excels\data.xlsx generated successfully.

```

### Stage 2: Execute UI Automation

Run the main engine. Use flags to customize the experience.

```bash
# Basic run
$ python main.py

# Verbose mode with silent notifications and fast start
$ python main.py --verbose --silent --short-wait

```

#### **Calibration Walkthrough**

If no `config.json` is found, the script will guide you through setup:

```text
$ python main.py

[?] Use the previous settings (Y/N)? n
[?] Enter the desired version (e.g., 802): 915
[!] Place your cursor on the UUID input field. Press Ctrl + Left Click.
    Position confirmed at (412, 530).

[!] Place your cursor on the Update button. Press Ctrl + Left Click.
    Position confirmed at (890, 530).

[!] Configuration saved to config.json.
[!] Once ready, press OK to start. You will have 5 seconds to open the window.

14:22:01 ComaxLogger INFO  Starting upgrade for UUID: 8832-XJ-11
14:22:03 ComaxLogger DEBUG Image Match Found: Update Window (Confidence: 0.94)
14:22:05 ComaxLogger INFO  UUID 8832-XJ-11 updated successfully. (14 Remaining)

```

---

## ⌨️ Command Line Arguments

| Flag | Short | Description |
| --- | --- | --- |
| `--silent` | `-s` | Disable all beep/audio notifications. |
| `--verbose` | `-v` | Output all debug logs to the terminal. |
| `--short-wait` |  | Use a 3-second countdown instead of 5 seconds. |

## 🛑 Safety & Controls

* **Emergency Stop**: Press `Ctrl + Esc` to kill the script immediately.
* **Calibration**: If your window moves, run the script and select `N` when asked to "Use previous settings" to re-map the coordinates.

---
