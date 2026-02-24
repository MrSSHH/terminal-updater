# Terminal Automation Suite 🚀

A high-efficiency automation pipeline designed to filter massive device inventories and perform bulk firmware updates via computer vision. This suite bypasses manual data entry by simulating human interaction with legacy UI software.

---

## 🏗️ Project Structure

* **`filter_data.py`**: The "Brain." It cleans the master Excel, filters by hardware/customer, and splits data into manageable chunks.
* **`main.py`**: The "Hands." The automation engine that handles UI clicks, image recognition, and version input.
* **`Assets/`**:
* `Excels/`: Your input `Data.xlsx` and the generated `toUpdate` batches.
* `Images/`: Reference screenshots for OpenCV to "see" the UI.
* `Sounds/`: Audio feedback files (`beep.mp3`, etc.).



---

## 🛠️ Features

### 1. Smart Data Filtering

* **Hardware Exclusion**: Automatically ignores specific models (e.g., serial numbers starting with `EF500`).
* **Version Logic**: Only targets devices where both current and pending versions are below the `TARGET_VERSION` (default: 808).
* **Automatic Chunking**: Large datasets are automatically split into batches of **300** to prevent software memory leaks or crashes.

### 2. Intelligent Automation

* **Visual Verification**: Uses **OpenCV** template matching to ensure the update window is open before clicking.
* **Coordinate Memory**: Calibration is saved to `config.json` so you only have to map your screen once.
* **Fail-Safes**:
* `Ctrl + Esc`: Emergency stop.
* `Ctrl + Left Click`: Precise coordinate mapping.
* Audio cues to keep you informed without looking at the screen.



---

## 🚀 Getting Started

### Prerequisites

* **Windows OS** (Required for Win32 API alerts and UI interaction).
* **Python 3.8+**
* **Display**: Set Windows scaling to **100%** for image recognition accuracy.

### Installation

```bash
git clone https://github.com/MrSSHH/terminal-updater.git
cd terminal-updater
pip install -r requirements.txt

```

---

## 🖥️ Terminal Usage

### Step 1: Filter & Prepare

Place your master file in `Assets\Excels\Data.xlsx` and run the filter:

```text
$ python filter_data.py

INFO: Created: Assets\Excels\toUpdate\output_0.xlsx
INFO: Created: Assets\Excels\toUpdate\output_1.xlsx
INFO: Created final: Assets\Excels\toUpdate\output_2.xlsx
INFO: Filtering complete!

```

### Step 2: Execute Update

Run the automation. Use `--verbose` to see real-time confidence scores for image matching.

```text
$ python main.py --verbose --short-wait

[?] Use the previous settings (Y/N)? n
[?] Enter the desired version (e.g., 802): 808
[!] Place your cursor on the UUID input field. Press Ctrl + Left Click.
    Position confirmed: (450, 310)

[!] Configuration saved to config.json. 
[!] Once ready, press OK. Starting in 3s...

10:15:02 ComaxLogger DEBUG Starting to upgrade UUID: 8832-XJ-11
10:15:03 ComaxLogger DEBUG Image recognition confidence: 0.98
10:15:04 ComaxLogger INFO  Click at (450, 310)
10:15:08 ComaxLogger DEBUG UUID 8832-XJ-11 updated successfully.

```

---

## ⌨️ Command Line Arguments

| Argument | Description |
| --- | --- |
| `-s`, `--silent` | Disable all sound notifications. |
| `-v`, `--verbose` | Output all debug logs and image matching scores to the terminal. |
| `--short-wait` | Reduces the initial startup countdown to 3 seconds. |

---

## 🛑 Safety & Control

* **Stop Process**: Hold `Ctrl + Esc` to kill the automation loop.
* **Re-Calibration**: If your UI layout changes, run the script and choose `N` at the "Use previous settings" prompt to reset `config.json`.
