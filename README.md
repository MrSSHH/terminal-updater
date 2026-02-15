# Comax Automation Suite 🚀

A two-stage automation pipeline designed to filter device inventory and perform bulk firmware updates via UI automation.

## 📋 Overview
This suite eliminates the manual labor of updating device versions. It first cleans and segments your data based on specific business rules, then uses computer vision to interact with the update software exactly like a human would.

---

## 🏗️ Project Structure
* **`filter_data.py`**: Pre-processes the master Excel sheet, applies filters, and splits data into batch files.
* **`main.py`**: The automation engine that handles UI clicks, image recognition, and version entry.
* **`Assets/`**: Contains the required Excel data and images for UI recognition.
* **`Sounds/`**: Audio cues for process status and alerts.

---

## 🛠️ Features

### 1. Smart Data Filtering (`filter_data.py`)
* **Exclusion Logic**: Automatically skips "forbidden" customer IDs and specific hardware (e.g., EF500).
* **Version Targeting**: Filters for devices below a specific version threshold (e.g., v808).
* **Batching**: Splits large datasets into chunks of 300 rows to ensure system stability during updates.

### 2. Intelligent Automation (`main.py`)
* **Coordinate Calibration**: Map your UI once; the script remembers coordinates via `config.json`.
* **Computer Vision**: Uses OpenCV to verify that the "Update Window" is actually visible before clicking.
* **Safety Kill-Switch**: Press `Ctrl + Esc` at any time to halt the process immediately.
* **Audio Feedback**: Uses beeps and alerts to notify you when manual intervention or calibration is needed.

---

## 🚀 Getting Started

### Prerequisites
* Windows OS (Required for `ctypes` alerts and UI focus).
* Python 3.8+

### Installation
1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt