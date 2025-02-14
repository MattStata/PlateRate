# PlateRate

PlateRate is a GUI-based tool designed to **rapidly find the best linear region** for rate calculations in kinetic data exported from a Thermo Scientific Varioskan Lux plate reader. It can **automatically** detect the region of highest linearity for each sample, or allow the user to **override** that region interactively.

---

## Overview

- **Primary Goal**: Semi-automated identification of a linear portion in absorbance vs. time data e.g for assays of enzyme activity rates.
- **Interactive GUI**: Lets you review and adjust the suggested region for each sample.
- **Multi-File Workflow**: You can process multiple data columns in multiple exported Excel files in a single session.
- **Final Output**: Combines all processed results (slope, start/end time, number of points used) into one results CSV file.

---

## Features

1. **Automatic Segment Detection**  
   - The script calculates a linear fit (slope, intercept) and compares its R² to a user-defined cutoff (default: 0.9995).  
   - It trims points from the front or back until that R² is reached or fewer than 2 points remain.  

2. **Interactive GUI**  
   - Refine or override the region by specifying new start/end times or by using key bindings to shift the boundaries.  
   - Displays updated slope and R² values immediately.

3. **Customizable R² Nudging**  
   - Press Ctrl+Up or Ctrl+Down to adjust the R² cutoff in small increments and see how it affects the chosen region.

4. **Multi-File Session**  
   - After finishing one file, the program prompts you to open another or finish and save results.

5. **Read/Write**  
   - Imports Excel files exported by the Varioskan Lux software.
   - Merges the results into a single CSV.

---

## Availability

1. **Python Script**  
   - Ideal for users who have Python 3.x installed. Clone this repository and run `PlateRate.py`.  
   - Dependencies include `numpy`, `pandas`, `matplotlib`, and `tkinter` (built into most Python distributions).

2. **Standalone Windows Application**  
   - For users without Python, PlateRate is also provided as a **compiled `.exe`** file for use in Windows.  
   - Double-clicking this executable will launch the same interactive GUI without needing Python installed.

---

## Usage

1. **Open** PlateRate (either the Python script or the `.exe`).  
2. **Select** an Excel file when prompted.  
3. **Review and refine** each sample’s linear segment in the GUI. You can:
   - **Press Enter** to accept the suggested region.
   - **Press Ctrl+Enter** or click “Apply Range” to override with a new start/end time.
   - **Use** arrow key bindings to shift region boundaries or nudge the R² cutoff.
4. **Repeat** for all columns in that file.
5. **Optionally** select more files.
6. **Save** combined results to CSV when done.

---

## Key Bindings (Summary)

- **Enter**: Accept current region  
- **Ctrl+Enter**: Apply Range override  
- **Ctrl+Left/Right**: Expand/Contract from the left boundary (one data point)  
- **Shift+Left/Right**: Contract/Expand from the right boundary (one data point)  
- **Ctrl+Up/Down**: Increase/Decrease R² cutoff  

---

## Requirements (Python Script Version)

- Python 3.x  
- `numpy`, `pandas`, `matplotlib`  
- The built-in `tkinter` library (usually included by default in standard Python)

---

## Credits

- **Author**: Matt Stata  
- **GitHub**: [MattStata/PlateRate](https://github.com/MattStata/PlateRate)
