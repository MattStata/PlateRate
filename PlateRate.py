#!/usr/bin/env python
"""
PlateRate: A script for rapid processing of kinetic data exported from the Varioskan Lux plate reader,
           to find the best linear region of each dataset for rate calculations.

Author: Matt Stata
GitHub: https://github.com/MattStata/PlateRate

Workflow Overview:
1) The user selects one or more Excel files exported from the Varioskan Lux.
2) For each file, this script:
   - Reads the data, identifying 'avg. time [s]' as the time index (if present).
   - Presents an interactive GUI to choose or override the linear region for each column (which represents a well or sample).
   - After processing all columns in a file, the user can select another file or stop.
3) Upon completion, the user is prompted to save a combined CSV of results.

Key Features:
- Time is assumed to be in a column named 'avg. time [s]'.
- Interactive GUI with the following key bindings:
   - Enter: Accept current segment
   - Ctrl+Enter: Apply Range override
   - Ctrl+Left/Right: Expand/Contract region from the left by 1 data point
   - Shift+Left/Right: Contract/Expand region from the right by 1 data point
   - Ctrl+Up/Down: Nudge R² cutoff up/down with a custom increment
- Note that if the user closes the GUI window (e.g., clicking the 'X'), no data will be saved.
"""

import sys
import re
import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("Agg")  # Embeds plots in a Tk window rather than using plt.show().
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


def process_excel(file_path):
    """
    Reads and cleans an Excel file:
    - Finds the header row containing 'Reading', uses that as the starting point for column names.
    - Drops columns named 'Unnamed' and any column literally named 'Reading'.
    - Strips leading/trailing whitespace from remaining column names.
    - If 'avg. time [s]' is present, set it as the index (time axis).
    - Drops any rows containing NaNs.
    """
    raw_data = pd.read_excel(file_path, header=None)
    header_row = None
    for i, row in raw_data.iterrows():
        if row[0] == "Reading":
            header_row = i
            break

    if header_row is None:
        raise ValueError("No header row starting with 'Reading' found in the file.")

    df = pd.read_excel(file_path, skiprows=header_row)
    df = df.loc[:, ~df.columns.astype(str).str.match(r'^Unnamed')]

    # Strip whitespace from column names only
    new_columns = [str(col).strip() for col in df.columns]
    df.columns = new_columns

    if "Reading" in df.columns:
        df.drop(columns=["Reading"], inplace=True)

    if "avg. time [s]" in df.columns:
        df.set_index("avg. time [s]", inplace=True)

    df.dropna(inplace=True)
    return df


def calculate_r2(x, y):
    """
    Calculates the R² value and linear fit (slope, intercept) for numeric arrays x and y.
    Returns (r_squared, (slope, intercept)).
    """
    if len(x) < 2:
        return 0.0, (0.0, 0.0)

    coeffs = np.polyfit(x, y, 1)  # slope, intercept
    fit_y = np.polyval(coeffs, x)
    ss_total = np.sum((y - np.mean(y))**2)
    ss_residual = np.sum((y - fit_y)**2)
    r_squared = 1 - (ss_residual / ss_total) if ss_total != 0 else 0.0
    return r_squared, coeffs


def find_best_linear_segment(time, values, r2_cutoff=0.9995, change_threshold=0.5):
    """
    Finds the best linear segment for a given dataset by iteratively removing front/back points
    until R² >= r2_cutoff or too few points remain. Then refines the ends if they have abnormally
    low observed changes compared to the fitted expectation.

    Returns:
    (start_idx, end_idx, slope, r2) for the chosen segment, or (None, None, None, None) if not found.
    """
    start_idx = 0
    end_idx = len(time)
    r2, coeffs = calculate_r2(time[start_idx:end_idx], values[start_idx:end_idx])

    # Iteratively remove the first or last point to increase R² until cutoff is reached or no data left.
    while r2 < r2_cutoff and (end_idx - start_idx) > 2:
        r2_first, _ = calculate_r2(time[start_idx+1:end_idx], values[start_idx+1:end_idx])
        r2_last, _ = calculate_r2(time[start_idx:end_idx-1], values[start_idx:end_idx-1])

        if r2_first > r2_last:
            start_idx += 1
            r2 = r2_first
        else:
            end_idx -= 1
            r2 = r2_last

    if r2 < r2_cutoff or (end_idx - start_idx) < 2:
        return None, None, None, None

    r2, coeffs = calculate_r2(time[start_idx:end_idx], values[start_idx:end:end])
    # Typo above, let's correct it to the standard approach
    r2, coeffs = calculate_r2(time[start_idx:end_idx], values[start_idx:end_idx])
    slope = coeffs[0]

    # Refine ends by removing points that have abnormally low observed change.
    while (end_idx - start_idx) > 2:
        # Check front
        expected_front = abs(np.polyval(coeffs, time[start_idx+1]) - np.polyval(coeffs, time[start_idx]))
        observed_front = abs(values[start_idx+1] - values[start_idx])
        if observed_front <= change_threshold * expected_front:
            start_idx += 1
            r2, coeffs = calculate_r2(time[start_idx:end_idx], values[start_idx:end_idx])
            slope = coeffs[0]
            continue

        # Check end
        expected_end = abs(np.polyval(coeffs, time[end_idx-2]) - np.polyval(coeffs, time[end_idx-1]))
        observed_end = abs(values[end_idx-2] - values[end_idx-1])
        if observed_end <= change_threshold * expected_end:
            end_idx -= 1
            r2, coeffs = calculate_r2(time[start_idx:end_idx], values[start_idx:end_idx])
            slope = coeffs[0]
            continue

        break

    return start_idx, end_idx, slope, r2


def get_r2_increment(r2_cutoff):
    """
    Custom increment logic for nudging the R² cutoff:
      - <= 0.9996 => 0.0001
      - <= 0.9997 => 0.00005
      - above 0.9997 => 0.000025
    """
    if r2_cutoff > 0.9997:
        return 0.000025
    elif r2_cutoff > 0.9996:
        return 0.00005
    else:
        return 0.0001


class TrendlineGUI:
    """
    PlateRate main GUI for interactive segment selection.

    - The user can accept or override the best linear region for each column,
      or nudge the R² cutoff dynamically.
    - The script updates the best segment in real time as the user modifies the region.
    - Final slopes and ranges are stored for each column.
    """

    def __init__(self, df, r2_cutoff=0.9995, change_threshold=0.5):
        self.df = df
        self.columns = df.columns.tolist()
        self.num_cols = len(self.columns)
        self.r2_cutoff = r2_cutoff
        self.change_threshold = change_threshold

        # Storage for final results
        self.results_dict = {}
        self.current_col_idx = 0

        # Remember last user-typed range
        self.last_start_val = None
        self.last_end_val = None

        self.root = tk.Tk()
        self.root.title("PlateRate - Interactive Kinetic Analysis")

        # If user closes the window, exit
        self.root.protocol("WM_DELETE_WINDOW", self.force_exit)

        # Key binding info
        info_text = (
            "PlateRate Key Bindings:\n"
            "Enter = Accept\n"
            "Ctrl+Enter = Apply Range\n"
            "Ctrl+Left/Right = Expand/Contract left by 1\n"
            "Shift+Left/Right = Contract/Expand right by 1\n"
            "Ctrl+Up/Down = Increase/Decrease R² cutoff"
        )
        info_frame = ttk.Frame(self.root, padding="5 5 5 5")
        info_frame.pack(side=tk.TOP, fill=tk.X)

        keybinding_label = ttk.Label(info_frame, text=info_text)
        keybinding_label.pack(side=tk.LEFT)

        mainframe = ttk.Frame(self.root, padding="5 5 5 5")
        mainframe.pack(fill=tk.BOTH, expand=True)

        # Figure & canvas
        self.fig = Figure(figsize=(6,4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=mainframe)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Control frame
        control_frame = ttk.Frame(mainframe, padding="5 5 5 5")
        control_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # Start/End range boxes
        ttk.Label(control_frame, text="Start time:").pack(side=tk.LEFT)
        self.start_entry = ttk.Entry(control_frame, width=10)
        self.start_entry.pack(side=tk.LEFT, padx=2)

        ttk.Label(control_frame, text="End time:").pack(side=tk.LEFT)
        self.end_entry = ttk.Entry(control_frame, width=10)
        self.end_entry.pack(side=tk.LEFT, padx=2)

        # "Apply Range" button
        apply_btn = ttk.Button(control_frame, text="Apply Range", command=self.apply_override)
        apply_btn.pack(side=tk.LEFT, padx=4)

        # Accept & Quit
        accept_btn = ttk.Button(control_frame, text="Accept", command=self.accept_segment)
        accept_btn.pack(side=tk.LEFT, padx=4)

        quit_btn = ttk.Button(control_frame, text="Quit", command=self.quit_all)
        quit_btn.pack(side=tk.RIGHT, padx=4)

        # Label for textual info
        self.col_label = ttk.Label(mainframe, text="", anchor="center", font=("Arial", 12, "bold"))
        self.col_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Storage for current data column
        self.time = None
        self.values = None
        self.start_idx = None
        self.end_idx = None
        self.slope = None
        self.r2_value = None

        # Key bindings
        self.root.bind("<Return>", self.enter_accept)
        self.root.bind("<Control-Return>", self.ctrl_enter)
        self.root.bind("<Control-Left>", self.ctrl_left)
        self.root.bind("<Control-Right>", self.ctrl_right)
        self.root.bind("<Shift-Left>", self.shift_left)
        self.root.bind("<Shift-Right>", self.shift_right)
        self.root.bind("<Control-Up>", self.ctrl_up)
        self.root.bind("<Control-Down>", self.ctrl_down)

        self.load_column(0)

    def force_exit(self):
        """Exit if user closes the window."""
        sys.exit(0)

    def load_column(self, col_idx):
        """Load a new column, auto-determine best segment, and display it."""
        if col_idx >= self.num_cols:
            self.finish_all()
            return

        self.current_col_idx = col_idx
        col_name = self.columns[col_idx]
        self.time = self.df.index.to_numpy(dtype=float)
        self.values = self.df[col_name].to_numpy(dtype=float)

        si, ei, slope, r2 = find_best_linear_segment(self.time, self.values, self.r2_cutoff, self.change_threshold)
        disp = f"{col_idx+1}/{self.num_cols}"

        if si is not None:
            self.start_idx = si
            self.end_idx = ei
            self.slope = slope
            self.r2_value = r2
            msg = (
                f"Column {disp}: {col_name}\n"
                f"{self.r2_cutoff} (R² cutoff)\n"
                "Best Linear Region (Auto):\n"
                f"{r2:.6f}\n"
                f"{slope:.3e}"
            )
        else:
            self.start_idx = None
            self.end_idx = None
            self.slope = None
            self.r2_value = 0.0
            msg = (
                f"Column {disp}: {col_name}\n"
                f"{self.r2_cutoff} (R² cutoff)\n"
                "Best Linear Region (Auto):\n"
                "No valid region\n"
                "0.00e+00"
            )

        self.col_label.config(text=msg)
        self.start_entry.delete(0, tk.END)
        self.end_entry.delete(0, tk.END)
        if self.last_start_val is not None:
            self.start_entry.insert(0, str(self.last_start_val))
        if self.last_end_val is not None:
            self.end_entry.insert(0, str(self.last_end_val))

        self.plot_current()

    def recompute_auto_segment(self):
        """Re-run the best-segment finder after changing r2_cutoff."""
        si, ei, slope, r2 = find_best_linear_segment(self.time, self.values, self.r2_cutoff, self.change_threshold)
        idx_disp = f"{self.current_col_idx+1}/{self.num_cols}"
        col_name = self.columns[self.current_col_idx]

        if si is not None:
            self.start_idx = si
            self.end_idx = ei
            self.slope = slope
            self.r2_value = r2
            msg = (
                f"Column {idx_disp}: {col_name}\n"
                f"{self.r2_cutoff} (R² cutoff)\n"
                "Best Linear Region (Auto):\n"
                f"{r2:.6f}\n"
                f"{slope:.3e}"
            )
        else:
            self.start_idx = None
            self.end_idx = None
            self.slope = None
            self.r2_value = 0.0
            msg = (
                f"Column {idx_disp}: {col_name}\n"
                f"{self.r2_cutoff} (R² cutoff)\n"
                "Best Linear Region (Auto):\n"
                "No valid region\n"
                "0.00e+00"
            )

        self.col_label.config(text=msg)
        self.plot_current()

    def plot_current(self):
        """Plot the data, highlight the chosen segment, and show the linear fit."""
        self.ax.clear()

        # Legend ordering: the segment first in the legend, then the unused data
        # We do that by plotting the segment first (blue) then the unused data (gray),
        # but it's simpler to do the reverse so let's just set the label order after.
        seg_t, seg_v = [], []
        if self.start_idx is not None and self.end_idx is not None:
            seg_t = self.time[self.start_idx:self.end_idx]
            seg_v = self.values[self.start_idx:self.end_idx]

        # Plot the segment first so it appears first in the legend
        self.ax.scatter(seg_t, seg_v, color='blue', alpha=0.8, label="Data Used for Trendline")

        # Plot all data as "Unused"
        # We'll just plot the entire set as gray, so the overlapping region is effectively replaced by the blue points
        self.ax.scatter(self.time, self.values, color='gray', alpha=0.6, label="Unused Data")

        # Compute the linear fit line if we have 2 or more points
        if len(seg_t) > 1:
            _, ccoeffs = calculate_r2(seg_t, seg_v)
            slope, intercept = ccoeffs
            x_min, x_max = np.min(self.time), np.max(self.time)
            x_span = x_max - x_min
            x_left = x_min - 0.05 * x_span
            x_right = x_max + 0.05 * x_span
            x_ext = np.linspace(x_left, x_right, 200)
            y_ext = slope * x_ext + intercept
            self.ax.plot(x_ext, y_ext, color='red', label="Trendline")

        # Axis labels
        self.ax.set_xlabel("Time (Seconds)")
        self.ax.set_ylabel("Absorbance")
        self.ax.legend()
        self.ax.grid(True)
        self.canvas.draw()

    def apply_override(self):
        """User changed the start/end range (Apply Range button or Ctrl+Enter)."""
        start_txt = self.start_entry.get().strip()
        end_txt   = self.end_entry.get().strip()
        if not start_txt or not end_txt:
            messagebox.showwarning("Invalid Input", "Please enter both start and end times.")
            return

        try:
            user_start = float(start_txt)
            user_end   = float(end_txt)
        except ValueError:
            messagebox.showwarning("Invalid Input", "Start/End must be numeric.")
            return

        if user_start > user_end:
            messagebox.showwarning("Invalid Range", "Start time must be <= end time.")
            return

        self.last_start_val = user_start
        self.last_end_val   = user_end

        def nearest_index_for_time(t):
            return int(np.argmin(np.abs(self.time - t)))

        si = nearest_index_for_time(user_start)
        ei = nearest_index_for_time(user_end)
        if si >= ei or (ei - si) < 2:
            messagebox.showwarning("Invalid Range", "Not enough points in that range.")
            return

        self.start_idx = si
        self.end_idx   = ei

        seg_t = self.time[si:ei]
        seg_v = self.values[si:ei]
        new_r2, ccoeffs = calculate_r2(seg_t, seg_v)
        slope = ccoeffs[0]
        self.slope = slope
        self.r2_value = new_r2

        idx_disp = f"{self.current_col_idx+1}/{self.num_cols}"
        col_name = self.columns[self.current_col_idx]
        msg = (
            f"Column {idx_disp}: {col_name}\n"
            f"{self.r2_cutoff} (R² cutoff)\n"
            "Best Linear Region (User Override):\n"
            f"{new_r2:.6f}\n"
            f"{slope:.3e}"
        )
        self.col_label.config(text=msg)
        self.plot_current()

    def accept_segment(self):
        """User accepted the current region (Enter)."""
        col_name = self.columns[self.current_col_idx]
        if (self.slope is not None) and (self.start_idx is not None) and (self.end_idx is not None):
            abs_slope = abs(self.slope)
            start_val = float(self.time[self.start_idx])
            end_val   = float(self.time[self.end_idx - 1])
            n_pts     = self.end_idx - self.start_idx
            self.results_dict[col_name] = {
                "Slope": abs_slope,
                "StartTime": start_val,
                "EndTime": end_val,
                "NPoints": n_pts
            }
        else:
            self.results_dict[col_name] = {
                "Slope": None,
                "StartTime": None,
                "EndTime": None,
                "NPoints": 0
            }

        next_idx = self.current_col_idx + 1
        if next_idx < self.num_cols:
            self.load_column(next_idx)
        else:
            self.finish_all()

    def finish_all(self):
        """All columns finished; close GUI."""
        messagebox.showinfo("Done", "All columns processed.")
        self.root.quit()
        self.root.destroy()

    def quit_all(self):
        """User clicked Quit; store None for current column, exit immediately."""
        col_name = self.columns[self.current_col_idx]
        self.results_dict[col_name] = {
            "Slope": None,
            "StartTime": None,
            "EndTime": None,
            "NPoints": 0
        }
        self.root.quit()
        self.root.destroy()

    def mainloop(self):
        self.root.mainloop()

    def force_exit(self):
        """If user closes window, forcibly exit."""
        sys.exit(0)

    def enter_accept(self, event):
        """Enter => Accept current region."""
        self.accept_segment()

    def ctrl_enter(self, event):
        """Ctrl+Enter => Apply Range override."""
        self.apply_override()

    def ctrl_left(self, event):
        """Expand left boundary by 1 data point."""
        if self.start_idx is not None and self.end_idx is not None:
            new_start = max(0, self.start_idx - 1)
            if new_start < self.end_idx - 1:
                self.start_idx = new_start
                self.update_segment_display("Best Linear Region (User Override)")

    def ctrl_right(self, event):
        """Contract left boundary by 1 data point."""
        if self.start_idx is not None and self.end_idx is not None:
            new_start = self.start_idx + 1
            if new_start < self.end_idx - 1:
                self.start_idx = new_start
                self.update_segment_display("Best Linear Region (User Override)")

    def shift_left(self, event):
        """Contract right boundary by 1 data point."""
        if self.start_idx is not None and self.end_idx is not None:
            new_end = self.end_idx - 1
            if new_end > self.start_idx + 1:
                self.end_idx = new_end
                self.update_segment_display("Best Linear Region (User Override)")

    def shift_right(self, event):
        """Expand right boundary by 1 data point."""
        if self.start_idx is not None and self.end_idx is not None:
            new_end = self.end_idx + 1
            if new_end <= len(self.time):
                self.end_idx = new_end
                self.update_segment_display("Best Linear Region (User Override)")

    def ctrl_up(self, event):
        """Increase R² cutoff, re-run best segment."""
        inc = get_r2_increment(self.r2_cutoff)
        self.r2_cutoff += inc
        self.r2_cutoff = round(self.r2_cutoff, 8)
        if self.r2_cutoff > 1:
            self.r2_cutoff = 1.0
        self.recompute_auto_segment()

    def ctrl_down(self, event):
        """Decrease R² cutoff, re-run best segment."""
        inc = get_r2_increment(self.r2_cutoff)
        new_val = self.r2_cutoff - inc
        if new_val < 0:
            new_val = 1e-6
        self.r2_cutoff = round(new_val, 8)
        self.recompute_auto_segment()

    def update_segment_display(self, prefix="Best Linear Region (User Override)"):
        """After shifting the segment by one data point, recalc slope & R², re-plot."""
        seg_t = self.time[self.start_idx:self.end_idx]
        seg_v = self.values[self.start_idx:self.end_idx]
        r2, ccoeffs = calculate_r2(seg_t, seg_v)
        slope = ccoeffs[0]
        self.slope = slope
        self.r2_value = r2

        idx_disp = f"{self.current_col_idx+1}/{self.num_cols}"
        col_name = self.columns[self.current_col_idx]
        msg = (
            f"Column {idx_disp}: {col_name}\n"
            f"{self.r2_cutoff} (R² cutoff)\n"
            f"{prefix}:\n"
            f"{r2:.6f}\n"
            f"{slope:.3e}"
        )
        self.col_label.config(text=msg)
        self.plot_current()


def run_app():
    """
    PlateRate main entry point for multi-file mode:
      1) User selects an Excel file exported from the Varioskan Lux.
      2) The file is processed, 'avg. time [s]' used as index if present.
      3) The user interacts with the GUI for each column to set the best linear region.
      4) The user can select more files or quit.
      5) All results are combined into one CSV at the end.
    """
    root = tk.Tk()
    root.withdraw()

    all_results = []

    while True:
        excel_path = filedialog.askopenfilename(
            title="Open Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not excel_path:
            ans = messagebox.askyesno("Question", "No file selected. Try another file?")
            if ans:
                continue
            else:
                break

        try:
            df = process_excel(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process file:\n{e}")
            ans = messagebox.askyesno("Question", "Try another file?")
            if ans:
                continue
            else:
                break

        gui = TrendlineGUI(df)
        gui.mainloop()

        results_dict = gui.results_dict
        for sample_name, info in results_dict.items():
            all_results.append({
                "File": excel_path,
                "Sample": sample_name,
                "Slope": info["Slope"],
                "StartTime": info["StartTime"],
                "EndTime": info["EndTime"],
                "NPoints": info["NPoints"]
            })

        ans = messagebox.askyesno("Question", "Do you want to process another file?")
        if not ans:
            break

    if not all_results:
        messagebox.showinfo("No Results", "No results to save. Exiting.")
        root.destroy()
        sys.exit(0)

    default_output = "combined_results.csv"
    save_path = filedialog.asksaveasfilename(
        title="Save Combined Results As",
        initialfile=default_output,
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if save_path:
        out_df = pd.DataFrame(all_results, columns=["File","Sample","Slope","StartTime","EndTime","NPoints"])
        try:
            out_df.to_csv(save_path, index=False)
            messagebox.showinfo("Saved", f"Results saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file:\n{e}")
    else:
        messagebox.showinfo("Not Saved", "Results not saved.")

    root.destroy()
    sys.exit(0)


if __name__ == "__main__":
    run_app()
