import os
import numpy as np
import originpro as op
from tkinter import Tk, filedialog, messagebox
import time

# Open file dialog
root = Tk()
root.withdraw()
while True:
    folder_path = filedialog.askdirectory(title="Select folder with Raman raw data")
    if folder_path and os.path.isdir(folder_path):
        break
    messagebox.showerror("Error", "Invalid folder selection! Please select a valid folder.")

# Normalize folder path
folder_path = os.path.normpath(folder_path)

# Get .txt files
file_paths = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".txt")]
if not file_paths:
    print("Error: No .txt files found.")
    exit()

# Create output folder
processed_folder = os.path.join(folder_path, "Processed Raman data")
os.makedirs(processed_folder, exist_ok=True)

# Verify folder creation
if not os.path.exists(processed_folder):
    print(f"Error: Processed folder was not created: {processed_folder}")
    exit()

# Define Origin project path
project_path = os.path.join(processed_folder, "Raman_Analysis.opju")

# Check folder permissions
test_file_path = os.path.join(processed_folder, "test.txt")
try:
    with open(test_file_path, "w") as test_file:
        test_file.write("Write test successful.")
    os.remove(test_file_path)
except Exception as e:
    print(f"Error: Cannot write to folder {processed_folder}: {e}")
    exit()

# Open Origin
op.new()

# Initialize error log file
error_log_path = os.path.join(processed_folder, "error_log.txt")

# Process files
for file_path in file_paths:
    file_path = os.path.normpath(file_path)
    try:
        filename = os.path.splitext(os.path.basename(file_path))[0]

        # Check if first row is a header
        with open(file_path, "r") as f:
            first_line = f.readline()
        skip_rows = 1 if any(c.isalpha() for c in first_line) else 0

        # Load data
        data = np.loadtxt(file_path, skiprows=skip_rows)
        if data.ndim != 2 or data.shape[1] < 2:
            error_message = f"Error: Invalid data format in {file_path}\n"
            with open(error_log_path, "a") as log_file:
                log_file.write(error_message)
            continue
        
        wave, intensity = data[:, 0], data[:, 1]

        # Normalize intensity
        max_intensity = np.max(intensity)
        if max_intensity == 0:
            warning_message = f"Warning: Maximum intensity is zero in {file_path}, skipping normalization.\n"
            with open(error_log_path, "a") as log_file:
                log_file.write(warning_message)
        else:
            intensity = intensity / max_intensity

        # Create new worksheet
        wks = op.new_sheet("w", filename)
        wks.from_list(0, wave, "Raman Shift (cm⁻¹)")
        wks.from_list(1, intensity, "Intensity (a.u.)")

        # Create and attach graph
        graph = op.new_graph(template="Line")
        layer = graph[0]
        plot = layer.add_plot(wks, 1, 0)
        plot.color = "red"
        plot.set_str("connect", "spline")  # Set connect type to b-spline
        plot.set_int("line.width", 2)
        layer.x_label = "Raman Shift (cm⁻¹)"
        layer.y_label = "Intensity (a.u.)"
        layer.rescale()
        layer.set_xlim(0, np.max(wave))
        layer.set_ylim(0, 1.1, 0.2)
    except Exception as e:
        error_message = f"Error processing {file_path}: {e}\n"
        with open(error_log_path, "a") as log_file:
            log_file.write(error_message)

# Save project with retry mechanism
if op.project:
    print(f"Attempting to save Origin project to: {project_path}")
    for attempt in range(3):
        try:
            op.project.save(project_path)
            time.sleep(2)
            if os.path.exists(project_path):
                print(f"Successfully saved Origin project to: {project_path}")
                break
            else:
                print(f"Warning: Attempt {attempt + 1} failed, retrying...")
        except Exception as save_error:
            error_message = f"Error saving project (attempt {attempt + 1}): {save_error}\n"
            with open(error_log_path, "a") as log_file:
                log_file.write(error_message)
            if attempt == 2:
                print("Failed to save project after 3 attempts.")
else:
    error_message = "Error: No active Origin project to save.\n"
    with open(error_log_path, "a") as log_file:
        log_file.write(error_message)

print("Closing Origin...")
time.sleep(2)
op.exit()
