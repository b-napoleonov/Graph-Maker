import os
import numpy as np
import originpro as op
from tkinter import Tk, filedialog
import time

# Open file dialog
root = Tk()
root.withdraw()
folder_path = filedialog.askdirectory(title="Select folder with Raman raw data")

# Check for valid selection
if not folder_path or not os.path.isdir(folder_path):
    print("Error: Invalid folder selection!")
    exit()

# Normalize folder path to ensure consistent use of forward slashes
folder_path = os.path.normpath(folder_path)

# Get .txt files
file_paths = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".txt")]
if not file_paths:
    print("Error: No .txt files found.")
    exit()

# Create output folder
processed_folder = os.path.join(folder_path, "Processed Raman data")
processed_folder = os.path.normpath(processed_folder)  # Normalize path

try:
    os.makedirs(processed_folder, exist_ok=True)
except Exception as e:
    print(f"Error: Could not create processed folder {processed_folder}: {e}")
    exit()

# Check if folder was created successfully
if not os.path.exists(processed_folder):
    print(f"Error: Processed folder was not created: {processed_folder}")
    exit()

# Define Origin project path
project_path = os.path.join(processed_folder, "Raman_Analysis.opju")
project_path = os.path.normpath(project_path)  # Normalize path

# Check folder permissions
test_file_path = os.path.join(processed_folder, "test.txt")
test_file_path = os.path.normpath(test_file_path)
try:
    with open(test_file_path, "w") as test_file:
        test_file.write("Write test successful.")
    os.remove(test_file_path)
except Exception as e:
    print(f"Error: Cannot write to folder {processed_folder}: {e}")
    exit()

# Open Origin
op.new()

# Process files
try:
    for file_path in file_paths:
        file_path = os.path.normpath(file_path)  # Normalize path
        try:
            filename = os.path.splitext(os.path.basename(file_path))[0]
            
            # Check if first row is a header
            with open(file_path, "r") as f:
                first_line = f.readline()
            skip_rows = 1 if any(c.isalpha() for c in first_line) else 0

            # Load data
            data = np.loadtxt(file_path, skiprows=skip_rows)
            wave, intensity = data[:, 0], data[:, 1]

            # Create new worksheet
            wks = op.new_sheet("w", filename)
            wks.from_list(0, wave, "Raman Shift (cm⁻¹)")
            wks.from_list(1, intensity, "Intensity (a.u.)")

            # Create and attach graph
            graph = op.new_graph(template="Line")
            layer = graph[0]
            layer.add_plot(wks, 1, 0)
            layer.x_label = "Raman Shift (cm⁻¹)"
            layer.y_label = "Intensity (a.u.)"
            layer.rescale()
            layer.set_xlim(0, "auto")  # Ensure no data below 0 is shown
        
        except Exception as e:
            print(f"Error processing {file_path}: {e}")

# Save project
finally:
    if op.project:
        print(f"Attempting to save Origin project to: {project_path}")
        try:
            # Ensure the project is saved with an explicit path
            op.project.save(project_path)  # Explicitly provide the path
            time.sleep(2)  # Allow save to complete
            
            # Check if the file was created
            if os.path.exists(project_path):
                print(f"Successfully saved Origin project to: {project_path}")
            else:
                print(f"Error: Origin project file was not created at {project_path}!")

        except Exception as save_error:
            print(f"Error saving project: {save_error}")
    else:
        print("Error: No active Origin project to save.")

    print("Closing Origin...")
    time.sleep(2)
    op.exit()
