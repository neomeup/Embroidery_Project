'''
Optimization script for the broderie program
'''
from datetime import datetime

c = datetime.now()

# Displays Time
current_time = c.strftime('%H:%M:%S')
print('Current Time is:', current_time)

import time

# Calculate the start time
start = time.time()

import optuna
from colormath.color_objects import sRGBColor, LabColor
from colormath.color_conversions import convert_color
from colormath.color_diff import delta_e_cie2000
from PIL import Image
import numpy as np
from pathlib import Path


import master_color_list

MASTER_LIST = master_color_list.MASTER_LIST


# Locate the project directory
BASE_DIR = Path(__file__).resolve().parent

# Input directory
INPUT_DIR = BASE_DIR / "input_photo"

# Output directory
OUTPUT_DIR = BASE_DIR / "output_templates"

# Create output directory if it doesn't exist
OUTPUT_DIR.mkdir(exist_ok=True)

# Find input photo
input_files = [file for file in INPUT_DIR.iterdir() if file.is_file()]

if len(input_files) == 0:
    raise FileNotFoundError(
        "No photo found in input folder. Please add a photo to convert."
    )

if len(input_files) > 1:
    raise ValueError(
        "Multiple files exist in input folder, please select a single photo to convert"
    )

INPUT_PATH = input_files[0]



# Function to calculate color match error (optimized with vectorization)
def calculate_error(input_image, master_list, Kl, Kc, Kh):
    # Open image and resize to speed up processing (downsample to smaller resolution, e.g., 100x125)
    image = Image.open(input_image).convert("RGB")
    #image = image.resize((100, 125))  # Adjust to a smaller resolution
    pixels = np.array(image).reshape(-1, 3)  # Flatten the image to (num_pixels, 3)

    # Convert master colors to Lab color space
    master_lab_list = []
    for master_color in master_list:
        master_rgb = master_color[0]
        master_lab = convert_color(sRGBColor(master_rgb[0] / 255, master_rgb[1] / 255, master_rgb[2] / 255), LabColor)
        master_lab_list.append(np.array([master_lab.lab_l, master_lab.lab_a, master_lab.lab_b]))

    # Vectorize the calculation of color differences
    input_rgb = pixels / 255.0  # Normalize input image pixels
    input_lab = np.array([convert_color(sRGBColor(r, g, b), LabColor).get_value_tuple() for r, g, b in input_rgb])
    input_lab = np.array(input_lab)

    # Calculate delta E for each pixel and all master colors using broadcasting
    input_l, input_a, input_b = np.split(input_lab, 3, axis=1)
    master_l, master_a, master_b = np.split(np.array(master_lab_list), 3, axis=1)

    # Calculate delta E for all pixels and all master colors
    diff_l = input_l - master_l.T
    diff_a = input_a - master_a.T
    diff_b = input_b - master_b.T
    delta_e = np.sqrt(diff_l ** 2 + diff_a ** 2 + diff_b ** 2)

    print(f"First few delta E values: {delta_e[:5]}","\n")  # Print the first few delta E values


    # For each pixel, find the minimum distance (best matching master color)
    min_distances = np.min(delta_e, axis=1)

    print(f"Error Stats - min: {np.min(min_distances)}, max: {np.max(min_distances)}, mean: {np.mean(min_distances)}","\n")  # Add this line
    
    # Return the average delta E (error) across all pixels
    return np.mean(min_distances)

# Objective function for Optuna (with debug prints)
def objective(trial):
    Kl = trial.suggest_float("Kl", 0.5, 2.0)
    Kc = trial.suggest_float("Kc", 0.5, 2.0)
    Kh = trial.suggest_float("Kh", 0.5, 2.0)

    # Calculate error and print trial information
    error = calculate_error(INPUT_PATH, MASTER_LIST, Kl, Kc, Kh)

    return error

# Run the optimization with parallelization and a reduced number of trials for testing
if __name__ == "__main__":
    start = time.time()

    # Create study and optimize using parallelization (use all CPU cores)
    study = optuna.create_study(direction="minimize")
    study.optimize(objective, n_trials=5, n_jobs=-1)  # Use 20 trials for testing and all cores

    # Best parameters
    print("Best parameters found:")
    print(study.best_params)
    print(f"Best error: {study.best_value}")


# Calculate the end time and time taken
end = time.time()
length = end - start

print("It took", length, "seconds!")
