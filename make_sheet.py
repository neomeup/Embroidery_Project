'''
To make an excel from a photo for broderie

- saves the excel sheet in this same folder

requires changing the source of the colormath package because of a depreciation 
-- here in the last line from asscalar to the item

# noinspection PyPep8Naming
def delta_e_cie2000(color1, color2, Kl=1, Kc=1, Kh=1):
    """
    Calculates the Delta E (CIE2000) of two colors.
    """
    color1_vector = _get_lab_color1_vector(color1)
    color2_matrix = _get_lab_color2_matrix(color2)
    delta_e = color_diff_matrix.delta_e_cie2000(
        color1_vector, color2_matrix, Kl=Kl, Kc=Kc, Kh=Kh)[0]
    return delta_e.item()
'''
import time

# Calculate the start time
start = time.time()


import openpyxl
from openpyxl.styles import PatternFill, Font
from PIL import Image
from colormath.color_objects import sRGBColor, LabColor
from colormath.color_conversions import convert_color
from colormath.color_diff import delta_e_cie2000
from sklearn.cluster import KMeans
import numpy as np
from pathlib import Path

import user_parameters
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

# Find the next available output filename
counter = 1

while True:
    OUTPUT_PATH = OUTPUT_DIR / f"output_{counter}.xlsx"

    if not OUTPUT_PATH.exists():
        break

    counter += 1


# Processed Master List: Filters based on ownership status
if user_parameters.FILTER_UNOWNED_COLORS:
    PROCESSED_MASTER_LIST = [color for color in MASTER_LIST if color[3]]
else:
    PROCESSED_MASTER_LIST = MASTER_LIST.copy()

def find_prominent_colors(image_path, num_colors):

    #Find the most prominent colors in the image using k-means clustering.

    image = Image.open(image_path).convert("RGB")
    image_array = np.array(image)
    pixels = image_array.reshape(-1, 3)

    kmeans = KMeans(n_clusters=num_colors, random_state=42)
    kmeans.fit(pixels)
    cluster_centers = kmeans.cluster_centers_

    return cluster_centers

def match_colors_to_master(cluster_centers, master_list):

    #Match cluster centers to the closest colors in the master list.

    matched_colors = []
    for center in cluster_centers:
        lab_center = convert_color(sRGBColor(center[0] / 255, center[1] / 255, center[2] / 255), LabColor)
        min_distance = float('inf')
        best_match = None

        for color in master_list:
            master_rgb = color[0]
            master_lab = convert_color(sRGBColor(master_rgb[0] / 255, master_rgb[1] / 255, master_rgb[2] / 255), LabColor)
            distance = delta_e_cie2000(lab_center, master_lab)
            if distance < min_distance:
                min_distance = distance
                best_match = color

        if best_match not in matched_colors:
            matched_colors.append(best_match)

    return matched_colors

# Modify the palette processing logic
if not user_parameters.USE_ALL_COLORS:
    # Find prominent colors in the input image
    cluster_centers = find_prominent_colors(INPUT_PATH, user_parameters.MAX_COLORS)

    # Match the cluster centers to the master list
    PROCESSED_MASTER_LIST = match_colors_to_master(cluster_centers, MASTER_LIST)

# Precompute LAB values for the updated PROCESSED_MASTER_LIST
PRECOMPUTED_LAB = {
    tuple(color[0]): convert_color(sRGBColor(color[0][0] / 255, color[0][1] / 255, color[0][2] / 255), LabColor)
    for color in PROCESSED_MASTER_LIST
}

def closest_color_lab(r, g, b):
    input_color = convert_color(sRGBColor(r / 255, g / 255, b / 255), LabColor)
    min_distance = float('inf')
    best_match = None

    for rgb, lab in PRECOMPUTED_LAB.items():
        distance = delta_e_cie2000(input_color, lab, Kl=1, Kc=1, Kh=2)
        if distance < min_distance:
            min_distance = distance
            best_match = rgb

    return best_match

def create_sheet(input_path, output_path, grid_rows, grid_cols):
    image = Image.open(input_path).convert("RGB")
    image_resized = image.resize((grid_cols, grid_rows))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Cross Stitch"

    used_colors = set()
    for row in range(grid_rows):
        for col in range(grid_cols):
            r, g, b = image_resized.getpixel((col, row))
            closest_rgb = closest_color_lab(r, g, b)

            hex_color = next(color[1] for color in PROCESSED_MASTER_LIST if tuple(color[0]) == closest_rgb)
            name = next(color[2] for color in PROCESSED_MASTER_LIST if tuple(color[0]) == closest_rgb)
            ownership = next(color[3] for color in PROCESSED_MASTER_LIST if tuple(color[0]) == closest_rgb)

            used_colors.add((name, hex_color, ownership))

            cell = sheet.cell(row=row + 1, column=col + 1)
            cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

    for col in range(1, grid_cols + 1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 2.5
    for row in range(1, grid_rows + 1):
        sheet.row_dimensions[row].height = 15

    key_start_col = grid_cols + 2
    sheet.cell(row=1, column=key_start_col).value = "Color Key"
    sheet.cell(row=1, column=key_start_col).font = Font(bold=True)

    key_row = 2
    max_name_length = 0

    for name, hex_color, ownership in sorted(used_colors):
        sheet.cell(row=key_row, column=key_start_col).value = name
        max_name_length = max(max_name_length, len(name))

        key_cell = sheet.cell(row=key_row, column=key_start_col + 1)
        key_cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
        ownership_value = "Owned" if ownership else "Unowned"
        sheet.cell(row=key_row, column=key_start_col + 2).value = ownership_value
        key_row += 1

    sheet.column_dimensions[openpyxl.utils.get_column_letter(key_start_col)].width = max_name_length + 2

    workbook.save(output_path)
    print(f"Excel file saved: {output_path}")

# Run the program
create_sheet(INPUT_PATH, OUTPUT_PATH, user_parameters.GRID_ROWS, user_parameters.GRID_COLS)

# Calculate the end time and time taken
end = time.time()
length = end - start

# Show the results : this can be altered however you like
print("It took", length, "seconds!")