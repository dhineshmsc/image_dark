import os
import cv2
import numpy as np
from openpyxl import Workbook
from datetime import datetime

# Source folder containing the images
source_folder = 'source_images'

# Output folder for the Excel file
output_folder = 'output'

if not os.path.exists(source_folder):
    os.makedirs(source_folder)
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# List files in the source folder
files = os.listdir(source_folder)

# Filter only JPG and PNG files
image_files = [file for file in files if file.lower().endswith(('.jpg', '.png'))]

# Create a workbook
wb = Workbook()
ws = wb.active

# Add headers
ws.append(["Image Name", "Average Pixel Intensity", "Status"])

# Loop through image files
for file_name in image_files:
    # Read the image
    image_path = os.path.join(source_folder, file_name)
    image = cv2.imread(image_path)
    
    # Convert the image to grayscale
    gray_image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    
    # Calculate average pixel intensity
    average_intensity = int(np.mean(gray_image))
    
    # Determine status based on average intensity
    if average_intensity >= 50:
        status = "Normal"
    elif average_intensity >= 20:
        status = "Dark"
    else:
        status = "Very_Dark"
    
    #print(f"Image: {file_name}, Average Pixel Intensity: {average_intensity}, Status: {status}")
    
    ws.append([file_name, f"{average_intensity}%", status])

# Generate a unique filename based on current date and time
output_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"

# Save the workbook to the output folder
output_path = os.path.join(output_folder, output_filename)
wb.save(output_path)

print(f"Successfully created image detail in Excel. Path is the output folder in named :'{output_filename}'.")

#print(f"Excel file '{output_filename}' created with image details in the output folder.")
