import os
import zipfile
import pandas as pd
from io import BytesIO

# Define the path to the ZIP file
zip_path = "./Reportes.zip"

# Check if the ZIP file exists
if not os.path.exists(zip_path):
    print(f"Error: The file {zip_path} is missing. Please provide the ZIP file and try again.")
    exit(1)

# Set to store unique movie names
unique_movies = set()

try:
    # Open the ZIP file and process each Excel file in memory
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        for file_name in zip_ref.namelist():
            if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
                try:
                    with zip_ref.open(file_name) as file:
                        print(f"Processing file: {file_name}")

                        # Read the Excel file from memory
                        df = pd.read_excel(BytesIO(file.read()), skiprows=3, engine='openpyxl')  # Skip first 3 rows to reach headers

                        # Ensure column B (index 1) exists
                        if df.shape[1] > 1:
                            unique_movies.update(df.iloc[:, 1].dropna().unique())  # Extract unique movie names

                except (ValueError, pd.errors.ExcelFileError) as e:
                    print(f"Warning: Could not process {file_name} due to an Excel error: {e}")

except FileNotFoundError:
    print(f"Error: The file {zip_path} was not found. Please check the file path and try again.")
    exit(1)
except zipfile.BadZipFile:
    print(f"Error: The file {zip_path} is not a valid ZIP archive or is corrupted.")
    exit(1)
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    exit(1)

# Write the unique movie names to a text file
output_file = "./unique_movies.txt"
try:
    with open(output_file, "w", encoding="utf-8") as f:
        for movie in sorted(unique_movies, key=lambda x: str(x).lower()):  # Sort alphabetically (case-insensitive)
            f.write(str(movie) + "\n")
    print(f"Unique movies list saved to {output_file}")
except IOError as e:
    print(f"Error: Could not write to {output_file}: {e}")
