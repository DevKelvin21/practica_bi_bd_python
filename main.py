import zipfile
import pandas as pd
from io import BytesIO

# Define the path to the ZIP file
zip_path = "./Reportes.zip"

# Set to store unique movie names
unique_movies = set()

# Open the ZIP file and process each Excel file in memory
with zipfile.ZipFile(zip_path, "r") as zip_ref:
    for file_name in zip_ref.namelist():
        if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
            with zip_ref.open(file_name) as file:
                print(f"Processing file: {file_name}")
                
                # Read the Excel file from memory
                df = pd.read_excel(BytesIO(file.read()), skiprows=3, engine='openpyxl')  # Skip first 3 rows to reach headers
                
                # Ensure column B (index 1) exists
                if df.shape[1] > 1:
                    unique_movies.update(df.iloc[:, 1].dropna().unique())  # Extract unique movie names

# Write the unique movie names to a text file
output_file = "./unique_movies.txt"
with open(output_file, "w", encoding="utf-8") as f:
    for movie in sorted(map(str, unique_movies)):
        f.write(movie + "\n")

print(f"Unique movies list saved to {output_file}")
