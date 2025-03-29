import os
import pandas as pd

# Define the folder containing the reports
folder_path = "./Reportes"

# Set to store unique movie names
unique_movies = set()

# Iterate over all Excel files in the folder
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") and not file.startswith("~$"):
        file_path = os.path.join(folder_path, file)
        print(f"Reading file: {file_path}")
        
        # Read the Excel file
        df = pd.read_excel(file_path, skiprows=3, engine='openpyxl')  # Skip first 3 rows to reach headers
        
        # Ensure column B (index 1) exists
        if df.shape[1] > 1:
            unique_movies.update(df.iloc[:, 1].dropna().unique())  # Extract unique movie names

# Write the unique movie names to a text file
output_file = "./unique_movies.txt"
with open(output_file, "w", encoding="utf-8") as f:
    for movie in sorted(map(str, unique_movies)):
        f.write(movie + "\n")

print(f"Unique movies list saved to {output_file}")
