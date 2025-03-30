import os
import zipfile
import subprocess
import sys
import pandas as pd
from io import BytesIO

# Function to install required dependencies
def install_dependencies():
    required_packages = ["pandas", "openpyxl"]
    for package in required_packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Installing missing package: {package}")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Function to process the ZIP file and extract unique movie names
def process_zip_file(zip_path):
    unique_movies = set()

    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
                    try:
                        with zip_ref.open(file_name) as file:
                            print(f"Processing file: {file_name}")

                            df = pd.read_excel(BytesIO(file.read()), skiprows=3, engine='openpyxl')  # Skip first 3 rows

                            if df.shape[1] > 1:
                                unique_movies.update(df.iloc[:, 1].dropna().unique())  # Extract unique movie names

                    except (ValueError, pd.errors.ExcelFileError) as e:
                        print(f"Warning: Could not process {file_name} due to an Excel error: {e}")

    except FileNotFoundError:
        print(f"Error: The file {zip_path} was not found. Please check the file path and try again.")
        sys.exit(1)
    except zipfile.BadZipFile:
        print(f"Error: The file {zip_path} is not a valid ZIP archive or is corrupted.")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

    return unique_movies

# Function to write unique movies to a text file
def write_output(output_file, unique_movies):
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            for movie in sorted(unique_movies, key=lambda x: str(x).lower()):  # Sort alphabetically (case-insensitive)
                f.write(str(movie) + "\n")
        print(f"Unique movies list saved to {output_file}")
    except IOError as e:
        print(f"Error: Could not write to {output_file}: {e}")

def main():
    # Step 1: Install dependencies
    install_dependencies()

    # Step 2: Define paths
    zip_path = "./Reportes.zip"
    output_file = "./unique_movies.txt"

    # Step 3: Process ZIP file
    if not os.path.exists(zip_path):
        print(f"Error: The file {zip_path} is missing. Please provide the ZIP file and try again.")
        sys.exit(1)

    unique_movies = process_zip_file(zip_path)

    # Step 4: Write results
    write_output(output_file, unique_movies)

if __name__ == "__main__":
    main()
