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

# Function to process the ZIP file and aggregate attendee data
def process_zip_file(zip_path):
    movie_attendees = {}

    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            for file_name in zip_ref.namelist():
                if file_name.endswith(".xlsx") and not file_name.startswith("~$"):
                    try:
                        with zip_ref.open(file_name) as file:
                            print(f"Processing file: {file_name}")

                            df = pd.read_excel(BytesIO(file.read()), skiprows=3, engine='openpyxl')  # Skip first 3 rows

                            if df.shape[1] > 25:  # Ensure all necessary columns exist
                                relevant_data = df.iloc[:, [1, 3, 17, 21, 25]].dropna()  # Columns B, D, R, V, Z

                                # Filter for release week (week number == 1)
                                release_week_data = relevant_data[relevant_data.iloc[:, 1] == 1]

                                # Filter for second week (week number == 2)
                                second_week_data = relevant_data[relevant_data.iloc[:, 1] == 2]

                                # Aggregate attendees by movie title
                                for _, row in release_week_data.iterrows():
                                    movie = row.iloc[0]
                                    day_attendees = row.iloc[2]
                                    weekend_attendees = row.iloc[3]
                                    week_attendees = row.iloc[4]

                                    if movie in movie_attendees:
                                        movie_attendees[movie][0] += day_attendees
                                        movie_attendees[movie][1] += weekend_attendees
                                        movie_attendees[movie][2] += week_attendees
                                    else:
                                        movie_attendees[movie] = [day_attendees, weekend_attendees, week_attendees, 0]  # Default week 2 attendees to 0

                                # Aggregate second week attendees
                                for _, row in second_week_data.iterrows():
                                    movie = row.iloc[0]
                                    week2_attendees = row.iloc[4]

                                    if movie in movie_attendees:
                                        movie_attendees[movie][3] += week2_attendees
                                    else:
                                        movie_attendees[movie] = [0, 0, 0, week2_attendees]  # Default other values to 0

                    except Exception as e:
                        print(f"Warning: Could not process {file_name} due to an error: {e}")

    except FileNotFoundError:
        print(f"Error: The file {zip_path} was not found. Please check the file path and try again.")
        sys.exit(1)
    except zipfile.BadZipFile:
        print(f"Error: The file {zip_path} is not a valid ZIP archive or is corrupted.")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

    return movie_attendees

# Function to write aggregated attendee data to a text file
def write_output(output_file, movie_attendees):
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("Movie, Release Day Attendees, Weekend Attendees, Weekly Attendees (Week 1), Weekly Attendees (Week 2)\n")
            for movie, attendees in sorted(movie_attendees.items()):
                f.write(f"{movie}, {attendees[0]}, {attendees[1]}, {attendees[2]}, {attendees[3]}\n")
        print(f"Aggregated movie attendees list saved to {output_file}")
    except IOError as e:
        print(f"Error: Could not write to {output_file}: {e}")

def main():
    # Step 1: Install dependencies
    install_dependencies()

    # Step 2: Define paths
    zip_path = "./Reportes.zip"
    output_file = "./movie_attendees.txt"

    # Step 3: Process ZIP file
    if not os.path.exists(zip_path):
        print(f"Error: The file {zip_path} is missing. Please provide the ZIP file and try again.")
        sys.exit(1)

    movie_attendees = process_zip_file(zip_path)

    # Step 4: Write results
    write_output(output_file, movie_attendees)

if __name__ == "__main__":
    main()
