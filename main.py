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

                                # Aggregate total attendees for each movie, ensuring valid weeks
                                for _, row in relevant_data.iterrows():
                                    movie = row.iloc[0]
                                    week_number = row.iloc[1]  # Week number
                                    release_day_attendees = row.iloc[2]  # Release day
                                    weekend_attendees = row.iloc[3]  # Weekend
                                    weekly_attendees = row.iloc[4]  # Weekly

                                    # Only aggregate for weeks where the week number >= 1
                                    if week_number >= 1:
                                        if movie not in movie_attendees:
                                            movie_attendees[movie] = [0, 0, 0, 0, 0]  # Initialize all columns
                                        
                                        # Aggregate the data for each movie
                                        if week_number == 1:
                                            movie_attendees[movie][0] += release_day_attendees  # Release day
                                            movie_attendees[movie][1] += weekend_attendees  # Weekend
                                            movie_attendees[movie][2] += weekly_attendees  # Week 1
                                        elif week_number == 2:
                                            movie_attendees[movie][3] += weekly_attendees  # Week 2

                                        # Aggregate total attendees (week >= 1)
                                        movie_attendees[movie][4] += release_day_attendees + weekend_attendees + weekly_attendees

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

    # Calculate remaining attendees and percentages after processing all movies
    for movie, attendees in movie_attendees.items():
        total = attendees[4]
        week1 = attendees[2]
        week2 = attendees[3]
        
        # Calculate remaining attendees
        remaining = max(0, total - (week1 + week2))
        movie_attendees[movie].append(remaining)

        # Calculate Week 1 and Week 2 percentages based on total
        week1_percentage = (week1 / total * 100) if total > 0 else 0
        week2_percentage = (week2 / total * 100) if total > 0 else 0
        
        # Add the percentages to the movie's data
        movie_attendees[movie].append(week1_percentage)
        movie_attendees[movie].append(week2_percentage)

    return movie_attendees

# Function to write aggregated attendee data to a text file
def write_output(output_file, movie_attendees):
    try:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("Película, Dia Inicial, Asistentes 1er Fin de Semana, Asistentes 1er Semana, Porcentaje Semana 1, Asistentes Semanales (Semana 2), Porcentaje Semana 2, Asistentes Restantes, Asistentes Totales\n")
            for movie, attendees in sorted(movie_attendees.items()):
                f.write(f"{movie}, {attendees[0]}, {attendees[1]}, {attendees[2]}, {attendees[6]:.2f}%, {attendees[3]}, {attendees[7]:.2f}%, {attendees[5]}, {attendees[4]}\n")
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
