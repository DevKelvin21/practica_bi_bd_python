# Cinema Attendance Data Processor

This script processes Excel reports containing weekly cinema attendance statistics across different countries. It extracts unique movie names and aggregates the attendance data for each movie from all `.xlsx` files inside `Reportes.zip`. The results are then saved to a text file with detailed statistics.

## Features
- Reads all `.xlsx` files inside `Reportes.zip` without extracting them to disk.
- Extracts unique movie names from column B, starting from row 4.
- Aggregates attendance data for the release day, 1st weekend, Week 1, and Week 2.
- Calculates percentages for the 1st and 2nd week relative to total attendance.
- Computes the total number of attendees for each movie and the remaining attendees (total minus Week 1 and Week 2 attendees).
- Saves the detailed movie attendance data to `attendance_summary.txt`.

## Requirements
- Python 3.x
- Required libraries: `pandas`, `openpyxl`

Install dependencies using:
```bash
pip install pandas openpyxl
```

## Usage
1. Ensure that `Reportes.zip` is in the same directory as `main.py`.
2. Run the script:
   ```bash
   python main.py
   ```
3. The aggregated movie attendance data will be saved in `attendance_summary.txt`.

## Notes
- The script processes files directly from memory, avoiding unnecessary disk usage.
- Ensure all Excel files inside `Reportes.zip` have a similar structure with movie names in column B and attendance data in the relevant columns (Release Day in column R, Weekend in column V, Weekly in column Z, and Week Number in column D).
- Temporary Excel files (`~$` prefix) are ignored automatically.

## Output File Format
The resulting output file, `attendance_summary.txt`, contains the following columns:

1. **Película** - Movie name
2. **Asistentes Día de Estreno** - Release day attendees
3. **Asistentes 1er Fin de Semana** - 1st weekend attendees
4. **1er Semana** - Week 1 attendees
5. **Porcentaje Semana 1** - Percentage of total for Week 1
6. **Asistentes Semanales (Semana 2)** - Weekly attendees for Week 2
7. **Porcentaje Semana 2** - Percentage of total for Week 2
8. **Asistentes Restantes** - Remaining attendees (total minus Week 1 and Week 2 attendees)
9. **Asistentes Totales** - Total attendees (sum of all valid attendance data)

## License
This project is open-source and available for personal or commercial use.