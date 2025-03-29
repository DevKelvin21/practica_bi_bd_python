# Cinema Attendance Data Processor

This script processes Excel reports containing weekly cinema attendance statistics across different countries. It extracts unique movie names from all `.xlsx` files inside `Reportes.zip` and saves them to a text file.

## Features
- Reads all `.xlsx` files inside `Reportes.zip` without extracting them to disk.
- Extracts unique movie names from column B, starting from row 4.
- Saves the list of unique movies to `unique_movies.txt`.

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
3. The unique movie names will be saved in `unique_movies.txt`.

## Notes
- The script processes files directly from memory, avoiding unnecessary disk usage.
- Ensure all Excel files inside `Reportes.zip` have a similar structure with movie names in column B.
- Temporary Excel files (`~$` prefix) are ignored automatically.

## License
This project is open-source and available for personal or commercial use.