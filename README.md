# Cinema Attendance Data Processor

This script processes Excel reports containing weekly cinema attendance statistics across different countries. It extracts unique movie names from all `.xlsx` files in the `Reportes` folder and saves them to a text file.

## Features
- Reads all `.xlsx` files in the `Reportes` folder.
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
1. Place all Excel files inside the `Reportes` folder.
2. Run the script:
   ```bash
   python main.py
   ```
3. The unique movie names will be saved in `unique_movies.txt`.

## Notes
- The script ignores temporary Excel files (`~$` prefix).
- Ensure all Excel files have a similar structure with movie names in column B.

## License
This project is open-source and available for personal or commercial use.
`