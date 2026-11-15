# Excel Cleaner 🧹

A Python tool that automatically cleans Excel files using Pandas.

## Features
- Checks if the Excel file exists.
- Reads Excel files and handles errors.
- Checks for required columns (`name` and `email`).
- Removes duplicate rows and missing values.
- Cleans names and email addresses.
- Validates email formats.
- Removes rows containing invalid email addresses.
- Saves the cleaned data into a new Excel file.

## Technologies Used
- Python
- Pandas
- OpenPyXL
- OS Module

## Input File
`user.xlsx`

## Output File
`cleaning_user.xlsx`

## How to Run

1. Install the required libraries:

   ```bash
   pip install pandas openpyxl
   ```

2. Place `user.xlsx` in the same folder as `frist.py`.

3. Run the program:

   ```bash
   python frist.py
   ```

## Project Goal

This project was built to practice Python, Pandas, file handling, and data cleaning.
