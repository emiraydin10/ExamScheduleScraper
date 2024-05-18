# Exam Schedule Scraper

This Python script scrapes exam schedule information from a PDF file and exports it to an Excel file. It searches for specified courses and retrieves details such as date, day, time range, and classroom.

## Output

The output Excel file contains the following columns:

Course
Date
Day
Time Range
Classroom
Dates are formatted as DD/MM/YYYY, and the file is sorted by date.

## Features

- Scrapes text from a PDF containing exam schedules.
- Searches for specific courses and extracts relevant details.
- Converts extracted data into a structured Excel file.
- Formats the Excel file for better readability.

## Requirements

- Python 3.x
- PyPDF2
- pandas
- openpyxl


## License

This project is licensed under the MIT License.

## Contact

For any questions or issues, please open an issue in the repository or contact the maintainer at:
E-mail: emiraydn2001@gmail.com

## Installation

To install the required packages, run the following command:
```bash
pip install PyPDF2 pandas openpyxl

