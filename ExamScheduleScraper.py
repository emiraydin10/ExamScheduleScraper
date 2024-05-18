import PyPDF2
import pandas as pd
from openpyxl import load_workbook

# Path to the PDF file
pdf_path = '' #Specify pdf path
reader = PyPDF2.PdfReader(pdf_path)

# List of courses to search for
myCourses = [''] #type the course codes you want to scrape like 'ABC 101'

# List to store extracted data
data = []

# Loop through each page of the PDF
for page in reader.pages:
    text = page.extract_text()
    for line in text.split('\n'):
        for course in myCourses:
            if course in line:
                # Split the line into parts
                parts = line.split()  # Split by whitespace
                if len(parts) >= 6:  # Ensure there are enough parts
                    # Use '-' character to split time range and class information
                    time_range = f"{parts[4]}-{parts[6]}"
                    classroom = ' '.join(parts[7:])

                    # Create a dictionary for the course information
                    course_info = {
                        "Course": f"{parts[0]} {parts[1]}",  # Course code and number
                        "Date": parts[2],  # Date
                        "Day": parts[3],  # Day
                        "Time Range": time_range,  # Time range
                        "Classroom": classroom  # Classroom (could be multiple)
                    }
                    data.append(course_info)
                break

# Convert the data into a Pandas DataFrame
df = pd.DataFrame(data)

# Convert the 'Date' column to datetime format
try:
    df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce')
except Exception as e:
    print(f"Date conversion error: {e}")

# Check for invalid dates and print a warning if found
invalid_dates = df[df['Date'].isna()]
if not invalid_dates.empty:
    print("The following dates could not be parsed and are set to NaT:")
    print(invalid_dates)

# Convert the 'Date' column back to string format
df['Date'] = df['Date'].dt.strftime('%d/%m/%Y')

# Sort the DataFrame by date
df = df.sort_values(by='Date')

# Write the DataFrame to an Excel file
excel_path = 'CourseExamDates.xlsx'
df.to_excel(excel_path, index=False)

# Open the Excel file using openpyxl and adjust column widths
workbook = load_workbook(excel_path)
worksheet = workbook.active

# Adjust column widths based on the maximum length of cell contents
for column in worksheet.columns:
    max_length = 0
    column = [cell for cell in column]
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = max_length + 2
    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

# Save the changes to the Excel file
workbook.save(excel_path)

print(f"Excel file created at {excel_path}")
