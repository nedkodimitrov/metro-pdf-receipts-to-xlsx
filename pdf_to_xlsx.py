import pdfplumber
import openpyxl
import re
import os
import sys


# column name, regex pattern, print in excel
columns = [
    ('Артикул номер',   r'\d+',         False),  # whole number
    ('Описание',        r'.{,25}',      True),  # up to 25 characters
    ('ПТ',              r'[А-Я]{2}',    False),  # 2 bulgarian letters
    ('Ед. цена',        r'[\d,]+',      True),  # decimal number
    ('Съд/бр.',         r'[\d,]+',      False),  # decimal number
    ('Цена',            r'[\d,]+',      False),  # decimal number
    ('МЕК-во',          r'\d+',         False),  # whole number
    ('Сума нето',       r'[\d,]+',      False),  # decimal number
    ('Сума ДДС',        r'[\d,]+',      False),  # decimal number
    ('ОТ',              r'\d*?',        False),  # whole number or empty   ?
    ('АКД ВКВ',         r'[A-Z]?',      False),  # 'P' or empty            ?
    ('ОБЩО BGN',        r'[\d,]+',      False),  # decimal number
    ('ДДС',             r'[А-Я]',       True)   # 1 bulgarian letter
]

# join all regex patterns in a single one
regex_pattern = r'\s*'.join(f'({pattern})' for _, pattern, _ in columns)

# get the indexes of columns whose 'use' (column[][2]) is True
used_columns_indexes = [i for i, (_, _, use) in enumerate(columns) if use]

# Check if a command-line argument for working directory is provided
if len(sys.argv) > 1:
    new_directory = sys.argv[1]

    # Change the current working directory
    try:
        os.chdir(new_directory)
        print(f"Changed current working directory to: {os.getcwd()}")
    except FileNotFoundError:
        print(f"Error: The specified directory '{new_directory}' does not exist.")

# Get a list of PDF files in the current directory
pdf_files = [file for file in os.listdir() if file.endswith('.pdf')]

# Iterate through each PDF file
for pdf_file_name in pdf_files:

    # Open the PDF file using pdfplumber
    with pdfplumber.open(pdf_file_name) as pdf:
        # Create a new Excel workbook and select the active sheet
        excel_workbook = openpyxl.Workbook()
        excel_sheet = excel_workbook.active

        # Set the first row as column names. Add those whose use is True.
        excel_sheet.append([columns[i][0] for i in used_columns_indexes] + ['Брой'])

        # Iterate through each page in the PDF
        for page in pdf.pages:
            text = page.extract_text()

            # Iterate through each line and match with the regex pattern
            for line in text.split('\n'):
                if match := re.search(regex_pattern, line):
                    # Add matched data to the Excel sheet. Add those whose use is True
                    quantity_per_package = float(match.group(5).replace(',', '.'))
                    num_packages = int(match.group(7))
                    excel_sheet.append([match.group(i+1) for i in used_columns_indexes] + [quantity_per_package * num_packages])

        # Extract the file name without the '.pdf' extension
        file_name_without_extension = os.path.splitext(pdf_file_name)[0]

        # Save the Excel file with the corresponding PDF file name
        excel_workbook.save(f'{file_name_without_extension}.xlsx')