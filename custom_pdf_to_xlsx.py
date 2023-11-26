import pdfplumber
import openpyxl
import re
import os
import sys


MAX_NUM_ROWS_PER_EXCEL = 20
COLUMN_NAMES = ['Описание', 'ДДС', 'Брой', 'Ед. цена']

regex_pattern = re.compile(r'''
    \d{5,}\s+  # UNUSED serial numer 
    (.{1,25})\s+  # description
    [А-Я]{2}\s+  # UNUSED type of package
    ([\d,]+)\s+  # single item price
    ([\d,]+)\s+  # quantity 1
    [\d,]+\s+  # UNUSED price
    (\d+)  # quantity 2
    .*
    ([А-Я])  # VAT
''', re.VERBOSE)


curr_dir = os.getcwd()  # Set to current directory

# Construct the path to the 'pdfs' directory inside 'curr_dir'
pdfs_directory = os.path.join(curr_dir, 'pdfs')

# Change the current working directory to 'pdfs'
try:
    os.chdir(pdfs_directory)
except FileNotFoundError:
    print(f"Error: The specified 'pdfs' directory '{pdfs_directory}' does not exist.")
    sys.exit(1)

# Create the 'xlsx' directory if it doesn't exist
xlsx_directory = os.path.join(curr_dir, 'xlsx')
os.makedirs(xlsx_directory, exist_ok=True)

# Get a list of PDF files in the current directory
pdf_files = [file for file in os.listdir() if file.endswith('.pdf')]

# Iterate through each PDF file
for pdf_file_name in pdf_files:

    # Open the PDF file using pdfplumber
    with pdfplumber.open(pdf_file_name) as pdf:

        excel_rows = []  # list of all the matching rows that will be added to excel file(s)

        # Iterate through each page in the PDF
        for page in pdf.pages:
            text = page.extract_text()

            # Iterate through each line on that page and match with the regex pattern
            for line in text.split('\n'):
                if match := regex_pattern.search(line):
                    description, single_item_price, quantity1, quantity2, VAT = match.groups()

                    quantity1 = float(quantity1.replace(',', '.'))
                    quantity2 = int(quantity2)
                    quantity = quantity1 * quantity2

                    # if the quantity is in the description (for example "180БР. ЯЙЦА L МCH", fix quantity and single price)
                    if match_description := re.search(r'(\d+)БР\.', description):
                        quantity_per_box = int(match_description.group(1))
                        quantity *= quantity_per_box
                        single_item_price = round(float(single_item_price.replace(',', '.')) / quantity_per_box, 9)

                    # save the row in the list
                    excel_rows.append([description, VAT, quantity, single_item_price])

        # Extract the file name without the '.pdf' extension
        file_name_without_extension = os.path.splitext(pdf_file_name)[0]

        # Iterate through the matching lines and save them in excel files
        for i in range(0, len(excel_rows), MAX_NUM_ROWS_PER_EXCEL):
            # Create a new Excel workbook and select the active sheet
            excel_workbook = openpyxl.Workbook()
            excel_sheet = excel_workbook.active

            # Set the first row as column names.
            excel_sheet.append(COLUMN_NAMES)

            # add up to MAX_NUM_ROWS_PER_EXCEL lines in a single xlsx file
            for row in excel_rows[i:i + MAX_NUM_ROWS_PER_EXCEL]:
                excel_sheet.append(row)

            # Save the Excel file with the corresponding PDF file name
            excel_workbook.save(os.path.join(xlsx_directory, f'{file_name_without_extension}_{i // MAX_NUM_ROWS_PER_EXCEL + 1}.xlsx'))