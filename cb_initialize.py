from docx import Document
import random

# Function to generate random code numbers
def generate_code_numbers():
    return [random.randint(1000000, 9999999) for _ in range(120)]

# Function to create a Word document with 4 columns
def create_word_document(file_path):
    # Generate code numbers
    codes = generate_code_numbers()

    # Create a new Document
    doc = Document()

    # Add a table with 4 columns
    table = doc.add_table(rows=0, cols=4)
    table.style = 'Table Grid'

    # Add code numbers to the table
    for i in range(0, len(codes), 4):
        row_cells = table.add_row().cells
        for j in range(4):
            row_cells[j].text = f"{i + j + 1}.\n{codes[i + j]}"

    # Save the document to the specified file path
    doc.save(file_path)

# Specify the file path where you want to save the document
file_path = "C:/Users/E457875/Downloads/code_numbers.docx"

# Create the Word document at the specified file path
create_word_document(file_path)
