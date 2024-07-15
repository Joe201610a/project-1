import os
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    
    for border_name, border_attrs in kwargs.items():
        border = OxmlElement(f"w:{border_name}")
        for attr_name, attr_value in border_attrs.items():
            border.set(qn(f"w:{attr_name}"), str(attr_value))
        tcPr.append(border)

def set_font(cell, font_name='Calibri (Body)', font_size=14):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), font_name)

# Create a new Document
doc = Document()

# Add a title to the document
doc.add_heading('Table Example', level=1)

# Add a picture to the top right
picture_path = r'C:\Users\Joe201610a\Downloads\project-1-main\project-1-main\Untitled.png'  # Correct path

# Check if the image path is valid
if not os.path.exists(picture_path):
    print("The specified image path does not exist.")
else:
    doc.add_picture(picture_path, width=Inches(1.5))
    last_paragraph = doc.paragraphs[-1]
    last_paragraph.alignment = 2  # Right alignment

# Create a table with 1 header row and 3 data rows, each with 6 columns
table = doc.add_table(rows=4, cols=6)

# Define headers and data
headers = ["Header 1", "Header 2", "Header 3", "Header 4", "Header 5", "Header 6"]
data = [
    ["Row 1, Col 1", "Row 1, Col 2", "Row 1, Col 3", "Row 1, Col 4", "Row 1, Col 5", "Row 1, Col 6"],
    ["Row 2, Col 1", "Row 2, Col 2", "Row 2, Col 3", "Row 2, Col 4", "Row 2, Col 5", "Row 2, Col 6"],
    ["Row 3, Col 1", "Row 3, Col 2", "Row 3, Col 3", "Row 3, Col 4", "Row 3, Col 5", "Row 3, Col 6"]
]

# Populate the header row
header_row = table.rows[0]
for col_idx, header in enumerate(headers):
    cell = header_row.cells[col_idx]
    cell.text = header
    
    # Set borders for the header cell
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "000000"},
        bottom={"sz": 12, "val": "single", "color": "000000"},
        start={"sz": 12, "val": "single", "color": "000000"},
        end={"sz": 12, "val": "single", "color": "000000"}
    )

    # Set font for the header cell
    set_font(cell, font_name='Calibri (Body)', font_size=14)

# Populate the data rows
for row_idx, row_data in enumerate(data, start=1):
    row = table.rows[row_idx]
    for col_idx, cell_data in enumerate(row_data):
        cell = row.cells[col_idx]
        cell.text = cell_data
        
        # Set borders for the data cell
        set_cell_border(
            cell,
            top={"sz": 12, "val": "single", "color": "000000"},
            bottom={"sz": 12, "val": "single", "color": "000000"},
            start={"sz": 12, "val": "single", "color": "000000"},
            end={"sz": 12, "val": "single", "color": "000000"}
        )

# Save the document
doc.save('table_with_header_and_picture.docx')

print("Document created successfully!")
