from docx import Document

# Create a new Document
doc = Document()

# Add a title to the document
doc.add_heading('Table Example', level=1)

# Create a table with 3 rows and 3 columns
table = doc.add_table(rows=3, cols=3)

# Populate the table with data
data = [
    ["Header 1", "Header 2", "Header 3"],
    ["Row 1, Col 1", "Row 1, Col 2", "Row 1, Col 3"],
    ["Row 2, Col 1", "Row 2, Col 2", "Row 2, Col 3"]
]

# Fill the table with data
for row_idx, row_data in enumerate(data):
    row = table.rows[row_idx]
    for col_idx, cell_data in enumerate(row_data):
        cell = row.cells[col_idx]
        cell.text = cell_data

# Save the document
doc.save('table_example.docx')

print("Document created successfully!")
