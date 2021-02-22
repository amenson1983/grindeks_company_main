import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Hello')

# Text with formatting.
worksheet.write('B1', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(1, 0, 123)
worksheet.write(1, 1, 123.456)

# Insert an image.
#worksheet.insert_image('B5', 'logo.png')

workbook.close()