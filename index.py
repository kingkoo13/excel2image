import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Create the images folder if it doesn't exist
if not os.path.exists('images'):
    os.makedirs('images')

# Open the Excel workbook
wb = openpyxl.load_workbook('your_excel_file.xlsx')
sheet = wb.active

# Assuming SKU, top_length, bottom_length, and brust_size are in columns A, B, C, and D respectively
sku_column = 'A'
top_length_column = 'B'
bottom_length_column = 'C'
brust_size_column = 'D'

# Define image dimensions and other parameters
image_width = 800
image_height = 600
background_color = (255, 255, 255)  # White
table_color = (0, 0, 0)  # Black
text_color = (0, 0, 0)  # Black
font_size = 18
font = ImageFont.truetype("arial.ttf", font_size)

# Define table parameters
num_columns = 4  
column_width = (image_width - 20) // num_columns  # Adjusted to include 20px space around the table
row_height = 50  # Adjusted row height to include space for text

# Calculate the position of the table to center it in the image with 20px space around
table_x = (image_width - num_columns * column_width) // 2
table_y = (image_height - row_height) // 2

# Create a new image for each row
for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
    # Create a new image for each row
    img = Image.new('RGB', (image_width, image_height), background_color)
    draw = ImageDraw.Draw(img)
    
    # Draw table headers
    headers = ['SKU', 'Top Length', 'Bottom Length', 'Brust Size']
    for col_idx, header in enumerate(headers):
        x0 = table_x + col_idx * column_width
        y0 = table_y
        x1 = table_x + (col_idx + 1) * column_width
        y1 = table_y + row_height
        draw.rectangle([x0, y0, x1, y1], outline=table_color, fill=table_color)
        text_width, text_height = draw.textsize(header, font=font)
        text_x = x0 + (column_width - text_width) // 2
        text_y = y0 + (row_height - text_height) // 2
        draw.text((text_x, text_y), header, fill=text_color, font=font)
    
    # Draw table data
    for col_idx, cell_value in enumerate(row):
        x0 = table_x + col_idx * column_width
        y0 = table_y + row_height
        x1 = table_x + (col_idx + 1) * column_width
        y1 = table_y + 2 * row_height
        draw.rectangle([x0, y0, x1, y1], outline=table_color, fill=None)
        text = str(cell_value)
        text_width, text_height = draw.textsize(text, font=font)
        text_x = x0 + (column_width - text_width) // 2
        text_y = y0 + (row_height - text_height) // 2
        draw.text((text_x, text_y), text, fill=text_color, font=font)
    
    # Save the image with the SKU as filename inside the "images" folder
    sku = row[0]
    image_filename = f"images/{sku}.png"
    img.save(image_filename)

# Close the workbook
wb.close()
