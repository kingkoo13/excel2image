import pandas as pd
import matplotlib.pyplot as plt
import os
from matplotlib.offsetbox import OffsetImage, AnnotationBbox

# Read the Excel file
file_path = 'Book2.xlsx'  # Replace with your actual file path
df = pd.read_excel(file_path)

# Group the data by 'fynd_uid'
grouped = df.groupby('fynd_uid')

# Directory to save images
output_dir = 'images'

# Create the directory if it doesn't exist
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Path to the background image
background_image_path = 'bg.jpeg'  # Replace with the path to your background image

# Function to create a size chart
def create_size_chart(data, first_sku, background_image_path):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.axis('off')

    # Load and plot the background image
    img = plt.imread(background_image_path)
    ax.imshow(img, aspect='auto', extent=[0, 10, 0, 4], alpha=0.7)

    # Create table data
    columns = ['Size', 'Shoulder', 'Chest', 'Top Length', 'Waist', 'Bottom Length']
    table_data = [columns] + data[columns].values.tolist()

    # Create the table
    table = ax.table(cellText=table_data, colLabels=None, cellLoc='center', loc='center')

    # Adjust table properties
    table.auto_set_font_size(False)
    table.set_fontsize(12)
    table.scale(1.2, 1.2)

    # Make table cells background transparent
    for cell in table.get_celld().values():
        cell.set_facecolor('none')
        cell.set_edgecolor('black')  # Keep the borders for visibility

    # Set title - removing Uid from title for now, can be replaced from Product name if required
    plt.title(f'Size Chart', fontsize=15)

    # Save the figure
    output_path = os.path.join(output_dir, f'size_chart_{first_sku}.png')
    plt.savefig(output_path, bbox_inches='tight', dpi=300)
    plt.close(fig)

# Create a size chart for each unique fynd_uid
for fynd_uid, data in grouped:
    first_sku = data['sku'].iloc[0]  # Get the first SKU for the current fynd_uid
    create_size_chart(data, first_sku, background_image_path)

print("Success -> Size charts created successfully in images folder")
