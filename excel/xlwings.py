import yaml
import xlwings as xw
from tabulate import tabulate

# Import data from YAML file
with open("excel/info.yaml", "r") as f:
    data = yaml.load(f, Loader=yaml.SafeLoader)

# Print data in table format
print(tabulate(data, headers="keys", tablefmt="grid"))

# Open the existing workbook
wb = xw.Book('excel/kursus.xlsx')
ws = wb.sheets['Forside']

# Define cell formatting
djof_color = 'ff00405f'  # Color in hex format

# Define formats
name_font = {'name': 'Arial', 'size': 20, 'color': (255, 255, 255)}  # RGB for white
education_font = {'name': 'Arial', 'size': 16, 'color': (255, 255, 255)}
email_font = {'name': 'Arial', 'size': 11, 'color': (255, 255, 255)}

# Define cells
cells = [
    (13, 3),  # C13
    (14, 3),  # C14
    (15, 3),  # C15
    (17, 3),  # C17
    (18, 3),  # C18
    (19, 3),  # C19
    (21, 3),  # C21
    (22, 3),  # C22
    (23, 3)   # C23
]

name_cells = cells[0::3]
education_cells = cells[1::3]
email_cells = cells[2::3]

print(name_cells)
print(education_cells)
print(email_cells)

# Apply formatting and write data to cells
for row, col in cells:
    cell = ws.range((row, col))
    cell.color = xw.utils.hex_to_rgb(djof_color)  # Convert hex color to RGB tuple
    # Note: Row height cannot be set directly with xlwings on macOS

# Update name cells
for i, (row, col) in enumerate(name_cells):
    cell = ws.range((row, col))
    cell.value = data[i]["name"]
    cell.api.Font.Name = name_font['name']
    cell.api.Font.Size = name_font['size']
    cell.api.Font.Color = xw.utils.rgb_to_int(name_font['color'])

print("All names have been updated")

# Update education cells
for i, (row, col) in enumerate(education_cells):
    cell = ws.range((row, col))
    cell.value = data[i]["education"]
    cell.api.Font.Name = education_font['name']
    cell.api.Font.Size = education_font['size']
    cell.api.Font.Color = xw.utils.rgb_to_int(education_font['color'])

print("All educations have been updated")

# Update email cells
for i, (row, col) in enumerate(email_cells):
    cell = ws.range((row, col))
    cell.value = data[i]["email"]
    cell.api.Font.Name = email_font['name']
    cell.api.Font.Size = email_font['size']
    cell.api.Font.Color = xw.utils.rgb_to_int(email_font['color'])
    cell.api.Hyperlinks.Add(cell.api, Address=f"mailto:{data[i]['email']}", TextToDisplay=data[i]["email"])

print("All emails have been updated")

# Save and close the workbook
wb.save('excel/kursus_output.xlsx')
wb.close()
