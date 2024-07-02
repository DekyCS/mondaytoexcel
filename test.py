from openpyxl import load_workbook

# Load the existing workbook
workbook_path = './projet.xlsx'
workbook = load_workbook(workbook_path)

sheet = workbook['ExperiSens']  # Replace 'Sheet1' with the name of your sheet

# Add data to specific cells
sheet['C4'] = 'New Data'

# Save the workbook
workbook.save(workbook_path)