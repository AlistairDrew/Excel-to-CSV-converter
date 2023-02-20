import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create a Tkinter window
root = tk.Tk()
root.withdraw()

# Ask the user to select an input file
file_path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xlsx')])

# Read the excel file
excel_file = pd.read_excel(file_path)

# Convert the data to a csv file
csv_file = excel_file.to_csv('output_file.csv', index=False)

print('Conversion completed successfully!')