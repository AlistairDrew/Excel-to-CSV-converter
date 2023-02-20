import pandas as pd
import tkinter as tk
from tkinter import filedialog

class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel to CSV Converter')
        
        # Create input file selection button
        self.input_file_path = tk.StringVar()
        input_file_button = tk.Button(self.root, text='Select Input File', command=self.select_input_file)
        input_file_button.pack()
        
        # Create output file selection button
        self.output_file_path = tk.StringVar()
        output_file_button = tk.Button(self.root, text='Select Output File', command=self.select_output_file)
        output_file_button.pack()
        
        # Create headers input field
        self.headers = tk.StringVar()
        headers_label = tk.Label(self.root, text='Headers (comma-separated):')
        headers_label.pack()
        headers_entry = tk.Entry(self.root, textvariable=self.headers)
        headers_entry.pack()
        
        # Create data input field
        self.data = tk.StringVar()
        data_label = tk.Label(self.root, text='Data (comma-separated rows):')
        data_label.pack()
        data_entry = tk.Entry(self.root, textvariable=self.data)
        data_entry.pack()
        
        # Create conversion button
        convert_button = tk.Button(self.root, text='Convert', command=self.convert)
        convert_button.pack()
    
    def select_input_file(self):
        input_file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if input_file_path:
            self.input_file_path.set(input_file_path)
    
    def select_output_file(self):
        output_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
        if output_file_path:
            self.output_file_path.set(output_file_path)
    
    def convert(self):
        input_file_path = self.input_file_path.get()
        output_file_path = self.output_file_path.get()
        headers = self.headers.get().split(',')
        data = [row.split(',') for row in self.data.get().split('\n')]
        
        if input_file_path and output_file_path and headers and data:
            try:
                # Read the Excel file into a pandas DataFrame
                df = pd.read_excel(input_file_path)
                
                # Create a new DataFrame with the specified headers and data
                new_df = pd.DataFrame(data, columns=headers)
                
                # Append the new DataFrame to the original DataFrame
                df = pd.concat([df, new_df], sort=False)
                
                # Write the combined DataFrame to a CSV file
                df.to_csv(output_file_path, index=False)
                
                tk.messagebox.showinfo('Conversion Successful', 'Excel file converted to CSV format.')
            except Exception as e:
                tk.messagebox.showerror('Error', str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
    root.mainloop()