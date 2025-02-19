import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox

def select_from_list(title, options):
    """Creates a pop-up window with a dropdown selection and returns the chosen option."""
    selected_value = None  # Local variable to store selection

    def set_choice():
        nonlocal selected_value  # Modify the local variable
        selected_value = var.get()
        popup.destroy()  # Close pop-up window

    popup = tk.Toplevel()
    popup.title(title)
    popup.geometry("300x150")
    
    tk.Label(popup, text=title).pack(pady=10)
    
    var = tk.StringVar(popup)
    var.set(options[0])  # Default to first option
    
    dropdown = tk.OptionMenu(popup, var, *options)
    dropdown.pack(pady=5)
    
    tk.Button(popup, text="Select", command=set_choice).pack(pady=10)
    
    popup.grab_set()  # Make window modal
    popup.wait_window()  # Wait for window to close before proceeding
    
    return selected_value  # Return user selection

# Create a hidden root window
root = tk.Tk()
root.withdraw()

# Select Old Excel file
file1 = filedialog.askopenfilename(title="Select the Old Excel file", filetypes=[("Excel Files", "*.xlsx")])

# Select New Excel file
file2 = filedialog.askopenfilename(title="Select the New Excel file", filetypes=[("Excel Files", "*.xlsx")])

# Ensure both files are selected
if not file1 or not file2:
    messagebox.showwarning("Warning", "You must select both an Old and New Excel file!")
    exit()

# Load sheet names from both files
try:
    xls1 = pd.ExcelFile(file1)
    xls2 = pd.ExcelFile(file2)
    sheets1 = xls1.sheet_names
    sheets2 = xls2.sheet_names
except Exception as e:
    messagebox.showerror("Error", f"Could not read Excel files:\n{e}")
    exit()

# Let user select a sheet from each file
sheet1 = select_from_list("Select a sheet from the Old file", sheets1)
sheet2 = select_from_list("Select a sheet from the New file", sheets2)

# Load columns from the selected sheet in the New file
try:
    df = pd.read_excel(xls2, sheet_name=sheet2)
    columns = df.columns.tolist()
except Exception as e:
    messagebox.showerror("Error", f"Could not read the selected sheet in the New Excel file:\n{e}")
    exit()

# Let user select a key column from the chosen sheet
key_column = select_from_list(f"Select a key column from {sheet2}", columns)

# Ask for output file save location
output_file = filedialog.asksaveasfilename(title="Save Output File As", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])

# Ensure an output file was selected
if not output_file:
    messagebox.showwarning("Warning", "You must select an output file!")
    exit()

# Show confirmation message
messagebox.showinfo(
    "Selections Confirmed",
    f"Old File: {file1} (Sheet: {sheet1})\nNew File: {file2} (Sheet: {sheet2})\nKey Column: {key_column}\nOutput File: {output_file}"
)






import pandas as pd

def compare_excel_files(file1, sheet1, file2, sheet2, key_column, output_file):
    # Load both Excel files
    df1 = pd.read_excel(file1, sheet_name=sheet1)
    df2 = pd.read_excel(file2, sheet_name=sheet2)
    
    # Ensure the key column exists in both files
    if key_column not in df1.columns or key_column not in df2.columns:
        raise ValueError(f"Key column '{key_column}' must exist in both files.")
    
    # Convert key column to a set for quick lookup
    existing_keys = set(df1[key_column].dropna())
    
    # Check for new values in df2 that are not in df1
    def check_new_value(x):
        if x not in existing_keys:
            return 'YES'
        else:
            return 'NO'
    
    df2['NEW'] = df2[key_column].apply(check_new_value)
    
    # Save the result to a new Excel file
    df2.to_excel(output_file, index=False)
    print(f"Comparison complete. Results saved to {output_file}")

# Example usage
#file2 = "C:\\Users\\apagano\\OneDrive - JBS International\\Documents\\testfile109.xlsx"
#file1 = "C:\\Users\\apagano\\OneDrive - JBS International\\Documents\\testfile106.xlsx"
#output_file = "C:\\Users\\apagano\\Downloads\\NAWS109.xlsx"
#key_column = "Variable Name"  # Adjust this to the relevant column

compare_excel_files(file1, sheet1, file2, sheet2, key_column, output_file)
