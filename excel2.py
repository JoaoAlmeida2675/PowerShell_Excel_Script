import pandas as pd
import tkinter as tk
from tkinter import filedialog

def letter_to_number(letter):
    # Convert a letter to its corresponding numeric value (A=1, B=2, ..., Z=26)
    return ord(letter.upper()) - ord('A') + 1

def run_script(lines_var, columns_var):
    try:
        lines = list(map(int, lines_var.get().split(',')))
        specified_columns = list(map(str.strip, columns_var.get().split(',')))

        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

        if file_path:
            df = pd.read_excel(file_path, sheet_name=0)

            # Check if specified columns exist in the DataFrame based on column positions
            num_columns = len(df.columns)
            valid_columns = [col for col in specified_columns if 1 <= letter_to_number(col) <= num_columns]

            if not valid_columns:
                raise ValueError("Invalid column positions")

            # Filter the DataFrame based on lines and specified columns
            df_filtered = df.iloc[lines, [letter_to_number(col) - 1 for col in valid_columns]]

            # Identify empty cells in the DataFrame
            empty_cells = df_filtered[df_filtered.isna().any(axis=1)]

            # Print a detailed report
            print("Script executed successfully!")
            print("Empty Cells:")
            for index, row in empty_cells.iterrows():
                line_number = lines[index]
                for col in valid_columns:
                    if pd.isna(row[col]):
                        print(f"Line {line_number} - Missing Data In Cell {col}")

    except Exception as e:
        print(f"Error running script: {e}")

# Create the main window
root = tk.Tk()
root.title("Script Runner")
root.geometry("400x300")

# Create entry boxes for lines and columns
lines_label = tk.Label(root, text="Lines to Search (comma-separated):")
lines_entry = tk.Entry(root)
lines_label.pack()
lines_entry.pack()

columns_label = tk.Label(root, text="Column Positions to Search (comma-separated):")
columns_entry = tk.Entry(root)
columns_label.pack()
columns_entry.pack()

# Create a button to run the script
run_button = tk.Button(root, text="Run", command=lambda: run_script(lines_entry, columns_entry))
run_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
