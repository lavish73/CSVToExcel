import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to convert CSV to Excel
def csv_to_excel(csv_file, excel_file):
    try:
        # Read the CSV file using pandas
        data = pd.read_csv(csv_file)
        
        # Write the data to an Excel file
        data.to_excel(excel_file, index=False)
        messagebox.showinfo("Success", f"CSV file has been successfully converted to {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to open the file dialog and select a CSV file
def browse_csv():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        entry_csv.delete(0, tk.END)
        entry_csv.insert(0, file_path)

# Function to open the file dialog and select a location to save the Excel file
def browse_excel():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, file_path)

# Function to trigger the conversion process
def convert():
    csv_file = entry_csv.get()
    excel_file = entry_excel.get()
    
    if not csv_file or not excel_file:
        messagebox.showwarning("Input Error", "Please select both the CSV file and the output location.")
    else:
        csv_to_excel(csv_file, excel_file)

# Create the main window
root = tk.Tk()
root.title("CSV to Excel Converter")
root.geometry("600x300")
root.configure(bg="#f0f0f0")  # Set background color

# Set a custom font
font_label = ("Helvetica", 12)
font_button = ("Helvetica", 10, "bold")

# Header label
header_label = tk.Label(root, text="CSV to Excel Converter", font=("Helvetica", 16, "bold"), fg="#4CAF50", bg="#f0f0f0")
header_label.place(x=160, y=20)

# CSV file selection label and button
label_csv = tk.Label(root, text="1. Select the CSV file to convert:", font=font_label, bg="#f0f0f0", anchor="w")
label_csv.place(x=30, y=70)
entry_csv = tk.Entry(root, width=30, font=font_label)
entry_csv.place(x=230, y=70)
button_browse_csv = tk.Button(root, text="Browse", command=browse_csv, font=font_button, bg="#4CAF50", fg="white", relief="raised", padx=10)
button_browse_csv.place(x=480, y=70)

# Excel file selection label and button
label_excel = tk.Label(root, text="2. Choose where to save the Excel file:", font=font_label, bg="#f0f0f0", anchor="w")
label_excel.place(x=30, y=120)
entry_excel = tk.Entry(root, width=30, font=font_label)
entry_excel.place(x=230, y=120)
button_browse_excel = tk.Button(root, text="Browse", command=browse_excel, font=font_button, bg="#4CAF50", fg="white", relief="raised", padx=10)
button_browse_excel.place(x=480, y=120)

# Convert button
button_convert = tk.Button(root, text="Convert", command=convert, font=font_button, bg="#2196F3", fg="white", relief="raised", padx=15, pady=5)
button_convert.place(x=230, y=180)

# Start the main event loop
root.mainloop()
