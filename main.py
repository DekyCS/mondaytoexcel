import tkinter as tk
from tkinter import filedialog

API_KEY = ""

def choose_folder():
    folder_path = filedialog.askdirectory()
    print(folder_path)

root = tk.Tk()

root.geometry("500x500")

root.title("Monday A Excel")

# Title Section
label = tk.Label(root, text="Monday À Excel", font=('Arial', 18))
label.pack(padx=20, pady=20)

# API Section
apiFrame = tk.Frame(root)

apiLabel = tk.Label(apiFrame, text="API Key: ", font=('Arial', 14))
apiLabel.grid(row=0, column=0, padx=10)

entry = tk.Entry(apiFrame, text="123", width=32, font=('Arial', 14))
entry.grid(row=0, column=1)

apiFrame.pack(pady=10, anchor="w")

# Add Default API Key
entry.insert(0, API_KEY)

# Excel Checkbox
excelCheckBox = tk.Checkbutton(root, text="Export Excel", font=('Arial', 14))
excelCheckBox.pack(padx=10, pady=10, anchor="w")

# Folder Chooser Excel
chooseBtn = tk.Button(root, text="Choose Folder", font=('Arial', 10), command=choose_folder)
chooseBtn.pack(padx=10, pady=10, anchor="w")

# PDF Checkbox
pdfCheckBox = tk.Checkbutton(root, text="Export PDF", font=('Arial', 14))
pdfCheckBox.pack(padx=10, pady=10, anchor="w")

# Export Button
exportBtn = tk.Button(root, text="Export", font=('Arial', 16))
exportBtn.pack(padx=20, pady=20)


root.mainloop()