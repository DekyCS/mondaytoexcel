import tkinter as tk
from tkinter import filedialog
from dotenv import load_dotenv
import os
import requests
import json

load_dotenv()

API_KEY = os.getenv("API_KEY")
BOARD_ID = os.getenv("BOARD_ID")
API_URL = os.getenv("API_URL")

# GraphQL query with corrected types
query = """
query($boardId: [ID!], $cursor: String) {
  boards(ids: $boardId) {
    items_page(limit: 100, cursor: $cursor) {
      cursor
      items {
        name
        column_values {
          column {
            title
          }
          text
        }
      }
    }
  }
}
"""

# Update saved API key
def update_api_key():
    global API_KEY
    file_path = ".env"
    new_key = entry.get()
    if new_key:  
        with open(file_path, 'r') as file:
            lines = file.readlines()
        with open(file_path, 'w') as file:
            for line in lines:
                if line.startswith("API_KEY"):
                    file.write(f'{"API_KEY"}={new_key}\n')
                else:
                    file.write(line)
        API_KEY = new_key
        tk.messagebox.showinfo("Success", "API Key updated successfully!")
    else:
        tk.messagebox.showerror("Error", "API Key cannot be empty.")

# Update saved Board ID
def update_board_id():
    global BOARD_ID
    file_path = ".env"
    new_key = boardEntry.get()
    if new_key:  
        with open(file_path, 'r') as file:
            lines = file.readlines()
        with open(file_path, 'w') as file:
            for line in lines:
                if line.startswith("BOARD_ID"):
                    file.write(f'{"BOARD_ID"}={new_key}\n')
                else:
                    file.write(line)
        BOARD_ID = new_key
        tk.messagebox.showinfo("Success", "Board ID updated successfully!")
    else:
        tk.messagebox.showerror("Error", "Board ID cannot be empty.")

# Check Checkbox
def on_check():
    if excelCheckState.get() == 1:
        chooseBtn.grid(padx=10, column=1, row=0)
    else:
        chooseBtn.grid_forget()

# Open File Explorer to choose path
def choose_folder():
    folder_path = filedialog.askdirectory()
    excelPath.delete(0, tk.END)
    excelPath.insert(0, folder_path)

# Export Button
def export_info():
    if excelCheckState.get() == 0 and pdfCheckState.get() == 0:
        tk.messagebox.showerror("Error", "You must select a method to export.")
    else:
        if excelCheckState.get() == 1:
            excelWorks = export_excel()
        if pdfCheckState.get() == 1:
            pass
        if excelWorks:
            tk.messagebox.showinfo("Success", "Successfully exported!")

# Export Excel
def export_excel():
    path = excelPath.get()
    if path:
        all_items = call_api()
        reformatted_items = reformat_items(all_items)
        print(json.dumps(reformatted_items, indent=2))
        return True
    else:
        tk.messagebox.showerror("Error", "You must select a folder to save")
    return False

# API Request
def call_api():
    # Headers for authentication
    headers = {
        'Authorization': API_KEY,
        'Content-Type': 'application/json'
    }

    items = []
    cursor = None  # Start with no cursor

    while True:
        # Dynamic variables for GraphQL query
        variables = {
            "boardId": [BOARD_ID],
            "cursor": cursor
        }

        # Make the API call
        response = requests.post(API_URL, headers=headers, json={'query': query, 'variables': variables})
        response_data = response.json()
        
        if response.status_code != 200 or 'errors' in response_data:
            print(f"Error fetching data: {response_data.get('errors')}")
            break

        # Extract items and cursor from the response
        board_data = response_data['data']['boards'][0]['items_page']
        items.extend(board_data['items'])
        cursor = board_data['cursor']
        
        # If cursor is None, it means there are no more pages to fetch
        if not cursor:
            break

    return items

# Function to reformat items into the desired structure
def reformat_items(items):
    formatted_items = []
    for item in items:
        formatted_item = {
            "name": item['name'],
            "column_values": {cv['column']['title']: cv['text'] for cv in item['column_values']}
        }
        formatted_items.append(formatted_item)
    return formatted_items


# Tkinter root setup

root = tk.Tk()

root.geometry("500x500")

root.title("Monday A Excel")
if os.path.exists("icon.ico"):
    root.iconbitmap("icon.ico")

# Title Section
label = tk.Label(root, text="Monday À Excel", font=('Arial', 18))
label.pack(padx=20, pady=20)

# API Section
apiFrame = tk.Frame(root)

apiLabel = tk.Label(apiFrame, text="API Key: ", font=('Arial', 14))
apiLabel.grid(row=0, column=0, padx=5)

entry = tk.Entry(apiFrame, width=28, font=('Arial', 14))
entry.grid(row=0, column=1, padx=5)

saveBtn = tk.Button(apiFrame, text="Save", font=('Arial', 10), command=update_api_key)
saveBtn.grid(row=0, column=2, padx=5)

apiFrame.pack(pady=10)

# Add Default API Key
entry.insert(0, API_KEY)

# Board ID Section
boardFrame = tk.Frame(root)

boardLabel = tk.Label(boardFrame, text="Board ID: ", font=('Arial', 14))
boardLabel.grid(row=0,column=0, padx=5)

boardEntry = tk.Entry(boardFrame, width=28, font=('Arial', 14))
boardEntry.grid(row=0, column=1, padx=5)

boardSaveBtn = tk.Button(boardFrame, text="Save", font=('Arial', 10), command=update_board_id)
boardSaveBtn.grid(row=0, column=2, padx=5)

boardFrame.pack(pady=10)

# Add Default Board ID
boardEntry.insert(0, BOARD_ID)

# Excel Checkbox
excelFrame = tk.Frame(root)

excelCheckState = tk.IntVar()
excelCheckBox = tk.Checkbutton(excelFrame, text="Export Excel", font=('Arial', 14), variable=excelCheckState, command=on_check)
excelCheckBox.grid(padx=10, column=0, row=0)

# Folder Chooser Excel
chooseBtn = tk.Button(excelFrame, text="Choose Folder", font=('Arial', 10), command=choose_folder)
excelPath = tk.Entry(root)

excelFrame.pack(pady=10, anchor="w")

# PDF Checkbox
pdfCheckState = tk.IntVar()
pdfCheckBox = tk.Checkbutton(root, text="Export PDF", font=('Arial', 14), variable=pdfCheckState)
pdfCheckBox.pack(padx=10, pady=10, anchor="w")

# Export Button
exportBtn = tk.Button(root, text="Export", font=('Arial', 16), command=export_info)
exportBtn.pack(padx=20, pady=20)


root.mainloop()