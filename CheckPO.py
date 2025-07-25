import tkinter as tk
from tkinter import filedialog
import os

def check_existence():
    folder_path =  r"C:\Feasibility\WhatsApp Order\Output WS"
    """Checks if a list of items or a single item exists in text files within a folder."""

    '''folder_path = filedialog.askdirectory(title="Select Folder Containing Text Files")
    if not folder_path:
        return  # User cancelled folder selection'''

    items_to_check = text_input.get("1.0", tk.END).strip().splitlines()
    if not items_to_check:
        result_label.config(text="Please enter items to check.")
        return

    results = {}
    for item in items_to_check:
        results[item] = []

    for filename in os.listdir(folder_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(folder_path, filename)
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    file_content = f.read()
                    for item in items_to_check:
                        if item.strip() and item.strip() in file_content: #check for empty item and if item is in file.
                            results[item].append(filename)
            except Exception as e:
                result_label.config(text=f"Error reading {filename}: {e}")
                return

    result_text = ""
    for item, found_in_files in results.items():
        result_text += f"'{item}': "
        if found_in_files:
            result_text += ", ".join(found_in_files) + "\n"
        else:
            result_text += "Not found.\n"

    result_label.config(text=result_text)

# UI Setup
root = tk.Tk()
root.title("WhatsApp Po-Order Checker")

text_input_label = tk.Label(root, text="Enter items to check (one per line):")
text_input_label.pack(pady=5)

text_input = tk.Text(root, height=10, width=50)
text_input.pack(pady=5)

check_button = tk.Button(root, text="Check Existence", command=check_existence)
check_button.pack(pady=10)

result_label = tk.Label(root, text="")
result_label.pack(pady=10)

root.mainloop()