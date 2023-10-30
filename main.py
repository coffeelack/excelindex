import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

global indexed_folder, indexed_files, indexed_dirs, result_label, search_entry, search_inside_files

def index_folder():
    global indexed_folder, indexed_files, indexed_dirs
    folder_path = filedialog.askdirectory(title="Select Folder to Index")
    if folder_path:
        indexed_folder = folder_path
        indexed_files = 0
        indexed_dirs = 0
        index_files_in_folder(folder_path)
        result_label.config(text=f"Indexed {indexed_files} files and {indexed_dirs} directories.")

def index_files_in_folder(folder):
    global indexed_files, indexed_dirs
    for item in os.listdir(folder):
        item_path = os.path.join(folder, item)
        if os.path.isfile(item_path) and item_path.endswith(('.xls', '.xlsx')):
            indexed_files += 1
        elif os.path.isdir(item_path):
            indexed_dirs += 1
            index_files_in_folder(item_path)

def search_files(result_listbox):
    search_term = search_entry.get().lower()
    if not search_term:
        messagebox.showinfo("Search Term Missing", "Please enter a search term.")
        return

    if not indexed_folder:
        messagebox.showinfo("Folder Not Indexed", "Please index a folder first.")
        return

    result_listbox.delete(0, tk.END)
    found = False
    for item in os.listdir(indexed_folder):
        item_path = os.path.join(indexed_folder, item)
        if search_term in item.lower() and os.path.isfile(item_path) and item_path.endswith(('.xls', '.xlsx')):
            result_listbox.insert(tk.END, item)
            found = True

        if search_inside_files.get() == 1 and os.path.isfile(item_path) and item_path.endswith(('.xls', '.xlsx')):
            try:
                wb = load_workbook(item_path, read_only=True)
                for sheet in wb:
                    for row in sheet.iter_rows():
                        for cell in row:
                            if search_term in str(cell.value).lower():
                                result_listbox.insert(tk.END, item)
                                found = True
                                break
            except Exception as e:
                print(f"Error while processing {item}: {e}")

    if not found:
        messagebox.showinfo("No Matches Found", "No files matching the search term were found.")

def open_file(result_listbox=None):
    selected_file = result_listbox.get(tk.ACTIVE)
    if selected_file:
        os.startfile(os.path.join(indexed_folder, selected_file))
    else:
        messagebox.showinfo("No File Selected", "Please select a file to open.")

def main():
    global indexed_folder, indexed_files, indexed_dirs, result_label, search_entry, search_inside_files

    # Create GUI
    root = tk.Tk()
    root.minsize(240, 300)
    root.title("File Indexer")

    index_button = tk.Button(root, text="Index Folder", command=index_folder)
    index_button.pack()

    search_label = tk.Label(root, text="Search:")
    search_label.pack()

    search_entry = tk.Entry(root)
    search_entry.pack()

    search_inside_files = tk.IntVar()  # Variable to hold checkbox state
    search_inside_checkbox = tk.Checkbutton(root, text="Search Inside Files", variable=search_inside_files)
    search_inside_checkbox.pack()

    search_button = tk.Button(root, text="Search", command=lambda: search_files(result_listbox))
    search_button.pack()

    result_listbox = tk.Listbox(root)
    result_listbox.pack()

    open_button = tk.Button(root, text="Open", command=lambda: open_file(result_listbox))
    open_button.pack()

    quit_button = tk.Button(root, text="Quit", command=root.quit)
    quit_button.pack()

    result_label = tk.Label(root, text="")
    result_label.pack()

    indexed_folder = ""
    indexed_files = 0
    indexed_dirs = 0

    root.mainloop()

if __name__ == "__main__":
    main()
