import os
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

global indexed_folder, indexed_files, indexed_dirs, result_label, search_entry, search_inside_files, indexed_folder_label


def index_folder():
    global indexed_folder, indexed_files, indexed_dirs
    folder_path = filedialog.askdirectory(title="Select Folder to Index")
    if folder_path:
        indexed_folder = folder_path
        indexed_files = 0
        indexed_dirs = 0
        index_files_in_folder(folder_path)
        result_label.config(text=f"Indexed {indexed_files} files and {indexed_dirs} directories.")
        indexed_folder_label.config(text=f"Indexed Folder: {indexed_folder}")


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

    for root, dirs, files in os.walk(indexed_folder):
        for item in files:
            item_path = os.path.join(root, item)

            if os.path.isfile(item_path) and item_path.endswith(('.xls', '.xlsx')):
                try:
                    wb = load_workbook(item_path, read_only=True)
                    for sheet in wb:
                        for row in sheet.iter_rows():
                            for cell in row:
                                if search_term in str(cell.value).lower():
                                    result_listbox.insert(tk.END, f"File: {item} | Path: {item_path}")
                                    break
                except Exception as e:
                    print(f"Error while processing {item}: {e}")

    if result_listbox.size() == 0:
        messagebox.showinfo("No Matches Found", "No files matching the search term were found.")



def open_file(result_listbox=None):
    selected_item = result_listbox.get(tk.ACTIVE)
    if selected_item:
        file_path = selected_item.split("| Path: ")[-1].strip()
        os.startfile(file_path)
    else:
        messagebox.showinfo("No File Selected", "Please select a file to open.")


def open_path(result_listbox=None):
    selected_item = result_listbox.get(tk.ACTIVE)
    if selected_item:
        file_path = selected_item.split("| Path: ")[-1].strip()
        os.startfile(os.path.dirname(file_path))
    else:
        messagebox.showinfo("No File Selected", "Please select a file to open the path.")

def show_license(root):
    license_text = """
    License - CC-BY-NC-SA 4.0

    This software is licensed under the Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License.
    
    You have the following freedoms:
    
    Use the software for any lawful purpose, including commercial.
    Change and adapt the software.
    Pass on the software.
    
    This license requires you to read the following terms and conditions:
    
    Attribution: You must credit the author's name in the manner specified by him/her.
    Non-Commercial: You may not use this software for commercial purposes.
    Share Alike: If you modify, modify, or use the Software as the basis for new software, you must distribute your contributions under the same license as the original.
    Disclaimer:
    The Software is provided “as is” without warranty of any kind. The author is not liable for any damage resulting from the use of this software.
    
    A notice:
    This License applies solely to the Software and does not affect any other portion of your project not covered by this License.
    
    For a detailed version of the license terms, please visit the official page of the Creative Commons license CC-BY-NC-SA 4.0.
    © 2023 Gabriel Unsinn
    """
    license_window = tk.Toplevel(root)
    license_window.title("License")
    license_label = tk.Label(license_window, text=license_text, justify=tk.LEFT)
    license_label.pack(padx=10, pady=10)
    license_window.iconbitmap('Shell-Icon.ico')

def main():
    # Declare global variables
    global indexed_folder, indexed_files, indexed_dirs, result_label, search_entry, search_inside_files, indexed_folder_label

    # Create GUI
    root = tk.Tk()
    root.minsize(240, 300)
    root.title("File Indexer")

    # Set icon
    root.iconbitmap('Shell-Icon.ico')

    # Create widget for path of indexed folder
    indexed_folder_label = tk.Label(root, text="")
    indexed_folder_label.pack()

    # Create widget for indexing and searching
    index_button = tk.Button(root, text="Index Folder", command=index_folder)
    index_button.pack()

    # Create widget for displaying the result of the indexing
    result_label = tk.Label(root, text="")
    result_label.pack()

    # Create widget for searching the indexed folder and files with a search term
    search_label = tk.Label(root, text="Search:")
    search_label.pack()

    # Create widget for entering the search term
    search_entry = tk.Entry(root)
    search_entry.pack()

    # Create widget so that the user can choose whether to search inside the files or not
    search_inside_files = tk.IntVar()
    search_inside_checkbox = tk.Checkbutton(root, text="Search Inside Files", variable=search_inside_files)
    search_inside_checkbox.pack()

    # Create widget for triggering the search
    search_button = tk.Button(root, text="Search", command=lambda: search_files(result_listbox))
    search_button.pack()

    # Create widget for displaying the search results
    result_listbox = tk.Listbox(root)
    result_listbox.pack(expand=True, fill=tk.BOTH)

    # Create widget frame for the following two buttons
    button_frame = tk.Frame(root)
    button_frame.pack()

    # Create widget for opening the selected file's path
    open_file_button = tk.Button(button_frame, text="Open File", command=lambda: open_file(result_listbox))
    open_file_button.pack(side=tk.LEFT)

    # Create widget for opening the selected file's path
    open_path_button = tk.Button(button_frame, text="Open Path", command=lambda: open_path(result_listbox))
    open_path_button.pack(side=tk.LEFT, padx=(5, 0))

    # Create widget for quitting the program
    quit_button = tk.Button(root, text="Quit", command=root.quit)
    quit_button.pack()

    # Create widget frame for the following two widgets
    frame = tk.Frame(root)
    frame.pack(pady=10)

    # Create widget for displaying the name of the author
    name_label = tk.Label(frame, text="© 2023 Gabriel Unsinn")
    name_label.pack(side=tk.LEFT)

    # Create widget for displaying the license
    license_button = tk.Button(frame, text="View License", command=lambda: show_license(root))
    license_button.pack(side=tk.LEFT, padx=(10, 0))

    # Initialize variables
    indexed_folder = ""
    indexed_files = 0
    indexed_dirs = 0

    # Start GUI
    root.mainloop()


if __name__ == "__main__":
    main()
