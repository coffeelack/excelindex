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
    Lizenztext - CC-BY-NC-SA 4.0

    Diese Software ist lizenziert unter der Creative Commons Namensnennung - Nicht-kommerziell - Weitergabe unter gleichen Bedingungen 4.0 International Lizenz.
    
    Sie haben folgende Freiheiten:
    
    Die Software für jeden legalen Zweck nutzen, auch kommerziell.
    Die Software verändern und anpassen.
    Die Software weitergeben.
    
    Dieser Lizenz müssen Sie folgende Bedingungen entnehmen:
    
    Namensnennung: Sie müssen den Namen des Autors/der Autorin in der von ihm/ihr festgelegten Weise nennen.
    Nicht-kommerziell: Sie dürfen diese Software nicht für kommerzielle Zwecke nutzen.
    Weitergabe unter gleichen Bedingungen: Wenn Sie die Software verändern, abwandeln oder als Grundlage für eine neue Software verwenden, müssen Sie Ihre Beiträge unter derselben Lizenz wie das Original verbreiten.
    Haftungsausschluss:
    Die Software wird "wie sie ist" bereitgestellt, ohne jegliche Gewährleistung. Der Autor/die Autorin haftet nicht für eventuelle Schäden, die aus der Nutzung dieser Software entstehen.
    
    Hinweis:
    Diese Lizenz gilt ausschließlich für die Software und hat keine Auswirkungen auf andere Teile Ihres Projekts, die nicht unter dieser Lizenz stehen.
    
    Für eine detaillierte Version der Lizenzbedingungen besuchen Sie bitte die offizielle Seite der Creative Commons-Lizenz CC-BY-NC-SA 4.0.
    
    © 2023 Gabriel Unsinn
    """
    license_window = tk.Toplevel(root)
    license_window.title("License")
    license_label = tk.Label(license_window, text=license_text, justify=tk.LEFT)
    license_label.pack(padx=10, pady=10)
    license_window.iconbitmap('Shell-Icon.ico')

def main():
    global indexed_folder, indexed_files, indexed_dirs, result_label, search_entry, search_inside_files, indexed_folder_label

    # Create GUI
    root = tk.Tk()
    root.minsize(240, 300)
    root.title("File Indexer")

    root.iconbitmap('Shell-Icon.ico')

    indexed_folder_label = tk.Label(root, text="")
    indexed_folder_label.pack()

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
    result_listbox.pack(expand=True, fill=tk.BOTH)

    button_frame = tk.Frame(root)
    button_frame.pack()

    open_file_button = tk.Button(button_frame, text="Open File", command=lambda: open_file(result_listbox))
    open_file_button.pack(side=tk.LEFT)

    open_path_button = tk.Button(button_frame, text="Open Path", command=lambda: open_path(result_listbox))
    open_path_button.pack(side=tk.LEFT)

    quit_button = tk.Button(root, text="Quit", command=root.quit)
    quit_button.pack()

    result_label = tk.Label(root, text="")
    result_label.pack()

    frame = tk.Frame(root)
    frame.pack(pady=10)

    name_label = tk.Label(frame, text="© 2023 Gabriel Unsinn")
    name_label.pack(side=tk.LEFT)

    license_button = tk.Button(frame, text="View License", command=lambda: show_license(root))
    license_button.pack(side=tk.LEFT, padx=(10, 0))

    indexed_folder = ""
    indexed_files = 0
    indexed_dirs = 0

    root.mainloop()


if __name__ == "__main__":
    main()
