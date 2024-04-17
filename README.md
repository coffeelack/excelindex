© 2023 Gabriel Unsinn / Github: [https://github.com/coffeelack](https://github.com/coffeelack/excelindex)

# File Indexer - Help

Welcome to the File Indexer application! This tool allows you to index a folder, 
search for different files, and their content.

The program is written in Python 3.11 and uses the following libraries:
os, threding, tkinter, openpyxl, docx2txt, PyPDF2

## Prerequisites:

For the application to work properly, you need to have the following installed on your system:

Please select the **latest version of Python 3.11** for your operating system.
While installing Python, make sure to check the box that says **"Add Python to PATH"**.
https://www.python.org/downloads/

NOTE: The application will not work with Python 3.10 or earlier versions or 3.12 of Python 3, only Python 3.11 will work.

## Start the Application:

### Windows:

Execute the `windows_starter.bat` file to start the application.

### Linux:

Mark the `linux_starter.py` file as executable:

      chmod +x linux_starter.py

Execute the 'linux_starter.py' to start the application.

      ./linux_starter.py


## Indexing a Folder:

1. Select a file extension you want to index.
2. Click the "Select Folder" button.
3. Select the folder you want to index.
4. The application will recursively count the number of files and directories inside the selected folder.

## Searching for Files:

1. Enter a search term in the "Search" field.
2. By default only file names are being searched.
   Optionally, check the "File content" box to search within the content of Excel files
   and/or check the "Directory names" box to search within subdirectory names.
4. Click the "Search" button.
5. The application will display a list of files that match the search criteria.

## Opening Files or Paths:

- Select a file from the search results.
- Click either "Open File" to open the file directly or "Open Path" to open the file's containing folder.

## Viewing License Information:

Click the "View License" button to view the software's licensing details.

## Viewing Help Information:

Click the "Help" button to view this help text.

## Quitting the Program:

Click the "Quit" button to exit the application.

Note: Please ensure that you have proper permissions to access and modify the files in the indexed folder.

For further assistance or inquiries, refer to the "View License" section or contact the author via GitHub.
