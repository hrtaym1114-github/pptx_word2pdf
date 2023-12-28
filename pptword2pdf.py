import os
import tkinter as tk
from tkinter import filedialog
import comtypes.client
import win32com.client
from pathlib import Path

def convert_to_pdf(input_file_path, output_file_path):
    file_extension = os.path.splitext(input_file_path)[1].lower()

    # Check if output directory exists and is writable
    output_dir = os.path.dirname(output_file_path)
    if not os.path.exists(output_dir) or not os.access(output_dir, os.W_OK):
        print(f"Output directory {output_dir} does not exist or is not writable.")
        return

    # Delete the output file if it already exists
    if os.path.exists(output_file_path):
        os.remove(output_file_path)

    if file_extension == '.pptx':
        # Powerpoint application setup
        application = win32com.client.Dispatch('Powerpoint.Application')

        # File path specification
        pptx = Path(input_file_path)
        pdf = Path(output_file_path)

        # Property specification
        read_only = True  # Read-only
        title = False  # Title setting
        window = False  # Window display

        # Open the Powerpoint file and save it in PDF format
        presentation = application.Presentations.Open(pptx, read_only, title, window)
        presentation.SaveAs(pdf, 32)  # 32=filetype

        # Application termination processing
        presentation.close()
        application.quit()
        presentation = None
        application = None

    elif file_extension == '.docx':
        # Word application setup
        word = win32com.client.Dispatch('Word.Application')

        # File path specification
        docx = Path(input_file_path)
        pdf = Path(output_file_path)

        # Open the Word file and save it in PDF format
        doc = word.Documents.Open(docx)
        doc.SaveAs(pdf, FileFormat=17)  # 17=filetype

        # Application termination processing
        doc.Close()
        word.Quit()
        doc = None
        word = None

def convert_folder_to_pdf(folder_path):
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path) and file_path.endswith(('.pptx', '.docx')):
            pdf_path = file_path + '.pdf'
            convert_to_pdf(file_path, pdf_path)
            print(f"Converted: {file_path} to {pdf_path}")

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    folder_path = filedialog.askdirectory()  # Show the folder selection dialog
    if folder_path:
        convert_folder_to_pdf(folder_path)
        print("All files have been converted.")
    else:
        print("No folder selected.")

select_folder()
