import sys, os, datetime, time
import traceback
from pathlib import Path

if sys.platform == "win32":
    import win32com.client

#import requests
#from html.parser import HTMLParser

from docx import Document
from docxcompose.composer import Composer

import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

DEBUG = False

if getattr(sys, 'frozen', False):  
    applicationPath = os.path.dirname(sys.executable)
else:
    applicationPath = os.path.dirname(os.path.abspath(__file__))

output_dir = Path(applicationPath, "saved")
output_dir.mkdir(parents=True, exist_ok=True)

output_dir = Path(applicationPath, "templates")
output_dir.mkdir(parents=True, exist_ok=True)

output_dir = Path(applicationPath, "resume")
output_dir.mkdir(parents=True, exist_ok=True)


def select_file(initialdir='/', filetypes = ('All files', '*.*')):
    """Opens a file dialog and returns the selected file path."""

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir=initialdir, 
        filetypes=filetypes
    )

    if filename:
        return filename
    else:
        showinfo(
            title='No File Selected',
            message='No file was chosen.'
        )
        return None


def select_directory(initialdir='/'):
    folder_path = fd.askdirectory(
    title="Select a folder",
    initialdir=initialdir
    )

    if folder_path:
            return folder_path
    else:
        showinfo(
            title='No folder Selected',
            message='No folder was chosen.'
        )
        return None

    
def replace_keyords(document, keyword, keyword_responses):
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if f"${keyword}" in run.text:
                run.text = run.text.replace(f"${keyword}", keyword_responses[keyword])
                if DEBUG:
                    print("=====================================") 
                    print(paragraph.text)

    return document


def start():

    filetypes = (
        ('Docx files', '*.docx'),
        ('All files', '*.*')
    )

    print("Select Cover Letter Template (must be .docx) in the file explorer window")
    selected_filepath = select_file(Path(applicationPath, "templates"), filetypes)
    if selected_filepath == None:
        print("Template file must be selected. Stoping Application.")
        return


    document = Document(selected_filepath)

    keywords = ["DATE", "ADDRESS", "POSITION", "COMPANY"]
    keyword_responses = {}

    # puts the responses in the dict
    for keyword in keywords:        
        if keyword == "DATE":
            insert_text = datetime.datetime.now().strftime("%B %d, %Y")
        elif keyword == "ADDRESS":
            insert_text = '\n' + input(f"Input for {keyword.lower()}: ").replace(", ", '\n', 1)
        else:
            insert_text = input(f"Input for {keyword.lower()}: ")

        keyword_responses[keyword] = insert_text

    for keyword in keywords:
        document = replace_keyords(document, keyword, keyword_responses)

    # appending the resume to the coverletter 
    print("Select Resume (must be .docx) in the file explorer window or exit to skip")
    selected_resume = select_file(Path(applicationPath, "resume"), filetypes)       

    if selected_resume is not None:
        # document.add_page_break()
        resume_document = Document(selected_resume)

        composer = Composer(document)
        composer.append(resume_document)

    # save the file
    file_name = input("File name:")

    print("Select save folder in the file explorer window or exit to put in default saved file")
    folder_path = select_directory(Path(applicationPath, "saved"))  
    if folder_path is None:
        folder_path = f"{applicationPath}/saved"
        os.makedirs(folder_path, exist_ok=True)

    file_dir = os.path.join(folder_path, f"{file_name}.docx")
    document.save(file_dir)
    time.sleep(2)

    # open the file in word 
    try:
        if sys.platform == "win32":
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True
            document = word.Documents.Open(file_dir)
    except Exception as e: 
        traceback.print_exc()
        print("something went wrong. try opening manually")


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window

    start()