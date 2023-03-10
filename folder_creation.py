import os
import shutil
from openpyxl import load_workbook
from tkinter import *
from tkinter import filedialog as fd
import datetime

xlsx_path = ''
folder_path = ''

root = Tk()
root.title("Excel manager")

root.resizable(False, False)

db_button = Button(root, text="Wybierz plik Excel", command=lambda: open_file())
db_button.grid(row=0, column=0)
db_name_label1 = Label(root, text='Plik Excel nie zostal wybrany')
db_name_label1.grid(row=2, column=0)

db_name_label = Label(root, text="Wpisz nazwe arkusza: ")
db_name_label.grid(row=4, column=0)
db_name = Entry(root, width=60)
db_name.grid(row=6, column=0)

open_button = Button(root, text="Wybierz folder z plikami", command=lambda: open_folder())
open_button.grid(row=8, column=0)
db_name_label2 = Label(root, text='Folder z plikami nie zostal wybrany')
db_name_label2.grid(row=10, column=0)

open_button = Button(root, text="Wykonaj", command=lambda: transform_excel())
open_button.grid(row=12, column=0)

error_label = Label(root, text='')
error_label.grid(row=14, column=0)


def open_file():
    global xlsx_path
    filetype=('excel file','*.xlsx;*.xlsm;*.xltx;*.xltm'), ('All files', '*.*')
    filepath=fd.askopenfilenames(filetypes=filetype)
    filepath2 = ''
    for i in filepath:
        filepath2 = filepath2 + i
        xlsx_path = filepath2
    db_name_label1.config(text=xlsx_path)


def open_folder():
    global folder_path
    mr_dir = fd.askdirectory()
    folder_path = mr_dir
    db_name_label2.config(text=folder_path)


def transform_excel():
    global xlsx_path
    global folder_path
    now = datetime.datetime.now()
    # Check if the XLSX file path is valid
    if not os.path.isfile(xlsx_path) or not xlsx_path.endswith('.xlsx'):
        with open('errors.txt', 'w') as f:
            f.write(f'Zly plik Excel\n {now}')
        error_label.config(text='Zly plik Excel')
        return
        
    # Check if the folder path is valid
    if not os.path.isdir(folder_path):
        with open('errors.txt', 'w') as f:
            f.write(f'Zla sciezka folderu z plikami\n {now}')
        error_label.config(text='Zla sciezka folderu z plikami')
        return
            
    workbook = load_workbook(xlsx_path)
    worksheet_name = db_name.get()
    
    try:
        worksheet = workbook[worksheet_name]
    except KeyError:
        with open('errors.txt', 'w') as f:
            f.write(f'Zla nazwa arkusza: {worksheet_name} {now}\n')
        error_label.config(text=f'Zla nazwa arkusza: {worksheet_name}')
        return

    # Loop through the rows of the worksheet
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        # Get the values from the columns
        try:
            col1, col2, col3, col4 = row
        except ValueError:
            with open('errors.txt', 'w') as f:
                f.write(f'Zle dane w wierszu: {row}\n {now}')
            error_label.config(text=f'Zle dane w wierszu: {row}')
            continue

        # Determine the child folder path
        if col2 and col3 and col4:
            child_folder = os.path.join(os.getcwd(), col3, str(col4))
        else:
            child_folder = os.path.join(os.getcwd(), 'no parameter')
        os.makedirs(child_folder, exist_ok=True)

        # Loop through the files in the specific folder and copy them to the child folder
        specific_folder = os.path.join(os.getcwd(), folder_path)
        for filename in os.listdir(specific_folder):
            if col1 in filename:
                src_path = os.path.join(specific_folder, filename)
                if col2:
                    dst_filename = str(col1) + '_' + str(col2) + os.path.splitext(filename)[1]
                else:
                    dst_filename = col1 + os.path.splitext(filename)[1]
                dst_path = os.path.join(child_folder, dst_filename)
                try:
                    shutil.copy(src_path, dst_path)
                    error_label.config(text='Pliki skopiowane poprawnie')
                except Exception as e:
                    with open('errors.txt', 'w') as f:
                        f.write(f'Blad kopiowania pliku {src_path} to {dst_path}: {e} {now}\n')
                    error_label.config(text=f'Blad kopiowania pliku {src_path} to {dst_path}: {e}')


root.mainloop()
