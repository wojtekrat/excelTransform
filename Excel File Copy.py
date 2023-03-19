import os
import shutil
from tkinter import *
from tkinter import filedialog as fd
import datetime
import pandas as pd
from customtkinter import *

root = CTk()
set_appearance_mode("System")  # Modes: system (default), light, dark
set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green
root.title("Excel manager")
root.resizable(False, False)

main_frame = CTkFrame(root)
main_frame.pack(padx=50, pady=50)

frame_excel = CTkFrame(main_frame)
frame_excel.pack(padx=5, pady=5)
db_button = CTkButton(frame_excel, text="Choose Excel file", command=lambda: open_file())
db_button.grid(row=1, column=0, padx=5, pady=5)
db_name_label1 = CTkLabel(frame_excel, text='Excel file not chosen')
db_name_label1.grid(row=2, column=0, padx=5, pady=5)

frame_excel2 = CTkFrame(main_frame)
frame_excel2.pack(padx=5, pady=5)
db_name_label = CTkLabel(frame_excel2, text="Choose sheet: ")
db_name_label.grid(row=1, column=0, padx=5, pady=5)
sheet_dropdown = CTkOptionMenu(frame_excel2)

frame_folder1 = CTkFrame(main_frame)
frame_folder1.pack(padx=5, pady=5)
open_button = CTkButton(frame_folder1, text="Choose folder with files", command=lambda: open_folder())
open_button.grid(row=1, column=0, padx=5, pady=5)
db_name_label2 = CTkLabel(frame_folder1, text='Folder with files has not been chosen')
db_name_label2.grid(row=2, column=0, padx=5, pady=5)

frame_folder2 = CTkFrame(main_frame)
frame_folder2.pack(padx=5, pady=5)
open_button2 = CTkButton(frame_folder2, text="Choose output folder", command=lambda: select_output_folder())
open_button2.grid(row=1, column=0, padx=5, pady=5)
db_name_label3 = CTkLabel(frame_folder2, text='Output folder has not been chosen')
db_name_label3.grid(row=2, column=0, padx=5, pady=5)

frame_progress = CTkFrame(main_frame)
frame_progress.pack(padx=5, pady=5)
progress_label = CTkLabel(frame_progress, text="Files copy progress")
progress_label.pack(pady=10)
progress_bar = CTkProgressBar(frame_progress, orientation=HORIZONTAL, width=200, height=20, mode='determinate')
progress_bar.set(0)
progress_bar.pack(pady=10)
progress_label = CTkLabel(frame_progress, text="Processed files: 0 / 0")
progress_label.pack(pady=10)

open_button3 = CTkButton(main_frame, text="Copy files", width=100, height=50, command=lambda: transform_excel())
open_button3.pack(padx=5, pady=5)

error_label = CTkLabel(main_frame, text='')
error_label.pack(padx=5, pady=5)

now = datetime.datetime.now()

def open_file():
    global xlsx_path, sheet_names, choice

    # Select file and update path label
    filetype = [('Excel files', '*.xlsx;*.xlsm;*.xls;*.xlsb;*.xltx;*.xltm;*.xlam;*.xla;*.xlw;*.xlr;*.xml')]
    filepath = fd.askopenfilenames(filetypes=filetype)
    filepath2 = ''
    for i in filepath:
        filepath2 = filepath2 + i
        xlsx_path = filepath2
    db_name_label1.configure(text=xlsx_path)

    # Load all sheet names from the Excel file
    xl = pd.ExcelFile(xlsx_path)
    sheet_names = xl.sheet_names

    # Create a dropdown menu to choose a sheet
    choice = StringVar()
    sheet_dropdown.configure(variable=choice, values=sheet_names)
    sheet_dropdown.grid(row=2, column=0, padx=5, pady=5)

    def update_choice(*args):
        global choice
        choice = choice.get()
    choice.trace('w', update_choice)
    choice.set(sheet_names[0])  # Set default value

def open_folder():
    global folder_path
    mr_dir = fd.askdirectory()
    folder_path = mr_dir
    db_name_label2.configure(text=folder_path)

def select_output_folder():
    global output_folder 
    output_folder = fd.askdirectory()
    db_name_label3.configure(text=output_folder)
    return output_folder

def validate_paths():
    global now
        # Check if the XLSX file path is valid
    if not os.path.isfile(xlsx_path) or not xlsx_path.endswith('.xlsx'):
        with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
            f.write(f'Invalid Excel file\n {now}')
        error_label.configure(text='Invalid Excel file')
        return False
        # Check if the folder path is valid
    if not os.path.isdir(folder_path):
        with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
            f.write(f'Invalid files path\n {now}')
        error_label.configure(text='Invalid files path')
        return False
    return True

def get_worksheet():
    global xlsx_path, worksheet, output_folder
    try:
        worksheet_name = choice
        
        # Load worksheet using pandas
        xl = pd.ExcelFile(xlsx_path)
        worksheet = xl.parse(worksheet_name)
        progress_label.configure(text=f"Processed files: 0 / {len(worksheet.index)}")
        
        # Fill NaN/empty values with empty string
        worksheet.fillna('', inplace=True)
        
        # Create parent folder
        excel_filename = os.path.basename(xlsx_path)
        parent_folder_name = os.path.splitext(excel_filename)[0] + '_' + now.date().isoformat()
        parent_folder_path = os.path.join(output_folder, parent_folder_name)
        os.makedirs(parent_folder_path, exist_ok=True)
        
        # Update child folder path
        output_folder = parent_folder_path
        return worksheet  
        
    except KeyError:
        with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
            f.write(f'Invalid sheet name: {worksheet_name} {now}\n')
        error_label.configure(text=f'Invalid sheet name: {worksheet_name}')
        return None
    
def transform_excel():
    global xlsx_path, folder_path, choice, output_folder, now, worksheet

    if validate_paths():
        get_worksheet()
        process_rows()

def process_rows():
    total_files = len(worksheet.index)
    processed_files = 0
    for i, row in worksheet.iterrows():
        progress_bar.set(processed_files)
        process_row(row)
        processed_files += 1
        progress_bar.update()
        progress_label.configure(text=f"Processed files: {processed_files} / {total_files}")

def process_row(row):
    try:
        col1, col2, col3, col4 = row
    except ValueError:
        handle_invalid_row(row)
        return

    child_folder = determine_child_folder(str(col2), str(col3), str(col4))
    os.makedirs(child_folder, exist_ok=True)

    process_files(col1, child_folder, col2)

def determine_child_folder(col2, col3, col4):
    if col2 and col3 and col4:
        return os.path.join(output_folder, str(col3), str(col4))
    else:
        return os.path.join(output_folder, 'no parameter')

def process_files(col1, child_folder, col2):
    specific_folder = os.path.join(os.getcwd(), folder_path)
    files_found = False

    for root, dirs, files in os.walk(specific_folder):
        for filename in files:
            if col1 in filename:
                src_path = os.path.join(root, filename)
                dst_filename = determine_destination_filename(col1, col2, filename)
                dst_path = os.path.join(child_folder, dst_filename)
                if os.path.exists(dst_path):
                    handle_existing_file(dst_path)
                    continue

                try:
                    copy_file(src_path, dst_path, col2, filename, child_folder)
                    files_found = True
                except Exception as e:
                    handle_copy_error(src_path, dst_path, e)

    if not files_found:
        handle_no_files_found_error(col1)

def copy_file(src_path, dst_path, col2, filename, child_folder):
    # Check if col2 is empty or not
    if col2:
        # Convert col2 to integer
        col2_int = int(col2)
        # Check if the file extension is .dxf
        if filename.lower().endswith('.dxf'):
            # Copy the file for col2 number of times
            for i in range(col2_int):
                # Add the col2 value to the filename
                new_filename = f"{os.path.splitext(filename)[0]}_{i+1}{os.path.splitext(filename)[1]}"
                # Construct the destination path for the new file
                new_dst_path = os.path.join(child_folder, new_filename)
                shutil.copy(src_path, new_dst_path)
                handle_successful_copy()
        else:
            # Copy the file just one time
            shutil.copy(src_path, dst_path)
            handle_successful_copy()
    else:
        # Copy the file to 'no parameter' folder
        shutil.copy(src_path, dst_path)
        handle_successful_copy()

def handle_no_files_found_error(col1):
    with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
        f.write(f"No files found for column 1 value: {col1} {now}\n")
    error_label.configure(text=f"No files found for column 1 value: {col1}")

def determine_destination_filename(col1, col2, filename):
    if col2:
        return f"{os.path.splitext(filename)[0]}_{col2}{os.path.splitext(filename)[1]}"
    else:
        return f"{os.path.splitext(filename)[0]}{os.path.splitext(filename)[1]}"

def handle_invalid_row(row):
    with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
        f.write(f'Invalid data in row: {row} {now}\n')
    error_label.configure(text=f'Invalid data in row: {row}')

def handle_existing_file(dst_path):
    with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
        f.write(f'File {dst_path} already exists\n {now}\n')
    error_label.configure(text=f'File {dst_path} already exists')

def handle_successful_copy():
    error_label.configure(text='Files copied properly')

def handle_copy_error(src_path, dst_path, e):
    with open(os.path.join(output_folder, 'errors.txt'), 'a+') as f:
        f.write(f'file copy error {src_path} to {dst_path}: {e} {now}\n')
    error_label.configure(text=f'File copy error {src_path} to {dst_path}: {e}')

root.mainloop()