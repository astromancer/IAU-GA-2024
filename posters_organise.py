# Script to process to raw dump of posters that we will receive.

import shutil
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog


def rename_pdfs(directory):
    # Loop through each file in the directory
    path = Path(directory)
    assert path.exists()

    # Get files starting with "submission_" and ending with ".pdf"
    for i, old_file in enumerate(path.glob('submission_*.pdf')):
        # Create the new file name by removing the "submission_" part
        filename = old_file.name
        new_filename = filename.remove_prefix('submission_')
        # Get the full path of the current and new file names
        new_file = path / new_filename
        # Rename the file
        old_file.rename(old_file, new_file)
        print(f'Renamed: {filename} -> {new_filename}')

    print(f'Renamed {i + 1} files.')


def sanitize_filename(filename):
    return filename.replace(' ', '_').replace(':', '').replace('\n', '')


def copy_and_assign_screens(df, pdf_directory, output_dir):
    # Extract unique dates and start times from the dataframe
    df['date'] = pd.to_datetime(df['date']).dt.date
    df['start time (local time)'] = pd.to_datetime(
        df['start time (local time)'], format='%H:%M:%S').dt.time

    unique_dates = df['date'].unique()

    output_dir = Path(output_dir)
    pdf_directory = Path(pdf_directory)

    session_labels = ['session 1', 'session 2']
    # Create directories for each unique date and start time
    for date in unique_dates:
        date_dir = output_dir / date.strftime('%Y-%m-%d')
        if not date_dir.exists():
            date_dir.mkdir()
            print(f'Created directory for date: {date_dir}')

        start_times = sorted(df[df['date'] == date]['start time (local time)'].unique())
        for session_label, start_time in zip(session_labels, start_times):
            # FIXME: `start_time` is not used in the loop below???
            start_time_dir = date_dir / session_label
            if start_time_dir.exists():
                start_time_dir.mkdir()
                print(f"Created directory: {start_time_dir}")

    # Assign screens and copy PDFs to the appropriate date and start time folder

    for date in unique_dates:
        date_str = date.strftime('%Y-%m-%d')
        for session_label in session_labels:
            screen_number = 1
            for index, row in df[(df['date'] == date) & (
                    df['start time (local time)'] == sorted(df[df['date'] == date]['start time (local time)'].unique())[
                        session_labels.index(session_label)])].iterrows():
                abstract_name = str(row['Abstract Submission ID'])
                pdf_filename = sanitize_filename(abstract_name) + '.pdf'
                pdf_path = pdf_directory / pdf_filename

                if pdf_path.exists():
                    new_pdf_filename = f'{screen_number}_{pdf_filename}'
                    destination_path = output_dir / date_str / session_label / new_pdf_filename
                    shutil.copy(pdf_path, destination_path)
                    print(f"Copied {pdf_filename} to {destination_path}")

                    # Update the dataframe with the screen number and rename the pdf
                    df.at[index, 'Screen number'] = screen_number
                    screen_number += 1
                else:
                    print(f'PDF {pdf_filename} not found in {pdf_directory}')

    print('All PDFs copied and screens assigned successfully.')
    return df


if __name__ == "__main__":
    # Create a Tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the root window

    # Open a file dialog to select the directory with PDF files
    pdf_directory = filedialog.askdirectory(title="Select Directory with PDF Files")

    # Check if a directory was selected
    if not pdf_directory:
        print("No directory selected.")
        exit(1)

    #
    rename_pdfs(pdf_directory)

    # Open a file dialog to select the Excel file
    excel_file_path = filedialog.askopenfilename(title="Select Excel File",
                                                 filetypes=[("Excel files", "*.xlsx *.xls")])

    # Check if an Excel file was selected
    if not excel_file_path:
        print("No Excel file selected.")
        exit(1)

    # Load the spreadsheet
    df = pd.read_excel(excel_file_path)

    # Open a file dialog to select the output directory for the poster sessions
    output_dir = filedialog.askdirectory(
        title="Select Output Directory for Poster Sessions")

    # Check if an output directory was selected
    if not output_dir:
        print("No output directory selected.")
        exit(1)

    # Main
    updated_df = copy_and_assign_screens(df, pdf_directory, output_dir)

    # Save the updated dataframe back to the Excel file
    updated_df.to_excel(excel_file_path, index=False)
    print("Updated Excel file with screen numbers.")
