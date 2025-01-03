from ics import Calendar
import csv
#import streamlit as st
import pandas as pd
from datetime import datetime
import locale
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import glob

def preprocess(file_object, output_path):
    """
    Adjust the formatting of LOCATION lines so that the next two lines are indented
    and replace occurrences of "Ny Stadion" with Ny Stadion.
    """
    #with open(file_object, 'r', encoding='utf-8') as f:
    #    lines = f.readlines()
    
    lines = file_object.read().decode('utf-8').split('\n')

    # Combine all lines into a single string for easier processing
    content = ''.join(lines)

    # Replace "Ny Stadion" with Ny Stadion
    content = content.replace('"Ny Stadion"', 'Ny Stadion')

    # Split the content back into lines
    lines = content.splitlines()

    processed_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        processed_lines.append(line + "\n")  # Ensure lines retain newline characters

        # Check if this line is a LOCATION line
        if line.startswith("LOCATION:"):
            # Indent the next two lines if they exist
            if i + 1 < len(lines) and not lines[i + 1].startswith(" "):
                processed_lines.append(" " + lines[i + 1].strip() + "\n")
                i += 1
            if i + 1 < len(lines) and not lines[i + 1].startswith(" "):
                processed_lines.append(" " + lines[i + 1].strip() + "\n")
                i += 1
        i += 1

    # Write the processed lines back to the output file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(processed_lines)

def parse_ics_to_csv_test(ics_files, csv_file):
    """Parse multiple ICS files and export events to CSV."""
    all_events = []

    # Handle single file or list of files
    if not isinstance(ics_files, list):
        ics_files = [ics_files]

    # Collect events from all files
    for ics_file in ics_files:
        with open(ics_file, 'r', encoding='utf-8') as f:
            content = f.read()
            calendar = Calendar(content)
            all_events.extend(calendar.events)

    # Sort events by date and time
    all_events.sort(key=lambda x: x.begin)

    # Write to CSV
    with open(csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Description']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for event in all_events:
            writer.writerow({
                'Description': event.description or ''
            })

def parse_ics_to_csv(ics_file, csv_file):
    """Parse ICS file and export events to CSV."""

    with open(ics_file, 'r', encoding='utf-8') as f:
        content = f.read()
        calendar = Calendar(content)

    # Open CSV file for writing
    with open(csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Description']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for event in calendar.events:
            writer.writerow({
                'Description': event.description or ''
            })
            

def parse_entry(content):
    try:
        # Regex patterns to extract fields
        kampnr_pattern = r"Kampnr (\d+)"
        date_time_pattern = r"(\d{2}-\d{2}-\d{4}) kl\. (\d{2}:\d{2})"

        # Extract data using regex
        kampnr_match = re.search(kampnr_pattern, content)
        date_time_match = re.search(date_time_pattern, content)

        # Assign default values if matches are missing
        kampnr = kampnr_match.group(1) if kampnr_match else "Unknown"
        kampnr = int(kampnr)
        dato, tidspunkt = (
            (date_time_match.group(1), date_time_match.group(2))
            if date_time_match
            else ("Unknown", "Unknown")
        )

        # Extract league name (Række) and Årgang
        række = content.split("\n")[0].strip()
        årgang = content.split("\n")[0].split(" ")[0]  

        # Parse the date
        date_object = datetime.strptime(dato, "%d-%m-%Y")

        # Get the week number
        ugenr = date_object.strftime("%W")
        ugenr = int(ugenr)

        # Get the day of the week
        dag = date_object.strftime("%A").capitalize()  # Full day name (e.g., Saturday)

        # Convert to danish
        dag = get_day_of_week_danish(dag)

        # Home and away teams
        hjem = content.split("\n")[3].split(" - ")[0]
        ude = content.split("\n")[3].split(" - ")[1]

        # Region
        postnr = content.split("\n")[7].split(" ")[0]
        postnr = int(postnr)
        region = (
            "Øst" if postnr < 5000 else "Vest"
        )

        return [årgang, dag, dato, ugenr, tidspunkt, række, kampnr, hjem, ude, region]
    except Exception as e:
        print(f"Error parsing entry: {content}\n{e}")
        return None

def get_day_of_week_danish(day_of_week):
    danish_days = {
        "Monday": "Mandag",
        "Tuesday": "Tirsdag",
        "Wednesday": "Onsdag",
        "Thursday": "Torsdag",
        "Friday": "Fredag",
        "Saturday": "Lørdag",
        "Sunday": "Søndag"
    }

    return danish_days.get(day_of_week, day_of_week)

def mk_df():
    df = pd.DataFrame(columns=["Årgang", "Dag", "Dato", "Uge", "Tidspunkt", "Række", "Kampnr", "Hjem", "Ude", "Region"])

    return df

def fill_df(df, input_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # Split the data into entries by matching content between double quotes
    entries = re.findall(r'"(.*?)"', raw_data, re.DOTALL)
    fixed_entries = [item.replace("*** IKKE TIDS FASTSAT ****\n\n", "") for item in entries]
    pattern = r'\nInkl\. straffe: \d+-\d+'
    cleaned_entries = [re.sub(pattern, "", item) for item in fixed_entries]
    for entry in cleaned_entries:
        df.loc[len(df)] = parse_entry(entry)

    df["Scout"] = ""

    return df

def fill_df_test(df, input_path):
    with open(input_path, 'r', encoding='utf-8') as f:
        raw_data = f.read()

    # Split the data into entries by matching content between double quotes
    entries = re.findall(r'"(.*?)"', raw_data, re.DOTALL)
    #print(entries)

    # Process each entry: remove unwanted text and clean up the data
    fixed_entries = [item.replace("*** IKKE TIDS FASTSAT ****\n\n", "") for item in entries]
    
    # Remove 'Inkl. straffe' part (match any score format like '6-7')
    pattern = r'\nInkl\. straffe: \d+-\d+'
    cleaned_entries = [re.sub(pattern, "", item) for item in fixed_entries]

    # Convert the cleaned entries into a format suitable for DataFrame (e.g., a dictionary or list)
    data_to_add = []
    for entry in cleaned_entries:
        # Call the function that extracts the structured data from each entry
        parsed_data = parse_entry(entry)  # Assuming parse_entry is already handling extraction of structured data
        print(parsed_data)
        data_to_add.append(parsed_data)

    # Assuming that parse_entry returns a dict or list matching the columns in your DataFrame
    # Add all parsed entries to the DataFrame
    df = df._append(data_to_add, ignore_index=True)
    
    # Optionally, add any default values or computed columns, e.g., the "Scout" column
    df["Scout"] = ""

    return df

def to_excel(df, output_path):
    df = df.map(lambda x: x.replace("��", "ø") if isinstance(x, str) else x)
    df['Dato'] = pd.to_datetime(df['Dato'], format='%d-%m-%Y').dt.date
    df.sort_values(by=["Dato", "Tidspunkt"], ascending=[True, True], inplace=True)
    df.to_excel(f"{output_path}", index=False)

    # Make columns wider
    wb = load_workbook(output_path)
    ws = wb.active

    ws.column_dimensions['C'].width = 11.5
    ws.column_dimensions['D'].width = 5
    ws.column_dimensions['F'].width = 33
    ws.column_dimensions['H'].width = 22
    ws.column_dimensions['I'].width = 22

    # Define the blue fill for every second row
    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    # Apply blue background to every second row starting from row 1 (Excel rows start at 1)
    for row in range(1, len(df) + 2, 2):  # Start from row 1 (Excel row 1 corresponds to pandas header)
        for cell in ws[row]:
            cell.fill = blue_fill

    wb.save(output_path)


def to_excel_test(df, buffer):
    """
    Converts a DataFrame to a styled Excel file and writes it to a BytesIO buffer.
    """
    # Replace problematic characters
    df = df.map(lambda x: x.replace("��", "ø") if isinstance(x, str) else x)

    # Process and sort the DataFrame
    df['Dato'] = pd.to_datetime(df['Dato'], format='%d-%m-%Y').dt.date
    df.sort_values(by=["Dato", "Tidspunkt"], ascending=[True, True], inplace=True)

    # Save DataFrame to an Excel buffer
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    # Load the workbook from the buffer
    buffer.seek(0)
    wb = load_workbook(buffer)
    ws = wb.active

    # Adjust column widths
    ws.column_dimensions['C'].width = 11.5
    ws.column_dimensions['E'].width = 33
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 22

    # Define the blue fill for every second row
    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")

    # Apply blue background to every second row starting from row 2
    for row in range(2, len(df) + 2, 2):  # Start from row 2 (header is row 1)
        for cell in ws[row]:
            cell.fill = blue_fill

    # Save the workbook back to the buffer
    buffer.seek(0)
    wb.save(buffer)
    buffer.seek(0)

def delete_files_in_folders(folder_names):
    """
    Deletes all files inside the specified folders.
    
    Parameters:
        folder_names (list): List of folder names (relative to script's directory).
    """
    for folder in folder_names:
        if os.path.exists(folder) and os.path.isdir(folder):
            # Find all files in the folder
            files = glob.glob(os.path.join(folder, '*'))
            for file in files:
                if os.path.isfile(file):  # Only delete files
                    try:
                        os.remove(file)
                        print(f"Deleted: {file}")
                    except Exception as e:
                        print(f"Error deleting {file}: {e}")
        else:
            print(f"Folder '{folder}' does not exist or is not a directory.")