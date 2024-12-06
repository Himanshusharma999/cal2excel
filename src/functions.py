from ics import Calendar
import csv
import streamlit as st
import pandas as pd

def preprocess(input_path, output_path):
    """
    Adjust the formatting of LOCATION lines so that the next two lines are indented.
    """
    with open(input_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    processed_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        processed_lines.append(line)

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

    # Write the processed lines to a new file
    with open(output_path, 'w', encoding='utf-8') as f:
        f.writelines(processed_lines)

def parse_ics_to_csv(ics_file, csv_file):
    """Parse ICS file and export events to CSV."""
    with open(ics_file, 'r', encoding='utf-8') as f:
        calendar = Calendar(f.read())

    # Open CSV file for writing
    with open(csv_file, mode='w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Description']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for event in calendar.events:
            writer.writerow({
                'Description': event.description or ''
            })

input_ics = "../icals/kampe5db2b479-c637-46cf-b23b-6fb2736f5352.ics"
fixed_ics = "fixed_calendar.ics"
output_csv = "calendar_events.csv"

preprocess(input_ics, fixed_ics)

parse_ics_to_csv(fixed_ics, output_csv)

df = pd.read_csv(output_csv, encoding='utf-8')

st.write(df)