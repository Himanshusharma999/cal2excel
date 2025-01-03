import streamlit as st
import utils
import components.file_uploader
from io import BytesIO
import zipfile
import requests

def main():
    st.title("DBU Kalender til Excel")
    run_app()



def run_app():
    input_method = st.radio("Hvordan vil du uploade kalendere?",
                             ["Upload filer", "Indtast URL'er"])

    if input_method == "Upload filer":
        uploaded_files = components.file_uploader.upload_files()
    elif input_method == "Indtast URL'er":
        uploaded_files = components.file_uploader.fetch_ical_urls()

    if uploaded_files:
        # Process all files into a single CSV
        utils.parse_ics_to_csv_test(uploaded_files, "/tmp/descripted.csv")

        # Create single DataFrame with all events
        df = utils.mk_df()
        df = utils.fill_df(df, "/tmp/descripted.csv")
        
        # Create Excel buffer
        buffer = BytesIO()
        utils.to_excel_test(df, buffer)
        buffer.seek(0)

        # Single download button for the combined Excel
        st.download_button(
            label="Download samlet Excel fil",
            data=buffer,
            file_name="combined_calendar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
if __name__ == "__main__":
    main()