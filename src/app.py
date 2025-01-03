import streamlit as st
import utils
import components.file_uploader
from io import BytesIO
from datetime import datetime

def main():
    st.title("DBU Kalender til Excel")
    run_app()

def run_app():
    input_method = st.radio("Hvordan vil du uploade kalendere?",
                             ["Upload filer", "Indtast URL'er"])

    # Add date filter option
    filter_option = st.radio("Hvilke kampe vil du se?",
                           ["Alle kampe", "Kun fremtidige kampe"])

    if input_method == "Upload filer":
        uploaded_files = components.file_uploader.upload_files()
    elif input_method == "Indtast URL'er":
        uploaded_files = components.file_uploader.fetch_ical_urls()

    if uploaded_files and len(uploaded_files) > 0:
        try:
            # Preprocess each file
            processed_files = []
            for file in uploaded_files:
                processed = BytesIO()
                utils.preprocess(file, processed)
                processed.seek(0)
                processed_files.append(processed)

            # Process all files into a single CSV
            utils.parse_ics_to_csv_test(processed_files, "/tmp/descripted.csv")

            # Create single DataFrame with all events
            df = utils.mk_df()
            df = utils.fill_df(df, "/tmp/descripted.csv")
            
            # Filter for future games if selected
            if filter_option == "Kun fremtidige kampe":
                today = datetime.now().date()
                df = df.loc[df['Dato'] >= today]
                
                if len(df) == 0:
                    st.warning("Ingen fremtidige kampe fundet")
                    st.stop()
            
            # Create Excel buffer
            buffer = BytesIO()
            utils.to_excel_test(df, buffer)
            buffer.seek(0)

            st.success(f"Successfully processed {len(df)} game(s)")

            # Single download button for the combined Excel
            st.download_button(
                label="Download samlet Excel fil",
                data=buffer,
                file_name="combined_calendar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    
if __name__ == "__main__":
    main()