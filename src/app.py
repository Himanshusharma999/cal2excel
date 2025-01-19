import streamlit as st
import utils
import components.file_uploader
from io import BytesIO
from datetime import datetime
import pandas as pd

def main():
    run_app()

def run_app():
    # Clear cache
    st.cache_data.clear()

    # Add a page title with description
    st.title("DBU Ical til Excel")
    st.write("Upload dine DBU Ical-filer/links og få dem samlet i en Excel-fil")

    # Create two columns for input options
    col1, col2 = st.columns(2)
    with col1:
        input_method = st.radio("Hvordan vil du uploade kalendere?",
                                ["Upload filer", "Indtast links"])
    with col2:
        filter_option = st.radio("Hvilke kampe skal betragtes?",
                               ["Alle kampe", "Kun fremtidige kampe"])

    if input_method == "Upload filer":
        uploaded_files = components.file_uploader.upload_files()
    elif input_method == "Indtast links":
        uploaded_files = components.file_uploader.fetch_ical_urls()

    if uploaded_files and len(uploaded_files) > 0:
        try:
            # Show processing status
            with st.spinner('Behandler kalendere...'):
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
                    df['Dato'] = pd.to_datetime(df['Dato'], format='%d-%m-%Y').dt.date
                    df = df[df['Dato'] >= today]
                    
                    if len(df) == 0:
                        st.warning("Ingen fremtidige kampe fundet")
                        st.stop()
                
                # Create Excel buffer
                buffer = BytesIO()
                utils.to_excel_test(df, buffer)
                buffer.seek(0)

            # Show success message with game count
            st.success(f"Behandlet {len(df)} kampe succesfuldt")

            # Show summary statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Antal kampe", len(df))
            with col2:
                st.metric("Antal Rækker", len(df['Række'].unique()))
            with col3:
                st.metric("Antal uger", len(df['Uge'].unique()))

            # Download button with clear call-to-action
            st.download_button(
                label="📥 Download samlet Excel fil",
                data=buffer,
                file_name="combined_calendar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Klik her for at downloade alle kampe i én Excel-fil"
            )

        except Exception as e:
            st.error(f"Der opstod en fejl: {str(e)}")
            # Log the full error for debugging
            st.exception(e)
    
if __name__ == "__main__":
    main()
