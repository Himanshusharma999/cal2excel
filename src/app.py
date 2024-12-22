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
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for idx, file in enumerate(uploaded_files):
                st.success(f"Uploaded: {file.name}")
                utils.delete_files_in_folders(["fixed", "csvs", "excels"])
                utils.preprocess(file, "/tmp/fixed_calendar.ics")
                utils.parse_ics_to_csv("/tmp/fixed_calendar.ics", "/tmp/descripted.csv")

                df = utils.mk_df()
                df = utils.fill_df(df, "/tmp/descripted.csv")
                excel_name = df["Række"].iloc[0]
                st.write(f"Debug - Excel name: {excel_name}")
                
                buffer = BytesIO()
                utils.to_excel_test(df, buffer)
                buffer.seek(0)

                zip_file.writestr(f"{excel_name}.xlsx", buffer.getvalue())

                st.download_button(
                    label=f"Download Excel fil",
                    data=buffer,
                    file_name=f"{excel_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{idx}"  # Unique key for each download button
                )                

        # Finalize the ZIP file in memory
        #zip_buffer.seek(0)

        # Single download button for the ZIP file
        st.download_button(
            label="Download alle excel filer i en ZIP fil",
            data=zip_buffer,
            file_name="calendars.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()