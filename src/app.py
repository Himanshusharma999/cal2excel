import streamlit as st
import utils
import components.file_uploader
from io import BytesIO
import zipfile 

def main():
    st.title("DBU Kalender til Excel")
    run_app()

def run_app():
    st.header("Upload ical filer")
    uploaded_files = components.file_uploader.upload_files()

    if uploaded_files:
        zip_buffer = BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for idx, file in enumerate(uploaded_files):
                st.success(f"Uploaded: {file.name}")
                utils.delete_files_in_folders(["fixed", "csvs", "excels"])
                utils.preprocess(file, "/tmp/fixed_calendar.ics")
                utils.parse_ics_to_csv("/tmp/fixed_calendar.ics", "/tmp/fixed_calendar1.csv")

                df = utils.mk_df()
                df = utils.fill_df(df, "/tmp/fixed_calendar.csv")
                excel_name = df["Række"].iloc[0]

                buffer = BytesIO()
                utils.to_excel_test(df, buffer)

                st.download_button(
                    label=f"Download Excel fil",
                    data=buffer,
                    file_name=f"{excel_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{idx}"  # Unique key for each download button
                )

                # Add the Excel file to the ZIP archive
                zip_file.writestr(f"{excel_name}.xlsx", buffer.getvalue())

        # Finalize the ZIP file in memory
        zip_buffer.seek(0)

        # Single download button for the ZIP file
        st.download_button(
            label="Download alle excel filer i en ZIP fil",
            data=zip_buffer,
            file_name="Samlet_kalendere.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()