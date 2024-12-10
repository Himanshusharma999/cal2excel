import streamlit as st
import utils
import components.file_uploader
from io import BytesIO 

def main():
    st.title("DBU Kalender til Excel")
    run_app()

def run_app():
    st.header("Upload ical filer")
    uploaded_files = components.file_uploader.upload_files()

    if uploaded_files:
        for file in uploaded_files:
            st.success(f"Uploaded: {file.name}")
            utils.preprocess(file, "fixed/fixed_calendar.ics")
            utils.parse_ics_to_csv("fixed/fixed_calendar.ics", "csvs/fixed_calendar.csv")

        df = utils.mk_df()
        df = utils.fill_df(df, "csvs/fixed_calendar.csv")

        buffer = BytesIO()
        utils.to_excel_test(df, buffer)
        st.download_button(
            label="Download Excel fil",
            data=buffer,
            file_name="DBU Kalender.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()