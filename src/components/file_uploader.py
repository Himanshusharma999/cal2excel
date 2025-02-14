import streamlit as st
import requests
from io import BytesIO

def upload_files():
    return st.file_uploader("Vælg flere ad gangen", type=['ics', 'ical'], accept_multiple_files=True)

def fetch_ical_urls():
    """
    Component for fetching iCal files from URLs.
    Returns a list of BytesIO objects containing the calendar data.
    """
    files_to_process = []
    
    urls = st.text_area("Indtast iCal URL (et link per linje), tryk (cmd + enter) for at eksekvere")
    
    if urls and st.button("Hent kalendere"):
        for url in urls.strip().split('\n'):
            try:
                # Convert webcal:// to https://
                url = url.strip()
                if url.startswith('webcal://'):
                    url = 'https://' + url[9:]
                
                response = requests.get(url)
                if response.status_code == 200:
                    content = response.content.decode('utf-8')
                    if 'BEGIN:VCALENDAR' not in content:
                        st.error(f"Invalid iCal file: {url}")
                        continue

                    file_like = BytesIO(response.content)
                    file_like.name = f"calendar_{len(files_to_process)}.ics"
                    files_to_process.append(file_like)

                else:
                    st.error(f"Failed to fetch calendar from: {url}")
            except Exception as e:
                st.error(f"Error fetching calendar from {url}: {str(e)}")
    
    return files_to_process
