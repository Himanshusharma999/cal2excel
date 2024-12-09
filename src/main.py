import functions as f
import pandas as pd
import streamlit as st

input_ics = "../icals/kampe5db2b479-c637-46cf-b23b-6fb2736f5352.ics"
fixed_ics = "fixed/fixed_calendar13s2222.ics"
output_csv = "csvs/calendar_events22s2222.csv"

f.preprocess(input_ics, fixed_ics)

f.parse_ics_to_csv(fixed_ics, output_csv)

df = f.mk_df()

df = f.fill_df(df, output_csv)

f.to_excel(df, "excels/test62222.xlsx")