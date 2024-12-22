import utils as f
import pandas as pd

input_ics = "../icals/kampec8e03235-0e48-4760-9cf9-fbd67799bb74.ics"
fixed_ics = "fixed/test.ics"
output_csv = "csvs/test.csv"
excel_output = "excels/test.xlsx"

f.delete_files_in_folders(["fixed", "csvs", "excels"])

f.preprocess(input_ics, fixed_ics)

f.parse_ics_to_csv(fixed_ics, output_csv)

df = f.mk_df()

df = f.fill_df(df, output_csv)

f.to_excel(df, excel_output)