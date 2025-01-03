import utils as f
import pandas as pd

input_ics = "../icals/kampe5db2b479-c637-46cf-b23b-6fb2736f5352.ics"
input_ics2 = "../icals/kampe9b1978b9-7540-4643-a79a-b4225180718b.ics"

fixed_ics = "fixed/test.ics"
fixed_ics2 = "fixed/test2.ics"

output_csv = "csvs/test.csv"
excel_output = "excels/test.xlsx"

f.delete_files_in_folders(["fixed", "csvs", "excels"])

f.preprocess(input_ics, fixed_ics)
f.preprocess(input_ics2, fixed_ics2)

f.parse_ics_to_csv_test([fixed_ics, fixed_ics2], output_csv)

df = f.mk_df()

df = f.fill_df(df, output_csv)

f.to_excel(df, excel_output)