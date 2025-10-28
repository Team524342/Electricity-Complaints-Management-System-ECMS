import pandas as pd

df1 = pd.read_excel('data/complaints.xlsx')
df2 = pd.read_excel("data/voiceComplaint.xlsx")

merged_df = pd.concat([df1, df2], ignore_index=True)

merged_df.to_excel('merged_file.xlsx', index=False)