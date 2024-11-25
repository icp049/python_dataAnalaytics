import pandas as pd
import os

# Load the data
df = pd.read_excel('data.xlsx')

# Ensure datetime conversion for relevant columns
df['pcrStartTime'] = pd.to_datetime(df['pcrStartTime'], errors='coerce')
df['pcrStopTime'] = pd.to_datetime(df['pcrStopTime'], errors='coerce')

# Clean up `pcrStopTime` entries
df['pcrStopTime'] = df['pcrStopTime'].replace(r'\\N', pd.NaT, regex=True)
df['pcrStartTime'] = df['pcrStartTime'].replace(r'\\N', pd.NaT, regex=True)

# Strip and clean `pcrArea` column
df['pcrArea'] = df['pcrArea'].str.strip()
df['pcrArea'] = df['pcrArea'].str.replace(r'\s+', ' ', regex=True)
df['pcrArea'] = df['pcrArea'].str.title()

# Recalculate `pcrStopTime` for all rows based on `pcrStartTime` and `pcrMinutesUsed`
df['pcrStopTime'] = df.apply(
    lambda row: row['pcrStartTime'] + pd.Timedelta(minutes=row['pcrMinutesUsed']) 
    if not pd.isna(row['pcrStartTime']) else pd.NaT, 
    axis=1
)

# Add a column for hours used (calculated from minutes used)
df['pcrHoursUsed'] = (df['pcrMinutesUsed'] / 60).round(2)

# Ensure `pcrUserData1` is treated as a string
df['pcrUserData1'] = df['pcrUserData1'].astype(str)

# Sort the DataFrame by `pcrArea` and `pcrPC`
df_sorted = df.sort_values(by=['pcrArea', 'pcrPC'], ascending=[True, True])

# Format `pcrStartTime` and `pcrStopTime` back to string for exporting
df_sorted['pcrStartTime'] = df_sorted['pcrStartTime'].dt.strftime('%Y-%m-%d %H:%M:%S')
df_sorted['pcrStopTime'] = df_sorted['pcrStopTime'].dt.strftime('%Y-%m-%d %H:%M:%S')

# Save the sorted DataFrame to a new Excel file in the current directory
output_file = os.path.join(os.getcwd(), 'modified_data_with_correct_stoptime.xlsx')
df_sorted.to_excel(output_file, index=False)

print(f"New Excel file saved to: {output_file}")
