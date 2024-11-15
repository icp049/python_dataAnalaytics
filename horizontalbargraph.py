import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
# Load the Excel file into a pandas DataFrame
df = pd.read_excel('data.xlsx')

# Convert 'pcrStartTime' and 'pcrStopTime' to datetime objects (ensuring proper conversion)
df['pcrStartTime'] = pd.to_datetime(df['pcrStartTime'], errors='coerce')
df['pcrStopTime'] = pd.to_datetime(df['pcrStopTime'], errors='coerce')

# Ensure that '\N' values are replaced with NaT (Not a Time) in case there are any such values
df['pcrStopTime'] = df['pcrStopTime'].replace(r'\\N', pd.NaT, regex=True)
df['pcrStartTime'] = df['pcrStartTime'].replace(r'\\N', pd.NaT, regex=True)



#clean the pcrArea for any formats so that fields will not double 
df['pcrArea'] = df['pcrArea'].str.strip()
df['pcrArea'] = df['pcrArea'].str.replace(r'\s+', ' ', regex=True)
df['pcrArea'] = df['pcrArea'].str.title()


# Convert 'pcrMinutesUsed' from minutes to hours (round to 2 decimal places)
df['pcrHoursUsed'] = (df['pcrMinutesUsed'] / 60).round(2) 

# Ensure that 'pcrUserData1' is treated as a string (so it doesn't get converted to scientific notation)
df['pcrUserData1'] = df['pcrUserData1'].astype(str)

# Sort the DataFrame by 'pcrArea' to group rows by location (branch)
df_sorted = df.sort_values(by=['pcrArea', 'pcrPC'], ascending=[True, True])

# Explicitly format the 'pcrStartTime' and 'pcrStopTime' columns to ensure they are saved correctly
df_sorted['pcrStartTime'] = df_sorted['pcrStartTime'].dt.strftime('%Y-%m-%d %H:%M:%S')
df_sorted['pcrStopTime'] = df_sorted['pcrStopTime'].dt.strftime('%Y-%m-%d %H:%M:%S')

# Specify the path to save the file in the current working directory
output_file = os.path.join(os.getcwd(), 'modified_data_sorted_by_area.xlsx')

# Save the new DataFrame to a new Excel file in the current directory
df_sorted.to_excel(output_file, index=False)

print(f"New Excel file saved to: {output_file}")

#create the graphs folder where graphs would be saved
save_dir = os.path.join(os.getcwd(), 'graphs')
os.makedirs(save_dir, exist_ok=True)


#bar graph horizontal
plt.figure(figsize=(10, 6))
area_hours = df.groupby('pcrArea')['pcrHoursUsed'].sum().sort_values()
area_hours.plot(kind='barh', color='skyblue')
plt.title('Total Hours Used per pcrArea')
plt.xlabel('Total Hours Used')
plt.ylabel('pcrArea')
plt.tight_layout()
plt.savefig(os.path.join(save_dir, 'total_hours_used_per_area.png'))
plt.close()  # Close the plot after saving


