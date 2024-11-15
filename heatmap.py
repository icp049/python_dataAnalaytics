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



#heatmap general
usage_by_area_and_hour = df.groupby(['pcrArea', 'hour_of_day']).size().unstack(fill_value=0)
plt.figure(figsize=(12, 8))
sns.heatmap(usage_by_area_and_hour, cmap='YlGnBu', annot=True, fmt='d', linewidths=.5)
plt.title('Heatmap of Usage by pcrArea and Hour of Day')
plt.xlabel('Hour of Day')
plt.ylabel('pcrArea')
plt.tight_layout()
plt.savefig(os.path.join(save_dir, f'heatmap_usage_all.png'))
plt.close()

#heatmap per area
for area in df['pcrArea'].unique():
    plt.figure(figsize=(12, 8))
    # Filter data for the specific area
    area_data = df[df['pcrArea'] == area]
    # Extract the hour from pcrStartTime to analyze usage at each hour
    area_data['hour_of_day'] = pd.to_datetime(area_data['pcrStartTime']).dt.hour
    # Create a pivot table to count the occurrences of pcrPC usage per hour
    usage_by_pc_and_hour = area_data.pivot_table(index='hour_of_day', columns='pcrPC', aggfunc='size', fill_value=0)
    
    # Plot heatmap showing usage for each pcrPC in this specific pcrArea
    sns.heatmap(usage_by_pc_and_hour, cmap='YlGnBu', annot=True, fmt='d', linewidths=.5)
    plt.title(f'Heatmap of Usage for {area} (pcrPC vs. Hour of Day)')
    plt.xlabel('pcrPC')
    plt.ylabel('Hour of Day')
    plt.tight_layout()
    plt.savefig(os.path.join(save_dir, f'heatmap_usage_{area}.png'))
    plt.close() # Show each heatmap after it's created
