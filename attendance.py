import pandas as pd
from datetime import date, datetime
import calendar

# Get user input for the month
month = input("Enter the month (e.g., 01 for January): ")

# Importing the biometric file
file = open("ALOG_001.txt", encoding="utf-16-le")
# Empty Data Array creation
data_arr = []
# Column Array for Dataframe creation
column_arr = file.readline().split("\t")
# Loop to read the data from the file
while True:
    data = file.readline()
    if data == "":
        break
    arr = data.split("\t")
    data_arr.append(arr)
data_df = pd.DataFrame(data_arr,columns = column_arr)

# Remove '\n' from the 'DateTime\n' column
data_df['DateTime\n'] = data_df['DateTime\n'].str.replace('\n', '')

# Split the 'DateTime\n' column into 'Date' and 'Time' columns.
data_df[['Date', 'Time']] = data_df['DateTime\n'].str.split(' ', expand=True)

# Convert 'Date' column to datetime objects
data_df['Date'] = pd.to_datetime(data_df['Date'])

# Filter data for the specified month
data_df = data_df[data_df['Date'].dt.strftime('%m') == month]

# Group the dataframe by 'EnNo' and 'Date' and get the minimum and maximum time for each group.
time_range = data_df.groupby(['EnNo', 'Date'])['Time'].agg(['min', 'max']).reset_index()

# Rename the columns to match the desired output.
time_range = time_range.rename(columns={'min': 'In Time', 'max': 'Out Time'})

# Merge the time_range dataframe with the original dataframe to get the 'Name' column.
new_df = time_range.merge(data_df[['EnNo', 'Date', 'Name']], on=['EnNo', 'Date'], how='left')

# Select the desired columns.
new_df = new_df[['Name', 'EnNo', 'Date', 'In Time', 'Out Time']]

# Remove duplicate values.
new_df = new_df.drop_duplicates()

# Convert 'In Time' and 'Out Time' columns to datetime objects
new_df['In_Time'] = pd.to_datetime(new_df['In Time'], format='%H:%M:%S')
new_df['Out_Time'] = pd.to_datetime(new_df['Out Time'], format='%H:%M:%S')

# Create 'Late Mark' column
new_df['Late Mark'] = new_df['In_Time'].apply(lambda x: x.time() > pd.to_datetime('10:15:59').time())

# Calculate 'Working Hours' column
new_df['Working Hours'] = (new_df['Out_Time'] - new_df['In_Time']).dt.total_seconds() / 3600

# Get the current date.
today = date.today()
date_string = today.strftime("%Y-%m-%d")

# Export to Excel.
new_df.to_excel(f"Attendance-{date_string}.xlsx", index=False)

# Create a summary dataframe
summary_df = new_df.groupby(['Name'])['Late Mark'].agg(['count', 'sum']).reset_index()

# Rename columns
summary_df = summary_df.rename(columns={'count': 'Total Records', 'sum': 'Number of Late Marks'})

# Calculate total working hours
total_working_hours = new_df.groupby(['Name'])['Working Hours'].sum().reset_index()

# Merge with summary_df
summary_df = summary_df.merge(total_working_hours, on='Name')

# Rename column
summary_df = summary_df.rename(columns={'Working Hours': 'Total Working Hours'})

# Round 'Total Working Hours' to 2 decimal places
summary_df['Total Working Hours'] = summary_df['Total Working Hours'].round(2)

# Display the summary dataframe
print(summary_df)

