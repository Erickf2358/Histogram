import pandas as pd
import matplotlib.pyplot as plt
from datetime import timedelta
from matplotlib.ticker import MaxNLocator
from matplotlib.ticker import FuncFormatter
import tkinter as tk
from tkinter import filedialog
import numpy as np

# Ask for file upload
root = tk.Tk()
root.withdraw()  # Hide the main window
file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

if not file_path:
    print("No file selected. Exiting...")
    exit()

# Full import statement
sheet_mapping = 'actlist'
df_CP_2 = pd.read_excel(file_path, sheet_mapping)

df_CP_2['EV to go'] = df_CP_2['Cost_Budget']
df_CP_2['Project Group'] = df_CP_2['UC8']
df_CP_2['Start Date'] = df_CP_2['Start']
df_CP_2['End Date'] = df_CP_2['End']

df_CP_2['EV to go'].sum()/1000000

# Select specific tables of the dataframe
df_CP_EV_togo = df_CP_2 [['Project Group','EV to go','Start Date','End Date']]

# Set the start of the upcoming month
df_CP_EV_togo['Today'] = '01/01/2024'

# Remove rows with EV to_go equals to cero
df_CP_EV_togo = df_CP_EV_togo[df_CP_EV_togo['EV to go'] != 0]

# Ensure the date columns are date type
df_CP_EV_togo['Start Date'] = pd.to_datetime(df_CP_EV_togo['Start Date'] , errors='coerce')
df_CP_EV_togo['End Date'] = pd.to_datetime(df_CP_EV_togo['End Date'] , errors='coerce')
df_CP_EV_togo['Today'] = pd.to_datetime(df_CP_EV_togo['Today'] , errors='coerce')
df_CP_EV_togo

# Ensure that the start date is greater or equals than the start of next month
df_CP_EV_togo['Start Date'] = df_CP_EV_togo.apply(lambda row: row['Start Date'] if row['Start Date'] > row['Today'] else row['Today'], axis=1)

# Ensure that the end date is greater or equals than the start of next month
df_CP_EV_togo['End Date'] = df_CP_EV_togo.apply(lambda row: row['End Date'] if row['End Date'] > row['Today'] else row['Today'], axis=1)

# Check the value of the column EV to_go
df_CP_EV_togo['EV to go'].sum()/1000000

df, cost_column, start_column, end_column = df_CP_EV_togo, 'EV to go', 'Start Date','End Date'

df[end_column] = pd.to_datetime(df[end_column], errors='coerce')
df[start_column] = pd.to_datetime(df[start_column], errors='coerce')
df['Duration'] = (df[end_column] - df[start_column]).dt.days + 1

# Calculate cost per day
df['Cost per day'] = (df[cost_column] / df['Duration']).astype(float).round(2)

# Define the step as a timedelta of 1 day
step = timedelta(days=1)

# Add a new column with a list of dates
df['Date_Range'] = df.apply(
        lambda row: pd.date_range(start=row[start_column], periods=row['Duration'], freq=step).tolist(), axis=1)

# Explode the 'Date_Range' column into separate rows
df_exploded = df.explode('Date_Range')

# Add a new column 'Start_Month' with the start of the month for each date in 'Date_Range'
df_exploded['Start_Month'] = df_exploded['Date_Range'].dt.to_period('M').dt.to_timestamp()

# Group by 'Start_Month' and calculate the sum of 'Cost per day'
df_group = df_exploded.groupby(['Start_Month','Project Group']).agg({'Cost per day': 'sum'}).reset_index()

# Rename 'Cost per day' to 'Cost'
df_group.rename(columns={'Cost per day': 'Cost'}, inplace=True)

df_pivot = df_group.pivot(index='Start_Month', columns='Project Group', values='Cost').reset_index()

df_pivot

df1 = df_pivot

# Convert 'Start_Month' to datetime
df1['Start_Month'] = pd.to_datetime(df1['Start_Month']).dt.strftime('%Y-%m')

# Get dynamic order from columns (exclude 'Start_Month')
order = [col for col in df1.columns if col != 'Start_Month']
print(f"Dynamic order: {order}")

# Function to format y-axis
def millions(x, pos):
    'The two args are the value and tick position'
    return '%1.1fM' % (x * 1e-6)

formatter = FuncFormatter(millions)

# Define base colors and create dynamic color list
base_colors = ['gray', 'orangered', 'orange', 'green', 'purple', 'deepskyblue', 'pink', 'limegreen', 'hotpink', 'orange']

# If we have more categories than base colors, extend the color list
if len(order) > len(base_colors):
    # Generate additional colors using a colormap
    additional_colors = plt.cm.tab20(np.linspace(0, 1, len(order) - len(base_colors)))
    additional_colors = [plt.colors.to_hex(color) for color in additional_colors]
    colors = base_colors + additional_colors
else:
    # Use only the needed base colors
    colors = base_colors[:len(order)]

print(f"Colors used: {colors}")

# Create subplots
fig, (ax1) = plt.subplots(1, 1, figsize=(10, 7), sharey=True)

# Plot the stacked bar chart
df1.set_index('Start_Month')[order].plot(kind='bar', stacked=True, ax=ax1, color=colors)
ax1.set_title('Reference dates from: Disciplines group')
ax1.set_xlabel('Start Month')
ax1.set_ylabel('Values')
ax1.grid(True)  # Add grid to the first subplot

# Set the maximum number of x-ticks
ax1.xaxis.set_major_locator(MaxNLocator(nbins=16))  

# Adjust layout
plt.gca().yaxis.set_major_formatter(formatter)

plt.tight_layout()
plt.show()