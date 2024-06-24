# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 13:02:24 2024

@author: aman_verma
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load the Excel file
file_path = R"C:\Users\dell\Documents\data\M&E Insight Analyst Assignment.xlsx"
excel_data = pd.ExcelFile(file_path)

# Display sheet names to understand the structure of the file
sheet_names = "M&E Insight Analyst Assignment"
print("Sheet names:", sheet_names)

# Load the 'Dataset' sheet to examine its content
data_df = pd.read_excel(file_path, sheet_name='Dataset')

# Clean the dataset by setting appropriate headers and removing any non-data rows
data_df.columns = data_df.iloc[0]
data_df = data_df[1:]

# Reset index for better handling
data_df.reset_index(drop=True, inplace=True)

# Rename the columns for better understanding
data_df.columns = ['District Name', 'School Name', 'Total Students Registered', 'M1 Active', 'M2 Active', 
                   'M3 Active', 'M4 Active', 'M5 Active', 'M6 Active', 'M7 Active', 'M8 Active', 'M9 Active', 
                   'M10 Active', 'M11 Active', 'M12 Active', 'M1 Avg Learning Time', 'M2 Avg Learning Time', 
                   'M3 Avg Learning Time', 'M4 Avg Learning Time', 'M5 Avg Learning Time', 'M6 Avg Learning Time', 
                   'M7 Avg Learning Time', 'M8 Avg Learning Time', 'M9 Avg Learning Time', 'M10 Avg Learning Time', 
                   'M11 Avg Learning Time', 'M12 Avg Learning Time']

# Convert relevant columns to numeric values
for col in data_df.columns[2:]:
    data_df[col] = pd.to_numeric(data_df[col], errors='coerce')

# Display cleaned dataframe
print(data_df.head())

# Descriptive analysis
total_students_registered = data_df['Total Number of Students Registered'].sum()
print("Total number of students registered on Khan Academy:", total_students_registered)

# Calculate the number of active students for each month
active_students_monthly = data_df[['M1', 'M2', 'M3', 'M4', 'M5', 
                                   'M6', 'M7', 'M8', 'M9', 'M10', 
                                   'M11', 'M12']].sum()
print("Number of active students each month:")
print(active_students_monthly)

# Calculate the average learning time per student for each month
avg_learning_time_monthly = data_df[['M1', 'M2', 'M3', 
                                     'M4', 'M5', 'M6', 
                                     'M7', 'M8', 'M9', 
                                     'M10', 'M11', 'M12']].mean()
print("Average learning time per student each month:")
print(avg_learning_time_monthly)

# Visualization
# Plotting the number of active students month-wise
plt.figure(figsize=(12, 6))
sns.barplot(x=active_students_monthly.index, y=active_students_monthly.values)
plt.title('Number of Active Students per Month')
plt.xlabel('Month')
plt.ylabel('Number of Active Students')
plt.xticks(rotation=45)
plt.show()

# Plotting the average learning time per student month-wise
plt.figure(figsize=(12, 6))
sns.lineplot(x=avg_learning_time_monthly.index, y=avg_learning_time_monthly.values, marker='o')
plt.title('Average Learning Time per Student per Month')
plt.xlabel('Month')
plt.ylabel('Average Learning Time (minutes)')
plt.xticks(rotation=45)
plt.show()
