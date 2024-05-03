### Run the TakeRate BAQ in Epicor to generate the TakeRateInput.xlsx ###

import pandas as pd

InputFile = 'TakeRateInput.xlsx'
OutputFile = 'TakeRateOutput.xlsx'

# Read data from an Excel file
df = pd.read_excel(InputFile)

# Filter out unapplied cash/ref lines
df = df[df['Part'] != '']

# Filter only options and whole goods
options_df = df[df['WGorOpt'] == 'Option']
whole_goods_df = df[df['WGorOpt'] == 'Whole Good']

# Get base whole goods by filtering for invoice line 1
base_whole_goods = whole_goods_df[['Part', 'Invoice']]

# Merge options with their associated base whole goods
merged_df = pd.merge(options_df, base_whole_goods, on='Invoice', suffixes=('_option', '_whole_good'))

# Convert 'Inv Date' to datetime if it's not already
merged_df['Inv Date'] = pd.to_datetime(merged_df['Inv Date'])

# Extract the year from 'Inv Date' and convert it to an integer
merged_df['Year'] = merged_df['Inv Date'].dt.year.astype(int)

# Group by year, whole good, option, and description, and calculate the count
grouped = merged_df.groupby(['Year', 'Part_whole_good', 'Part_option', 'Description']).size().reset_index(name='Option_Count')

# Calculate the total count of options for each whole good within a year
grouped['Total_Option_Count'] = grouped.groupby(['Year', 'Part_whole_good'])['Option_Count'].transform('sum')

# Calculate the percentage of each option based on the total count for its whole good within a year
grouped['Percentage_of_Total_Option_Count'] = ((grouped['Option_Count'] / grouped['Total_Option_Count']) * 100).round(2)

# Rank the options within each whole good group based on count
grouped['Rank'] = grouped.groupby(['Year', 'Part_whole_good'])['Percentage_of_Total_Option_Count'].rank(ascending=False, method='dense')

# Calculate the count of how many times a whole good was purchased within a year
whole_good_counts = merged_df.groupby(['Year', 'Part_whole_good'])['Invoice'].nunique().reset_index(name='Times_Whole_Good_sold')

# Merge the whole good counts with the grouped data
grouped = pd.merge(grouped, whole_good_counts, on=['Year', 'Part_whole_good'])

# Sort by 'Year', 'Part_whole_good', and 'Rank'
grouped = grouped.sort_values(by=['Year', 'Part_whole_good', 'Rank'])

# Add the count of whole goods sold 
grouped.loc[grouped.groupby(['Year', 'Part_whole_good']).head(1).index, 'Times_Whole_Good_sold'] = grouped['Times_Whole_Good_sold']

# Calculate Option taken percentage Option_count / Times_Whole_Good_Sold
grouped['Percentage_of_times_sold_with_Whole_Good'] = (grouped['Option_Count'] / grouped['Times_Whole_Good_sold'] * 100).round(2)

# Move the 'Times_Whole_Good_sold' column to appear right after the 'Part_whole_good' column
column_order = ['Year', 'Part_whole_good', 'Times_Whole_Good_sold', 'Part_option', 'Description', 'Option_Count', 'Total_Option_Count', 'Percentage_of_Total_Option_Count', 'Percentage_of_times_sold_with_Whole_Good']
grouped = grouped[column_order]

# Write to Excel workbook with each year's data in a separate sheet
with pd.ExcelWriter(OutputFile) as writer:
    for year in grouped['Year'].unique():
        year_df = grouped[grouped['Year'] == year]
        sheet_name = f'options_{year}'
        year_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Options ranked by whole good for {year} have been saved to sheet '{sheet_name}' in {OutputFile}")
