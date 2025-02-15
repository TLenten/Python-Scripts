import pandas as pd


input_file = "PartBOMBottomUp.xlsx"
part_char_10 = "X:\\DMT\\DMT Part\\Character10\\part_char10_import.csv"
ud24 = "UD24_import.csv"


#####################################
#Part Character10 field formatter
#####################################
# Read the input data from the file
#input_file = "PartBOMBottomUp.xlsx"

print("Reading Input File")
df = pd.read_excel(input_file)

print("Generating Part character10 DMT")

# Drop duplicate rows based on "Part Part Num", "Part1 Part Num", and "Part1 Prod Code"
df_unique = df.drop_duplicates(subset=["Part Part Num", "Part1 Part Num", "Part1 Prod Code"])

# Group by "Part Part Num" and "Part1 Prod Code" and concatenate "Part1 Prod Code" with count
grouped = df_unique.groupby(["Part Part Num", "Part1 Prod Code"])["Part1 Prod Code"].apply(lambda x: ', '.join([f"{prod_group}({count})" for prod_group, count in x.value_counts().items()])).reset_index(name="Character10")

# Further group by "PartNum" and concatenate "Character10" values within each group
grouped_partnum = grouped.groupby("Part Part Num")["Character10"].apply(lambda x: ', '.join(x)).reset_index()

# Create the new dataframe with the desired format
new_df = pd.DataFrame(columns=["Company", "PartNum", "Character10"])
new_df["Company"] = ["HSM"] * len(grouped_partnum)
new_df["PartNum"] = grouped_partnum["Part Part Num"]
new_df["Character10"] = grouped_partnum["Character10"]

# Output the new dataframe to a new Excel file or overwrite existing
new_df.to_csv(part_char_10, index=False)

print("Part Character10 Complete")


#####################################
#UD24 Formatter
#####################################
# Read the input data from the file #
#input_file = "PartBOMBottomUp.xlsx"
print("Generate UD24")

df = pd.read_excel(input_file)

# Group by "Part Part Num" and "Part1 Prod Code" and aggregate "Part1 Class ID"
grouped = df.groupby("Part Part Num")["Part1 Part Num"].apply(lambda x: '/'.join(sorted(set(x)))).reset_index()

# Create the new dataframe with the desired format
new_df = pd.DataFrame(columns=["Company", "Key1", "Key2", "Key3", "Key4", "Key5", "Character01"])
new_df["Company"] = ["HSM"] * len(grouped)
new_df["Company"] = new_df["Company"]
new_df["Key1"] = grouped["Part Part Num"]
new_df["Key2"] = ""
new_df["Key3"] = ""
new_df["Key4"] = ""
new_df["Key5"] = ""
new_df["Character01"] = grouped["Part1 Part Num"]

# Output the new dataframe to a new Excel file or overwrite existing
new_df.to_csv(ud24, index=False)

print("UD24 Complete")