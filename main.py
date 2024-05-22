import pandas as pd

df = pd.read_excel('firstXL.xlsx', sheet_name='Table_1')
print(df)

# Extraction of location and ids
locations = df.columns[2:]
location_ids = df.iloc[0, 2:]

# Initialize lists 
rec_locs = []
rec_loc_ids = []
del_locs = []
del_loc_ids = []
distances = []

# Iterate over each row in the df
for index, row in df.iterrows():
    if index == 0:  
        continue
    rec_loc = row[0] 
    rec_loc_id = row[1]  

    # Iterate over each column starting from the third column
    for i, val in enumerate(row[2:]):
        if pd.notna(val):  # Skip empty cells
            rec_locs.append(rec_loc)
            rec_loc_ids.append(rec_loc_id)
            del_locs.append(locations[i])
            del_loc_ids.append(location_ids[i])
            distances.append(f"{val:.5f}") 

# Create a DataFrame from the transformed data
transformed_df = pd.DataFrame({
    'Rec LocID':rec_locs ,
    'Receipt Map Meters': rec_loc_ids,
    'Del LocID': del_loc_ids,
    'Delivery Map Meters': del_locs,
    'Distance': distances
})

# checking that no ro is dropped
expected_rows = 529
if len(transformed_df) != expected_rows:
    print(f"Warning: Expected {expected_rows} rows, but got {len(transformed_df)} rows")

print(transformed_df)

# Save the transformed DataFrame to a new Excel file
transformed_df.to_excel('secondXL.xlsx', index=False)
