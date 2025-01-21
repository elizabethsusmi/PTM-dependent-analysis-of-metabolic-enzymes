import pandas as pd

file_path = r"input_filepath"
df = pd.read_excel(file_path, sheet_name="Sheet1")

# Process each 'Type' (modification) separately
df_list = []

for type_name, type_df in df.groupby('Type'):
    # Calculate the total count for each gene within each 'Type'
    type_df['total_count'] = type_df.groupby('Gene symbol')['reference_count'].transform('sum')
    
    # Sort by 'reference_count' within each gene in descending order
    type_df = type_df.sort_values(by=['Gene symbol', 'reference_count'], ascending=[True, False])

    # Calculate cumulative frequency within each 'Gene symbol'
    type_df['cumulative_frequency'] = type_df.groupby('Gene symbol')['reference_count'].cumsum()

    # Calculate cumulative percentage
    type_df['cumulative_percentage'] = (type_df['cumulative_frequency'] / type_df['total_count']) * 100

    # Append processed data for each 'Type' to the list
    df_list.append(type_df)

# Concatenate all processed data back into a single DataFrame
final_df = pd.concat(df_list)

output_path = 'output_filepath'
final_df.to_excel(output_path, index=False)
