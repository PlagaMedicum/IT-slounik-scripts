import pandas as pd

file_path = 'input.xlsm'
sheet_name = 'Зводны слоўнік'

df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

# Step 1: Process Rows with "Eng Term Tildas"
def process_eng_term_tildas(df):
    for i in range(len(df)):
        if pd.notna(df.at[i, 'Eng Term Tildas']):
            if pd.isna(df.at[i, 'Eng Term']):
                for j in range(i - 1, -1, -1):  # Find the first previous row with non-empty "Eng Term"
                    if pd.notna(df.at[j, 'Eng Term']):
                        df.at[i, 'Eng Term'] = df.at[j, 'Eng Term']
                        break
    return df

# Step 2: Merge Rows with Similar English Terms
def merge_similar_english_terms(df):
    required_columns = ['Eng Term', 'Eng Part of speech']

    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"Column '{col}' not found in DataFrame")

    # Group rows by 'Eng Term' and 'Eng Part of speech'
    grouped = df.groupby(required_columns, as_index=False)

    merged_rows_list = []
    seen_indexes = set()  # To track rows that have been processed and merged

    for name, group in grouped:
        if len(group) > 1:  # Only merge groups with more than one row
            merged_row = group.iloc[0].copy()
            for col in df.columns:
                if col not in required_columns:
                    merged_row[col] = '\n'.join(group[col].dropna().unique())
            merged_rows_list.append(merged_row)
            seen_indexes.update(group.index)  # Mark indexes as processed
        else:
            merged_rows_list.append(group.iloc[0])
            seen_indexes.update(group.index)

    merged_rows = pd.DataFrame(merged_rows_list)

    # Combine the processed merged rows with the original DataFrame rows that not in seen_indexes
    final_df = pd.concat(
            [df[~df.index.isin(seen_indexes)], merged_rows],
            ignore_index=True
        ).drop_duplicates().sort_values(
            by=['Eng Term', 'Eng Part of speech', 'Eng Term Tildas']
        ).reset_index(drop=True)

    return final_df

df = process_eng_term_tildas(df)
df = merge_similar_english_terms(df)

output_path = 'output.xlsm'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Processing complete. The file has been saved as 'output.xlsm'.")
