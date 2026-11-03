import os
import time
import pandas as pd
import config


def read_source_file():
    source_path = config.Source_file_path

    if not os.path.exists(source_path):
        print(f"Error: Source file not found at path: {source_path}")
        exit()

    df = pd.read_excel(source_path)

    return df, source_path


def remove_exact_duplicates(df):
    before_rows = len(df)

    df = df.drop_duplicates()

    after_rows = len(df)
    deleted_rows = before_rows - after_rows

    if deleted_rows > 0:
        print(f"{deleted_rows} duplicated rows are deleted...")

    return df


def remove_duplicate_numbers(df, mapping_df):
    # Access source unique identifier (Number) from mapping file
    number_row = mapping_df[mapping_df["Final_Column_Name"] == "Incident Number"]

    # number_source_col = number_row["Source_Column_Name"]
    number_source_col = number_row["Source_Column_Name"].iloc[0]

    if number_source_col not in df.columns:
        print(f"ERROR: Source column '{number_source_col}' not found in source data.")
        exit()

    dup_mask = df[number_source_col].duplicated()
    df.loc[dup_mask, number_source_col] = pd.NA

    return df, number_source_col


def apply_filtering(df, number_source_col):
    before_rows = len(df)

    df = df[df[number_source_col].notna()]

    after_rows = len(df)
    duplicated_numbers = before_rows - after_rows

    if duplicated_numbers > 0:
        print(f"{duplicated_numbers} duplicate numbers are found in your data and excluded them...")

    return df


def transform_columns(df, mapping_df):
    # Build output dataframe
    output_df = pd.DataFrame()  # Creating empty dataframe
    used_source_cols = []

    for index, row in mapping_df.iterrows():
        source_col = row["Source_Column_Name"]
        final_col = row["Final_Column_Name"]

        # Main Logic
        # Copy if source column exists, else create empty column
        if pd.notna(source_col) and source_col in df:
            output_df[final_col] = df[source_col]
            used_source_cols.append(source_col)
        else:
            output_df[final_col] = pd.NA

    # Adding remaining source columns
    for col in df.columns:
        if col not in used_source_cols:
            output_df[col] = df[col]

    # Date handling (type conversions)
    date_columns = ["Reported Date", "Resolved Date"]
    for col in date_columns:
        if col in output_df.columns:
            output_df[col] = pd.to_datetime(output_df[col], errors="coerce")

    return output_df


def write_output(df, source_path):
    source_dir = os.path.dirname(config.Source_file_path)
    source_file_name = os.path.basename(config.Source_file_path)
    source_name_without_ext = os.path.splitext(source_file_name)[0]

    output_file_name = f"{source_name_without_ext}-transformed.xlsx"
    output_file_path = os.path.join(source_dir, output_file_name)

    df.to_excel(output_file_path, index=False)
    print("Success: file transformed successfully.")
    print(f"Output path: {output_file_path}")


def main():
    start_time = time.time()

    # 1) Read mapping file
    if not os.path.exists(config.Mapping_file):
        print("Error: Mapping file not found.")
        exit()

    mapping_df = pd.read_excel(config.Mapping_file)

    # Read source file
    source_df, source_path = read_source_file()

    # Remove exact duplicate rows
    source_df = remove_exact_duplicates(source_df)

    # Remove duplicate numbers
    source_df, number_source_col = remove_duplicate_numbers(source_df, mapping_df)

    # Apply filtering
    source_df = apply_filtering(source_df, number_source_col)

    # Transform columns
    output_df = transform_columns(source_df, mapping_df)

    # Write output to excel
    write_output(output_df, source_path)

    end_time = time.time()
    execution_time = end_time - start_time

    print(f"execution completed in {execution_time:.2f} seconds")


if __name__ == "__main__":
    main()
