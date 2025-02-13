# slide_8.py
# update the data table on slide 8
from pandas import DataFrame
import pandas as pd
from helper_modules import get_table_object
from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR


def slide_8_updater(df, prs):
    print('slide_8_updater')
    slide_index = 7
    slide = prs.slides[slide_index]

    table = get_table_object(slide)
    if not table:
        print('No table found on Slide 8')
        return
    # print(table)

    # extract table dimensions
    num_rows = len(table.rows)
    num_cols = len(table.columns)

    # Initialize variable to store table data
    table_data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
    table_df: DataFrame = pd.DataFrame(table_data[1:], columns=table_data[0])
    table_df.set_index(table_df.columns[0], inplace=True)

    # drop oldest quarter data
    # table_df_current = table_df.drop(table_df.columns[0], axis=1)
    table_df_current = table_df.drop(columns=[table_df.columns[0]])

    q19_result = df['Q19'].dropna().value_counts()
    q19_df =pd.DataFrame({f'{REPORTING_PERIOD} {REPORTING_YEAR}': q19_result}).fillna(0)
    # d = {f'{REPORTING_PERIOD} {REPORTING_YEAR}': q19_result}

    q19_df_combined = pd.concat([table_df, q19_df], axis=1).fillna(0)

    updated_rows, updated_cols = q19_df_combined.shape

    # Resize table if necessary
    while len(table.rows) < updated_rows + 1:  # +1 to account for header
        table.add_row()
    while len(table.columns) < updated_cols + 1:  # +1 to account for index
        table.add_column()

    # Update table headers
    for col_idx, col_name in enumerate(q19_df_combined.columns, start=1):  # Start at 1 to skip index column
        table.cell(0, col_idx).text = col_name

    # Update table contents
    for row_idx, (index_name, row_values) in enumerate(q19_df_combined.iterrows(), start=1):  # Skip header row
        table.cell(row_idx, 0).text = str(index_name)  # Update index column
        for col_idx, value in enumerate(row_values, start=1):
            table.cell(row_idx, col_idx).text = str(int(value))  # Convert float to int for clean formatting

    print("Slide 8 table updated successfully.")

    # # Extract data row by row
    # for row in table.rows:
    #     row_data = [cell.text.strip() for cell in row.cells]  # Get row values
    #     table_data.append(row_data)  # Append row to list
    # # print(f'{type(table_data) = }')
    #
    # # Convert table data to a DataFrame
    # # print(f'{table_df = }')
    #
    # # print(f'{table_df_current = }')
    #
    #
    # # print(f'{q19_result = }')
    # # print(q19_result)
    #
    # q19_df = pd.DataFrame(data=d)
    # # print(f'{q19_df['Q4 2024'] = }')
    # q19_df = q19_df.drop(q19_df.index[0])
    #
    # # print(f'{q19_df = }')
    #
    # # table_df_current_joined = table_df.join(table_df_current)
    # # print(f'{table_df_current_joined = }')
    #
    # q19_df_combined = pd.concat([table_df_current, q19_df], axis=1).fillna(0)
    # print(f'{q19_df_combined = }')
    # # # Display DataFrame
    # # import ace_tools as tools
    # # tools.display_dataframe_to_user(name="Extracted Table", dataframe=df)
    #
    # # pull data from existing table
    #
    # # drop oldest quarter of data
    #
    # # update table with new quarter data
    #
    # # replace table data with updated
    #
