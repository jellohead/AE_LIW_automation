# slide_8.py
# update the data table on slide 8
from pandas import DataFrame
import pandas as pd
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from helper_modules import get_table_object
from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR


def slide_8_updater(df, prs):
    print('\n\tslide_8_updater\n')
    slide_index = 7
    slide = prs.slides[slide_index]

    table = get_table_object(slide)
    if not table:
        print('No table found on Slide 8')
        return

    # extract table dimensions
    num_rows = len(table.rows)
    num_cols = len(table.columns)

    # Get existing data from old table
    table_data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
    table_df: DataFrame = pd.DataFrame(table_data[1:], columns=table_data[0])
    table_df.set_index(table_df.columns[0], inplace=True)

    # drop oldest quarter data
    table_df_current = table_df.drop(columns=[table_df.columns[0]])
    print(f'{table_df_current = }')

    # get current quarter data from dataset
    q19_result = df['Q19'].replace('', float('nan')).dropna().value_counts()
    # append Base value to the series
    q19_result.at['Base:'] = len(q19_result)
    q19_df =pd.DataFrame({f'{REPORTING_PERIOD} {REPORTING_YEAR}': q19_result}).fillna(0)

    # combine old and new data, convert new data to integers
    q19_df_combined = pd.concat([table_df_current, q19_df], axis=1).fillna(0).astype(int)

    # sort the combined df and put Base at the bottom
    base_row = q19_df_combined[q19_df_combined.index == 'Base:']
    not_base_rows = q19_df_combined[q19_df_combined.index != 'Base:'].sort_values(by=f'{REPORTING_PERIOD} {REPORTING_YEAR}', ascending=False)
    q19_df_combined = pd.concat([not_base_rows, base_row], axis=0).fillna(0).astype(int)    # q19_df_other_rows = q19_df_combined.loc[q19_df_combined.index != 'Base:'].sort_values(f'{REPORTING_PERIOD} {REPORTING_YEAR}', ascending=False)

    # Step 1: Remove existing table (if any)
    shapes = slide.shapes
    for shape in shapes:
        if shape.has_table:  # Check if shape is a table
            sp = shape._element  # Get the XML element of the shape
            slide.shapes._spTree.remove(sp)  # Remove the shape

    # works as expected up to this point

    # Step 3: Add a new table to the slide
    # rows, cols = q19_df_combined.shape
    # table_shape = slide.shapes.add_table(rows + 1, cols + 1, Inches(1), Inches(1.5), Inches(8), Inches(3)).table
    #
    #
    # # Step 4: Insert column headers
    # table_shape.cell(0, 0).text = 'Category'
    # for col_idx, col_name in enumerate(q19_df_combined.columns):
    #     table_shape.cell(0, col_idx + 1).text = col_name  # First row as header
    #
    # #Step 5: Insert data rows
    # for row_idx, (index_value, row) in enumerate(q19_df_combined.iterrows()):
    #     table_shape.cell(row_idx +1, 0).text = str(index_value)
    #     for col_idx, value in enumerate(row):
    #         table_shape.cell(row_idx + 1, col_idx + 1).text = str(value)

    # TODO: updated table is not including the index column of the dataframe
    # Step 3: Add a new table to the slide
    rows, cols = q19_df_combined.shape

    # Add one more column for the index
    table_shape = slide.shapes.add_table(rows + 1, cols + 1, Inches(1), Inches(1.5), Inches(8), Inches(3)).table

    # Step 4: Insert column headers (including index column)
    table_shape.cell(0, 0).text = "Category"  # Name for the index column
    for col_idx, col_name in enumerate(q19_df_combined.columns):
        table_shape.cell(0, col_idx + 1).text = col_name  # Shifted by +1

    # Step 5: Insert data rows (including index values)
    for row_idx, (index_value, row) in enumerate(q19_df_combined.iterrows()):
        table_shape.cell(row_idx + 1, 0).text = str(index_value)  # First column for index
        for col_idx, value in enumerate(row):
            table_shape.cell(row_idx + 1, col_idx + 1).text = str(value)  # Shifted by +1
