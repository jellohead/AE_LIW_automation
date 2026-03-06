# slide_56.py
# Demographics Education and Anyone Under 18 Living in Home

import typing
import logging
from unittest.mock import inplace

from pandas import DataFrame, Series
import pandas as pd
from pptx.util import Inches
from ae_liw_automation.helper_modules import get_table_object, combine_multiple_questions, \
    get_table_shape_by_name, build_pptx_table
from ae_liw_automation.config import REPORTING_PERIOD, REPORTING_YEAR

logger = logging.getLogger(__name__)


# TODO: recode no/none/nothing responses to group them
# TODO: add logic to handle if table requires sorting by values


def slide_56_updater(meta, df, df_labeled, prs):
    slide_index = 54

    msg = f"Updating slide {slide_index + 1}"
    width = 40
    print(f"\n{'=' * width}\n{' ' + msg + ' ':=^{width}}\n{'=' * width}\n")
    logger.info(msg)

    slide = prs.slides[slide_index]

    question_dict = {
                    'Table 1':'D6',
                     'Table 5': 'D7'}

    education_labels_list = [
        'Some high school',
        'Graduated high school',
        'Some college',
        'Graduated college',
        'Post-graduate work',
        'Do not know/unsure',
        'Refused',
        'Base:',
    ]

    upper_table = list(question_dict.keys())[0]

    new_quarter_label = f'{REPORTING_PERIOD} {REPORTING_YEAR}'
    label_sub_dict = {'Prefer not to respond': 'Refused',
                      'DK/unsure': 'Do not know/unsure'}
    last_rows: list[str] = ['Do not know/unsure','Other', 'Refused', 'Base:']
    # table_names: list[str] = ['Table 1', 'Table 2']

    for table_name, question in question_dict.items():
        table_shape = get_table_shape_by_name(slide, table_name)
        if not table_shape:
            print(f'No table found on {slide_index + 1}')
            logger.info(f'No table found on {slide_index + 1}')
            return

        old_table_name = table_shape.name
        logger.info(f'Updating "{old_table_name}"')
        table = table_shape.table

        # get a reference to the table shape for the old table
        table_shape_element = table_shape._element

        # Get existing data from old table
        table_data = [[cell.text.strip() for cell in row.cells] for row in table.rows]
        table_df: DataFrame = pd.DataFrame(table_data[1:], columns=table_data[0])
        table_df.set_index(table_df.columns[0], inplace=True)

        # drop oldest quarter data
        table_df_existing = table_df.drop(columns=[table_df.columns[0]])

        index_name = table_df_existing.index.name

        # get the existing base row and then drop it from the working dataframe
        base_row_df = table_df_existing.loc[['Base:']]
        base_row_df[new_quarter_label] = len(df)

        # get current quarter data from dataset
        current_quarter_result_series = df_labeled[question].value_counts(normalize=True).map("{:.0%}".format)
        current_quarter_result_df = pd.DataFrame({new_quarter_label: current_quarter_result_series})
        current_quarter_result_df.rename(index=label_sub_dict, inplace=True)
        current_quarter_result_df.loc['Base:'] = len(df)

        final_result_df_combined = pd.concat(
            [table_df_existing, current_quarter_result_df],
            axis=1
        ).fillna(f'0%')

        # reindex the dataframe to get in education_labels_list order
        # final_result_df_combined = final_result_df_combined.reindex(education_labels_list)

        # remove rows where all values are 0% (excluding Base: row)
        non_base_mask = final_result_df_combined.index != 'Base:'
        all_zero_mask = (final_result_df_combined == '0%').all(axis=1)
        # this is the final form of the dataframe that will update table
        final_result_df_combined = final_result_df_combined[~(non_base_mask & all_zero_mask)]

        final_result_df_combined.rename_axis(index_name, inplace=True)

        print(final_result_df_combined)

        base_row = final_result_df_combined[final_result_df_combined.index == 'Base:']

        # provide a vertical offset so tables do not overlap in case of two tables being updated
        top_offset = Inches(1.7) if table_name == upper_table else Inches(4.5)

        build_pptx_table(slide, table_shape_element, final_result_df_combined,
                         base_row, index_name, old_table_name, top_offset)

    logger.info(
        f'Update of slide {slide_index + 1} complete.\nManually adjust position and size of the table.')
