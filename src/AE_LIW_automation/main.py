import logging

import pandas as pd
import pyreadstat
import pptx
from config.constants import DATASET_FILE_PATH, PPTX_INPUT_FILE, PPTX_OUTPUT_FILE, EXCEL_FILE
from config.logging_config import setup_logging
from helper_modules.read_data import read_data
from helper_modules.get_chart_object_by_name import get_chart_object_by_name
from slide_updaters import (slide_1_updater, slide_3_updater, slide_4_updater, slide_6_updater, slide_7_updater,
                            slide_8_updater, slide_9_updater, slide_10_updater, slide_11_updater, slide_12_updater,
                            slide_13_updater,
                            slide_14_updater,
                            slide_15_updater, slide_16_updater,
                            slide_17_updater, slide_18_updater, slide_19_updater, slide_21_updater, slide_22_updater,
                            slide_23_updater,
                            slide_24_updater,
                            slide_25_updater, slide_26_updater, slide_27_updater, slide_28_updater, slide_29_updater,
                            slide_30_updater, slide_31_updater, slide_32_updater, slide_33_updater, slide_35_updater,
                            slide_36_updater, slide_37_updater,
                            slide_38_updater, slide_39_updater, slide_41_updater,
                            slide_40_updater,
                            slide_43_updater, slide_44_updater, slide_45_updater, slide_46_updater, slide_48_updater,
                            slide_50_updater, slide_52_updater, slide_60_updater,
                            )

logger = logging.getLogger(__name__)


# TODO: slides with stacked bar charts showing T3B values and totals, like slide 48, need rounding functions added to display total value that matches sum of T3B ratings

def main():
    setup_logging()
    logger.info('Starting AE LIW Automation')
    print(f'In main')

    df, meta, df_labeled = read_data(DATASET_FILE_PATH)
    prs = pptx.Presentation(PPTX_INPUT_FILE)
    slide_1_updater(df, prs)
    slide_3_updater(df, prs)
    slide_4_updater(meta, df, df_labeled, prs)
    slide_6_updater(df, prs)
    slide_7_updater(df, prs)
    slide_8_updater(df, prs)
    slide_9_updater(df, prs)
    slide_10_updater(df, prs)
    slide_11_updater(df, prs)
    slide_12_updater(df, prs)
    slide_13_updater(df, prs)
    slide_14_updater(df, prs)
    slide_15_updater(df, prs)
    slide_16_updater(meta, df, df_labeled, prs)
    slide_17_updater(df, prs)
    slide_18_updater(df, prs)
    slide_19_updater(meta, df, df_labeled, prs)
    slide_21_updater(df, prs)
    slide_22_updater(df, prs)
    slide_23_updater(df, prs)
    slide_24_updater(df, prs)
    slide_25_updater(df, prs)
    slide_26_updater(meta, df, df_labeled, prs)
    slide_27_updater(df, prs)
    slide_28_updater(df, prs)
    slide_29_updater(df, prs)
    slide_30_updater(df, prs)
    slide_31_updater(df, prs)
    slide_32_updater(meta, df, df_labeled, prs)
    slide_33_updater(meta, df, df_labeled, prs)
    slide_35_updater(df, prs)
    slide_36_updater(df, prs)
    slide_37_updater(meta, df, df_labeled, prs)
    slide_38_updater(df, prs)
    slide_39_updater(meta, df, df_labeled, prs)
    slide_40_updater(df, prs)
    slide_41_updater(meta, df, df_labeled, prs)
    slide_43_updater(df, meta, df_labeled, prs)
    slide_44_updater(df, meta, df_labeled, prs)
    slide_45_updater(df, meta, df_labeled, prs)
    slide_46_updater(meta, df, df_labeled, prs)
    slide_48_updater(df, prs)
    slide_50_updater(meta, df, df_labeled, prs)
    slide_52_updater(meta, df, df_labeled, prs)
    # slide_60_updater(df, meta, df_labeled, prs)

    prs.save(PPTX_OUTPUT_FILE)
    # df_labeled.to_excel(EXCEL_FILE)


if __name__ == '__main__':
    main()
