import pandas as pd
import pyreadstat
import pptx
from config.constants import DATASET_FILE_PATH, PPTX_INPUT_FILE, PPTX_OUTPUT_FILE, EXCEL_FILE
from helper_modules.read_data import read_data
from helper_modules.get_chart_object_by_name import get_chart_object_by_name
from slide_updaters import slide_1_updater, slide_3_updater, slide_4_updater


def main():
    df, meta, df_labeled = read_data(DATASET_FILE_PATH)
    prs = pptx.Presentation(PPTX_INPUT_FILE)
    slide_1_updater(df, prs)
    slide_3_updater(df, prs)
    slide_4_updater(df, prs)

    prs.save(PPTX_OUTPUT_FILE)
    # df_labeled.to_excel(EXCEL_FILE)



if __name__ == '__main__':
    main()