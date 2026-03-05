import pandas as pd
from typing import List, Dict
import logging


logger = logging.getLogger(__name__)


# TODO: add some logic to determine base cell value for various scenarios

def combine_multiple_questions(
    df: pd.DataFrame,
    question_list: List[str],
    label_sub_dict: Dict[str, str] = None,
    clean_strings: bool = True,
    include_base: bool = True,
    base_label: str = "Base:",
    base_calc_method: str = "population",

) -> pd.Series:
    """
    Summarizes multiple categorical survey questions by combining their value_counts,
    optionally applying label substitutions and returning a single Series of total counts.

    Parameters:
        df              : The source DataFrame
        question_list   : List of column names (questions) to include
        label_sub_dict  : Optional dict to rename or merge responses (e.g., {'All other': 'Other'})
        clean_strings   : If True, convert responses to str before processing
        include_base    : If True, adds a 'Base:' row with total count of non-blank responses
        base_label      : Label to use for the base count row
        base_calc_method: Sets how to generate base value (e.g., use entire population)

    Returns:
        pd.Series: Combined value counts with optional label substitutions and base count
    """
    series_list = []

    # use stack() to combine multiple questions
    series2 = df[question_list].stack().value_counts()
    if clean_strings:
        series2 = series2.astype(str)

    print(series2.unique())
    if label_sub_dict:
        series2.replace(label_sub_dict, inplace=True)

    for question in question_list:
        series = df[question].dropna()

        if clean_strings:
            series = series.astype(str)

        if label_sub_dict:
            series = series.replace(label_sub_dict)

        counts = series.value_counts()
        if not counts.empty:
            series_list.append(counts)

    if not series_list:
        return pd.Series(dtype=int)

    result = pd.concat(series_list).groupby(level=0).sum()

    if include_base:
        if base_calc_method == 'population':
            result.at[base_label] = len(df)
        else:
            result.at[base_label] = sum(series.sum() for series in series_list)

    logger.info(f'Manually verify {base_label} value is correct')

    return result
