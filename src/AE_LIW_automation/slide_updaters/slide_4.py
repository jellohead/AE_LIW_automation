# slide_4.py
# This file contains the functions for updating the slide 4 of the Powerpoint file.


from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
# from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR

from config import REPORTING_PERIOD, REPORTING_YEAR, CURRENT_MONTH_TEXT, CURRENT_YEAR
from helper_modules import get_chart_object_by_name


def slide_4_updater(df, prs) -> object:
    print('updating slide 4')
    slide_index = 3
    slide = prs.slides[slide_index]

    q24_TopBox_result = (df['Q24']
                         .dropna()
                         .apply(lambda x: 'TopBox'if x in [8, 9, 10] else 'BottomBox')
                         .value_counts(normalize=True)
                         )['TopBox']

    paragraph_strings = [f'Overall satisfaction score with energy savings was {q24_TopBox_result:.0%} in {REPORTING_PERIOD} {REPORTING_YEAR}, up from 29% in Q2 2024.',
                         f'Contractor and customer service ratings remained relatively high for all attributes.',
                         f'Customers appeared to be satisfied with the follow-up phone calls and indicated that the Austin Energy staff member/contractor did an overall good job on the work done at their homes.',
                         f'For this quarter, Austin Energy’s website, friends, family, and word of mouth, as well as utility bill inserts were the top responses for how customers first learned about the weatherization program.',
                         f'For this quarter, due to the small sample size, none of the changes can be deemed significant.'
                         ]

    text_holder = None
    # get shape object by name
    for shape in slide.shapes:
        if shape.name == 'Rectangle 3':
            shape.text_frame.clear()
            text_holder = shape.text_frame

    p = text_holder.paragraphs[0]

    for para_string in paragraph_strings:
        print(para_string)
        clean_text = para_string.replace('-', '')
        p = text_holder.add_paragraph()
        p.text = clean_text.strip()
        p.alignment = PP_ALIGN.LEFT

        if para_string.startswith('-'):
            p.level = 1
        else:
            p.level = 0


        run = p.runs[0]
        run.font.name = 'Tahoma'
        run.font.size = Pt(18 if p.level == 0 else 16)
