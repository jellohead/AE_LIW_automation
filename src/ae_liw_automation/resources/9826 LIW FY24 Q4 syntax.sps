* Encoding: UTF-8.
* Frequencies for everything.

FREQUENCIES VARIABLES=sys_RespNum to D12
    /ORDER=ANALYSIS.

RECODE Q11 (1 THRU 7=0) (11=0) (8 THRU 10=1) INTO Q11_TopBox.
VARIABLE LABELS Q11_TopBox 'Q11 - 11. How would you rate your satisfaction with the Austin Energy weatherization program?'.
EXECUTE.
FREQUENCIES VARIABLES=Q11_Topbox
    /ORDER=ANALYSIS.

RECODE Q18 (1 THRU 7=0) (11=0) (8 THRU 10=1) INTO Q18_TopBox.
VARIABLE LABELS Q18_TopBox 'Q18 - 18. How important is it to you that Austin Energy offers its customer assistance with home weatherization?.'.
EXECUTE.
FREQUENCIES VARIABLES=Q18_Topbox
    /ORDER=ANALYSIS.


RECODE Q22 (1 THRU 7=0) (11=0) (8 THRU 10=1) INTO Q22_TopBox.
VARIABLE LABELS Q22_TopBox 'Q22 - 22. What is your level of understanding of your utility bill and the energy savings related to home improvements?'.
EXECUTE.
FREQUENCIES VARIABLES=Q22_Topbox
    /ORDER=ANALYSIS.

RECODE Q24 (1 THRU 7=0) (11=0) (8 THRU 10=1) INTO Q24_TopBox.
VARIABLE LABELS Q24_TopBox 'Q24 - 24. How satisfied are you with the amount of energy savings you are seeing on your bill since your energy improvements were completed?'.
EXECUTE.
FREQUENCIES VARIABLES=Q24_Topbox
    /ORDER=ANALYSIS. 

RECODE Q31 (1 THRU 7=0) (11=0) (8 THRU 10=1) INTO Q31_TopBox.
VARIABLE LABELS Q31_TopBox 'Q31 - 31. How satisfied are you with Austin Energy?'.
EXECUTE.
FREQUENCIES VARIABLES=Q31_Topbox
    /ORDER=ANALYSIS. 


* slide 28.
CROSSTABS 
  /TABLES=Q27 BY Q28 Q28_1 
  /FORMAT=AVALUE TABLES 
  /CELLS=COUNT 
  /COUNT ROUND CELL.


* Calculate average survey duration.
* Add SPSS python wrapper.
BEGIN PROGRAM PYTHON3.
import spss,spssaux
import pyreadstat
from datetime import datetime, timedelta

def extract_time(input_string):
        time = input_string.split(" - ")[1].split(" ")[0]
        return time

def subtract_time(time1, time2):
    time1 = datetime.strptime(time1, "%H:%M:%S")
    time2 = datetime.strptime(time2, "%H:%M:%S")
    difference = time2 - time1
    return difference.total_seconds()/60

def main():
    Data_Info=str(spssaux.GetDatasetInfo())
    df, meta = pyreadstat.read_sav(Data_Info)

    totalTime = 0
    validSurveys = 0
    totalTimeUsing_strptime = 0
    recordCount = 0
    
    for row in df.itertuples():
        recordCount += 1
        surveyDuration = subtract_time(extract_time(df.at[row.Index, "sys_StartTime"]), extract_time(df.at[row.Index, "sys_EndTime"]))
        print ('Record #', recordCount, 'survey duration = ', surveyDuration)

        # do not include survey durations that are negative or > 1 hour long
        if surveyDuration <= 60 and surveyDuration > 0:
            validSurveys += 1 
            totalTime += surveyDuration
    
    print(f"\nTotal Time = {totalTime} minutes")
    print(f"Total number of surveys = {len(df)}")
    print(f"Number of surveys > 0 minutes  and <= 1 hour long = {validSurveys}")
    print(f"Time per survey = {totalTime/validSurveys} minutes\n")

if __name__ == "__main__":
    main()

END PROGRAM.



USE ALL.
COMPUTE filter_$=(Q27<8).
VARIABLE LABELS filter_$ 'Q27<8 (FILTER)'.
VALUE LABELS filter_$ 0 'Not Selected' 1 'Selected'.
FORMATS filter_$ (f1.0).
FILTER BY filter_$.
EXECUTE.



FREQUENCIES VARIABLES=Q28_1
  /ORDER=ANALYSIS.

Filter OFF.

* Export Output in Excel format.
BEGIN PROGRAM PYTHON3.
import spss, spssaux
import os

# Get dataset information
Data_Info_path = str(spssaux.GetDatasetInfo())
print(f'Data_Info_path: {Data_Info_path}')

# Extract the directory path and modify the filename
directory_path = os.path.dirname(Data_Info_path)
file_name = os.path.basename(Data_Info_path).replace('.sav', ' Excel output.xlsx')
data_info_path_revised = os.path.join(directory_path, file_name)
print(f'data_info_path_revised: {data_info_path_revised}')

# Define the SPSS syntax for exporting output to Excel
spss_syntax = f"""
OUTPUT EXPORT
  /CONTENTS EXPORT=ALL LAYERS=PRINTSETTING MODELVIEWS=PRINTSETTING
  /XLSX DOCUMENTFILE='{data_info_path_revised}'
     OPERATION=CREATEFILE
     LOCATION=LASTCOLUMN NOTESCAPTIONS=YES.
"""

# Submit the SPSS syntax
spss.Submit(spss_syntax)

END PROGRAM.

