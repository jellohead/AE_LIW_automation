* Encoding: UTF-8.
* Frequencies for everything.
FREQUENCIES VARIABLES=sys_RespNum sys_StartTime sys_EndTime sys_ElapsedTime S1 S2 S3 Q1 Q1_7_other 
    Q2_1 Q2_2 Q2_3 Q2_4 Q2_5 Q2_6 Q2_7 Q2_8 Q2_9 Q2_10 Q2_11 Q2_12 Q2_12_other Q3_r1 Q3_r2 Q3_r3 Q3_r4 
    Q3_r5 Q3_r6 Q4 Q5 Q6 Q6_1 Q6_2 Q6_3 Q7_r1 Q7_r2 Q7_r3 Q7_r4 Q7_r5 Q8 Q9 Q10 Q10_1 Q11 Q12 Q12_1 
    Q12_2 Q13 Q13_1 Q13_2 Q13_3 Q14_r1 Q14_r2 Q14_r3 Q14_r4 Q14_r5 Q14_r6 Q14_r7 Q14_r8 Q14_r9 Q15 Q16 
    Q17 Q18 Q19 Q20 Q21 Q22 Q23_1 Q23_2 Q23_3 Q23_4 Q23_5 Q23_6 Q23_7 Q23_8 Q23_8_other Q24 Q25 Q25_1 
    Q25_2 Q26 Q26_1 Q27 Q28 Q28_1 Q29_1 Q29_2 Q29_3 Q29_4 Q29_5 Q29_6 Q29_7 Q29_8 Q29_8_other Q30 Q30_1 
    Q31 Q32 D1 D2 D3 D4 D5 D5_7_other D6 D7 D8 D9 D9DK D10 D10DK D11 D12 ContactInfo_first 
    ContactInfo_last ContactInfo_email ContactInfo_Phone InterID Account_Name PHONE Program Premise_Id 
    Service_Address City State Zip Quarter District Total_Sq_Ft Building_Type_Name Occupancy_Status 
    IDNum FirstName LastName CustomerName PhoneNumber 
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
BEGIN PROGRAM PYTHON3.
import spss,spssaux

# start with calcSurveyLength.py from last report

import pandas as pd
from datetime import datetime, timedelta

Data_Info=str(spssaux.GetDatasetInfo())
Slash_Pos=Data_Info.rfind("/")
Ext_Pos=Data_Info.rfind(".")
SPSS_Name = Data_Info[Slash_Pos+1:Ext_Pos]
SPSS_Path=Data_Info[:Slash_Pos+1]

df = pd.read_spss(SPSS_Path + SPSS_Name + '.sav')

totalTime = 0
validSurveys = 0
totalTimeUsing_strptime = 0

def extract_time(input_string):
    time = input_string.split(" - ")[1].split(" ")[0]
    return time


def subtract_time(time1, time2):
    time1 = datetime.strptime(time1, "%H:%M:%S")
    time2 = datetime.strptime(time2, "%H:%M:%S")
    difference = time2 - time1
    return difference.total_seconds()/60

for row in df.itertuples():
    surveyDuration = subtract_time(extract_time(df.at[row.Index, "sys_StartTime"]), extract_time(df.at[row.Index, "sys_EndTime"]))
    print ('survey duration = ', surveyDuration, 'minutes')

    if surveyDuration <= 60:
        validSurveys += 1 
        totalTime += surveyDuration
 
print("Total Time = ", totalTime , "minutes, or")
print("Total number of surveys = ", len(df))
print('Surveys <= 1 hour long = ', validSurveys)
print("Time per survey = ", totalTime/validSurveys)

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
