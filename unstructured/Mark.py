import pandas as pd

#Read the Excel file that has student Answers
Studentscript = pd.read_excel("StudentScript.xlsx", sheet_name="Sheet1")

#Read the Excel file that has the correct Answers
Memo = pd.read_excel("Memo.xlsx", sheet_name="Sheet1")

#Merge the two dataframes
Compare = pd.merge(Studentscript, Memo, on="QuestionNumber", suffixes=('_Student', '_Memo')) 

Compare['Mark']= Compare.apply(lambda row: 2 if row['Answers_Student'] == row['Answers_Memo'] else 0, axis=1)

Compare.to_excel("Marking.xlsx", index=False)