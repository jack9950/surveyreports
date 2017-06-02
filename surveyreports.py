import openpyxl
import pandas as pd
from data import agentNameMap, agentToSupervisorMap

fileLocation = 'C:\\Users\\Jackson.Ndiho\\Documents\\Surveys\\';
fileName = 'MTD Detail - Jackson_18633_1495875294_2017-05-01_2017-05-27.csv'
filePath = fileLocation + fileName;

#The cvs file column headings are off by one column
#Add a column named 'Inserted Fix Column' to fix issue
correctedColNames = ['Department', 'Date', 'Agent Name', 'Csuid', 'Agent Id',
                     'Customer Id', 'Phone Number', 'Inserted Fix Column',
                     'Overall Recommend (NPS)', 'Effort:  Made Easy (CE)',
                     'Overall Experience (CSAT)', 'Agent Satisfaction (ASAT)',
                     'First Contact', 'Why This Score Care', 'Verbatim Comment',
                     'Used Web Prior to Call', 'Issue Solved', 'Friendly',
                     'Had Authority', 'Clear and Easy', 'Genuine Concern',
                     'Overall Relationship with Company', 'Please Explain',
                     'Region', 'Brand', 'Cust Lname', 'Call Date',
                     'Event Date', 'Cust Fname', 'Email']

svy = pd.read_csv(filePath, index_col=False)

svy.columns = correctedColNames

svy['Agent Names Proper'] = svy['Agent Name'].map(agentNameMap)

svy['Supervisor'] = svy['Agent Names Proper'].map(agentToSupervisorMap)

print('\n')

print(svy.loc[:, ['Agent Name', 'Agent Names Proper', 'Supervisor']])

# print(svy.columns)
# print(svy.loc[:, ['Date', 'Agent Name', 'Agent Names Proper', 'Supervisor']])

# agentData = openpyxl.load_workbook('C:\\Users\\\Jackson.Ndiho\\Documents
#                                       \\Surveys\\agentdata.xlsx')
# agentDataSheets = agentData.get_sheet_names()
# agentDataFirstSheetName = agentDataSheets[0]
# agentDataFirstSheet = agentData.get_sheet_by_name(agentDataFirstSheetName)
#
#
# # print(firstSheetName)
# agentNames = []
# for row in range(2, 60):
#     agentIDCell = "A" + str(row)
#     # agent_id = template_first_sheet[agent_id_cell].value
#     cell = agentDataFirstSheet[agentIDCell]
#     if cell.value:
#         print(cell.value)
#
# surveyData = openpyxl.load_workbook('C:\\Users\\\Jackson.Ndiho\\Documents\\Surveys\\surveys052617.xlsx')
# surveyDataSheets = surveyData.get_sheet_names()
# surveyDataFirstSheetName = surveyDataSheets[0]
# surveyDataFirstSheet = surveyData.get_sheet_by_name(surveyDataFirstSheetName)
#
# surveyInfo = []
# for row in range(2, 500):
#     surveyCell = "C" + str(row)
#     cell = surveyDataFirstSheet[surveyCell]
#     if cell.value:
#         print(cell.value)
