#Header Files Import
import pandas as pd
from openpyxl import load_workbook
import math

#Put excel sheet file path here
file_path = "accounts.xlsx"

#Generating list of unique cities 
project_data = pd.read_excel(file_path)
unique_projects = project_data['Project'].unique()
print(unique_projects)

#Creating new sheets for each city and new data frames for each city's project details and adding it to the newly created sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    for i in unique_projects:
        sheet_name = ""
        if pd.isna(i):
            sheet_name = 'Unknown_Projects'
            projectdf = project_data[project_data['Project'].isna()]
            projectdf.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            sheet_name = i+"_Projects"[:31]
            projectdf = project_data[project_data['Project']==i]
            projectdf.to_excel(writer, sheet_name=sheet_name, index=False)
