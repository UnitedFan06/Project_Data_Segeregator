#Header Files Import
import pandas as pd
from openpyxl import load_workbook

#Put excel sheet file path here
file_path = "project_data.xlsx"

#Generating list of unique cities 
project_data = pd.read_excel(file_path)
unique_cities = project_data['City'].unique()

#Creating new sheets for each city and new data frames for each city's project details and adding it to the newly created sheet
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    for i in unique_cities:
        sheet_name = i+"_Projects"[:31]
        citydf = project_data[project_data['City']==i]
        citydf.to_excel(writer, sheet_name=sheet_name, index=False)