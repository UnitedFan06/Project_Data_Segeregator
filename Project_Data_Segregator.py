import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import os

file_path = "project_data.xlsx"
main_sheet_name = 0

try:
    project_data = pd.read_excel(file_path, sheet_name=main_sheet_name)
    # FIX: A DataFrame does not have a 'name' attribute. Use main_sheet_name instead.
    # If main_sheet_name is an integer (like 0), its string representation is fine here.
    print(f"Successfully loaded main sheet '{main_sheet_name}' from '{file_path}'.")
except FileNotFoundError:
    print(f"Error: The file '{file_path}' was not found. Please ensure it exists.")
    exit()
except Exception as e:
    print(f"Error reading main Excel file: {e}")
    exit()

unique_projects = project_data['Project'].unique()

# Initialize all_sheets_data dictionary.
# If main_sheet_name was an integer, we need to get its actual name from the loaded DataFrame
# to use as the key for the dictionary if we want consistency with string sheet names.
# However, for simply storing the main DataFrame, just using the integer key is fine too
# or explicitly getting the sheet name if it's guaranteed to be a string name.
# For simplicity and to match the previous version's implicit main sheet handling,
# we'll use a placeholder key or simply add the main sheet to the dictionary by its resolved name
# after checking if it's a string name or an integer index.

# Get the actual name of the main sheet after loading
if isinstance(main_sheet_name, int):
    # When reading by index, pandas sets the sheet_name attribute on the returned DataFrame
    actual_main_sheet_name = project_data.attrs.get('sheet_name', str(main_sheet_name))
else:
    actual_main_sheet_name = main_sheet_name

all_sheets_data = {actual_main_sheet_name: project_data}


existing_workbook_sheets = []
if os.path.exists(file_path):
    try:
        wb = load_workbook(file_path, read_only=True)
        existing_workbook_sheets = wb.sheetnames
        wb.close()
    except InvalidFileException:
        print(f"Warning: '{file_path}' is not a valid Excel file or is corrupted. Will attempt to create/overwrite all sheets.")
    except Exception as e:
        print(f"An error occurred while inspecting existing sheets: {e}")

print("\nProcessing projects:")
for project_name in unique_projects:
    sheet_name_for_project = str(project_name) + "_Projects"
    if len(sheet_name_for_project) > 31:
        sheet_name_for_project = sheet_name_for_project[:31]

    new_project_data_df = project_data[project_data['Project'] == project_name].copy()

    df_to_write_to_sheet = new_project_data_df

    if sheet_name_for_project in existing_workbook_sheets:
        print(f"  - Appending data to existing sheet: '{sheet_name_for_project}'")
        try:
            existing_sheet_df = pd.read_excel(file_path, sheet_name=sheet_name_for_project)

            combined_df = pd.concat([existing_sheet_df, new_project_data_df], ignore_index=True)

            df_to_write_to_sheet = combined_df.drop_duplicates()

        except Exception as e:
            print(f"    Warning: Could not read existing sheet '{sheet_name_for_project}'. Overwriting instead. Error: {e}")
            df_to_write_to_sheet = new_project_data_df
    else:
        print(f"  - Creating new sheet: '{sheet_name_for_project}'")

    all_sheets_data[sheet_name_for_project] = df_to_write_to_sheet

try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
        for sheet_name, df_data in all_sheets_data.items():
            df_data.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"\nExcel file '{file_path}' updated successfully with all project data (appended where applicable).")
except Exception as e:
    print(f"\nError writing Excel file: {e}")
