# Importing all libraries
import easygui
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side


# Main function starts here
if __name__ == "__main__":

    # Select the twb file to extract sheet names
    file = easygui.fileopenbox(filetypes=['twb'])
    
    # Open the file
    File_object = open(file)
    file = File_object.read() 

    # Intialise list
    calculation_name = []
    calculation_id = []
    calculation_formula = []
    used_in_calculations = []
    used_in_sheets = []
    datasource_name = []
    worksheet_name = []
    columns_in_worksheet = []
    calcs_in_worksheet = []
    
    
    # Define column names
    columns = ['Calculation Name', 'Data Source Name', 'Calculation ID', 'Formula']

    # Create an empty DataFrame with specified columns
    df = pd.DataFrame(columns=columns)

    # Read twb file as soup object
    soup = BeautifulSoup(file, 'xml')

    # Find the names of datasources and store it in a list.
    datasources = soup.find('datasources')
    alldatasource = datasources.findAll('datasource')
    for datasource in alldatasource:
       
        # Find the columns from xml
        columns = datasource.findAll('column')
        
        # Iterate through each column
        for column in columns:
            
            column_name = column['name']
            if "Calculation" in column_name:  
                formula_tag = column.find('calculation')
                calculation_formula.append(formula_tag['formula'])      
                calculation_name.append(column['caption'])
                calculation_id.append(column['name'])
                datasource_name.append(datasource['caption'])
            else:
                pass

    # Add value from list to dataframe
    df['Calculation Name'] = calculation_name
    df['Data Source Name'] = datasource_name
    df['Calculation ID'] = calculation_id
    df['Formula'] = calculation_formula

    # Find if the calculation is used in another calculation
    for index, row in df.iterrows():
        calc_id = row['Calculation ID']
        contains_calc = []
        for formula in calculation_formula:           
            if calc_id in formula:
                contains_calc.append(True)
            else:
                contains_calc.append(False)
        contains_calc = any(contains_calc)
        used_in_calculations.append(contains_calc)
    
    # Add the new column to dataframe
    df['used_in_calculations'] = used_in_calculations
   
    # Find the Worksheets name
    worksheets = soup.find('worksheets')
    allworksheets = worksheets.findAll('worksheet')
    for worksheet in allworksheets:
        worksheet_name.append(worksheet['name'])

    # Find the list of columns from Worksheet
    for worksheet in worksheet_name:
        found_tag = soup.find('worksheet', {'name': worksheet})
        datasource_dependencies_tag = found_tag.findAll('datasource-dependencies')
        for datasource_dependency in datasource_dependencies_tag:
            column_tag = datasource_dependency.findAll('column')
            for column in column_tag:
                columns_in_worksheet.append(column['name'])

                # Remove brackets, commas, and single quotes using list comprehension
                columns_in_worksheet = [item.replace('[', '').replace(']', '').replace(',', '').replace("'", '') for item in columns_in_worksheet]

    # Get the calculations from columns 
    for calcs in columns_in_worksheet:
        if 'Calculation' in calcs:
            calcs_in_worksheet.append(calcs)

    # Remove duplicates from the list
    calcs_in_worksheet = list(set(calcs_in_worksheet))

    # Find the calculations in sheet
    for index, row in df.iterrows():
        calc_id = row['Calculation ID']
        calc_id = calc_id.replace('[', '').replace(']', '')
        sheet_contains_calc = []
        for calculation in calcs_in_worksheet:
            if calc_id == calculation:
              sheet_contains_calc.append(True)
            else:
               sheet_contains_calc.append(False)
        sheet_contains_calc = any(sheet_contains_calc)
        used_in_sheets.append(sheet_contains_calc)

    # Add the column to dataframe
    df['used_in_sheets'] = used_in_sheets

    # Filter the row for the condition
    filtered_df = df[(df['used_in_sheets'] == False) & (df['used_in_calculations'] == False)]

    # Excel file path
    excel_file_path = 'unused_calculations.xlsx'

    # Write dataframe to excel
    filtered_df.to_excel(excel_file_path, index=False, engine='openpyxl')