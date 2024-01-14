# Importing all libraries
import easygui
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side


# Main function starts here
if __name__ == "__main__":

    # Intialise lists and variables
    worksheet_names = []
    filters = []
    fields = []
    calculated_fields = []

    # Define column names
    columns = ['Worksheet Name', 'Filters', 'Fields', 'Calculations']

    # Create an empty DataFrame with specified columns
    df = pd.DataFrame(columns=columns)

    # Select the twb file to extract sheet names
    file = easygui.fileopenbox(filetypes=['twb'])
    
    # Open the file
    File_object = open(file)
    file = File_object.read() 

    # Read twb file as soup object
    soup = BeautifulSoup(file, 'xml')
    
    # Find the Worksheets tag in XML
    worksheets = soup.find('worksheets')
    allworksheets = worksheets.findAll('worksheet')
    for worksheet in allworksheets:
        worksheet_names.append(worksheet['name'])

    # Append worksheet names to the dataframe
    for value in worksheet_names:
        df = df.append({'worksheet': value}, ignore_index=True)

    # Iterate through dataframe to add field names
    for index, row in df.iterrows():
        found_tag = soup.find('worksheet', {'name': row['worksheet']})
        datasource_dependencies_tag = found_tag.findAll('datasource-dependencies')
        for datasource_dependency in datasource_dependencies_tag:
            column_tag = datasource_dependency.findAll('column')
            for column in column_tag:
                fields.append(column['name'])

                # Remove brackets, commas, and single quotes using list comprehension
                fields = [item.replace('[', '').replace(']', '').replace(',', '').replace("'", '') for item in fields]

                # Concatenate the elements into a single string for field names
                concatenated_string_fields = ','.join(fields)

                # Check if the field names contain calculation
                if "Calculation" in column['name']:
                    calculated_fields.append(column['caption'])
                else:
                    calculated_fields.append('')

                # Concatenate the elements into a single string for field names
                concatenated_string_calcs = ','.join(calculated_fields)

        # Find filters in the worksheet.
        filter = found_tag.findAll('filter')
        for group_filter in filter:
            group_filter_tag = group_filter.findAll('groupfilter', {'function': 'level-members'})
            for final_group_filter in group_filter_tag:

                filters.append(final_group_filter['level'])

                # Concatenate the elements into a single string for field names
                concatenated_string_filters = ','.join(filters)
     

        # Assign fields names to each row
        row['Fields'] = concatenated_string_fields
        row['Calculations'] = concatenated_string_calcs
        row['Filters'] = concatenated_string_filters


        # Clear the list after extracting details from a worksheet
        fields.clear()
        calculated_fields.clear()
        filters.clear()

    # Cleaning the calculated field column
    df['Calculations'] = df['Calculations'] = df['Calculations'].str.strip(',')
    
    # Cleaning filter field
    df['Filters'] = df['filters'].str.replace('nk','')
    df['Filters'] = df['filters'].str.replace('none','')
    df['Filters'] = df['filters'].str.replace(':','')
    df['Filters'] = df['filters'].str.replace('[','')
    df['Filters'] = df['filters'].str.replace(']','')
    
    # Excel file path
    excel_file_path = 'worksheet_details.xlsx'

    # Write dataframe to excel
    df.to_excel(excel_file_path, index=False, engine='openpyxl')

    # Open the excel file
    workbook = load_workbook(filename=excel_file_path)

    # Active sheet
    sheet = workbook.active 
    
    # Create a Font object with the desired font style
    font_header = Font(name='Arial', size=12, bold=True, italic=False)
    font_cell_range = Font(name='Arial', size=11, bold=False, italic=False)

    # Set the font style for header
    sheet['A1'].font = font_header
    sheet['B1'].font = font_header
    sheet['C1'].font = font_header
    sheet['D1'].font = font_header
  
    # Set the font style for range of Cells
    for row in sheet['A1:D100']:
        for cell in row:
            cell.font = font_cell_range

    # Set width for the columns in export
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 30
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 30

    for rows in  sheet['A1:D100']:
        for cell in rows:
            cell.alignment = Alignment(wrapText=True)

    workbook.save(excel_file_path)