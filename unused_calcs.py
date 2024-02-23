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
    used_in_calculations = []
    used_in_sheets = []

    