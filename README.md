# Tableau-Extractors
Python code to understand Tableau workbooks and better workbooks

1) Worksheet Extractor
----------------------
  
    With Enterprise Tableau Workbook the number of worksheets and its associated items can increase with complexity.This code helps extract Worksheet and its associated information such as
    
    - Filters
    - Calculated Fields
    - Fields
    
    How to use : 
    -----------
    
    1) Open command promt and navigate to respective directory.
    2) Type "python worksheet_extractors.py"
    3) Upload twb file in the dailog that appears.
    4) A excel sheet will be generated in the same directory as the Code is downloaded.



Unused Calculation
---------------------

   Calculations in workbook are used either in a worksheet or used to create other calculated fields. Often during analysis developers create calculations and they are not used in tableau workbook. It is hard to identify the calculations that are used in Tableau workbook, it often makes it hard to maintain them.

This code helps us identify the calculations that are not used in workbook.

How to Use
----------

  1) Open command promt and navigate to respective directory.
  2) Type "python unused.py"
  3) Upload twb file in the dailog that appears.
  4) A excel sheet will be generated in the same directory as the Code is downloaded.
