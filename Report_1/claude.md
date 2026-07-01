# Project Rules

This is an AI Agent that performs reporting, based on an Excel file.
I would like to find out the details of my profit and losses. This data is in the Excel file

# About Me

I have creat3ed an Excel file containing my Stock values and performances. It is continuosly updated.

# Rules

- The excel filename is MerrylAccount.xlsx. the file location is:
        C:\Users\johna\Documents
- The workbook name is Merrill.
- You may open the file in READ ONLY mode.
- If the file is already open, you may continue to read the file.
- Any report generated will be stored in the output/ folder

# Questions
- Ask me questions where there is ambiguity

# Actions
- open the excel file MerrylAccount.xslx. go to workbook named Merrill
- look at the rows that have letter 'f' in column A
- All rows with this value will be processed.
- For each such rows, do the followings:
    The row containg the letter 'f' in column A is marked as current row
    - Read the stock ticker
        
        This is done by moving over to columnB and looking up the value in column B in the row above current row. if the value is blank then continue looking up the row above it, until the value is not empty. 
        This represents the value of the ticker another
- Capture information
    - Capture the count from Column D in the current row, capture unit price from column E
    - Write to the output file named profit_loss_YYYY_MM_DD:HH:MM.pdf
# new instructions #1
- Modify the generated PDF so that each data row displays the current price extracted from Column F, Total Cost from Column G and Current Value from Cclumn H, and also display the Profit / Loss, and the percentage of change.
# new instructions #2
- Modify the generate_report.py to incclude the row Number for each ticker. Also order the report by Ascending Order for Profit / Loss
# new instructions #3
- Modify the generate_report.py: move the column named 'Sheet Row' one column left. AT the end of report generate a summary report that provides the total cost of the columns with negative Profit / Loss, and the total amount of Loss and calculate the percentage of loss. Also do the same with cost of columns with positive profit / Loss
 # new instructions #4
 - Modify genarate_report.py so that if the app is passed parameters'OBT' then the report should be Orderd By Ticker
  # new instructions #5
  - Modify generate_report.py: add a new column called Rating. This column shoulf be to the left of the ticker column. The value in this column should be   in column B of the current row. Also the file name should start with profit_ and not profi_
  # new instructions #6
  - Modify genarate_report.py : write the report to the 'output' directory in the parent directory of the file directory
  # new instructions #7
 - Modify genarate_report.py so that if the app is passed parameters'OBT' then the report name have the letters 'OBT' after the date portion