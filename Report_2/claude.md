# Project Rules

This is an AI Agent that performs reporting, based on an Excel file.
I have purchased stocks for 3 different accounts, named:E,R N
I would like to find out the details of my profit and losses. This data is in the Excel file

# About Me

I have creat3ed an Excel file containing my Stock values and performances. It is continuosly updated.

# Rules

- The excel filename is MerrylAccount.xlsx. the file location is:
        C:\Users\johna\Documents
- The workbook name is Merrill.
- You may open the file in READ ONLY mode.
- If the file is already open, you may continue to read the file.
- Any report generated will be stored in the output/ folder of the parent of the current directory

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
        This represents the value of the ticker. Call this row processing_row
- Capture information
    - the account for which this ticker is purchased is shown in column C. The value is
     one of the following: E or N or R.
    - Capture the count from Column D in the current row, capture unit price from column E
    - This information will be written to the output file named profit_loss_Individual_YYYY_MM_DD:HH:MM.pdf
# new instructions #1
- additionally, each row in the generated PDF also displays the current price extracted from Column F, Total Cost from Column G and Current Value from Column H, and also display the Profit / Loss, and the percentage of change.
# new instructions #2
- addionally each row in the generated report displays the relevant row number in the excel workbook. Also order the report by Ascending Order for Profit / Loss
# new instructions #3
-divide the report into 3 groups, based on the Account: 
N account named Edge-N
R account named Roth
E account named Edge
# new instructions #4
Modify the generated report so that the header  for each account includes the Total Loss, and Total Profit for that account.
Also the final Profit/Loss  table should include  one column for the total of  losses and one column for the total of profits for that account.
Also, if the user passes the parameters OBT then each report should be ordered by ticker name
