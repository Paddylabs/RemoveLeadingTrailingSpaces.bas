# RemoveLeadingTrailingSpaces

Visual Basic Macro that will remove any 'unseen' leading or trailing spaces from a selection of data in Excel.
Quite often you can be dealing with data that has been supplied to you in Excel or a CSV file.  Leading and / or trailing spaces can be very difficult to spot and can cause issue when using that data.  Think searches failing because the search criteria is "User1 " and not "User1" or scripts creating / amending users with incorrect values.  Select the data you are going to save as CSV and run this macro.  Any internal space is left alone i.e "   Surname, FirstName  " will be "Surname, FirstName" after running the macro.

To have this available at all times import the .bas file via the Developer Tab, Visual Basic Button. After its been imported you need to close Visual Basic and return to your Workbook, click the Macros Button and save the CSV Macro to your Personal Macro Workbook. (If you do not have a personal macro workbook you will need to create one).  The CSV Macro should now be available in all new workbooks.
