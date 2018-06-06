# RemoveLeadingTrailingSpaces

Visual Basic Macro that will remove any 'unseen' leading or trailing spaces from a selection of data in Excel.
Quite often you can be dealing with data that has been supplied to you in Excel or a CSV file.  Leading and / or trailing spaces can be very difficult to spot and can cause issue when using that data.  Think searches failing because the search criteria is "User1 " and not "User1" or scripts creating / amending users with incorrect values.  Select the data you are going to save as CSV and run this macro.  Any internal space is left alone i.e "   Surname, FirstName  " will be "Surname, FirstName" after running the macro.

To have this available at all times import the .bas file via the Developer Tab, Visual Basic Button. After its been imported you need to close Visual Basic and return to your Workbook, click the Macros Button and save the CSV Macro to your Personal Macro Workbook. (If you do not have a personal macro workbook you will need to create one).  The CSV Macro should now be available in all new workbooks.

Disclaimer: This script and any content is offered "as is" with no warranty. While this script was tested and working in my environment, it is vital that you test it in a test environment before using in your production environment. I am not responsible for any outcome that arises from you using these scripts.

My GitHub repositories are really just a collection of scripts and tools I've found and amended to suit my purposes or ones I've attempted to write from scratch to fulfil a requirement I've had at a particular time.

I started to use GitHub as I realised that I needed somewhere I could keep code so I could refer back to, to jog my memory on how to do things or to reuse so that I wasn't reinventing my wheel. (I say 'my' wheel as I'm not arrogant enough to think that any of my scripts could be 'the' wheel.)

Also, any suggestions on improvements to these scripts or better ways to do things are greatly appreciated. We're all standing on the shoulders on giants.... and I need more giants!
