# SalesTrackerScript
Uses Google Scripts with Google Sheets. Sorts inventory by Part No and creates individual sheets for each Part No from the master sheet. 

Google Scripts is JavaScript for Google Suite.

This program reads in a master program containing sales information for a company. It then sorts the master file into individual newly
created sheets for each product type. Example: All ES004 get placed in a sheet called ES004.
Once placed in a sheet the parts are then sorted by the last three digits of their product ID:

Example: A full part ID is ES004-101. The 101 is what is used to sort by least to greatest so ES004-101 is ahead of ES004-119 or ES004-110


The sales tracker is organized by five categories: Order Rec'D Date, Cus Name, Part No, Quantity, Sub-Totals (Quantity sum of last three digits)

Add in data from another source (I use Quickbooks) and then manipulate the financial data with the indiviudal part number spreadsheets.

Features to come: -Automatically create graphs and charts
                  ✅-Add in financial data to analyze and manipulate
                  
                  
To use: Modify spreadsheets and variables as needed. Then use the "STARTPROGRAM" to run the application.


**Quick tips
**Make sure to name your original spreadsheet "From_MRP" and financial data "From_QBs" or adjust the variables accordingly in the program.
**Change email in email function to whatever address you want
**To change format to desired ways, go to format function, go to this web address for formatting help with google scripts (https://developers.google.com/apps-script/reference/spreadsheet/range) or just google and stackoverflow
**Be careful when using the DeleteSheets() function because once it deletes something it is permanent unless you can recover it successfully with the undo button
**Contact me with any questions or concerns
**07/13/2018
