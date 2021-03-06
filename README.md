# SalesTrackerScript
Uses Google Scripts with Google Sheets. Sorts inventory by Part No and creates individual sheets for each Part No from the master sheet. 

Google Scripts is JavaScript for Google Suite.

Check the latest release for the most stable and current version (MASTER Template)

This program reads in a master program containing sales information for a company. It then sorts the master file into individual newly
created sheets for each product type. Example: All ES004 get placed in a sheet called ES004.
Once placed in a sheet the parts are then sorted by the last three digits of their product ID:

Example: A full part ID is ES004-101. The 101 is what is used to sort by least to greatest so ES004-101 is ahead of ES004-119 or ES004-110


The sales tracker is organized by five categories: Order Rec'D Date, Cus Name, Part No, Quantity, Sub-Totals (Quantity sum of last three digits)

Add in data from another source (I use Quickbooks) and then manipulate the financial data with the indiviudal part number spreadsheets.

Automatically sorts sheets within file in ascending or descending order.

Running total on sales for each individual part.

Creates a financial summary tab that shows each part number along with the sales and quantity sold per that part. There is also a total sales number shown and total quantity shown. The date range of sales is included at the top. Now includes a biggest part order section where it shows largest quantity of a part ordered along with which company ordered it.

Good for financial analysis on a company's sales for each part type.

Features to come: -Automatically create graphs and charts
                  ✅-Add in financial data to analyze and manipulate
                  
                  
To use: Modify spreadsheets and variables as needed. Then use the "STARTPROGRAM" to run the application.

Script worked with data from a 2017 sheet and a 2018 sheet so as long as the format of the spreadsheets is proper (correct naming and format of the From_MRP and From_QBs) the script will work for anything.


**Quick tips
**Make sure to name your original spreadsheet "From_MRP" and financial data "From_QBs" or adjust the variables accordingly in the program.
**Make sure to delete old sheets before sorting again or running STARTME again... hopefully this is fixed soon
**Change email in email function to whatever address you want
**To change format to desired ways, go to format function, go to this web address for formatting help with google scripts (https://developers.google.com/apps-script/reference/spreadsheet/range) or just google and stackoverflow
**Be careful when using the DeleteSheets() function because once it deletes something it is permanent unless you can recover it successfully with the undo button
**Contact me with any questions or concerns
**07/19/2018
***Author: Ryan Regier
