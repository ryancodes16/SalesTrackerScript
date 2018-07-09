# SalesTrackerScript
Uses Google Scripts with Google Sheets. Sorts inventory by Part No and creates individual sheets for each Part No from the master sheet. 

Google Scripts is JavaScript for Google Suite.

This program reads in a master program containing sales information for a company. It then sorts the master file into individual newly
created sheets for each product type. Example: All ES004 get placed in a sheet called ES004.
Once placed in a sheet the parts are then sorted by the last three digits of their product ID:

Example: A full part ID is ES004-101. The 101 is what is used to sort by least to greatest so ES004-101 is ahead of ES004-119 or ES004-110


The sales tracker is organized by five categories: Order Rec'D Date, Cus Name, Part No, Quantity, Sub-Totals (Quantity sum of last three digits)

Features to come: -Automatically create graphs and charts
                  -Add in financial data to analyze and manipulate
