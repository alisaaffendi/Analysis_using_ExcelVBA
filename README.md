# Analysis_using_ExcelVBA
This project is using Excel, to analyse data on sales, create receipt automatically and update stocks in inventory.

There are three Excel files included as a part of this project. There are:
1. Attendance_Project.xlsm
2. Inventory_Project.xlsm
3. Receipt.xlsm

For the first file - Attendance_Project.xlsm:
- It has data for staff of sales
- It has attandance data for each staff from Jan to Dec, with analysis ('Analysis' sheet) for each late check-in or early check-out.

For Inventory_Project.xlsm:
- Reads data from Attendance_Project.xlsm file
- Has inventory and item list for every products sold and stocks left (with details of salesperson)
- Has 'Receipt' sheet, which records sales transaction then proceed to export receipt as pdf and sheets in excel, update new stocks in inventory.

For Receipt.xlsm:
- This is the file that recorded receipt made in 'Receipt' sheet in Inventory_Project.xlsm file.
- The receipt is created with unique identifier (the receipt name) for each transactions and exported pdf will renamed using its identifier (to avoid replicates)
