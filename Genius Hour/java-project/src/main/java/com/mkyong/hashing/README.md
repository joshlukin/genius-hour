# Web Scraper in Java

This program scrapes data from the NHL API based on specific criteria and saves it to an Excel file. It uses the following dependencies:

- `org.apache.poi.ss.usermodel`: Used to read/write Microsoft Excel files (.xls and .xlsx)
- `org.json`: Used to work with JSON data

The program performs the following actions:

1. Creates an instance of `Workbook` to work with Excel files.
2. Creates a dedicated sheet and header row for all of the stats to be in one place.
3. Connects to NHL stats to determine how many records they have.
4. Loops through NHL's pages of JSONs and collects data from each record within the JSON.
5. Creates a header row for posts stat sheet.
6. Adds the collected stats to a Player Posts sheet.
7. Determines the total number of records from the first page of NHL's data.
8. Loops through the number of pages from NHL.
9. Creates a sheet for the shot and goal totals.
10. Adds headers for the shot and goal sheet.
11. Collects stats from individual players on each page.
12. Adds stats to the spreadsheet.
13. Transfers all data to the All Stats sheet at the appropriate columns.
14. Calculates posts + crossbars hit and (goals + posts) / shots percentages using `FormulaEvaluator`.
15. Writes the data to the Excel file.
16. Closes the workbook.