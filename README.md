# ExtractPowerPivot

## Tool for easy PowerPivot data extraction (models embedded inside Excel files)

1. Can open PowerPivot 2013 models in Excel 2010 (creates a new file by injecting the 2013 ABF backup file into an empty 2010 Excel file). Also does the reverse. Empty files are stored as BASE64 encoded strings.
2. Can query 2008 RTM, 2010 and 2013 PowerPivot models.
3. Can get table and column metadata, unique values in each column.
4. Can retrieve entire tables.
5. Can retrieve subsets of data (conditioning on one or more values).
6. Can work with row numbers, which in principle don't exist in PowerPivot/SSAS world (RowNumber is a hidden internal column that can only be accessed using a special SQL-like query).
7. Can write required settings to registry if using Excel 2013 (needs to be done only once). Check if settings present before 1st use.

Disclaimer: Never really tested with Excel 2016.

Sample PowerPivot data used in the xlsm file can be found here: https://www.microsoft.com/en-us/download/details.aspx?id=102

## TO-DO
Minor fixes required:

1. no 30 second wait (find a more clever way)
2. open pre2008 model in 2013
3. "Open Excel2013zip For Binary Lock Read Write As #1"
