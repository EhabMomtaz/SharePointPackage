SharePointAPI

Description : 

SharePontAPI package enables you to access files on Sharepoint via Microsoft Graph API only to Sharepoint Sites which accessed by our Azur Application.

If you want to use the Package on other different Sharepoint Sites than RPA Site, you should contact IT and require adding the Site URL with permissions 
to Read and Write on the Azur App.

The Package enables you to do some operations on Sharepoint files.

Available activities to do on Sharepoint Excel files are (Download one file,Download many files,Write Range,Append Range,Read Range).

You also can download any files not only Excel files.


How to Install?
Please Search "SharePointAPI' in our local packages and download last version.
The Package contains one activity, which can do all above-mentioned available activities by controlling Arguments (details below).

---------------------------------------------------------------------------------------------------------------------------------------------------------------------

Mandatory Arguments:

RequiredExcelActivity - type: String (the required excel activity) values details below:

- If you want to download one file from Sharepoint, please give value of "RequiredExcelActivity" argument "one"
- If you want to download many files from one sharepoint folder, this argument must be "many".
- If you want to Write Range on file Excel exist on SharePoint, this argument must be "Write Range".
- If you want to Append Range on Excel file exist on SharePoint, this argument must be "Append Range".
- If you want to Read Range on Excel file exist on SharePoint, this argument must be"Read Range".

SiteURL - type String (the Sharepoint Site URL where file/s located).
SharePointFolder - type String (the Sharepoint folder where you file/s located).
FileName - type String (the file name you want to Download,read,write or append , this argument is not required when you want to download many files must be empty).
---------------------------------------------------------------------------------------------------------------------------------------------------------------------

----------------------------------------------------------Activities Arguments---------------------------------------------------------------------------------------

Dwonload file/s arguments:
LocalFilePath - type String (path where you want file/s to be downloaded).
---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Write Range arguments:
StartingCellToWrite - type - String (the starting cell to write your datatable, if left empty data will be written from A1).
SheetName - type - String (the Sheet name where you want your datatable to be written, If left empty the activity will throw an exception).
DataTableToWrite - type - DataTable (the Datatable you want to be written on file, If left empty the activity will throw an exception).
---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Append Range arguments:
SheetName - type - String (the Sheet name where you want your datatable to be written, If left empty the activity will throw an exception).
DataTableToWrite - type - DataTable (the Datatable you want to be written on file, if left empty the activity will throw an exception).
---------------------------------------------------------------------------------------------------------------------------------------------------------------------

Read Range arguments:
SheetName - type - String (the Sheet name where you want your datatable to be written, If left empty the activity will throw an exception).
RangeToRead - type - String (the range to be read, if left empty the whole sheet will be read).
DataTableOut - type - DataTable (the datatable out argument to be retrieved).
---------------------------------------------------------------------------------------------------------------------------------------------------------------------
