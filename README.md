# VBA_NSlookup
Allows NSLookup as a formula/function. 

Original code is from https://www.geekality.net/2016/03/07/excel-function-for-nslookup-in-worksheet/

Add the code to a module in the Excel Visual Basic editor.
Save as an Excel Add-in - recommended location at the tgime of writing is %appdata%\Roaming\Microsoft\AddIns

Enable the add-in in Excel - File > Options > Add-ins 
At the bottom of the add-ins windows make sure "Excel Add-Ins" is selected in the Manage drop down box and click go
In the box that opens select (tick) "My Functions" 
click OK

In the spreadsheet with IP addresses, you should now have user defined functions NSLookup and FindIP available, you can check this by opening a sheet, clicking in a cell and starting to type either, you should see the autocomplete appear:\

or you can go to Formulas > Insert function > User Defined Functions

usage is 

=NSlookup(Value_To_Lookup,ReturnType)

Return types are

1 = IP Address
2 = Host Name (FQDN)
