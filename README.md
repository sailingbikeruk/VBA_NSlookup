# VBA_NSlookup
Allows NSLookup as a formula/function, but uses a pre-defined list of name servers rather than the endpoints default DNS

_You will need to change the code on lines 5-11 to use the name servers you wish to use and change the For loop so that iterates across the correct number of servers._

Original code is from https://www.geekality.net/2016/03/07/excel-function-for-nslookup-in-worksheet/

Add the code to a module in the Excel Visual Basic editor.

Save as an Excel Add-in, the recommended location at the time of writing is %appdata%\Roaming\Microsoft\AddIns

- Enable the add-in in Excel - File > Options > Add-ins
- At the bottom of the add-ins windows make sure "Excel Add-Ins" is selected in the Manage drop down box
- click go
- In the box that opens select (tick) the entry that relates to your saved add-ins workbook 
- click OK

In the spreadsheet with IP addresses, you should now have user defined functions NSLookup and FindIP available, you can check this by opening a sheet, clicking in a cell and starting to type either, you should see the autocomplete appear:\

or you can go to Formulas > Insert function > User Defined Functions

## Usage 

=NSlookup(Value_To_Lookup,ReturnType)

Return types are

- 1 = IP Address
- 2 = Host Name (FQDN)

## Examples

- =NSLookup(10.10.10.1,2)
- =NSlookup(A1,2)
- =NSlookup("yahoo.co.uk",1)
