# vlookupincell
Excel function to do a vlookup on multiple values within a single cell

# What this function does:

It takes the value of a cell and splits is up into pieces. Next is does a vlookup of each piece using a range you define. Finally it exports the newly lookuped up values in the same format as the original cell.

Example:
Input cell value: "thing-1,thing-2,thing-3"

Lookup Range:
        Column A | Column B
ROW 1   thing-1  |  14
ROW 2   thing-2  |  21
ROW 3   thing-3  |  7
ROW 4   thing-4  |  1001

Deliminator: ","

=vlookupincell(cell, range, deliminator)

For this example the output would be: 14,21,7,1001

# VB code

Function vlookupInCell(rng As Excel.Range, lookup As Excel.Range, delim As String) As String
  'This function designed to take in the value from a cell, split it using a specified deliminator, then do a vlookup on each
  'piece the output is is a string in the same format as the original cell but replaces all the original values with the vlookup values
  Dim tempValue As String
  Dim arrValues() As String
  ' split input cell into an array using delim
  arrValues = Split(rng, delim)
  ' loop through each value, do a vlookup using the lookup range and build output string
  For Each Item In arrValues
    tempValues = tempValues & Application.WorksheetFunction.VLookup(Item, lookup, 2, False) & delim
  Next
  vlookupInCell = Left(tempValues, Len(tempValues) - Len(delim))
End Function

#You can install this code in one of two ways

1 - Open a worksheet, click on the Developer tab, Click on Visual Basic, Find your worksheet in left menu and click on it, next click on the insert menu, then Module. Copy the VB into the newly created module
      
2 - Double click on the vlookupincell.xla file
  (see also http://www.cpearson.com/excel/createaddin.aspx for more details regarding excel add-ins)
