# vlookupincell
Excel function to do a vlookup on multiple values within a single cell

# What this function does:

It takes the value of a cell and splits is up into pieces. Next is does a vlookup of each piece using a range you define. Finally it exports the newly lookuped up values in the same format as the original cell.

Example:
Input cell value: "thing-1,thing-2,thing-3"

Lookup Range:
thing-1  |  14
thing-2  |  21
thing-3  |  7
thing-4  |  1001

Column: 2

Deliminator: ","

=vlookupincell(cell, range, column, deliminator)

For this example the output would be: 14,21,7,1001

#You can install this code in one of two ways

1 - Open a worksheet, click on the Developer tab, Click on Visual Basic, Find your worksheet in left menu and click on it, next click on the insert menu, then Module. Copy the VB into the newly created module
      
2 - Double click on the vlookupincell.xla file
  (see also http://www.cpearson.com/excel/createaddin.aspx for more details regarding excel add-ins)
