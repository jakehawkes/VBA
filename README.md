# VBA Modules, Functions and Classes

A collection of VBA code that I have collected over the years. 

## Lookup Functions

Probably the most heavily used of all my VBA.

**ConcatenateCells()**

Assign this to a keyboard shortcut, and then merge all the selected cells into the top left of the range by concantenating all the cells into a single string seperated with spaces. Excellent when copying and pasting proposal RFPs into excel..

**TextToSingleLine()**

Assign this to a keyboard shortcut, and then for each selected cell, converts CR to space.

**MatrixLU()**

A shortcut for the MATCH() and INDEX() method. Use it to lookup a value from a table given the row and column headings.

**lastRow()** and **lastCol()**

Returns the number of the last row or column of the sheet that has data in it

**getURL()**

Extracts the plain text URL for the given cell

**FOREX()**

A shortcut for vlookup(). Pass it the name of the currency "XXX", and the range of the table that contains the data.

## Chart Automation 

Code to automatically modify the axis max/min settings of chart objects. Careful with this one, something in it may sometimes break Excel's undo function or its goal seek function. A bug was reported to Microsoft about it, and there solution was not to use this function if you want those things to work! However, I think I fixed it by commenting out the `Application.Volatile True` statement. [According to Microsoft] [ms-kb] this turns off the execution of this function whenever anything changes in the workbook, which is fine by me, and apparantly is the source of the above bugs...

[ms-kb]: https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.Volatile "Microsoft VBA Documentation"

## Lat Long Functions

Mostly not working, this was an initial attempt to do great circle calculations in excel.

## Date Time Class & Functions

Used in my timesheet template, these functions will translate between month numbers and abbreviations and vice versa.
Also some other handy date math functions that come in handy.

## Copy Template

Automated sheet formatting and data filling using a template. Probably not usefull for anyone else, or even me. It was made during the RWIS Expansion project.
The comments in the header say:

    ' Specially designed to copy the table from the "Template" tab,  into the "Output" tab.
    ' to the output of the Equipment Schedule for the AT RWIS Expansion project
    '
    ' Copies from A1 in the Template tab to A1 in the Output tab
    ' Copies from A1 down to wherever "END OF TEMPLATE" is located, in whichever column it is in
    ' Copies all formatting values and column widths. 
