# Guide

This guide explains mainly a little of the `SpreadsheetML` standard
Because the code can be confusing because of lack of knowledge of the standard.


## Key concepts

> 🤓 An excel file is just a zip file of multiple XML documents.


[The standard explained by Microsoft](https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document?tabs=cs)

- This is the basic structure of a `SpreadsheetML`:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes" ?> 
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
            <sheet name="MySheet1" sheetId="1" r:id="rId1" /> 
            <sheet name="MySheet2" sheetId="2" r:id="rId2" /> 
        </sheets>
    </workbook>
```
-  A separate XML file is created for each `worksheet`.

- The `worksheet` XML files contain one or more block level elements such as `SheetData` represents the cell table and contains one or more `Row` elements. A `row` contains one or more `Cell` elements. 
Each `cell` contains a `CellValue` element that represents the value of the cell.

```xml
<?xml version="1.0" encoding="UTF-8" ?> 
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
            <row r="1">
                <c r="A1">
                    <v>100</v> 
                </c>
            </row>
        </sheetData>
    </worksheet>
```
