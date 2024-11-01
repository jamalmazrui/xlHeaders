# xlHeaders

Beta version 0.5
Copyright 2024 by Jamal Mazrui
MIT License

xlHeaders.vbs is a Windows command-line program for adding metadata to an Excel workbook (.xlsx file extension) in order that screen reader software automatically identifies column or row headers, as appropriate, when navigating through regions of tabular data. The program is written in the VBScript language (Visual Basic, Scripting Edition), which has built-in Windows support.

To reach the Windows command line, you can press the Windows+R key to invoke the Windows Run dialog, and then type "cmd" followed by Enter. The syntax for running the program is as follows:

`cscript.exe /nologo <FileName.xlsx>`

A single command-line parameter is passed to the program, specifying the name of the Excel file to process. For convenience, the batch file `run.bat` may be executed instead. It assumes an original file name of `source.xlsx`, which is copied to `target.xlsx` before processing that target file. This makes it easy to view changes in the target file without replacing the source file.

The processing of an Excel file involves the following steps:

* Iterate through each worksheet of the workbook.

* For each sheet, iterate through each region. Excel considers a region to be a rectangular area of cells surrounded by worksheet borders, empty columns, or empty rows.

* For a region to be considered tabular data, the program checks that it has at least two columns, at least two rows, and either a set of column headers or row headers.

* Column headers are assumed to exist if each cell of the topmost row of the region has a non-empty, unique value. Similarly, row headers exist if each cell of the leftmost column has a non-empty, unique value. Also, the number of columns in the top row should be the same as the maximum number of columns in the region as a whole. Likewise, the number of rows in the leftmost column should be the same as the maximum number of rows in the region as a whole.

* The program ensures that the left, top cell of each tabular region has a defined name in the workbook. The name starts with "ColumnTitle", "RowTitle", or "Title", depending on whether the region has column headers, row headers, or both. A two-digit numeric suffix completes the title, e.g., "01", "02", etc., ensuring that each title is unique in the workbook. If a tabular region already has such a defined name, no additional name is created.

After the program completes, the processed file may be opened for inspection. When using either the JAWS or NVDA screen reader, an associated column header, or row header, should be spoken as one navigates through cells in a tabular region. With JAWS, a setting should be turned off, which causes a different method of identifying headers is used instead. Press Control+F3 to review or edit the set of defined names in the workbook.

