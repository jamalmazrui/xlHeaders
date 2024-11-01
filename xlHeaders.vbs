' xlHeaders
' Version 0.5
' Copyright 2024 by Access Success LLC
' MIT License

option explicit
wscript.echo "Launching XlHeaders"

dim bFound, bColumnHeaders, bRowHeaders, bHeaders, bLoop, bName, bContinue
dim dHeaders, dRegions
dim iColumn, iRow, iCount, iSequence, iRegion, iRegionCellCount
dim oBook, oSheet, oName, oTitleCell, oArea, oCell, oExcel, oRegion, oUsedCell, oShell, oSystem
dim sTitleCellAddress, sValue, sSequence, sTitle, sFileXlsx, sAddress, sRegionAddress, sTitleStem, sUsedCellAddress
dim vValue

if wscript.Arguments.Count = 0 then
wscript.echo "Pass an Excel file name as a parameter"
wscript.quit
end if 'Arguments.Count

Set oShell =CreateObject("Wscript.Shell")
set oSystem = createObject("Scripting.FileSystemObject")
sFileXlsx = wscript.Arguments(0)
if instr(sFileXlsx, "\") = 0 then sFileXlsx = oSystem.BuildPath(oShell.CurrentDirectory, sFileXlsx)

set oExcel = wscript.CreateObject("Excel.Application")
oExcel.Visible = false
oExcel.DisplayAlerts = false
set oBook = oExcel.WorkBooks.Open(sFileXlsx)

set dRegions = nothing
set dRegions = CreateObject("Scripting.Dictionary")
dRegions.CompareMode = vbTextCompare

wscript.echo "Sheets " & oBook.WorkSheets.Count
for each oSheet in oBook.WorkSheets
wscript.echo ""
wscript.echo "Sheet " & oSheet.Name
for each oUsedCell in oSheet.UsedRange.Cells
bContinue = true
sUsedCellAddress = oUsedCell.Address
set oRegion = oUsedCell.CurrentRegion
sValue = trim(oUsedCell.Value)
if len(sValue) = 0 or oRegion.Columns.Count < 2 or oRegion.Rows.Count < 2 then bContinue = false
' if oColumns.Count <> oRows(1).Cells.Count then bContinue = false
' if oRows.Count <> oColumns(1).Cells.Count then bContinue = false

if bContinue then
sRegionAddress = oRegion.Address
iRegionCellCount = oRegion.Cells.Count
sTitleCellAddress = oRegion.Cells(1,1).Address
if dRegions.Exists(sRegionAddress) then
if iRegionCellCount > oSheet.Range(sRegionAddress).Cells.Count then dRegions(sRegionAddress) = sUsedCellAddress
else
dRegions.Add sRegionAddress, sUsedCellAddress
end if 'Exists
end if 'bContinue
next 'Cells

' wscript.echo "Regions " & dRegions.Count
for each sRegionAddress in dRegions.Keys

' wscript.echo "sRegionAddress " & sRegionAddress
next

for each sRegionAddress in dRegions.Keys
sUsedCellAddress = dRegions(sRegionAddress)
set oRegion = nothing
' set oRegion = oSheet.Range(sRegionAddress)
set oRegion = oSheet.Range(sUsedCellAddress).CurrentRegion

bContinue = true
set dHeaders = nothing
set dHeaders = CreateObject("Scripting.Dictionary")
dHeaders.CompareMode = vbTextCompare

bColumnHeaders = true
' wscript.echo "Columns " & oRegion.Rows(1).Cells.Count
for each oCell in oRegion.Rows(1).Cells
vValue = oCell.Value
sValue = trim(vValue)
' wscript.echo "Value " & sValue
if len(sValue) = 0 or dHeaders.exists(sValue) then 
bColumnHeaders = false
else
dHeaders.Add sValue, ""
end if 
next '' cells
if oRegion.Columns.Count <> oRegion.Rows(1).Cells.Count then bColumnHeaders = false

set dHeaders = nothing
set dHeaders = CreateObject("Scripting.Dictionary")
dHeaders.CompareMode = vbTextCompare

bRowHeaders = true
' wscript.echo "Rows " & oRegion.Columns(1).Cells.Count
for each oCell in oRegion.Columns(1).Cells
vValue = oCell.Value
sValue = trim(vValue)
' wscript.echo "Value " & sValue
if len(sValue) = 0 or dHeaders.exists(sValue) then 
bRowHeaders = false
else
dHeaders.Add sValue, ""
end if
next 'cells
if oRegion.Rows.Count <> oRegion.Columns(1).Cells.Count then bRowHeaders = false

bHeaders = true
' wscript.echo "Column headers " & bColumnHeaders
' wscript.echo "Row headers" & bRowHeaders

if bColumnHeaders and not bRowHeaders then sTitleStem = "ColumnTitle"
if not bColumnHeaders and bRowHeaders then sTitleStem = "RowTitle"
if bColumnHeaders and bRowHeaders then sTitleStem = "Title"
if not bColumnHeaders and not bRowHeaders then
sTitleStem = ""
bHeaders = false
bContinue = false
end if

' wscript.echo "Headers " & bHeaders
' wscript.echo "Header title stem" & sTitleStem

if bContinue then
iRegion = iRegion + 1
set oTitleCell = oRegion.Cells(1, 1)
sTitleCellAddress = oTitleCell.Address
' wscript.echo "Title cell address " & sTitleCellAddress
' wscript.echo "Region " & iRegion & " " & sTitleCellAddress

bFound = false
' wscript.echo "Names " & oBook.Names.Count
for each oName in oBook.Names
sAddress = oName.RefersToRange.Address
' wscript.echo "Address " & sAddress
' if sTitleCellAddress = oName.RefersToRange.Address and instr(oRegion.Name, "Title") > 0 then bFound = true
if sTitleCellAddress = oName.RefersToRange.Address then bFound = true
next 'names
if bFound then bContinue = false
end if 'bContinue

if bContinue then
bLoop = true
iSequence = 1
do while bLoop
sSequence = Right("0" & iSequence, 2)
' wscript.echo "Sequence " & sSequence
sTitle = sTitleStem & sSequence

bFound = false
for each oName in oBook.Names
if oName.Name = sTitle then bFound = true
next 'names

if bFound then
iSequence = iSequence + 1
else
bLoop = false
end if
loop

' wscript.echo "Title header " & sTitle
wscript.echo ""
wscript.echo "Region " & iRegion & " " & sTitle & " "
wscript.echo "Columns " & oRegion.Columns.Count & ", " & "Rows " & oRegion.Rows.Count
wscript.echo "Range " & sRegionAddress  
oTitleCell.Name = sTitle
end if 'bContinue
' next
next 'Keys
next 'worksheets

wscript.echo ""
wscript.echo "Saving changes"
oBook.save
oBook.Close(true)
oExcel.Quit
