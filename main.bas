Sub macrotemplate()

' SHEET VARIABLES

activeRowCount = ActiveSheet.UsedRange.Rows.Count
activeColumnCount = ActiveSheet.UsedRange.Columns.Count

'''''''''''''''''''''''''''''''''''''''''''''
' INPUT CONFIGURATION

headerRow = 1 ' <-- Row containing column titles.
startingRow = 2 ' <-- Row where actual data begins.

lastRow = activeRowCount ' <-- Last row of data. Leave to default "activeRowCount" if no unnecessary text at bottom of report.

' REFERENCE CONFIGURATION


' OUTPUT CONFIGRATION


'''''''''''''''''''''''''''''''''''''''''''''

' INPUT & DATA READING

Dim sheetHeaders As Variant
Dim sheetData As Variant

sheetHeaders = returnLine(headerRow, activeColumnCount)
Call arrayPrint(sheetHeaders, LBound(sheetHeaders), UBound(sheetHeaders))

Dim row As New Line
row.DbgPrint

sheetData = buildDictionary(startingRow, lastRow, sheetHeaders)

' READING DEBUG

''''''''''''''''''''''''''''''''''''''''''''''

' DATA MANIPULATION

End Sub