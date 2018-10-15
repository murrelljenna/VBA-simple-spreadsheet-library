''''''''''''''''''''''''''''''''''''''''
' Function Library
''''''''''''''''''''''''''''''''''''''''

' DATA READING FUNCTIONS

Function returnLine(row, activeColumnCount) As Variant
    Dim arr As Variant ' <-- Array of column headers
    
    ReDim arr(0 To activeColumnCount)
    For i = 0 To activeColumnCount
        c = i + 1
        
        arr(i) = Cells(row, c).value
    Next i

    returnLine = arr
End Function

Function buildDictionary(startingRow, lastRow, headers) As Variant
    Dim sheetData As Variant
    ReDim sheetData(1 To 1)
    
    Index = 1
    
    For R = startingRow To lastRow
    
        ReDim Preserve sheetData(1 To Index)
        Set sheetData(Index) = CreateObject("Scripting.Dictionary")
    
        For c = 1 To UBound(headers) ' <-- Assigns key:value pair for every line item on row.
            
            lineItem = Cells(R, c).value
            sheetData(Index).Add headers(c), lineItem
            
        Next c
        
        Index = Index + 1
        
    Next R ' <-- Moves on to next row.

    buildDictionary = sheetData

End Function

Function continueDictionary(Dictionary, startingRow, lastRow, headers)
    Index = (UBound(Dictionary) + 1)
    
    For R = startingRow To lastRow
    
        ReDim Preserve Dictionary(1 To Index)
        Set Dictionary(Index) = CreateObject("Scripting.Dictionary")
    
        For c = 1 To UBound(headers) ' <-- Assigns key:value pair for every line item on row.
            
            lineItem = Cells(R, c).value
            Dictionary(Index).Add headers(c), lineItem
            
        Next c
        
        Index = Index + 1
        
    Next R ' <-- Moves on to next row.

    continueDictionary = Dictionary

End Function

Function dataPrint(haystack, start, last)
    Debug.Print ("dataPrint")

    For Index = start To last
    
        On Error Resume Next
    
        For Each key In haystack(Index).keys()

            Debug.Print ("Row: " & Index & "," & "Column:" & key & "," & haystack(Index)(key))
            
        Next
    Next Index
End Function

Function arrayPrint(haystack, start, last)

    For Index = start To last

            Debug.Print (haystack(Index))
            
    Next Index

End Function

Function isDuplicate(needle, haystackArr)
    isDuplicate = False
    
    For Index = LBound(haystackArr) To UBound(haystackArr)
        If (needle = haystackArr(Index)) Then
            isDuplicate = True
            Exit For
        End If
    Next Index
    
End Function

' ADVANCED READING FUNCTIONS

Function combineArray(domArray, subArray)

    combinedLength = UBound(domArray) + UBound(subArray)
    ReDim Preserve domArray(1 To combinedLength)

    For Index = 1 To UBound(subArray)

        domArray(UBound(domArray) + Index) = subArray(Index)

    Next Index

    combineArray = domArray

End Function

Function inputPair(start, last, key, value, haystack) ' <-- Inputs individual key:value pair into dictionary.

    For Index = start To last
        
        haystack(Index).Add key, value
        
    Next Index

End Function

' DATA MANIPULATION FUNCTIONS

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Function cleaveName(x As String) As Variant
    
    x = replace(x, "-", "(")
    x = replace(x, " ", "(")
    cleaveName = Split(x, "(")

End Function

' REFERENCING FUNCTIONS

Function readFiles(path)

    Dim fileArray As Variant
    ReDim fileArray(0 To 0)
    
    Dim oFSO As Object
    Dim directory As Object
    Dim Files As Object
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set directory = oFSO.GetFolder(path)
    Set dirFiles = directory.Files
    
    ReDim fileArray(1 To dirFiles.Count)
    
    Index = 1
    For Each File In dirFiles
        fileArray(Index) = File.Name
        Index = Index + 1
    Next File

    readFiles = fileArray

End Function
