Option Explicit

' ============================================
' Module   : Utl_Table
' Layer    : Common / Utility
' Purpose  : Table operations (Tbl_Start marker search, read/write)
' Version  : 1.0.0
' Created  : 2026-02-02
' ============================================

' ============================================
' FindTblStartRow
' Search for Tbl_Start:<markerName> in column A
'
' Args:
'   ws: Target worksheet
'   markerName: Marker name (e.g., "header_info", "TaskList")
'   maxRows: Maximum rows to search (default 100)
'
' Returns:
'   Row number where marker found, or 0 if not found
' ============================================
Public Function FindTblStartRow(ws As Worksheet, _
                                 markerName As String, _
                                 Optional maxRows As Long = 100) As Long
    Dim searchText As String
    Dim i As Long
    Dim cellValue As Variant

    searchText = "Tbl_Start:" & markerName
    FindTblStartRow = 0

    For i = 1 To maxRows
        cellValue = ws.Cells(i, 1).Value
        If Not IsEmpty(cellValue) Then
            If InStr(1, CStr(cellValue), searchText, vbTextCompare) > 0 Then
                FindTblStartRow = i
                Exit Function
            End If
        End If
    Next i
End Function

' ============================================
' ReadKeyValueTable
' Read key-value table starting from specified header row
' Expects: Column A = Key, Column B = Value
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of table header
'   maxRows: Maximum data rows to read (default 100)
'
' Returns:
'   Dictionary with {key: value} pairs
' ============================================
Public Function ReadKeyValueTable(ws As Worksheet, _
                                   headerRow As Long, _
                                   Optional maxRows As Long = 100) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim keyVal As Variant
    Dim valueVal As Variant

    ' Data starts from headerRow + 1
    For i = headerRow + 1 To headerRow + maxRows
        keyVal = ws.Cells(i, 1).Value

        ' Stop at empty key
        If IsEmpty(keyVal) Or Trim(CStr(keyVal)) = "" Then
            Exit For
        End If

        valueVal = ws.Cells(i, 2).Value

        ' Skip formula values (return empty string)
        If IsFormula(ws.Cells(i, 2)) Then
            ' Read calculated value instead
            valueVal = ws.Cells(i, 2).Value
        End If

        dict(CStr(keyVal)) = valueVal
    Next i

    Set ReadKeyValueTable = dict
End Function

' ============================================
' ReadTableData
' Read table data with headers
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of table header
'   maxRows: Maximum data rows to read (default 500)
'   maxCols: Maximum columns to read (default 50)
'
' Returns:
'   Array(0) = Headers array (1-based)
'   Array(1) = Collection of Dictionary rows
' ============================================
Public Function ReadTableData(ws As Worksheet, _
                               headerRow As Long, _
                               Optional maxRows As Long = 500, _
                               Optional maxCols As Long = 50) As Variant
    Dim headers() As String
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim rows As Collection
    Dim rowDict As Object
    Dim cellVal As Variant
    Dim hasData As Boolean

    ' Read headers
    colCount = 0
    ReDim headers(1 To maxCols)

    For i = 1 To maxCols
        cellVal = ws.Cells(headerRow, i).Value
        If IsEmpty(cellVal) Or Trim(CStr(cellVal)) = "" Then
            Exit For
        End If
        colCount = colCount + 1
        headers(colCount) = CStr(cellVal)
    Next i

    ' Resize headers array
    If colCount > 0 Then
        ReDim Preserve headers(1 To colCount)
    Else
        ReDim headers(1 To 1)
        headers(1) = ""
    End If

    ' Read data rows
    Set rows = New Collection

    For i = headerRow + 1 To headerRow + maxRows
        hasData = False
        Set rowDict = CreateObject("Scripting.Dictionary")

        For j = 1 To colCount
            cellVal = ws.Cells(i, j).Value

            If Not IsEmpty(cellVal) Then
                hasData = True
            End If

            rowDict(headers(j)) = cellVal
        Next j

        ' Stop at empty row
        If Not hasData Then
            Exit For
        End If

        rows.Add rowDict
    Next i

    Dim result(0 To 1) As Variant
    result(0) = headers
    Set result(1) = rows

    ReadTableData = result
End Function

' ============================================
' ClearTableData
' Clear data rows in table (preserve header)
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of table header
'   colCount: Number of columns
'   maxRows: Maximum rows to clear (default 500)
'
' Returns:
'   Number of rows cleared
' ============================================
Public Function ClearTableData(ws As Worksheet, _
                                headerRow As Long, _
                                colCount As Long, _
                                Optional maxRows As Long = 500) As Long
    Dim i As Long, j As Long
    Dim hasData As Boolean
    Dim cleared As Long

    cleared = 0

    For i = headerRow + 1 To headerRow + maxRows
        hasData = False

        For j = 1 To colCount
            If Not IsEmpty(ws.Cells(i, j).Value) Then
                hasData = True
                ws.Cells(i, j).Value = Empty
                ' Clear hyperlink if exists
                On Error Resume Next
                ws.Cells(i, j).Hyperlinks.Delete
                On Error GoTo 0
            End If
        Next j

        If Not hasData Then
            Exit For
        End If

        cleared = cleared + 1
    Next i

    ClearTableData = cleared
End Function

' ============================================
' WriteTableRow
' Write a single row to table
'
' Args:
'   ws: Target worksheet
'   rowNum: Row number to write
'   headers: Array of header names
'   data: Dictionary with {header: value} pairs
'   linkColumn: Optional column name for hyperlink
' ============================================
Public Sub WriteTableRow(ws As Worksheet, _
                         rowNum As Long, _
                         headers As Variant, _
                         data As Object, _
                         Optional linkColumn As String = "")
    Dim i As Long
    Dim headerName As String
    Dim cellValue As Variant
    Dim lb As Long, ub As Long

    lb = LBound(headers)
    ub = UBound(headers)

    For i = lb To ub
        headerName = headers(i)

        If data.Exists(headerName) Then
            cellValue = data(headerName)
        Else
            cellValue = ""
        End If

        ' Column index is i - lb + 1 (1-based)
        Dim colIdx As Long
        colIdx = i - lb + 1

        ws.Cells(rowNum, colIdx).Value = cellValue

        ' Add hyperlink for link column
        If headerName = linkColumn And Len(CStr(cellValue)) > 0 Then
            On Error Resume Next
            ws.Hyperlinks.Add _
                Anchor:=ws.Cells(rowNum, colIdx), _
                Address:="", _
                SubAddress:="'" & CStr(cellValue) & "'!A1", _
                TextToDisplay:=CStr(cellValue)
            On Error GoTo 0
        End If
    Next i
End Sub

' ============================================
' GetTableHeaders
' Get headers from table
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of header
'   maxCols: Maximum columns to read (default 50)
'
' Returns:
'   Array of header names
' ============================================
Public Function GetTableHeaders(ws As Worksheet, _
                                 headerRow As Long, _
                                 Optional maxCols As Long = 50) As Variant
    Dim headers() As String
    Dim colCount As Long
    Dim i As Long
    Dim cellVal As Variant

    colCount = 0
    ReDim headers(1 To maxCols)

    For i = 1 To maxCols
        cellVal = ws.Cells(headerRow, i).Value
        If IsEmpty(cellVal) Or Trim(CStr(cellVal)) = "" Then
            Exit For
        End If
        colCount = colCount + 1
        headers(colCount) = CStr(cellVal)
    Next i

    If colCount > 0 Then
        ReDim Preserve headers(1 To colCount)
    Else
        ReDim headers(1 To 1)
        headers(1) = ""
    End If

    GetTableHeaders = headers
End Function

' ============================================
' ResizeListObject
' Resize Excel table (ListObject) to new range
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of header
'   dataRowCount: Number of data rows
'   colCount: Number of columns
'
' Returns:
'   True if successful
' ============================================
Public Function ResizeListObject(ws As Worksheet, _
                                   headerRow As Long, _
                                   dataRowCount As Long, _
                                   colCount As Long) As Boolean
    On Error GoTo ErrHandler

    Dim lo As ListObject
    Dim newRange As Range
    Dim endRow As Long

    ' Get first ListObject in sheet
    If ws.ListObjects.Count = 0 Then
        ResizeListObject = False
        Exit Function
    End If

    Set lo = ws.ListObjects(1)

    ' Calculate new end row (header + data rows, minimum 1 data row)
    endRow = headerRow + IIf(dataRowCount > 0, dataRowCount, 1)

    ' Create new range
    Set newRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(endRow, colCount))

    ' Resize table
    lo.Resize newRange

    ResizeListObject = True
    Exit Function

ErrHandler:
    ResizeListObject = False
End Function

' ============================================
' IsFormula
' Check if cell contains a formula
' ============================================
Private Function IsFormula(rng As Range) As Boolean
    On Error Resume Next
    IsFormula = rng.HasFormula
    If Err.Number <> 0 Then
        IsFormula = False
    End If
    On Error GoTo 0
End Function

' ============================================
' GetColumnIndex
' Get column index by header name (case-insensitive)
'
' Args:
'   headers: Array of header names
'   headerName: Header name to find
'
' Returns:
'   Column index (1-based), or 0 if not found
' ============================================
Public Function GetColumnIndex(headers As Variant, headerName As String) As Long
    Dim i As Long
    Dim lb As Long, ub As Long

    lb = LBound(headers)
    ub = UBound(headers)

    For i = lb To ub
        ' Case-insensitive comparison with trimmed values
        If StrComp(Trim(CStr(headers(i))), Trim(headerName), vbTextCompare) = 0 Then
            GetColumnIndex = i - lb + 1
            Exit Function
        End If
    Next i

    GetColumnIndex = 0
End Function

' ============================================
' UpdateKeyValueTable
' Update value in key-value table
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of table header
'   keyName: Key name to update
'   newValue: New value to set
'   maxRows: Maximum rows to search (default 100)
'
' Returns:
'   True if updated, False if key not found
' ============================================
Public Function UpdateKeyValueTable(ws As Worksheet, _
                                     headerRow As Long, _
                                     keyName As String, _
                                     newValue As Variant, _
                                     Optional maxRows As Long = 100) As Boolean
    Dim i As Long
    Dim keyVal As Variant

    For i = headerRow + 1 To headerRow + maxRows
        keyVal = ws.Cells(i, 1).Value

        If IsEmpty(keyVal) Or Trim(CStr(keyVal)) = "" Then
            Exit For
        End If

        If CStr(keyVal) = keyName Then
            ws.Cells(i, 2).Value = newValue
            UpdateKeyValueTable = True
            Exit Function
        End If
    Next i

    UpdateKeyValueTable = False
End Function

' ============================================
' ClearKeyValueTableValues
' Clear all values in key-value table (preserve keys)
'
' Args:
'   ws: Target worksheet
'   headerRow: Row number of table header
'   maxRows: Maximum rows to clear (default 100)
'
' Returns:
'   Number of values cleared
' ============================================
Public Function ClearKeyValueTableValues(ws As Worksheet, _
                                          headerRow As Long, _
                                          Optional maxRows As Long = 100) As Long
    Dim i As Long
    Dim keyVal As Variant
    Dim cleared As Long

    cleared = 0

    For i = headerRow + 1 To headerRow + maxRows
        keyVal = ws.Cells(i, 1).Value

        If IsEmpty(keyVal) Or Trim(CStr(keyVal)) = "" Then
            Exit For
        End If

        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            ws.Cells(i, 2).Value = Empty
            cleared = cleared + 1
        End If
    Next i

    ClearKeyValueTableValues = cleared
End Function

' ============================================
' LookupTableValue
' Lookup value from a table by key column and value column names
' Uses Tbl_Start marker and header row to identify columns dynamically
'
' Args:
'   ws: Target worksheet
'   markerName: Tbl_Start marker name (e.g., "Parameter")
'   keyColName: Column name for key (e.g., "name")
'   valueColName: Column name for value (e.g., "value")
'   keyToFind: Key value to search for
'   maxRows: Maximum rows to search (default 100)
'
' Returns:
'   Value if found, Empty if not found
' ============================================
Public Function LookupTableValue(ws As Worksheet, _
                                  markerName As String, _
                                  keyColName As String, _
                                  valueColName As String, _
                                  keyToFind As String, _
                                  Optional maxRows As Long = 100) As Variant
    LookupTableValue = Empty

    ' Find Tbl_Start marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, markerName)

    If markerRow = 0 Then
        Exit Function
    End If

    ' Header row is next row after marker
    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Find key and value column indices
    Dim headers As Variant
    headers = GetTableHeaders(ws, headerRow)

    Dim keyColIdx As Long
    Dim valueColIdx As Long
    keyColIdx = GetColumnIndex(headers, keyColName)
    valueColIdx = GetColumnIndex(headers, valueColName)

    If keyColIdx = 0 Or valueColIdx = 0 Then
        Exit Function
    End If

    ' Search for key in data rows
    Dim i As Long
    Dim cellKey As Variant

    For i = headerRow + 1 To headerRow + maxRows
        cellKey = ws.Cells(i, keyColIdx).Value

        ' Stop at empty key
        If IsEmpty(cellKey) Or Trim(CStr(cellKey)) = "" Then
            Exit For
        End If

        ' Case-insensitive comparison with trimmed values
        If StrComp(Trim(CStr(cellKey)), Trim(keyToFind), vbTextCompare) = 0 Then
            LookupTableValue = ws.Cells(i, valueColIdx).Value
            Exit Function
        End If
    Next i
End Function

