Option Explicit

' ============================================
' Module   : Pst_IndexUpdate
' Layer    : Presentation
' Purpose  : Update UI_Index sheet with all sheet information
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "IndexUpdate"

' Special columns (not from header_info)
Private Const COL_NO As String = "no"
Private Const COL_SHEET_NAME As String = "sheet_name"

' ============================================
' IndexUpdate
' Collect all sheet info and write to UI_Index table
' ============================================
Public Sub IndexUpdate()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "IndexUpdate: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Updating Index..."
    Application.ScreenUpdating = False

    ' Check UI_Index sheet exists
    If Not SheetExists(SHEET_UI_INDEX) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_UI_INDEX
        MsgBox "Sheet not found: " & SHEET_UI_INDEX, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsIndex As Worksheet
    Set wsIndex = ThisWorkbook.Worksheets(SHEET_UI_INDEX)

    ' Find IndexTable marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(wsIndex, TBL_INDEX_TABLE)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:IndexTable not found in " & SHEET_UI_INDEX
        MsgBox "Tbl_Start:IndexTable not found in " & SHEET_UI_INDEX, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Get index table headers
    Dim headers As Variant
    headers = GetTableHeaders(wsIndex, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    LogInfo TOOL_NAME, "Index headers: " & Join(headers, ", ")

    ' Collect all sheet info
    Dim sheetsData As Collection
    Set sheetsData = ParseAllSheets(headers)

    LogInfo TOOL_NAME, "Found " & sheetsData.Count & " sheets"

    ' Sort by sheet_name
    Dim sortedData As Collection
    Set sortedData = SortBySheetName(sheetsData)

    ' Clear existing data
    Dim cleared As Long
    cleared = ClearTableData(wsIndex, headerRow, colCount)
    LogInfo TOOL_NAME, "Cleared " & cleared & " existing rows"

    ' Write data
    Dim dataRow As Long
    dataRow = headerRow + 1

    Dim rowNum As Long
    rowNum = 0

    Dim item As Object
    For Each item In sortedData
        rowNum = rowNum + 1
        item(COL_NO) = rowNum

        WriteTableRow wsIndex, dataRow, headers, item, COL_SHEET_NAME
        dataRow = dataRow + 1
    Next item

    LogInfo TOOL_NAME, "Written " & rowNum & " rows"

    ' Resize table
    If Not ResizeListObject(wsIndex, headerRow, rowNum, colCount) Then
        LogWarn TOOL_NAME, "Failed to resize table, manual adjustment may be needed"
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "IndexUpdate: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Index update completed." & vbCrLf & _
           rowNum & " sheets indexed.", vbInformation, "Complete"

    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' ParseAllSheets
' Collect info from all sheets (except UI_Index)
'
' Args:
'   targetColumns: Array of column names to collect
'
' Returns:
'   Collection of Dictionary with sheet data
' ============================================
Private Function ParseAllSheets(targetColumns As Variant) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Skip UI_Index itself
        If ws.Name = SHEET_UI_INDEX Then
            GoTo NextSheet
        End If

        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        dict(COL_SHEET_NAME) = ws.Name

        ' Try to read header_info
        Dim markerRow As Long
        markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

        If markerRow > 0 Then
            ' Read header_info table
            Dim headerInfo As Object
            Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

            LogDebug TOOL_NAME, "  " & ws.Name & " header_info keys: " & Join(GetDictKeys(headerInfo), ", ")

            ' Copy target columns
            Dim i As Long
            Dim colName As String

            For i = LBound(targetColumns) To UBound(targetColumns)
                colName = targetColumns(i)

                ' Skip special columns
                If colName = COL_NO Or colName = COL_SHEET_NAME Then
                    GoTo NextColumn
                End If

                If headerInfo.Exists(colName) Then
                    Dim val As Variant
                    val = headerInfo(colName)

                    ' Skip formula values (string starting with =)
                    If Not IsEmpty(val) Then
                        If VarType(val) = vbString Then
                            If Left(CStr(val), 1) = "=" Then
                                val = Empty
                            End If
                        End If
                    End If

                    dict(colName) = val
                End If

NextColumn:
            Next i
        End If

        result.Add dict
        LogInfo TOOL_NAME, "Collected: " & ws.Name

NextSheet:
    Next ws

    Set ParseAllSheets = result
End Function

' ============================================
' SortBySheetName
' Sort collection by sheet_name
'
' Args:
'   data: Collection of Dictionary
'
' Returns:
'   Sorted Collection
' ============================================
Private Function SortBySheetName(data As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    If data.Count = 0 Then
        Set SortBySheetName = result
        Exit Function
    End If

    ' Convert to array for sorting
    Dim arr() As Variant
    ReDim arr(1 To data.Count)

    Dim i As Long
    For i = 1 To data.Count
        Set arr(i) = data(i)
    Next i

    ' Bubble sort by sheet_name
    Dim j As Long
    Dim swapped As Boolean
    Dim temp As Object

    For i = 1 To data.Count - 1
        swapped = False
        For j = 1 To data.Count - i
            If CStr(arr(j)(COL_SHEET_NAME)) > CStr(arr(j + 1)(COL_SHEET_NAME)) Then
                Set temp = arr(j)
                Set arr(j) = arr(j + 1)
                Set arr(j + 1) = temp
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i

    ' Convert back to collection
    For i = 1 To data.Count
        result.Add arr(i)
    Next i

    Set SortBySheetName = result
End Function

' ============================================
' GetDictKeys
' Get all keys from Dictionary as array
'
' Args:
'   dict: Dictionary object
'
' Returns:
'   Array of keys
' ============================================
Private Function GetDictKeys(dict As Object) As Variant
    Dim keys() As String
    Dim keyCount As Long
    keyCount = dict.Count

    If keyCount = 0 Then
        GetDictKeys = Array("")
        Exit Function
    End If

    ReDim keys(0 To keyCount - 1)

    Dim i As Long
    Dim key As Variant
    i = 0
    For Each key In dict.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next key

    GetDictKeys = keys
End Function
