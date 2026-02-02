Option Explicit

' ============================================
' Module   : Pst_ProjectIndexUpdate
' Layer    : Presentation
' Purpose  : Update UI_ProjectIndex with PJ sheet header info
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "ProjectIndexUpdate"

' Special columns
Private Const COL_NO As String = "no"
Private Const COL_SHEET_NAME As String = "sheet_name"

' ============================================
' ProjectIndexUpdate
' Collect header info from PJ sheets and write to project_index table
' ============================================
Public Sub ProjectIndexUpdate()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "ProjectIndexUpdate: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Updating Project Index..."
    Application.ScreenUpdating = False

    ' Check UI_ProjectIndex sheet exists
    If Not SheetExists(SHEET_UI_PROJECT_INDEX) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_UI_PROJECT_INDEX
        MsgBox "Sheet not found: " & SHEET_UI_PROJECT_INDEX, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsIndex As Worksheet
    Set wsIndex = ThisWorkbook.Worksheets(SHEET_UI_PROJECT_INDEX)

    ' Find project_index table
    Dim markerRow As Long
    markerRow = FindTblStartRow(wsIndex, TBL_PROJECT_INDEX)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_PROJECT_INDEX & " not found"
        MsgBox "Tbl_Start:" & TBL_PROJECT_INDEX & " not found", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Get table headers
    Dim headers As Variant
    headers = GetTableHeaders(wsIndex, headerRow)

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    LogInfo TOOL_NAME, "Table headers: " & Join(headers, ", ")

    ' Collect from all PJ sheets
    Dim projects As Collection
    Set projects = CollectAllProjects(headers)

    LogInfo TOOL_NAME, "Found " & projects.Count & " projects"

    ' Sort by project_id
    Dim sortedProjects As Collection
    Set sortedProjects = SortByProjectId(projects)

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
    For Each item In sortedProjects
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
    LogInfo TOOL_NAME, "ProjectIndexUpdate: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Project index update completed." & vbCrLf & _
           rowNum & " projects indexed.", vbInformation, "Complete"

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
' CollectAllProjects
' Collect header info from all PJ sheets
'
' Args:
'   targetColumns: Array of column names to collect
'
' Returns:
'   Collection of Dictionary with project data
' ============================================
Private Function CollectAllProjects(targetColumns As Variant) As Collection
    Dim result As Collection
    Set result = New Collection

    ' Get PJ sheets (exclude templates)
    Dim pjSheets As Collection
    Set pjSheets = FilterSheetsByPrefix(PREFIX_PROJECT)

    LogInfo TOOL_NAME, "Found " & pjSheets.Count & " sheets with prefix '" & PREFIX_PROJECT & "'"

    Dim sheetName As Variant
    For Each sheetName In pjSheets
        ' Skip templates
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE_PROJECT)) = PREFIX_TEMPLATE_PROJECT Then
            GoTo NextSheet
        End If

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))

        ' Find header_info marker
        Dim markerRow As Long
        markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

        If markerRow = 0 Then
            LogWarn TOOL_NAME, "Tbl_Start:header_info not found in " & sheetName
            GoTo NextSheet
        End If

        ' Read header_info
        Dim headerInfo As Object
        Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

        LogDebug TOOL_NAME, "  " & sheetName & " header_info keys: " & headerInfo.Count

        ' Create project data dictionary
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")

        dict(COL_SHEET_NAME) = CStr(sheetName)

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
                dict(colName) = headerInfo(colName)
            End If

NextColumn:
        Next i

        result.Add dict
        LogInfo TOOL_NAME, "Collected: " & sheetName

NextSheet:
    Next sheetName

    Set CollectAllProjects = result
End Function

' ============================================
' SortByProjectId
' Sort collection by project_id
'
' Args:
'   data: Collection of Dictionary
'
' Returns:
'   Sorted Collection
' ============================================
Private Function SortByProjectId(data As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    If data.Count = 0 Then
        Set SortByProjectId = result
        Exit Function
    End If

    ' Convert to array for sorting
    Dim arr() As Variant
    ReDim arr(1 To data.Count)

    Dim i As Long
    For i = 1 To data.Count
        Set arr(i) = data(i)
    Next i

    ' Bubble sort by project_id
    Dim j As Long
    Dim swapped As Boolean
    Dim temp As Object

    For i = 1 To data.Count - 1
        swapped = False
        For j = 1 To data.Count - i
            Dim pid1 As String, pid2 As String

            If arr(j).Exists("project_id") Then
                pid1 = CStr(arr(j)("project_id"))
            Else
                pid1 = ""
            End If

            If arr(j + 1).Exists("project_id") Then
                pid2 = CStr(arr(j + 1)("project_id"))
            Else
                pid2 = ""
            End If

            If pid1 > pid2 Then
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

    Set SortByProjectId = result
End Function
