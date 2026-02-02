Option Explicit

' ============================================
' Module   : Pst_UpdatePersonalTask
' Layer    : Presentation
' Purpose  : Collect Doing tasks for owner and update PersonalTask table
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "UpdatePersonalTask"

' ============================================
' UpdatePersonalTask
' Collect tasks with Kanban_Status=Doing for current PT sheet owner
' Must be run from a PT-* sheet
' ============================================
Public Sub UpdatePersonalTask()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdatePersonalTask: Started"
    LogInfo TOOL_NAME, "========================================"

    ' Check current sheet is PT-*
    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_PERSONAL)) <> PREFIX_PERSONAL Then
        MsgBox "Please run from PT-* sheet.", vbExclamation, "Error"
        LogError TOOL_NAME, "Not a PT sheet: " & currentSheet
        Exit Sub
    End If

    Application.StatusBar = "Collecting tasks..."
    Application.ScreenUpdating = False

    Dim wsPT As Worksheet
    Set wsPT = ThisWorkbook.Worksheets(currentSheet)

    ' Get owner_name from header_info
    Dim ownerName As String
    ownerName = ParseOwnerName(wsPT)

    LogInfo TOOL_NAME, "Owner name: " & ownerName

    If Len(ownerName) = 0 Then
        LogWarn TOOL_NAME, "owner_name is empty, will clear existing data"
    End If

    ' Collect all tasks from PJ sheets
    Dim allTasks As Collection
    Set allTasks = ParsePJTasks()

    LogInfo TOOL_NAME, "Collected " & allTasks.Count & " total tasks"

    ' Filter by owner and Kanban_Status=Doing
    Dim personalTasks As Collection
    Set personalTasks = ComputePersonalTasks(allTasks, ownerName)

    LogInfo TOOL_NAME, "Filtered " & personalTasks.Count & " tasks for owner"

    ' Find PersonalTask table marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(wsPT, TBL_PERSONAL_TASK)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_PERSONAL_TASK & " not found in " & currentSheet
        MsgBox "Tbl_Start:PersonalTask not found", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Get headers
    Dim headers As Variant
    headers = GetPersonalTaskHeaders()

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Clear existing data
    Dim cleared As Long
    cleared = ClearTableData(wsPT, headerRow, colCount)
    LogInfo TOOL_NAME, "Cleared " & cleared & " existing rows"

    ' Write data
    Dim dataRow As Long
    dataRow = headerRow + 1

    Dim rowNum As Long
    rowNum = 0

    Dim item As Object
    For Each item In personalTasks
        rowNum = rowNum + 1
        item("no") = rowNum

        WriteTableRow wsPT, dataRow, headers, item, ""
        dataRow = dataRow + 1
    Next item

    LogInfo TOOL_NAME, "Written " & rowNum & " rows"

    ' Resize table
    If Not ResizeListObject(wsPT, headerRow, rowNum, colCount) Then
        LogWarn TOOL_NAME, "Failed to resize table, manual adjustment may be needed"
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdatePersonalTask: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Personal task update completed." & vbCrLf & _
           rowNum & " tasks collected.", vbInformation, "Complete"

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
' ParseOwnerName
' Get owner_name from header_info
'
' Returns:
'   owner_name or ""
' ============================================
Private Function ParseOwnerName(ws As Worksheet) As String
    ParseOwnerName = ""

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:header_info not found in " & ws.Name
        Exit Function
    End If

    Dim headerInfo As Object
    Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

    If headerInfo.Exists("owner_name") Then
        Dim val As Variant
        val = headerInfo("owner_name")
        If Not IsEmpty(val) Then
            ParseOwnerName = CStr(val)
        End If
    End If
End Function

' ============================================
' ParsePJTasks
' Collect all tasks from PJ sheets
'
' Returns:
'   Collection of Dictionary with task data
' ============================================
Private Function ParsePJTasks() As Collection
    Dim result As Collection
    Set result = New Collection

    ' Get PJ sheets (exclude templates)
    Dim pjSheets As Collection
    Set pjSheets = FilterSheetsByPrefix(PREFIX_PROJECT)

    Dim sheetName As Variant
    For Each sheetName In pjSheets
        ' Skip templates
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE_PROJECT)) = PREFIX_TEMPLATE_PROJECT Then
            GoTo NextSheet
        End If

        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName))

        ' Get project_id from header_info
        Dim projectId As String
        projectId = CStr(sheetName)

        Dim headerMarker As Long
        headerMarker = FindTblStartRow(ws, TBL_HEADER_INFO)

        If headerMarker > 0 Then
            Dim headerInfo As Object
            Set headerInfo = ReadKeyValueTable(ws, headerMarker + 1)

            If headerInfo.Exists("project_id") Then
                Dim pid As Variant
                pid = headerInfo("project_id")
                If Not IsEmpty(pid) Then
                    projectId = CStr(pid)
                End If
            End If
        End If

        ' Find TaskList marker
        Dim taskMarker As Long
        taskMarker = FindTblStartRow(ws, TBL_TASK_LIST)

        If taskMarker = 0 Then
            GoTo NextSheet
        End If

        ' Read tasks
        Dim tableData As Variant
        tableData = ReadTableData(ws, taskMarker + 1)

        Dim headers As Variant
        headers = tableData(0)

        Dim rows As Collection
        Set rows = tableData(1)

        Dim row As Object
        For Each row In rows
            ' Add source info
            row("_sheet_name") = CStr(sheetName)
            row("_project_id") = projectId

            result.Add row
        Next row

NextSheet:
    Next sheetName

    LogInfo TOOL_NAME, "Collected " & result.Count & " tasks from " & pjSheets.Count & " PJ sheets"

    Set ParsePJTasks = result
End Function

' ============================================
' ComputePersonalTasks
' Filter tasks by owner and Kanban_Status=Doing
'
' Returns:
'   Collection of Dictionary in PersonalTask format
' ============================================
Private Function ComputePersonalTasks(allTasks As Collection, ownerName As String) As Collection
    Dim result As Collection
    Set result = New Collection

    If Len(ownerName) = 0 Then
        Set ComputePersonalTasks = result
        Exit Function
    End If

    Dim ownerLower As String
    ownerLower = LCase(ownerName)

    Dim task As Object
    For Each task In allTasks
        ' Check Kanban_Status = Doing
        Dim status As String
        status = ""
        If task.Exists("Kanban_Status") Then
            status = CStr(task("Kanban_Status"))
        End If

        If status <> KANBAN_DOING Then
            GoTo NextTask
        End If

        ' Check owner (owner_primary or owner_secondary contains owner_name)
        Dim ownerPrimary As String, ownerSecondary As String
        ownerPrimary = ""
        ownerSecondary = ""

        If task.Exists("owner_primary") Then
            ownerPrimary = LCase(CStr(task("owner_primary")))
        End If

        If task.Exists("owner_secondary") Then
            ownerSecondary = LCase(CStr(task("owner_secondary")))
        End If

        If InStr(1, ownerPrimary, ownerLower, vbTextCompare) = 0 And _
           InStr(1, ownerSecondary, ownerLower, vbTextCompare) = 0 Then
            GoTo NextTask
        End If

        ' Convert to PersonalTask format
        Dim personalTask As Object
        Set personalTask = CreateObject("Scripting.Dictionary")

        personalTask("src_project_id") = SanitizeValue(task("_project_id"))
        personalTask("src_sheet_name") = SanitizeValue(task("_sheet_name"))

        ' Map fields
        MapTaskField personalTask, task, "task_id"
        MapTaskField personalTask, task, "task_name"
        MapTaskField personalTask, task, "description"
        MapTaskField personalTask, task, "owner_primary"
        MapTaskField personalTask, task, "owner_secondary"
        MapTaskField personalTask, task, "Kanban_Status"
        MapTaskField personalTask, task, "MoSCoW_Priority"
        MapTaskField personalTask, task, "story_point"
        MapTaskField personalTask, task, "DoD_check"
        MapTaskField personalTask, task, "start_date"
        MapTaskField personalTask, task, "end_date"
        MapTaskField personalTask, task, "last_update"
        MapTaskField personalTask, task, "update_flag"
        MapTaskField personalTask, task, "feature_key"
        MapTaskField personalTask, task, "sprint_id"

        result.Add personalTask

NextTask:
    Next task

    LogInfo TOOL_NAME, "Filtered " & result.Count & " tasks for owner: " & ownerName

    Set ComputePersonalTasks = result
End Function

' ============================================
' MapTaskField
' Copy field from source to target, sanitizing value
' ============================================
Private Sub MapTaskField(target As Object, source As Object, fieldName As String)
    If source.Exists(fieldName) Then
        target(fieldName) = SanitizeValue(source(fieldName))
    Else
        target(fieldName) = ""
    End If
End Sub

' ============================================
' SanitizeValue
' Convert value to cell-writable format
' ============================================
Private Function SanitizeValue(val As Variant) As Variant
    If IsEmpty(val) Or IsNull(val) Then
        SanitizeValue = ""
        Exit Function
    End If

    ' Skip formulas
    If VarType(val) = vbString Then
        If Left(CStr(val), 1) = "=" Then
            SanitizeValue = ""
            Exit Function
        End If
    End If

    SanitizeValue = val
End Function
