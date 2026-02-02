Option Explicit

' ============================================
' Module   : Pst_UpdateTaskUrgent
' Layer    : Presentation
' Purpose  : Collect urgent tasks (overdue/approaching) to OUT_TaskUrgent
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "UpdateTaskUrgent"

' ============================================
' UpdateTaskUrgent
' Collect tasks with end_date <= today+3 days (not Done)
' Write to OUT_TaskUrgent sheet
' ============================================
Public Sub UpdateTaskUrgent()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateTaskUrgent: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Collecting urgent tasks..."
    Application.ScreenUpdating = False

    ' Check OUT_TaskUrgent sheet exists
    If Not SheetExists(SHEET_OUT_TASK_URGENT) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_OUT_TASK_URGENT
        MsgBox "Sheet not found: " & SHEET_OUT_TASK_URGENT, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets(SHEET_OUT_TASK_URGENT)

    Dim today As Date
    today = Date

    LogInfo TOOL_NAME, "Today: " & Format(today, "yyyy-mm-dd")
    LogInfo TOOL_NAME, "Threshold: " & URGENCY_THRESHOLD_DAYS & " days"

    ' Collect all tasks from PJ sheets
    Dim allTasks As Collection
    Set allTasks = ParsePJTasks()

    LogInfo TOOL_NAME, "Collected " & allTasks.Count & " total tasks"

    ' Filter urgent tasks
    Dim urgentTasks As Collection
    Set urgentTasks = ComputeUrgentTasks(allTasks, today)

    LogInfo TOOL_NAME, "Found " & urgentTasks.Count & " urgent tasks"

    ' Find TaskUrgent table marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(wsOut, TBL_TASK_URGENT)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_TASK_URGENT & " not found"
        MsgBox "Tbl_Start:TaskUrgent not found", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    ' Get headers
    Dim headers As Variant
    headers = GetTaskUrgentHeaders()

    Dim colCount As Long
    colCount = UBound(headers) - LBound(headers) + 1

    ' Clear existing data
    Dim cleared As Long
    cleared = ClearTableData(wsOut, headerRow, colCount)
    LogInfo TOOL_NAME, "Cleared " & cleared & " existing rows"

    ' Write data
    Dim dataRow As Long
    dataRow = headerRow + 1

    Dim rowNum As Long
    rowNum = 0

    Dim item As Object
    For Each item In urgentTasks
        rowNum = rowNum + 1
        item("no") = rowNum

        WriteTableRow wsOut, dataRow, headers, item, ""
        dataRow = dataRow + 1
    Next item

    LogInfo TOOL_NAME, "Written " & rowNum & " rows"

    ' Resize table
    If Not ResizeListObject(wsOut, headerRow, rowNum, colCount) Then
        LogWarn TOOL_NAME, "Failed to resize table, manual adjustment may be needed"
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateTaskUrgent: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Urgent task update completed." & vbCrLf & _
           rowNum & " urgent tasks found.", vbInformation, "Complete"

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
' ComputeUrgentTasks
' Filter tasks: not Done, end_date <= today + threshold
'
' Returns:
'   Collection of Dictionary in TaskUrgent format (sorted by end_date)
' ============================================
Private Function ComputeUrgentTasks(allTasks As Collection, today As Date) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim thresholdDate As Date
    thresholdDate = today + URGENCY_THRESHOLD_DAYS

    Dim task As Object
    For Each task In allTasks
        ' Check Kanban_Status != Done
        Dim status As String
        status = ""
        If task.Exists("Kanban_Status") Then
            status = CStr(task("Kanban_Status"))
        End If

        If status = KANBAN_DONE Then
            GoTo NextTask
        End If

        ' Check end_date
        Dim endDate As Variant
        endDate = Empty

        If task.Exists("end_date") Then
            endDate = task("end_date")
        End If

        ' Parse end_date
        Dim parsedDate As Date
        parsedDate = ParseDate(endDate)

        If parsedDate = 0 Then
            ' No valid date, skip
            GoTo NextTask
        End If

        ' Check if overdue or approaching
        If parsedDate > thresholdDate Then
            GoTo NextTask
        End If

        ' Calculate days remaining
        Dim daysRemaining As Long
        daysRemaining = DateDiff("d", today, parsedDate)

        ' Convert to TaskUrgent format
        Dim urgentTask As Object
        Set urgentTask = CreateObject("Scripting.Dictionary")

        urgentTask("src_project_id") = SanitizeValue(task("_project_id"))
        urgentTask("src_sheet_name") = SanitizeValue(task("_sheet_name"))
        urgentTask("task_id") = SanitizeValue(task("task_id"))
        urgentTask("task_name") = SanitizeValue(task("task_name"))
        urgentTask("summary") = SanitizeValue(task("summary"))
        urgentTask("owner_primary") = SanitizeValue(task("owner_primary"))
        urgentTask("owner_secondary") = SanitizeValue(task("owner_secondary"))
        urgentTask("Kanban_Status") = SanitizeValue(task("Kanban_Status"))
        urgentTask("MoSCoW_Priority") = SanitizeValue(task("MoSCoW_Priority"))
        urgentTask("story_point") = SanitizeValue(task("story_point"))
        urgentTask("DoD_check") = SanitizeValue(task("DoD_check"))
        urgentTask("start_date") = task("start_date")
        urgentTask("end_date") = task("end_date")
        urgentTask("last_update") = task("last_update")
        urgentTask("update_flag") = SanitizeValue(task("update_flag"))
        urgentTask("feature_key") = SanitizeValue(task("feature_key"))
        urgentTask("sprint_id") = SanitizeValue(task("sprint_id"))

        ' Internal sorting key
        urgentTask("_days_remaining") = daysRemaining

        result.Add urgentTask

NextTask:
    Next task

    ' Sort by days_remaining (overdue first)
    Set result = SortByDaysRemaining(result)

    LogInfo TOOL_NAME, "Found " & result.Count & " urgent tasks (overdue or within " & URGENCY_THRESHOLD_DAYS & " days)"

    Set ComputeUrgentTasks = result
End Function

' ============================================
' ParseDate
' Parse date value
'
' Returns:
'   Date value, or 0 if invalid
' ============================================
Private Function ParseDate(val As Variant) As Date
    ParseDate = 0

    If IsEmpty(val) Or IsNull(val) Then
        Exit Function
    End If

    If IsDate(val) Then
        ParseDate = CDate(val)
        Exit Function
    End If

    If VarType(val) = vbString Then
        Dim s As String
        s = CStr(val)

        ' Skip formulas
        If Left(s, 1) = "=" Then
            Exit Function
        End If

        ' Skip empty
        If Len(Trim(s)) = 0 Then
            Exit Function
        End If

        ' Try to parse
        On Error Resume Next
        ParseDate = CDate(s)
        On Error GoTo 0
    End If
End Function

' ============================================
' SortByDaysRemaining
' Sort collection by _days_remaining
'
' Returns:
'   Sorted Collection
' ============================================
Private Function SortByDaysRemaining(data As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    If data.Count = 0 Then
        Set SortByDaysRemaining = result
        Exit Function
    End If

    ' Convert to array for sorting
    Dim arr() As Variant
    ReDim arr(1 To data.Count)

    Dim i As Long
    For i = 1 To data.Count
        Set arr(i) = data(i)
    Next i

    ' Bubble sort by _days_remaining
    Dim j As Long
    Dim swapped As Boolean
    Dim temp As Object

    For i = 1 To data.Count - 1
        swapped = False
        For j = 1 To data.Count - i
            If arr(j)("_days_remaining") > arr(j + 1)("_days_remaining") Then
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

    Set SortByDaysRemaining = result
End Function

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
