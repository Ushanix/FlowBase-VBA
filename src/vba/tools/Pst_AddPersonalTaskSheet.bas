Option Explicit

' ============================================
' Module   : Pst_AddPersonalTaskSheet
' Layer    : Presentation
' Purpose  : Create new personal task sheet from template
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "AddPersonalTaskSheet"

' ============================================
' AddPersonalTaskSheet
' Create new personal task sheet from template
' Reads parameters from UI_AddSheet, copies template, updates header_info
' ============================================
Public Sub AddPersonalTaskSheet()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddPersonalTaskSheet: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.ScreenUpdating = False

    ' Check UI_AddSheet exists
    If Not SheetExists(SHEET_UI_ADD_SHEET) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_UI_ADD_SHEET
        MsgBox "Sheet not found: " & SHEET_UI_ADD_SHEET, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim wsAdd As Worksheet
    Set wsAdd = ThisWorkbook.Worksheets(SHEET_UI_ADD_SHEET)

    ' Parse AddPersonalTaskSheet table
    Dim params As Object
    Set params = ParseAddPersonalTaskTable(wsAdd)

    If params.Count = 0 Then
        LogError TOOL_NAME, "Failed to parse AddPersonalTaskSheet table"
        MsgBox "Failed to parse AddPersonalTaskSheet table", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get owner_name (required)
    Dim ownerName As String
    If params.Exists("owner_name") Then
        ownerName = CStr(params("owner_name"))
    End If

    If Len(ownerName) = 0 Then
        LogError TOOL_NAME, "owner_name is empty"
        MsgBox "owner_name is empty", vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "Owner name: " & ownerName

    ' Get template name
    Dim templateName As String
    templateName = GetTemplateName()
    LogInfo TOOL_NAME, "Template: " & templateName

    ' Check template exists
    If Not SheetExists(templateName) Then
        LogError TOOL_NAME, "Template not found: " & templateName
        MsgBox "Template not found: " & templateName, vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Generate new sheet name
    Dim newSheetName As String
    newSheetName = GenerateSheetName(ownerName)

    LogInfo TOOL_NAME, "New sheet name: " & newSheetName

    ' Check sheet doesn't already exist
    If SheetExists(newSheetName) Then
        LogError TOOL_NAME, "Sheet already exists: " & newSheetName
        MsgBox "Sheet already exists: " & newSheetName, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Application.StatusBar = "Creating new personal task sheet..."

    ' Copy template sheet
    Dim newWs As Worksheet
    Set newWs = CopySheet(templateName, newSheetName)

    If newWs Is Nothing Then
        LogError TOOL_NAME, "Failed to copy template"
        MsgBox "Failed to copy template", vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "Template copied successfully"

    ' Update header_info in new sheet
    Dim updated As Long
    updated = UpdateHeaderInfo(newWs, params, ownerName)
    LogInfo TOOL_NAME, "Updated " & updated & " parameters in header_info"

    ' Clear AddPersonalTaskSheet values
    Dim cleared As Long
    cleared = ClearAddPersonalTaskTableValues(wsAdd)
    LogInfo TOOL_NAME, "Cleared " & cleared & " values in AddPersonalTaskSheet"

    ' Activate new sheet
    newWs.Activate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddPersonalTaskSheet: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Personal task sheet created successfully." & vbCrLf & _
           "New sheet: " & newSheetName, vbInformation, "Complete"

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
' ParseAddPersonalTaskTable
' Parse Tbl_Start:AddParsonalTaskSheet table
'
' Returns:
'   Dictionary with {parameter: value}
' ============================================
Private Function ParseAddPersonalTaskTable(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_PERSONAL_TASK_SHEET)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_ADD_PERSONAL_TASK_SHEET & " not found"
        ' Debug: scan for markers
        LogInfo TOOL_NAME, "Scanning for Tbl_Start markers..."
        Dim i As Long
        Dim val As Variant
        For i = 1 To 50
            val = ws.Cells(i, 1).Value
            If Not IsEmpty(val) Then
                If InStr(1, CStr(val), "Tbl_Start", vbTextCompare) > 0 Then
                    LogDebug TOOL_NAME, "  Row " & i & ": " & CStr(val)
                End If
            End If
        Next i
        Set ParseAddPersonalTaskTable = dict
        Exit Function
    End If

    Set dict = ReadKeyValueTable(ws, markerRow + 1)
    LogInfo TOOL_NAME, "Parsed AddPersonalTaskSheet: " & dict.Count & " parameters"

    Set ParseAddPersonalTaskTable = dict
End Function

' ============================================
' GetTemplateName
' Get template sheet name from DEF_Parameter or use default
'
' Returns:
'   Template sheet name
' ============================================
Private Function GetTemplateName() As String
    GetTemplateName = DEFAULT_PERSONAL_TEMPLATE

    If Not SheetExists(SHEET_DEF_PARAMETER) Then
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_PARAMETER, "name", "value", PARAM_PERSONAL_TEMPLATE)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        GetTemplateName = CStr(result)
    End If
End Function

' ============================================
' GenerateSheetName
' Generate new sheet name
'
' Args:
'   ownerName: Owner name
'
' Returns:
'   Sheet name like "PT-Ushas"
' ============================================
Private Function GenerateSheetName(ownerName As String) As String
    GenerateSheetName = PREFIX_PERSONAL & ownerName
End Function

' ============================================
' GenerateSummary
' Generate default summary text
'
' Args:
'   ownerName: Owner name
'
' Returns:
'   Summary text
' ============================================
Private Function GenerateSummary(ownerName As String) As String
    GenerateSummary = ownerName & "の個人タスク管理シート"
End Function

' ============================================
' UpdateHeaderInfo
' Update header_info table in new sheet
'
' Returns:
'   Number of parameters updated
' ============================================
Private Function UpdateHeaderInfo(ws As Worksheet, params As Object, ownerName As String) As Long
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:header_info not found in " & ws.Name
        UpdateHeaderInfo = 0
        Exit Function
    End If

    Dim headerRow As Long
    headerRow = markerRow + 1

    Dim updated As Long
    updated = 0

    ' Set fixed values: sheet_role
    If UpdateKeyValueTable(ws, headerRow, "sheet_role", FIXED_SHEET_ROLE) Then
        LogInfo TOOL_NAME, "  Updated: sheet_role = " & FIXED_SHEET_ROLE
        updated = updated + 1
    End If

    ' Set owner_name
    If UpdateKeyValueTable(ws, headerRow, "owner_name", ownerName) Then
        LogInfo TOOL_NAME, "  Updated: owner_name = " & ownerName
        updated = updated + 1
    End If

    ' Set summary if not provided
    Dim summaryValue As String
    If params.Exists("summary") And Len(CStr(params("summary"))) > 0 Then
        summaryValue = CStr(params("summary"))
    Else
        summaryValue = GenerateSummary(ownerName)
    End If

    If UpdateKeyValueTable(ws, headerRow, "summary", summaryValue) Then
        LogInfo TOOL_NAME, "  Updated: summary = " & summaryValue
        updated = updated + 1
    End If

    ' Update other parameters from input
    Dim key As Variant
    For Each key In params.Keys
        Dim keyStr As String
        keyStr = CStr(key)

        ' Skip already handled
        If keyStr = "owner_name" Or keyStr = "sheet_role" Or keyStr = "summary" Then
            GoTo NextParam
        End If

        Dim val As Variant
        val = params(key)

        If UpdateKeyValueTable(ws, headerRow, keyStr, val) Then
            LogInfo TOOL_NAME, "  Updated: " & keyStr & " = " & CStr(val)
            updated = updated + 1
        Else
            LogWarn TOOL_NAME, "  Parameter '" & keyStr & "' not found in header_info"
        End If

NextParam:
    Next key

    UpdateHeaderInfo = updated
End Function

' ============================================
' ClearAddPersonalTaskTableValues
' Clear Value column in AddPersonalTaskSheet table
'
' Returns:
'   Number of values cleared
' ============================================
Private Function ClearAddPersonalTaskTableValues(ws As Worksheet) As Long
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_PERSONAL_TASK_SHEET)

    If markerRow = 0 Then
        ClearAddPersonalTaskTableValues = 0
        Exit Function
    End If

    ClearAddPersonalTaskTableValues = ClearKeyValueTableValues(ws, markerRow + 1)
End Function
