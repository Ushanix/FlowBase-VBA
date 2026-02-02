Option Explicit

' ============================================
' Module   : Pst_AddProjectSheet
' Layer    : Presentation
' Purpose  : Create new project sheet from template
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "AddProjectSheet"

' Parameters not written to header_info
Private Const SKIP_PARAM_FINANCIAL_YEAR As String = "financial_year"

' ============================================
' AddProjectSheet
' Create new project sheet from template
' Reads parameters from UI_AddSheet, copies template, updates header_info
' ============================================
Public Sub AddProjectSheet()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddProjectSheet: Started"
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

    ' Parse AddProjectManagementSheet table
    Dim params As Object
    Set params = ParseAddProjectTable(wsAdd)

    If params.Count = 0 Then
        LogError TOOL_NAME, "Failed to parse AddProjectManagementSheet table"
        MsgBox "Failed to parse AddProjectManagementSheet table", vbExclamation, "Error"
        GoTo Cleanup
    End If

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

    ' Get project_category and lookup category_code
    Dim projectCategory As String
    projectCategory = CStr(params("project_category"))

    If Len(projectCategory) = 0 Then
        LogError TOOL_NAME, "project_category is empty"
        MsgBox "project_category is empty", vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim categoryCode As String
    categoryCode = GetCategoryCode(projectCategory)

    If Len(categoryCode) = 0 Then
        LogError TOOL_NAME, "category_code not found for: " & projectCategory
        MsgBox "category_code not found for: " & projectCategory, vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "Category: " & projectCategory & " -> " & categoryCode

    ' Get financial year
    Dim financialYear As String
    If params.Exists("financial_year") And Len(CStr(params("financial_year"))) > 0 Then
        financialYear = CStr(params("financial_year"))
    Else
        financialYear = CalculateFiscalYear()
    End If
    LogInfo TOOL_NAME, "Financial year: " & financialYear

    ' Find max SEQ
    Dim maxSeq As Long
    maxSeq = FindMaxSeq(categoryCode, financialYear)
    LogInfo TOOL_NAME, "Max SEQ: " & maxSeq

    Dim newSeq As Long
    newSeq = maxSeq + 1

    ' Generate names
    Dim newSheetName As String
    Dim newProjectId As String
    newSheetName = GenerateSheetName(categoryCode, financialYear, newSeq)
    newProjectId = newSheetName

    LogInfo TOOL_NAME, "New sheet name: " & newSheetName
    LogInfo TOOL_NAME, "New project_id: " & newProjectId

    ' Validate sheet name
    Dim validationError As String
    validationError = ValidateSheetName(newSheetName)

    If Len(validationError) > 0 Then
        LogError TOOL_NAME, "Invalid sheet name: " & validationError
        MsgBox "Invalid sheet name: " & newSheetName & vbCrLf & vbCrLf & _
               validationError & vbCrLf & vbCrLf & _
               "Please check the category_code setting.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Check sheet doesn't already exist
    If SheetExists(newSheetName) Then
        LogError TOOL_NAME, "Sheet already exists: " & newSheetName
        MsgBox "Sheet already exists: " & newSheetName, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Application.StatusBar = "Creating new project sheet..."

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
    updated = UpdateHeaderInfo(newWs, params, newProjectId)
    LogInfo TOOL_NAME, "Updated " & updated & " parameters in header_info"

    ' Clear AddProjectManagementSheet values
    Dim cleared As Long
    cleared = ClearAddProjectTableValues(wsAdd)
    LogInfo TOOL_NAME, "Cleared " & cleared & " values in AddProjectManagementSheet"

    ' Activate new sheet
    newWs.Activate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "AddProjectSheet: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Project sheet created successfully." & vbCrLf & _
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
' ParseAddProjectTable
' Parse Tbl_Start:AddProjectManagementSheet table
'
' Returns:
'   Dictionary with {parameter: value}
' ============================================
Private Function ParseAddProjectTable(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_PROJECT_SHEET)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_ADD_PROJECT_SHEET & " not found"
        Set ParseAddProjectTable = dict
        Exit Function
    End If

    Set dict = ReadKeyValueTable(ws, markerRow + 1)
    LogInfo TOOL_NAME, "Parsed AddProjectManagementSheet: " & dict.Count & " parameters"

    Set ParseAddProjectTable = dict
End Function

' ============================================
' GetTemplateName
' Get template sheet name from DEF_Parameter or use default
'
' Returns:
'   Template sheet name
' ============================================
Private Function GetTemplateName() As String
    GetTemplateName = DEFAULT_PROJECT_TEMPLATE

    If Not SheetExists(SHEET_DEF_PARAMETER) Then
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_PARAMETER, "name", "value", PARAM_PROJECT_TEMPLATE)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        GetTemplateName = CStr(result)
    End If
End Function

' ============================================
' GetCategoryCode
' Get category_code from DEF_project_category sheet
'
' Args:
'   projectCategory: Category ID (e.g., system_development)
'
' Returns:
'   Category code (e.g., DEV) or ""
' ============================================
Private Function GetCategoryCode(projectCategory As String) As String
    GetCategoryCode = ""

    If Not SheetExists(SHEET_DEF_PROJECT_CATEGORY) Then
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PROJECT_CATEGORY)

    ' Column B: category_id, Column C: category_code
    Dim row As Long
    For row = 2 To 100
        Dim categoryId As Variant
        categoryId = ws.Cells(row, 2).Value

        If IsEmpty(categoryId) Then
            Exit For
        End If

        If CStr(categoryId) = projectCategory Then
            Dim categoryCode As Variant
            categoryCode = ws.Cells(row, 3).Value
            If Not IsEmpty(categoryCode) Then
                GetCategoryCode = CStr(categoryCode)
            End If
            Exit Function
        End If
    Next row
End Function

' ============================================
' CalculateFiscalYear
' Calculate current fiscal year (April start)
'
' Returns:
'   Fiscal year string like "FY25"
' ============================================
Private Function CalculateFiscalYear() As String
    Dim yr As Long
    Dim mon As Long

    yr = Year(Now)
    mon = Month(Now)

    If mon < 4 Then
        yr = yr - 1
    End If

    CalculateFiscalYear = "FY" & Format(yr Mod 100, "00")
End Function

' ============================================
' FindMaxSeq
' Find maximum SEQ number for given category and fiscal year
'
' Args:
'   categoryCode: Category code (e.g., INFRA)
'   fy: Fiscal year (e.g., FY25)
'
' Returns:
'   Maximum SEQ found (0 if none)
' ============================================
Private Function FindMaxSeq(categoryCode As String, fy As String) As Long
    Dim pattern As String
    pattern = PREFIX_PROJECT & categoryCode & "-" & fy & "-"

    Dim maxSeq As Long
    maxSeq = 0

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, Len(pattern)) = pattern Then
            ' Extract SEQ from sheet name
            Dim seqPart As String
            seqPart = Mid(ws.Name, Len(pattern) + 1)

            ' Check if it's a number
            If IsNumeric(seqPart) Then
                Dim seq As Long
                seq = CLng(seqPart)
                If seq > maxSeq Then
                    maxSeq = seq
                End If
                LogDebug TOOL_NAME, "  Found: " & ws.Name & " (SEQ=" & seq & ")"
            End If
        End If
    Next ws

    LogInfo TOOL_NAME, "Max SEQ for " & pattern & "*: " & maxSeq
    FindMaxSeq = maxSeq
End Function

' ============================================
' GenerateSheetName
' Generate new sheet name
'
' Returns:
'   Sheet name like "PJ-INFRA-FY25-03"
' ============================================
Private Function GenerateSheetName(categoryCode As String, fy As String, seq As Long) As String
    GenerateSheetName = PREFIX_PROJECT & categoryCode & "-" & fy & "-" & Format(seq, "00")
End Function

' ============================================
' UpdateHeaderInfo
' Update header_info table in new sheet
'
' Returns:
'   Number of parameters updated
' ============================================
Private Function UpdateHeaderInfo(ws As Worksheet, params As Object, projectId As String) As Long
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

    ' Update project_id
    If UpdateKeyValueTable(ws, headerRow, "project_id", projectId) Then
        LogInfo TOOL_NAME, "  Updated: project_id = " & projectId
        updated = updated + 1
    End If

    ' Update other parameters
    Dim key As Variant
    For Each key In params.Keys
        Dim keyStr As String
        keyStr = CStr(key)

        ' Skip parameters not meant for header_info
        If keyStr = SKIP_PARAM_FINANCIAL_YEAR Then
            GoTo NextParam
        End If

        ' project_id already handled
        If keyStr = "project_id" Then
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
' ClearAddProjectTableValues
' Clear Value column in AddProjectManagementSheet table
'
' Returns:
'   Number of values cleared
' ============================================
Private Function ClearAddProjectTableValues(ws As Worksheet) As Long
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_ADD_PROJECT_SHEET)

    If markerRow = 0 Then
        ClearAddProjectTableValues = 0
        Exit Function
    End If

    ClearAddProjectTableValues = ClearKeyValueTableValues(ws, markerRow + 1)
End Function

' ============================================
' ValidateSheetName
' Validate sheet name for Excel restrictions
'
' Args:
'   sheetName: Sheet name to validate
'
' Returns:
'   Error message if invalid, empty string if valid
' ============================================
Private Function ValidateSheetName(sheetName As String) As String
    Const MAX_SHEET_NAME_LENGTH As Long = 31
    Const INVALID_CHARS As String = ":\/?*[]"

    ValidateSheetName = ""

    ' Check length
    If Len(sheetName) > MAX_SHEET_NAME_LENGTH Then
        ValidateSheetName = "Sheet name exceeds " & MAX_SHEET_NAME_LENGTH & " characters. " & _
                            "(Current: " & Len(sheetName) & " characters)"
        Exit Function
    End If

    ' Check empty
    If Len(sheetName) = 0 Then
        ValidateSheetName = "Sheet name cannot be empty."
        Exit Function
    End If

    ' Check invalid characters
    Dim i As Long
    Dim c As String
    Dim foundChars As String
    foundChars = ""

    For i = 1 To Len(INVALID_CHARS)
        c = Mid(INVALID_CHARS, i, 1)
        If InStr(sheetName, c) > 0 Then
            If Len(foundChars) > 0 Then
                foundChars = foundChars & " "
            End If
            foundChars = foundChars & c
        End If
    Next i

    If Len(foundChars) > 0 Then
        ValidateSheetName = "Sheet name contains invalid characters: " & foundChars & vbCrLf & _
                            "Characters : \ / ? * [ ] cannot be used in sheet names."
        Exit Function
    End If
End Function
