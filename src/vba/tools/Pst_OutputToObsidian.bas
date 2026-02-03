Option Explicit

' ============================================
' Module   : Pst_OutputToObsidian
' Layer    : Presentation
' Purpose  : Export project/task info to Obsidian notes
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "OutputToObsidian"
Private Const MAPPING_SHEET As String = "M_Cov_WBS-Obsidian"

' ============================================
' OutputToObsidian
' Export current PJ sheet's project info and tasks to Obsidian
' Must be run from a PJ-* sheet
' ============================================
Public Sub OutputToObsidian()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidian: Started"
    LogInfo TOOL_NAME, "========================================"

    ' Check current sheet is PJ-*
    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    If Left(currentSheet, Len(PREFIX_PROJECT)) <> PREFIX_PROJECT Then
        MsgBox "Please run from PJ-* sheet.", vbExclamation, "Error"
        LogError TOOL_NAME, "Not a PJ sheet: " & currentSheet
        Exit Sub
    End If

    ' Skip templates
    If Left(currentSheet, Len(PREFIX_TEMPLATE_PROJECT)) = PREFIX_TEMPLATE_PROJECT Then
        MsgBox "Cannot export template sheet.", vbExclamation, "Error"
        LogError TOOL_NAME, "Cannot export template: " & currentSheet
        Exit Sub
    End If

    Application.StatusBar = "Exporting to Obsidian..."
    Application.ScreenUpdating = False

    ' Execute export for single sheet
    Dim result As String
    result = OutputToObsidianSheet(currentSheet)

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidian: Completed"
    LogInfo TOOL_NAME, "========================================"

    If Left(result, 5) = "ERROR" Then
        MsgBox result, vbExclamation, "Error"
    Else
        MsgBox result, vbInformation, "Complete"
    End If

    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' OutputToObsidianAll
' Export all PJ sheets to Obsidian (for batch processing)
' Called from UpdateAll
' ============================================
Public Sub OutputToObsidianAll()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidianAll: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Exporting all PJ sheets to Obsidian..."
    Application.ScreenUpdating = False

    ' Get Obsidian base path first
    Dim basePath As String
    basePath = GetObsidianBasePath()

    If Len(basePath) = 0 Then
        LogError TOOL_NAME, PARAM_OBSIDIAN_PATH & " not configured"
        MsgBox PARAM_OBSIDIAN_PATH & " not configured", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get all PJ sheets
    Dim pjSheets As Collection
    Set pjSheets = FilterSheetsByPrefix(PREFIX_PROJECT)

    Dim successCount As Long
    Dim skipCount As Long
    Dim errorCount As Long
    successCount = 0
    skipCount = 0
    errorCount = 0

    Dim sheetName As Variant
    For Each sheetName In pjSheets
        ' Skip templates
        If Left(CStr(sheetName), Len(PREFIX_TEMPLATE_PROJECT)) = PREFIX_TEMPLATE_PROJECT Then
            skipCount = skipCount + 1
            LogInfo TOOL_NAME, "Skipped template: " & CStr(sheetName)
            GoTo NextSheet
        End If

        Application.StatusBar = "Exporting: " & CStr(sheetName) & "..."

        Dim result As String
        result = OutputToObsidianSheet(CStr(sheetName))

        If Left(result, 5) = "ERROR" Or Left(result, 4) = "SKIP" Then
            If Left(result, 4) = "SKIP" Then
                skipCount = skipCount + 1
            Else
                errorCount = errorCount + 1
            End If
        Else
            successCount = successCount + 1
        End If

NextSheet:
    Next sheetName

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "OutputToObsidianAll: Completed"
    LogInfo TOOL_NAME, "  Success: " & successCount & ", Skip: " & skipCount & ", Error: " & errorCount
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Obsidian export completed." & vbCrLf & vbCrLf & _
           "Success: " & successCount & " sheets" & vbCrLf & _
           "Skipped: " & skipCount & " sheets" & vbCrLf & _
           "Errors: " & errorCount & " sheets", vbInformation, "Complete"

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
' OutputToObsidianSheet
' Export a single PJ sheet to Obsidian (internal function)
'
' Args:
'   sheetName: Name of the PJ sheet to export
'
' Returns:
'   Result message (starts with "ERROR" or "SKIP" on failure)
' ============================================
Private Function OutputToObsidianSheet(sheetName As String) As String
    On Error GoTo EH

    LogInfo TOOL_NAME, "Processing sheet: " & sheetName

    ' Get Obsidian base path from DEF_Parameter
    Dim basePath As String
    basePath = GetObsidianBasePath()

    If Len(basePath) = 0 Then
        OutputToObsidianSheet = "ERROR: " & PARAM_OBSIDIAN_PATH & " not configured"
        Exit Function
    End If

    Dim wsPJ As Worksheet
    Set wsPJ = ThisWorkbook.Worksheets(sheetName)

    ' Get project info from header_info
    Dim projectInfo As Object
    Set projectInfo = ParseProjectInfo(wsPJ)

    If projectInfo.Count = 0 Then
        LogError TOOL_NAME, "Failed to read header_info: " & sheetName
        OutputToObsidianSheet = "ERROR: Failed to read header_info"
        Exit Function
    End If

    ' Get vault folder from project info
    Dim vaultFolder As String
    If projectInfo.Exists("obsidian_path_form_vault_folder") Then
        vaultFolder = CStr(projectInfo("obsidian_path_form_vault_folder"))
    End If

    If Len(vaultFolder) = 0 Then
        LogInfo TOOL_NAME, "Skipped (no vault folder): " & sheetName
        OutputToObsidianSheet = "SKIP: obsidian_path_form_vault_folder not set"
        Exit Function
    End If

    ' Construct output path
    Dim outputDir As String
    outputDir = BuildFilePath(basePath, vaultFolder)

    LogInfo TOOL_NAME, "Output directory: " & outputDir

    ' Load field mappings
    Dim headerMapping As Object
    Dim taskMapping As Object
    ParseFieldMapping headerMapping, taskMapping

    ' Export project note
    Dim projectWritten As Boolean
    projectWritten = ApplyProjectNote(outputDir, projectInfo, headerMapping)

    ' Get tasks
    Dim tasks As Collection
    Set tasks = ParseTasks(wsPJ)

    LogInfo TOOL_NAME, "Found " & tasks.Count & " tasks in " & sheetName

    ' Export task notes
    Dim taskWritten As Long
    taskWritten = ApplyToObsidian(outputDir, projectInfo, tasks, headerMapping, taskMapping)

    LogInfo TOOL_NAME, "Completed: " & sheetName & " (1 project, " & taskWritten & " tasks)"

    OutputToObsidianSheet = "Export completed: " & sheetName & vbCrLf & _
                            "1 project note, " & taskWritten & " task notes."
    Exit Function

EH:
    LogError TOOL_NAME, "Error in " & sheetName & ": " & Err.Description
    OutputToObsidianSheet = "ERROR: " & Err.Description
End Function

' ============================================
' GetObsidianBasePath
' Get Obsidian base path from DEF_Parameter
'
' Returns:
'   Base path or ""
' ============================================
Private Function GetObsidianBasePath() As String
    GetObsidianBasePath = ""

    If Not SheetExists(SHEET_DEF_PARAMETER) Then
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_PARAMETER, "name", "value", PARAM_OBSIDIAN_PATH)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        GetObsidianBasePath = CStr(result)
    End If
End Function

' ============================================
' ParseProjectInfo
' Get project info from header_info
'
' Returns:
'   Dictionary with project info
' ============================================
Private Function ParseProjectInfo(ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:header_info not found"
        Set ParseProjectInfo = dict
        Exit Function
    End If

    Set dict = ReadKeyValueTable(ws, markerRow + 1)
    LogInfo TOOL_NAME, "Loaded project info: " & dict.Count & " fields"

    Set ParseProjectInfo = dict
End Function

' ============================================
' ParseTasks
' Get tasks from TaskList
'
' Returns:
'   Collection of Dictionary with task data
' ============================================
Private Function ParseTasks(ws As Worksheet) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_TASK_LIST)

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:TaskList not found"
        Set ParseTasks = result
        Exit Function
    End If

    Dim tableData As Variant
    tableData = ReadTableData(ws, markerRow + 1)

    Dim rows As Collection
    Set rows = tableData(1)

    ' Filter: only tasks with task_id
    Dim row As Object
    For Each row In rows
        If row.Exists("task_id") Then
            Dim taskId As Variant
            taskId = row("task_id")
            If Not IsEmpty(taskId) And Len(CStr(taskId)) > 0 Then
                result.Add row
            End If
        End If
    Next row

    LogInfo TOOL_NAME, "Found " & result.Count & " valid tasks"

    Set ParseTasks = result
End Function

' ============================================
' ParseFieldMapping
' Load field mappings from M_Cov_WBS-Obsidian sheet
' ============================================
Private Sub ParseFieldMapping(ByRef headerMapping As Object, ByRef taskMapping As Object)
    Set headerMapping = CreateObject("Scripting.Dictionary")
    Set taskMapping = CreateObject("Scripting.Dictionary")

    If Not SheetExists(MAPPING_SHEET) Then
        LogWarn TOOL_NAME, MAPPING_SHEET & " not found, using default mapping"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MAPPING_SHEET)

    Dim currentSection As String
    currentSection = ""

    Dim row As Long
    For row = 1 To ws.UsedRange.rows.Count
        Dim col1 As Variant, col2 As Variant
        col1 = ws.Cells(row, 1).Value
        col2 = ws.Cells(row, 2).Value

        If Not IsEmpty(col1) Then
            Dim col1Str As String
            col1Str = CStr(col1)

            ' Check for section markers
            If InStr(1, col1Str, "Tbl_Start:", vbTextCompare) > 0 Then
                If InStr(1, col1Str, "header_info", vbTextCompare) > 0 Then
                    currentSection = "header"
                ElseIf InStr(1, col1Str, "TaskList", vbTextCompare) > 0 Then
                    currentSection = "task"
                End If
                GoTo NextRow
            End If

            ' Read mapping
            If Len(currentSection) > 0 And Not IsEmpty(col2) Then
                Dim obsValue As String
                obsValue = Trim(CStr(col2))

                ' Skip "-" or "none"
                If obsValue = "-" Or LCase(obsValue) = "none" Then
                    GoTo NextRow
                End If

                ' Skip header rows
                If col1Str = "ProjectManager" Or obsValue = "Obsidian" Then
                    GoTo NextRow
                End If

                If currentSection = "header" Then
                    headerMapping(col1Str) = obsValue
                ElseIf currentSection = "task" Then
                    taskMapping(col1Str) = obsValue
                End If
            End If
        End If

NextRow:
    Next row

    LogInfo TOOL_NAME, "Loaded mapping: header=" & headerMapping.Count & ", task=" & taskMapping.Count
End Sub

' ============================================
' ApplyProjectNote
' Write project note to Obsidian
'
' Returns:
'   True if successful
' ============================================
Private Function ApplyProjectNote(outputDir As String, projectInfo As Object, headerMapping As Object) As Boolean
    ApplyProjectNote = False

    Dim projectId As String
    If projectInfo.Exists("project_id") Then
        projectId = CStr(projectInfo("project_id"))
    End If

    If Len(projectId) = 0 Then
        LogWarn TOOL_NAME, "project_id not found, skipping project note"
        Exit Function
    End If

    ' Create output directory
    If Not CreateFolder(outputDir) Then
        LogError TOOL_NAME, "Failed to create folder: " & outputDir
        Exit Function
    End If

    ' Generate frontmatter
    Dim frontmatter As String
    frontmatter = GenerateProjectFrontmatter(projectInfo, headerMapping)

    ' Build filename
    Dim projectName As String
    If projectInfo.Exists("project_name") Then
        projectName = CStr(projectInfo("project_name"))
    End If

    Dim filenameBase As String
    If Len(projectName) > 0 Then
        filenameBase = projectId & "_" & projectName
    Else
        filenameBase = projectId
    End If

    Dim filename As String
    filename = SanitizeFilename(filenameBase) & ".md"

    Dim filepath As String
    filepath = BuildFilePath(outputDir, filename)

    ' Read existing content or create new
    Dim content As String
    If FileExists(filepath) Then
        Dim existing As String
        existing = ReadTextFile(filepath)
        Dim body As String
        body = ExtractBodyAfterFrontmatter(existing)
        content = frontmatter & body
    Else
        Dim summary As String
        If projectInfo.Exists("summary") Then
            summary = CStr(projectInfo("summary"))
        End If

        content = frontmatter
        If Len(projectName) > 0 Then
            content = content & "# " & projectName & vbLf & vbLf
        End If
        If Len(summary) > 0 Then
            content = content & summary & vbLf
        End If
    End If

    ' Write file
    If WriteTextFile(filepath, content, False) Then
        LogInfo TOOL_NAME, "Written project note: " & filepath
        ApplyProjectNote = True
    Else
        LogError TOOL_NAME, "Failed to write: " & filepath
    End If
End Function

' ============================================
' ApplyToObsidian
' Write task notes to Obsidian
'
' Returns:
'   Number of files written
' ============================================
Private Function ApplyToObsidian(outputDir As String, projectInfo As Object, tasks As Collection, _
                                  headerMapping As Object, taskMapping As Object) As Long
    ApplyToObsidian = 0

    If Not CreateFolder(outputDir) Then
        LogError TOOL_NAME, "Failed to create folder: " & outputDir
        Exit Function
    End If

    Dim written As Long
    written = 0

    Dim task As Object
    For Each task In tasks
        Dim taskId As String
        Dim taskName As String

        If task.Exists("task_id") Then
            taskId = CStr(task("task_id"))
        End If

        If task.Exists("task_name") Then
            taskName = CStr(task("task_name"))
        End If

        ' Skip if no task_name
        If Len(taskName) = 0 Then
            GoTo NextTask
        End If

        ' Generate frontmatter
        Dim frontmatter As String
        frontmatter = GenerateTaskFrontmatter(projectInfo, task, headerMapping, taskMapping)

        ' Build filename
        Dim filenameBase As String
        If Len(taskId) > 0 Then
            filenameBase = taskId & "_" & taskName
        Else
            filenameBase = taskName
        End If

        Dim filename As String
        filename = SanitizeFilename(filenameBase) & ".md"

        Dim filepath As String
        filepath = BuildFilePath(outputDir, filename)

        ' Read existing or create new
        Dim content As String
        If FileExists(filepath) Then
            Dim existing As String
            existing = ReadTextFile(filepath)
            Dim body As String
            body = ExtractBodyAfterFrontmatter(existing)
            content = frontmatter & body
        Else
            Dim description As String
            If task.Exists("description") Then
                description = CStr(task("description"))
            End If

            content = frontmatter
            If Len(taskName) > 0 Then
                content = content & "# " & taskName & vbLf & vbLf
            End If
            If Len(description) > 0 Then
                content = content & description & vbLf
            End If
        End If

        ' Write file
        If WriteTextFile(filepath, content, False) Then
            written = written + 1
        End If

NextTask:
    Next task

    LogInfo TOOL_NAME, "Written " & written & " files to " & outputDir
    ApplyToObsidian = written
End Function

' ============================================
' GenerateProjectFrontmatter
' Generate YAML frontmatter for project note
' ============================================
Private Function GenerateProjectFrontmatter(projectInfo As Object, headerMapping As Object) As String
    Dim lines As Collection
    Set lines = New Collection

    lines.Add "---"
    lines.Add "role: ""project"""

    ' Dataview required fields
    If projectInfo.Exists("project_name") Then
        lines.Add "project: " & FormatYamlValue(projectInfo("project_name"))
    End If

    If projectInfo.Exists("project_category") Then
        lines.Add "domain: " & FormatYamlValue(projectInfo("project_category"))
    End If

    If projectInfo.Exists("summary") Then
        lines.Add "summary: " & FormatYamlValue(projectInfo("summary"))
    End If

    If projectInfo.Exists("status") Then
        lines.Add "status: " & FormatYamlValue(projectInfo("status"))
    End If

    ' Other mapped fields
    Dim skipWbs As Object
    Set skipWbs = CreateObject("Scripting.Dictionary")
    skipWbs("project_name") = True
    skipWbs("project_category") = True
    skipWbs("summary") = True
    skipWbs("status") = True

    Dim skipObs As Object
    Set skipObs = CreateObject("Scripting.Dictionary")
    skipObs("role") = True
    skipObs("project") = True
    skipObs("domain") = True
    skipObs("summary") = True
    skipObs("status") = True

    Dim wbsKey As Variant
    For Each wbsKey In headerMapping.Keys
        If Not skipWbs.Exists(CStr(wbsKey)) Then
            Dim obsKey As String
            obsKey = CStr(headerMapping(wbsKey))

            If Not skipObs.Exists(obsKey) Then
                If projectInfo.Exists(CStr(wbsKey)) Then
                    Dim val As Variant
                    val = projectInfo(CStr(wbsKey))
                    Dim formatted As String
                    formatted = FormatYamlValue(val)
                    If Len(formatted) > 0 Then
                        lines.Add obsKey & ": " & formatted
                    End If
                End If
            End If
        End If
    Next wbsKey

    lines.Add "---"
    lines.Add ""

    ' Join lines
    Dim result As String
    Dim line As Variant
    For Each line In lines
        result = result & line & vbLf
    Next line

    GenerateProjectFrontmatter = result
End Function

' ============================================
' GenerateTaskFrontmatter
' Generate YAML frontmatter for task note
' ============================================
Private Function GenerateTaskFrontmatter(projectInfo As Object, task As Object, _
                                          headerMapping As Object, taskMapping As Object) As String
    Dim lines As Collection
    Set lines = New Collection

    lines.Add "---"
    lines.Add "role: ""task"""

    ' Get task obs_keys (task takes priority)
    Dim taskObsKeys As Object
    Set taskObsKeys = CreateObject("Scripting.Dictionary")

    Dim key As Variant
    For Each key In taskMapping.Keys
        taskObsKeys(CStr(taskMapping(key))) = True
    Next key

    ' Project info (skip if in task mapping)
    Dim skipObs As Object
    Set skipObs = CreateObject("Scripting.Dictionary")
    skipObs("role") = True

    For Each key In headerMapping.Keys
        Dim obsKey As String
        obsKey = CStr(headerMapping(key))

        If Not skipObs.Exists(obsKey) And Not taskObsKeys.Exists(obsKey) Then
            If projectInfo.Exists(CStr(key)) Then
                Dim val As Variant
                val = projectInfo(CStr(key))
                Dim formatted As String
                formatted = FormatYamlValue(val)
                If Len(formatted) > 0 Then
                    lines.Add obsKey & ": " & formatted
                End If
            End If
        End If
    Next key

    ' Task info
    For Each key In taskMapping.Keys
        obsKey = CStr(taskMapping(key))

        If Not skipObs.Exists(obsKey) Then
            If task.Exists(CStr(key)) Then
                val = task(CStr(key))
                formatted = FormatYamlValue(val)
                If Len(formatted) > 0 Then
                    lines.Add obsKey & ": " & formatted
                End If
            End If
        End If
    Next key

    lines.Add "---"
    lines.Add ""

    ' Join lines
    Dim result As String
    Dim line As Variant
    For Each line In lines
        result = result & line & vbLf
    Next line

    GenerateTaskFrontmatter = result
End Function

' ============================================
' FormatYamlValue
' Format value for YAML frontmatter
' ============================================
Private Function FormatYamlValue(val As Variant) As String
    FormatYamlValue = ""

    If IsEmpty(val) Or IsNull(val) Then
        Exit Function
    End If

    If IsDate(val) Then
        FormatYamlValue = Format(val, "yyyy-mm-dd")
        Exit Function
    End If

    If VarType(val) = vbBoolean Then
        FormatYamlValue = IIf(val, "true", "false")
        Exit Function
    End If

    If IsNumeric(val) Then
        FormatYamlValue = CStr(val)
        Exit Function
    End If

    Dim s As String
    s = CStr(val)

    ' Skip formulas
    If Left(s, 1) = "=" Then
        Exit Function
    End If

    ' Quote if contains special chars
    If InStr(s, ":") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, """") > 0 Then
        s = Replace(s, """", "\""")
        FormatYamlValue = """" & s & """"
    Else
        FormatYamlValue = s
    End If
End Function

' ============================================
' ExtractBodyAfterFrontmatter
' Extract content after YAML frontmatter
' ============================================
Private Function ExtractBodyAfterFrontmatter(content As String) As String
    Dim lines() As String
    lines = Split(content, vbLf)

    If UBound(lines) < 0 Then
        ExtractBodyAfterFrontmatter = content
        Exit Function
    End If

    ' Check for opening ---
    If Trim(lines(0)) <> "---" Then
        ExtractBodyAfterFrontmatter = content
        Exit Function
    End If

    ' Find closing ---
    Dim i As Long
    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "---" Then
            ' Return everything after
            Dim result As String
            Dim j As Long
            For j = i + 1 To UBound(lines)
                If j = i + 1 Then
                    result = lines(j)
                Else
                    result = result & vbLf & lines(j)
                End If
            Next j
            ExtractBodyAfterFrontmatter = result
            Exit Function
        End If
    Next i

    ExtractBodyAfterFrontmatter = content
End Function
