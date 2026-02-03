Option Explicit

' ============================================
' Module   : Pst_UpdateLastMtgDate
' Layer    : Presentation
' Purpose  : Update LAST-MTG-DATE parameter to today's date
' Version  : 1.0.0
' Created  : 2026-02-03
' Note     : Called from UI_Dashboard button
' ============================================

Private Const TOOL_NAME As String = "UpdateLastMtgDate"
Private Const PARAM_KEY_LAST_MTG_DATE As String = "LAST-MTG-DATE"

' ============================================
' UpdateLastMtgDate
' Set DEF_Parameter LAST-MTG-DATE value to today's date
' ============================================
Public Sub UpdateLastMtgDate()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateLastMtgDate: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Updating LAST-MTG-DATE..."
    Application.ScreenUpdating = False

    ' Check DEF_Parameter sheet exists
    If Not SheetExists(SHEET_DEF_PARAMETER) Then
        LogError TOOL_NAME, "Sheet not found: " & SHEET_DEF_PARAMETER
        MsgBox "Sheet not found: " & SHEET_DEF_PARAMETER, vbExclamation, "Error"
        GoTo Cleanup
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    ' Get today's date
    Dim today As Date
    today = Date

    LogInfo TOOL_NAME, "Setting LAST-MTG-DATE to: " & Format(today, "yyyy-mm-dd")

    ' Update LAST-MTG-DATE value using column name based lookup
    Dim updated As Boolean
    updated = UpdateParameterValue(ws, PARAM_KEY_LAST_MTG_DATE, today)

    If Not updated Then
        LogError TOOL_NAME, "Key not found: " & PARAM_KEY_LAST_MTG_DATE
        MsgBox "Parameter key not found: " & PARAM_KEY_LAST_MTG_DATE & vbCrLf & _
               "Please add this key to " & SHEET_DEF_PARAMETER & " sheet.", _
               vbExclamation, "Error"
        GoTo Cleanup
    End If

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateLastMtgDate: Completed"
    LogInfo TOOL_NAME, "========================================"

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "LAST-MTG-DATE updated to: " & Format(today, "yyyy-mm-dd"), _
           vbInformation, "Complete"

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
' UpdateParameterValue
' Update value in DEF_Parameter table by key name
' Uses "name" and "value" column names (same as LookupTableValue)
'
' Args:
'   ws: DEF_Parameter worksheet
'   keyName: Parameter key to update (e.g., "LAST-MTG-DATE")
'   newValue: New value to set
'
' Returns:
'   True if updated, False if key not found
' ============================================
Private Function UpdateParameterValue(ws As Worksheet, _
                                      keyName As String, _
                                      newValue As Variant) As Boolean
    UpdateParameterValue = False

    LogInfo TOOL_NAME, "[DEBUG] UpdateParameterValue started"
    LogInfo TOOL_NAME, "[DEBUG] Sheet name: " & ws.Name
    LogInfo TOOL_NAME, "[DEBUG] Looking for key: " & keyName

    ' Find Tbl_Start:Parameter marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_PARAMETER)

    LogInfo TOOL_NAME, "[DEBUG] FindTblStartRow result: " & markerRow

    If markerRow = 0 Then
        LogError TOOL_NAME, "Tbl_Start:" & TBL_PARAMETER & " not found"
        ' Debug: scan first 50 rows for any Tbl_Start markers
        LogInfo TOOL_NAME, "[DEBUG] Scanning for Tbl_Start markers in column A..."
        Dim scanRow As Long
        Dim scanVal As Variant
        For scanRow = 1 To 50
            scanVal = ws.Cells(scanRow, 1).Value
            If Not IsEmpty(scanVal) Then
                If InStr(1, CStr(scanVal), "Tbl_Start", vbTextCompare) > 0 Then
                    LogInfo TOOL_NAME, "[DEBUG]   Row " & scanRow & ": " & CStr(scanVal)
                End If
            End If
        Next scanRow
        Exit Function
    End If

    ' Header row is next row after marker
    Dim headerRow As Long
    headerRow = markerRow + 1

    LogInfo TOOL_NAME, "[DEBUG] Header row: " & headerRow

    ' Debug: show header row contents
    Dim colIdx As Long
    Dim headerVal As Variant
    LogInfo TOOL_NAME, "[DEBUG] Header row contents:"
    For colIdx = 1 To 10
        headerVal = ws.Cells(headerRow, colIdx).Value
        If Not IsEmpty(headerVal) Then
            LogInfo TOOL_NAME, "[DEBUG]   Col " & colIdx & ": " & CStr(headerVal)
        End If
    Next colIdx

    ' Find "name" and "value" column indices
    Dim headers As Variant
    headers = GetTableHeaders(ws, headerRow)

    LogInfo TOOL_NAME, "[DEBUG] GetTableHeaders returned " & (UBound(headers) - LBound(headers) + 1) & " columns"

    Dim nameColIdx As Long
    Dim valueColIdx As Long
    nameColIdx = GetColumnIndex(headers, "name")
    valueColIdx = GetColumnIndex(headers, "value")

    LogInfo TOOL_NAME, "[DEBUG] nameColIdx: " & nameColIdx & ", valueColIdx: " & valueColIdx

    If nameColIdx = 0 Or valueColIdx = 0 Then
        LogError TOOL_NAME, "Column 'name' or 'value' not found in header"
        Exit Function
    End If

    ' Search for key in data rows
    Dim i As Long
    Dim cellKey As Variant
    Dim maxRows As Long
    maxRows = 100

    LogInfo TOOL_NAME, "[DEBUG] Scanning data rows for key: " & keyName

    Dim emptyCount As Long
    emptyCount = 0

    For i = headerRow + 1 To headerRow + maxRows
        cellKey = ws.Cells(i, nameColIdx).Value

        ' Skip empty rows but continue scanning (allow gaps)
        If IsEmpty(cellKey) Or Trim(CStr(cellKey)) = "" Then
            emptyCount = emptyCount + 1
            ' Stop after 5 consecutive empty rows
            If emptyCount >= 5 Then
                LogInfo TOOL_NAME, "[DEBUG] 5 consecutive empty rows at row " & i & ", stopping scan"
                Exit For
            End If
            GoTo NextRow
        End If

        emptyCount = 0  ' Reset empty counter
        LogInfo TOOL_NAME, "[DEBUG]   Row " & i & " name: [" & CStr(cellKey) & "]"

        ' Case-insensitive comparison
        If StrComp(Trim(CStr(cellKey)), Trim(keyName), vbTextCompare) = 0 Then
            ' Found the key, update the value
            ws.Cells(i, valueColIdx).Value = newValue
            LogInfo TOOL_NAME, "[DEBUG] MATCH FOUND! Updated row " & i & ": " & keyName & " = " & CStr(newValue)
            UpdateParameterValue = True
            Exit Function
        End If

NextRow:
    Next i

    LogError TOOL_NAME, "Key '" & keyName & "' not found in Parameter table"
    LogInfo TOOL_NAME, "[DEBUG] Total rows scanned: " & (i - headerRow - 1)
End Function
