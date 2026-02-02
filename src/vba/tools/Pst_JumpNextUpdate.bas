Option Explicit

' ============================================
' Module   : Pst_JumpNextUpdate
' Layer    : Presentation
' Purpose  : Jump to next PJ sheet with update_flag=YES
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "JumpNextUpdate"

' ============================================
' JumpNextUpdate
' Find next PJ sheet with update_flag=YES and activate it
' Searches from current sheet position, wraps around
' ============================================
Public Sub JumpNextUpdate()
    On Error GoTo EH

    Dim currentSheet As String
    currentSheet = ActiveSheet.Name

    Application.StatusBar = "Searching for next update_flag=YES..."
    LogInfo TOOL_NAME, "Starting from: " & currentSheet

    ' Get all PJ sheets with update_flag status
    Dim pjSheets As Collection
    Set pjSheets = GetPJSheetsWithUpdateFlag()

    If pjSheets.Count = 0 Then
        Application.StatusBar = False
        MsgBox "No PJ sheets found.", vbInformation, "Info"
        LogInfo TOOL_NAME, "No PJ sheets found"
        Exit Sub
    End If

    LogInfo TOOL_NAME, "Found " & pjSheets.Count & " PJ sheets"

    ' Find next sheet with update_flag=YES
    Dim nextSheet As String
    nextSheet = FindNextUpdateSheet(pjSheets, currentSheet)

    Application.StatusBar = False

    If Len(nextSheet) > 0 Then
        ' Activate the target sheet
        On Error Resume Next
        ThisWorkbook.Worksheets(nextSheet).Activate
        If Err.Number <> 0 Then
            MsgBox "Sheet not found: " & nextSheet, vbExclamation, "Error"
            LogError TOOL_NAME, "Sheet not found: " & nextSheet
        Else
            LogInfo TOOL_NAME, "Jumped to: " & nextSheet
        End If
        On Error GoTo EH
    Else
        MsgBox "No sheet with update_flag=YES found.", vbInformation, "Info"
        LogInfo TOOL_NAME, "No sheet with update_flag=YES found"
    End If

    Exit Sub

EH:
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

' ============================================
' GetPJSheetsWithUpdateFlag
' Get all PJ sheets with their update_flag status
'
' Returns:
'   Collection of Dictionary with {name, has_update_flag}
' ============================================
Private Function GetPJSheetsWithUpdateFlag() As Collection
    Dim result As Collection
    Set result = New Collection

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Filter PJ- prefix, exclude templates
        If Left(ws.Name, Len(PREFIX_PROJECT)) = PREFIX_PROJECT Then
            If Left(ws.Name, Len(PREFIX_TEMPLATE_PROJECT)) <> PREFIX_TEMPLATE_PROJECT Then
                Dim dict As Object
                Set dict = CreateObject("Scripting.Dictionary")

                dict("name") = ws.Name
                dict("has_update_flag") = CheckUpdateFlag(ws)

                result.Add dict

                LogDebug TOOL_NAME, "  " & ws.Name & ": update_flag=" & dict("has_update_flag")
            End If
        End If
    Next ws

    Set GetPJSheetsWithUpdateFlag = result
End Function

' ============================================
' CheckUpdateFlag
' Check if sheet has update_flag=YES
'
' Args:
'   ws: Worksheet to check
'
' Returns:
'   True if update_flag=YES
' ============================================
Private Function CheckUpdateFlag(ws As Worksheet) As Boolean
    CheckUpdateFlag = False

    ' Find Tbl_Start:header_info marker
    Dim markerRow As Long
    markerRow = FindTblStartRow(ws, TBL_HEADER_INFO)

    If markerRow = 0 Then
        Exit Function
    End If

    ' Read header_info table
    Dim headerInfo As Object
    Set headerInfo = ReadKeyValueTable(ws, markerRow + 1)

    ' Check update_flag value
    If headerInfo.Exists("update_flag") Then
        Dim flagValue As Variant
        flagValue = headerInfo("update_flag")

        If Not IsEmpty(flagValue) Then
            If UCase(CStr(flagValue)) = "YES" Then
                CheckUpdateFlag = True
            End If
        End If
    End If
End Function

' ============================================
' FindNextUpdateSheet
' Find next sheet with update_flag=YES
' Searches from current position, wraps around
'
' Args:
'   pjSheets: Collection of {name, has_update_flag}
'   currentSheet: Current sheet name
'
' Returns:
'   Sheet name, or "" if not found
' ============================================
Private Function FindNextUpdateSheet(pjSheets As Collection, currentSheet As String) As String
    FindNextUpdateSheet = ""

    If pjSheets.Count = 0 Then
        Exit Function
    End If

    ' Find current sheet index
    Dim currentIdx As Long
    Dim i As Long

    currentIdx = -1
    For i = 1 To pjSheets.Count
        If pjSheets(i)("name") = currentSheet Then
            currentIdx = i
            Exit For
        End If
    Next i

    If currentIdx = -1 Then
        ' Current sheet not in list, start from beginning
        currentIdx = 0
    End If

    ' Search from current position, wrap around
    Dim searchIdx As Long
    Dim total As Long
    total = pjSheets.Count

    For i = 1 To total
        searchIdx = ((currentIdx - 1 + i) Mod total) + 1

        If pjSheets(searchIdx)("has_update_flag") Then
            FindNextUpdateSheet = pjSheets(searchIdx)("name")
            Exit Function
        End If
    Next i
End Function
