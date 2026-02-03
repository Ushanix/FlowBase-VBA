Option Explicit

' ============================================
' Module   : Pst_UpdateAll
' Layer    : Presentation
' Purpose  : Execute all update tools in sequence
' Version  : 1.0.0
' Created  : 2026-02-03
' Note     : For end-of-day batch processing
' ============================================

Private Const TOOL_NAME As String = "UpdateAll"

' ============================================
' UpdateAll
' Execute all update tools in sequence
' Shows confirmation dialog before execution
' ============================================
Public Sub UpdateAll()
    On Error GoTo EH

    ' Show confirmation dialog
    Dim confirmResult As VbMsgBoxResult
    confirmResult = MsgBox( _
        "以下の処理を一括実行します：" & vbCrLf & vbCrLf & _
        "  1. IndexUpdate (インデックス更新)" & vbCrLf & _
        "  2. ProjectIndexUpdate (プロジェクトインデックス更新)" & vbCrLf & _
        "  3. UpdateTaskUrgent (緊急タスク更新)" & vbCrLf & _
        "  4. SortSheets (シート並び替え)" & vbCrLf & _
        "  5. OutputToObsidianAll (全PJシートObsidian出力) ※パス設定時のみ" & vbCrLf & vbCrLf & _
        "※ UpdatePersonalTask はPT-シートから個別実行してください" & vbCrLf & vbCrLf & _
        "処理に時間がかかります。実行しますか？", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "UpdateAll - 一括更新確認")

    If confirmResult <> vbYes Then
        Exit Sub
    End If

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateAll: Started"
    LogInfo TOOL_NAME, "========================================"

    Dim startTime As Double
    startTime = Timer

    Application.StatusBar = "UpdateAll: Processing..."
    Application.ScreenUpdating = False

    Dim successCount As Long
    Dim skipCount As Long
    Dim errorCount As Long
    Dim results As String

    successCount = 0
    skipCount = 0
    errorCount = 0
    results = ""

    ' 1. IndexUpdate
    results = results & ExecuteTool("IndexUpdate", "IndexUpdate")

    ' 2. ProjectIndexUpdate
    results = results & ExecuteTool("ProjectIndexUpdate", "ProjectIndexUpdate")

    ' 3. UpdateTaskUrgent
    results = results & ExecuteTool("UpdateTaskUrgent", "UpdateTaskUrgent")

    ' 4. SortSheets
    results = results & ExecuteTool("SortSheets", "SortSheets")

    ' 5. OutputToObsidianAll (skip if path not configured)
    If IsObsidianPathConfigured() Then
        results = results & ExecuteTool("OutputToObsidianAll", "OutputToObsidianAll")
    Else
        results = results & "  [SKIP] OutputToObsidianAll: パス未設定" & vbCrLf
        skipCount = skipCount + 1
        LogInfo TOOL_NAME, "Skipped OutputToObsidianAll: path not configured"
    End If

    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "UpdateAll: Completed in " & Format(elapsedTime, "0.0") & " seconds"
    LogInfo TOOL_NAME, "========================================"

    ' Show summary
    MsgBox "一括更新が完了しました。" & vbCrLf & vbCrLf & _
           results & vbCrLf & _
           "処理時間: " & Format(elapsedTime, "0.0") & " 秒", _
           vbInformation, "UpdateAll - 完了"

    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "UpdateAll - Error"
End Sub

' ============================================
' ExecuteTool
' Execute a single tool and return result string
'
' Args:
'   toolName: Name for logging
'   procName: Procedure name to call
'
' Returns:
'   Result string for display
' ============================================
Private Function ExecuteTool(toolName As String, procName As String) As String
    On Error GoTo ToolError

    LogInfo TOOL_NAME, "Executing: " & toolName
    Application.StatusBar = "UpdateAll: " & toolName & "..."

    Application.Run procName

    LogInfo TOOL_NAME, "Completed: " & toolName
    ExecuteTool = "  [OK] " & toolName & vbCrLf
    Exit Function

ToolError:
    LogError TOOL_NAME, "Error in " & toolName & ": " & Err.Description
    ExecuteTool = "  [ERROR] " & toolName & ": " & Err.Description & vbCrLf
End Function

' ============================================
' IsObsidianPathConfigured
' Check if Obsidian path is configured in DEF_Parameter
'
' Returns:
'   True if path is configured, False otherwise
' ============================================
Private Function IsObsidianPathConfigured() As Boolean
    IsObsidianPathConfigured = False

    If Not SheetExists(SHEET_DEF_PARAMETER) Then
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_DEF_PARAMETER)

    Dim result As Variant
    result = LookupTableValue(ws, TBL_PARAMETER, "name", "value", PARAM_OBSIDIAN_PATH)

    If Not IsEmpty(result) And Len(CStr(result)) > 0 Then
        IsObsidianPathConfigured = True
    End If
End Function
