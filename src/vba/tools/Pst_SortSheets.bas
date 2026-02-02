Option Explicit

' ============================================
' Module   : Pst_SortSheets
' Layer    : Presentation
' Purpose  : Sort all sheets by prefix priority
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : VBA native implementation (no Python)
' ============================================

Private Const TOOL_NAME As String = "SortSheets"

' ============================================
' SortSheets
' Sort all sheets by prefix priority from DEF_SheetPrefix
' ============================================
Public Sub SortSheets()
    On Error GoTo EH

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "SortSheets: Started"
    LogInfo TOOL_NAME, "========================================"

    Application.StatusBar = "Sorting sheets..."
    Application.ScreenUpdating = False

    ' Load prefix sort order from DEF_SheetPrefix
    Dim prefixOrder As Object
    Set prefixOrder = LoadPrefixSortOrder()

    If prefixOrder.Count = 0 Then
        LogWarn TOOL_NAME, "No prefix definitions found, sheets will be sorted alphabetically"
    Else
        LogInfo TOOL_NAME, "Loaded " & prefixOrder.Count & " prefix definitions"

        ' Log the definitions
        Dim key As Variant
        For Each key In prefixOrder.Keys
            LogDebug TOOL_NAME, "  '" & key & "' -> " & prefixOrder(key)
        Next key
    End If

    ' Compute sorted order
    Dim sortedNames As Variant
    sortedNames = ComputeSheetOrder(prefixOrder)

    ' Log the computed order
    LogInfo TOOL_NAME, "Computed order:"
    Dim i As Long
    For i = LBound(sortedNames) To UBound(sortedNames)
        Dim sortKey As Long
        sortKey = GetSheetSortKey(CStr(sortedNames(i)), prefixOrder)
        LogInfo TOOL_NAME, "  " & Format(i, "00") & ". [" & Format(sortKey, "0000") & "] " & sortedNames(i)
    Next i

    ' Apply the order
    Dim moved As Long
    moved = ApplySheetOrder(sortedNames)

    LogInfo TOOL_NAME, "Moved " & moved & " sheets"

    Application.ScreenUpdating = True
    Application.StatusBar = False

    LogInfo TOOL_NAME, "========================================"
    LogInfo TOOL_NAME, "SortSheets: Completed"
    LogInfo TOOL_NAME, "========================================"

    MsgBox "Sheet sorting completed." & vbCrLf & _
           moved & " sheets were reordered.", vbInformation, "Complete"

    Exit Sub

EH:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    LogError TOOL_NAME, "Error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub
