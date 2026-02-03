VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ============================================
' ThisWorkbook Module
' Purpose  : Workbook-level event handling
' Version  : 1.0.0
' Created  : 2026-02-03
' Note     : Handles last_update auto-update for all PJ- sheets
' ============================================

' ============================================
' Workbook_SheetChange
' Auto-update last_update column when task data changes
' Works for all PJ- sheets (including copied sheets)
' ============================================
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    On Error GoTo SafeExit

    ' Only process PJ- sheets (exclude templates)
    Dim sheetName As String
    sheetName = Sh.Name

    If Left(sheetName, 3) <> "PJ-" Then GoTo SafeExit
    If Left(sheetName, 7) = "DEF_PJ-" Then GoTo SafeExit

    ' Find ListObject with last_update column
    Dim lo As ListObject
    Set lo = FindTaskListTable(Sh)

    If lo Is Nothing Then GoTo SafeExit

    ' Ignore changes outside table data range
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit
    If Intersect(target, lo.DataBodyRange) Is Nothing Then GoTo SafeExit

    ' Get last_update column
    Dim colLast As ListColumn
    On Error Resume Next
    Set colLast = lo.ListColumns("last_update")
    On Error GoTo SafeExit

    If colLast Is Nothing Then GoTo SafeExit

    Dim rngLast As Range
    Set rngLast = colLast.DataBodyRange

    If rngLast Is Nothing Then GoTo SafeExit

    ' Skip if only last_update column was changed (prevent recursion)
    If Not Intersect(target, rngLast) Is Nothing Then
        If Intersect(target, lo.DataBodyRange).Address = Intersect(target, rngLast).Address Then GoTo SafeExit
    End If

    Application.EnableEvents = False

    ' Get unique rows that were changed
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim c As Range
    For Each c In Intersect(target, lo.DataBodyRange).Cells
        dict(CStr(c.Row)) = True
    Next c

    ' Update last_update for each changed row
    Dim k As Variant, r As Long
    For Each k In dict.Keys
        r = CLng(k)

        ' Calculate row index within table
        Dim idx As Long
        idx = r - lo.DataBodyRange.Row + 1

        ' Bounds check
        If idx >= 1 And idx <= lo.DataBodyRange.Rows.Count Then
            rngLast.Cells(idx, 1).Value = Now
        End If
    Next k

SafeExit:
    Application.EnableEvents = True
End Sub

' ============================================
' FindTaskListTable
' Find ListObject with last_update column
'
' Args:
'   ws: Target worksheet
'
' Returns:
'   ListObject with last_update column, or Nothing if not found
' ============================================
Private Function FindTaskListTable(ws As Object) As ListObject
    Set FindTaskListTable = Nothing

    On Error Resume Next
    Dim loCount As Long
    loCount = ws.ListObjects.Count
    On Error GoTo 0

    If loCount = 0 Then
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        ' Check if table has last_update column
        On Error Resume Next
        Dim col As ListColumn
        Set col = Nothing
        Set col = lo.ListColumns("last_update")
        On Error GoTo 0

        If Not col Is Nothing Then
            Set FindTaskListTable = lo
            Exit Function
        End If
    Next lo
End Function
