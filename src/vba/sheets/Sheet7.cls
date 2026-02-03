VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    On Error GoTo SafeExit

    ' Find ListObject with last_update column (handles copied sheets with renamed tables)
    Dim lo As ListObject
    Set lo = FindTaskListTable(Me)

    If lo Is Nothing Then GoTo SafeExit

    ' Ignore changes outside table data range
    If lo.DataBodyRange Is Nothing Then GoTo SafeExit
    If Intersect(target, lo.DataBodyRange) Is Nothing Then GoTo SafeExit

    ' Get last_update column
    Dim colLast As ListColumn
    Set colLast = lo.ListColumns("last_update")

    Dim rngLast As Range
    Set rngLast = colLast.DataBodyRange

    ' last_update�񂾂����ύX���ꂽ�ꍇ�͖����i�ċA�h�~�j
    If Not Intersect(target, rngLast) Is Nothing Then
        If Intersect(target, lo.DataBodyRange).Address = Intersect(target, rngLast).Address Then GoTo SafeExit
        ' ������Ɠ����ύX�i�\��t�����j�̏ꍇ�͍X�V�͑��s
    End If

    Application.EnableEvents = False

    ' �ύX���ꂽ�s����ӂɒ��o���� last_update ���X�V
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim c As Range
    For Each c In Intersect(target, lo.DataBodyRange).Cells
        dict(CStr(c.row)) = True
    Next c

    Dim k As Variant, r As Long
    For Each k In dict.keys
        r = CLng(k)

        ' ���̍s���e�[�u�����̉��s�ڂ��iDataBodyRange�擪����̃I�t�Z�b�g�j
        Dim idx As Long
        idx = r - lo.DataBodyRange.row + 1

        ' �e�[�u���O�i�O�̂��߁j��e��
        If idx >= 1 And idx <= lo.DataBodyRange.rows.count Then
            rngLast.Cells(idx, 1).Value = Now   ' ���t�����Ȃ� Date
        End If
    Next k

SafeExit:
    Application.EnableEvents = True
End Sub

' ============================================
' FindTaskListTable
' Find ListObject with last_update column
' Handles copied sheets where table name changes (TaskListTable -> TaskListTable2, etc.)
'
' Args:
'   ws: Target worksheet
'
' Returns:
'   ListObject with last_update column, or Nothing if not found
' ============================================
Private Function FindTaskListTable(ws As Worksheet) As ListObject
    Set FindTaskListTable = Nothing

    If ws.ListObjects.Count = 0 Then
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        ' Check if table has last_update column
        On Error Resume Next
        Dim col As ListColumn
        Set col = lo.ListColumns("last_update")
        On Error GoTo 0

        If Not col Is Nothing Then
            Set FindTaskListTable = lo
            Exit Function
        End If

        Set col = Nothing
    Next lo
End Function

