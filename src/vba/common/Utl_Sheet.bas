Option Explicit

' ============================================
' Module   : Utl_Sheet
' Layer    : Common / Utility
' Purpose  : Sheet operations (filter, sort, copy)
' Version  : 1.0.0
' Created  : 2026-02-02
' ============================================

' ============================================
' FilterSheetsByPrefix
' Get all sheet names starting with specified prefix
'
' Args:
'   wb: Target workbook (use ThisWorkbook if Nothing)
'   prefix: Prefix to filter (e.g., "PJ-")
'
' Returns:
'   Collection of sheet names
' ============================================
Public Function FilterSheetsByPrefix(prefix As String, _
                                      Optional wb As Workbook = Nothing) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            result.Add ws.Name
        End If
    Next ws

    Set FilterSheetsByPrefix = result
End Function

' ============================================
' SheetExists
' Check if sheet exists in workbook
'
' Args:
'   sheetName: Sheet name to check
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   True if sheet exists
' ============================================
Public Function SheetExists(sheetName As String, _
                             Optional wb As Workbook = Nothing) As Boolean
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    On Error Resume Next
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets(sheetName)
    SheetExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' ============================================
' GetSheetSortKey
' Get sort key for sheet based on prefix priority
'
' Args:
'   sheetName: Sheet name
'   prefixOrder: Dictionary with {prefix: sort_order}
'
' Returns:
'   Sort order (Long), or DEFAULT_SORT_ORDER if no match
' ============================================
Public Function GetSheetSortKey(sheetName As String, _
                                  prefixOrder As Object) As Long
    Dim matchedPrefix As String
    Dim matchedOrder As Long
    Dim prefix As Variant
    Dim prefixLen As Long

    matchedPrefix = ""
    matchedOrder = DEFAULT_SORT_ORDER

    ' Find longest matching prefix
    Dim key As Variant
    For Each key In prefixOrder.Keys
        prefix = CStr(key)
        prefixLen = Len(prefix)

        If Left(sheetName, prefixLen) = prefix Then
            If prefixLen > Len(matchedPrefix) Then
                matchedPrefix = prefix
                matchedOrder = CLng(prefixOrder(key))
            End If
        End If
    Next key

    GetSheetSortKey = matchedOrder
End Function

' ============================================
' LoadPrefixSortOrder
' Load prefix sort order from DEF_SheetPrefix sheet
'
' Args:
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   Dictionary with {prefix: sort_order}
' ============================================
Public Function LoadPrefixSortOrder(Optional wb As Workbook = Nothing) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    ' Check sheet exists
    If Not SheetExists(SHEET_DEF_SHEET_PREFIX, targetWb) Then
        Set LoadPrefixSortOrder = dict
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = targetWb.Worksheets(SHEET_DEF_SHEET_PREFIX)

    ' Find columns: sheet_prefix, sort_order
    Dim prefixCol As Long, orderCol As Long
    Dim col As Long
    Dim headerVal As Variant

    prefixCol = 0
    orderCol = 0

    For col = 1 To 20
        headerVal = ws.Cells(1, col).Value
        If Not IsEmpty(headerVal) Then
            Select Case CStr(headerVal)
                Case "sheet_prefix": prefixCol = col
                Case "sort_order": orderCol = col
            End Select
        End If
    Next col

    If prefixCol = 0 Or orderCol = 0 Then
        Set LoadPrefixSortOrder = dict
        Exit Function
    End If

    ' Read data rows
    Dim row As Long
    Dim prefix As Variant
    Dim order As Variant
    Dim orderVal As Long

    For row = 2 To 100
        prefix = ws.Cells(row, prefixCol).Value

        If IsEmpty(prefix) Or Trim(CStr(prefix)) = "" Then
            Exit For
        End If

        order = ws.Cells(row, orderCol).Value

        On Error Resume Next
        orderVal = CLng(order)
        If Err.Number <> 0 Then
            orderVal = DEFAULT_SORT_ORDER
        End If
        On Error GoTo 0

        dict(CStr(prefix)) = orderVal
    Next row

    Set LoadPrefixSortOrder = dict
End Function

' ============================================
' ComputeSheetOrder
' Compute sorted order of sheets based on prefix priority
'
' Args:
'   wb: Target workbook (use ThisWorkbook if Nothing)
'   prefixOrder: Dictionary with {prefix: sort_order}
'
' Returns:
'   Array of sheet names in sorted order
' ============================================
Public Function ComputeSheetOrder(prefixOrder As Object, _
                                   Optional wb As Workbook = Nothing) As Variant
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim sheetCount As Long
    sheetCount = targetWb.Worksheets.Count

    If sheetCount = 0 Then
        ComputeSheetOrder = Array()
        Exit Function
    End If

    ' Build array of (sortKey, sheetName)
    Dim sheetsWithKeys() As Variant
    ReDim sheetsWithKeys(1 To sheetCount, 1 To 2)

    Dim i As Long
    Dim ws As Worksheet
    i = 0
    For Each ws In targetWb.Worksheets
        i = i + 1
        sheetsWithKeys(i, 1) = GetSheetSortKey(ws.Name, prefixOrder)
        sheetsWithKeys(i, 2) = ws.Name
    Next ws

    ' Sort by (sortKey, sheetName) using bubble sort
    Dim j As Long
    Dim tempKey As Long
    Dim tempName As String
    Dim swapped As Boolean

    For i = 1 To sheetCount - 1
        swapped = False
        For j = 1 To sheetCount - i
            ' Compare: first by sortKey, then by name
            Dim doSwap As Boolean
            doSwap = False

            If sheetsWithKeys(j, 1) > sheetsWithKeys(j + 1, 1) Then
                doSwap = True
            ElseIf sheetsWithKeys(j, 1) = sheetsWithKeys(j + 1, 1) Then
                If sheetsWithKeys(j, 2) > sheetsWithKeys(j + 1, 2) Then
                    doSwap = True
                End If
            End If

            If doSwap Then
                tempKey = sheetsWithKeys(j, 1)
                tempName = sheetsWithKeys(j, 2)
                sheetsWithKeys(j, 1) = sheetsWithKeys(j + 1, 1)
                sheetsWithKeys(j, 2) = sheetsWithKeys(j + 1, 2)
                sheetsWithKeys(j + 1, 1) = tempKey
                sheetsWithKeys(j + 1, 2) = tempName
                swapped = True
            End If
        Next j
        If Not swapped Then Exit For
    Next i

    ' Extract sorted names
    Dim sortedNames() As String
    ReDim sortedNames(1 To sheetCount)

    For i = 1 To sheetCount
        sortedNames(i) = sheetsWithKeys(i, 2)
    Next i

    ComputeSheetOrder = sortedNames
End Function

' ============================================
' ApplySheetOrder
' Move sheets to match specified order
'
' Args:
'   sortedNames: Array of sheet names in desired order
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   Number of sheets moved
' ============================================
Public Function ApplySheetOrder(sortedNames As Variant, _
                                  Optional wb As Workbook = Nothing) As Long
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim moved As Long
    moved = 0

    Dim targetIdx As Long
    Dim ws As Worksheet
    Dim lb As Long, ub As Long

    lb = LBound(sortedNames)
    ub = UBound(sortedNames)

    For targetIdx = lb To ub
        Dim sheetName As String
        sheetName = sortedNames(targetIdx)

        On Error Resume Next
        Set ws = targetWb.Worksheets(sheetName)
        If Err.Number <> 0 Then
            On Error GoTo 0
            GoTo NextSheet
        End If
        On Error GoTo 0

        Dim currentIdx As Long
        currentIdx = GetSheetIndex(ws, targetWb)

        Dim desiredIdx As Long
        desiredIdx = targetIdx - lb + 1

        If currentIdx <> desiredIdx Then
            If desiredIdx = 1 Then
                ws.Move Before:=targetWb.Worksheets(1)
            Else
                ws.Move After:=targetWb.Worksheets(desiredIdx - 1)
            End If
            moved = moved + 1
        End If

NextSheet:
    Next targetIdx

    ApplySheetOrder = moved
End Function

' ============================================
' GetSheetIndex
' Get current index of worksheet in workbook
' ============================================
Private Function GetSheetIndex(ws As Worksheet, wb As Workbook) As Long
    Dim i As Long
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = ws.Name Then
            GetSheetIndex = i
            Exit Function
        End If
    Next i
    GetSheetIndex = 0
End Function

' ============================================
' CopySheet
' Copy template sheet to new sheet
'
' Args:
'   templateName: Template sheet name
'   newName: New sheet name
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   New worksheet, or Nothing if failed
' ============================================
Public Function CopySheet(templateName As String, _
                           newName As String, _
                           Optional wb As Workbook = Nothing) As Worksheet
    On Error GoTo ErrHandler

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    ' Check template exists
    If Not SheetExists(templateName, targetWb) Then
        Set CopySheet = Nothing
        Exit Function
    End If

    ' Check new name doesn't exist
    If SheetExists(newName, targetWb) Then
        Set CopySheet = Nothing
        Exit Function
    End If

    Dim templateWs As Worksheet
    Set templateWs = targetWb.Worksheets(templateName)

    ' Copy sheet to end
    templateWs.Copy After:=targetWb.Worksheets(targetWb.Worksheets.Count)

    Dim newWs As Worksheet
    Set newWs = targetWb.Worksheets(targetWb.Worksheets.Count)

    ' Rename tables to avoid conflicts
    RenameTablesInSheet newWs

    ' Rename sheet
    newWs.Name = newName

    Set CopySheet = newWs
    Exit Function

ErrHandler:
    Set CopySheet = Nothing
End Function

' ============================================
' RenameTablesInSheet
' Rename all ListObjects in sheet to unique names
'
' Args:
'   ws: Target worksheet
' ============================================
Public Sub RenameTablesInSheet(ws As Worksheet)
    On Error Resume Next

    Dim lo As ListObject
    Dim oldName As String
    Dim newName As String
    Dim suffix As Long

    For Each lo In ws.ListObjects
        oldName = lo.Name
        suffix = 1

        Do
            newName = oldName & "_" & suffix

            ' Try to rename
            Err.Clear
            lo.Name = newName

            If Err.Number = 0 Then
                Exit Do
            End If

            suffix = suffix + 1
            If suffix > 100 Then Exit Do
        Loop
    Next lo

    On Error GoTo 0
End Sub

' ============================================
' GetAllSheetNames
' Get all sheet names in workbook
'
' Args:
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   Collection of sheet names
' ============================================
Public Function GetAllSheetNames(Optional wb As Workbook = Nothing) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        result.Add ws.Name
    Next ws

    Set GetAllSheetNames = result
End Function

' ============================================
' CountSheetsByPrefix
' Count sheets starting with specified prefix
'
' Args:
'   prefix: Prefix to count
'   wb: Target workbook (use ThisWorkbook if Nothing)
'
' Returns:
'   Number of matching sheets
' ============================================
Public Function CountSheetsByPrefix(prefix As String, _
                                     Optional wb As Workbook = Nothing) As Long
    Dim targetWb As Workbook
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If

    Dim count As Long
    count = 0

    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        If Left(ws.Name, Len(prefix)) = prefix Then
            count = count + 1
        End If
    Next ws

    CountSheetsByPrefix = count
End Function
