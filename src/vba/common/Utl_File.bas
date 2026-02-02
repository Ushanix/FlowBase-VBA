Option Explicit

' ============================================
' Module   : Utl_File
' Layer    : Common / Utility
' Purpose  : File operations using FileSystemObject
' Version  : 1.0.0
' Created  : 2026-02-02
' Note     : Uses late binding for FSO (no reference required)
' ============================================

' ============================================
' CreateFolder
' Create folder if it doesn't exist (recursive)
'
' Args:
'   folderPath: Full path to folder
'
' Returns:
'   True if folder exists or was created
' ============================================
Public Function CreateFolder(folderPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(folderPath) Then
        CreateFolder = True
        Exit Function
    End If

    ' Create parent folders recursively
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        If Not CreateFolder(parentPath) Then
            CreateFolder = False
            Exit Function
        End If
    End If

    ' Create the folder
    fso.CreateFolder folderPath
    CreateFolder = True
    Exit Function

ErrHandler:
    CreateFolder = False
End Function

' ============================================
' WriteTextFile
' Write text content to file (UTF-8 with BOM)
'
' Args:
'   filePath: Full path to file
'   content: Text content to write
'   append: True to append, False to overwrite
'
' Returns:
'   True if successful
' ============================================
Public Function WriteTextFile(filePath As String, content As String, Optional append As Boolean = False) As Boolean
    On Error GoTo ErrHandler

    ' Ensure parent folder exists
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(filePath)

    If Not CreateFolder(parentPath) Then
        WriteTextFile = False
        Exit Function
    End If

    ' Use ADODB.Stream for UTF-8 output
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open

        If append And fso.FileExists(filePath) Then
            ' Read existing content first
            Dim existingStream As Object
            Set existingStream = CreateObject("ADODB.Stream")
            existingStream.Type = 2
            existingStream.Charset = "UTF-8"
            existingStream.Open
            existingStream.LoadFromFile filePath
            Dim existing As String
            existing = existingStream.ReadText
            existingStream.Close
            Set existingStream = Nothing

            .WriteText existing
        End If

        .WriteText content
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With

    Set stream = Nothing
    WriteTextFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    WriteTextFile = False
End Function

' ============================================
' ReadTextFile
' Read text content from file (UTF-8)
'
' Args:
'   filePath: Full path to file
'
' Returns:
'   File content, or "" if error
' ============================================
Public Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        ReadTextFile = ""
        Exit Function
    End If

    ' Use ADODB.Stream for UTF-8 input
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    With stream
        .Type = 2 ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadTextFile = .ReadText
        .Close
    End With

    Set stream = Nothing
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    ReadTextFile = ""
End Function

' ============================================
' FileExists
' Check if file exists
'
' Args:
'   filePath: Full path to file
'
' Returns:
'   True if file exists
' ============================================
Public Function FileExists(filePath As String) As Boolean
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FileExists = fso.FileExists(filePath)
End Function

' ============================================
' FolderExists
' Check if folder exists
'
' Args:
'   folderPath: Full path to folder
'
' Returns:
'   True if folder exists
' ============================================
Public Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderPath)
End Function

' ============================================
' SanitizeFilename
' Remove invalid characters from filename
'
' Args:
'   filename: Original filename
'
' Returns:
'   Sanitized filename
' ============================================
Public Function SanitizeFilename(filename As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim invalidChars As String

    invalidChars = "<>:""/\|?*"
    result = filename

    For i = 1 To Len(invalidChars)
        char = Mid(invalidChars, i, 1)
        result = Replace(result, char, "_")
    Next i

    SanitizeFilename = result
End Function

' ============================================
' BuildFilePath
' Combine folder path and filename
'
' Args:
'   folderPath: Folder path
'   filename: Filename
'
' Returns:
'   Combined path
' ============================================
Public Function BuildFilePath(folderPath As String, filename As String) As String
    If Right(folderPath, 1) = "\" Then
        BuildFilePath = folderPath & filename
    Else
        BuildFilePath = folderPath & "\" & filename
    End If
End Function

' ============================================
' GetFileExtension
' Get file extension
'
' Args:
'   filePath: File path
'
' Returns:
'   Extension (e.g., ".txt")
' ============================================
Public Function GetFileExtension(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    GetFileExtension = fso.GetExtensionName(filePath)
    If Len(GetFileExtension) > 0 Then
        GetFileExtension = "." & GetFileExtension
    End If
End Function

' ============================================
' GetFilenameWithoutExtension
' Get filename without extension
'
' Args:
'   filePath: File path
'
' Returns:
'   Filename without extension
' ============================================
Public Function GetFilenameWithoutExtension(filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim baseName As String
    baseName = fso.GetBaseName(filePath)

    GetFilenameWithoutExtension = baseName
End Function

' ============================================
' DeleteFile
' Delete file if it exists
'
' Args:
'   filePath: File path
'
' Returns:
'   True if deleted or didn't exist
' ============================================
Public Function DeleteFile(filePath As String) As Boolean
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If

    DeleteFile = (Err.Number = 0)
End Function
