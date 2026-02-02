Option Explicit

' ============================================
' Module   : Utl_Logger
' Layer    : Common / Utility
' Purpose  : Logger utility functions (dependency of Mgr_Logger.cls)
' Version  : 1.0.0
' Created  : 2026-02-02
' ============================================

Private Const TEMP_FOLDER_NAME As String = "FlowBase"

' ============================================
' GetUModelTempPath
' Get temp directory path for tool logs
'
' Args:
'   toolName: Name of the tool (used in path)
'
' Returns:
'   Path to temp directory (e.g., %TEMP%\FlowBase\toolname\)
' ============================================
Public Function GetUModelTempPath(ByVal toolName As String) As String
    Dim basePath As String
    basePath = Environ("TEMP")

    If Right(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If

    basePath = basePath & TEMP_FOLDER_NAME & "\" & LCase(toolName) & "\"

    ' Create directory if not exists
    On Error Resume Next
    MkDir Environ("TEMP") & "\" & TEMP_FOLDER_NAME
    MkDir basePath
    On Error GoTo 0

    GetUModelTempPath = basePath
End Function

' ============================================
' DeleteFileIfTooLarge
' Delete file if it exceeds size limit
'
' Args:
'   filePath: Path to the file
'   maxSizeBytes: Maximum file size in bytes
' ============================================
Public Sub DeleteFileIfTooLarge(ByVal filePath As String, ByVal maxSizeBytes As Long)
    If IsFileTooLarge(filePath, maxSizeBytes) Then
        SafeKill filePath
    End If
End Sub

' ============================================
' IsFileTooLarge
' Check if file exceeds size limit
'
' Args:
'   filePath: Path to the file
'   maxSizeBytes: Maximum file size in bytes
'
' Returns:
'   True if file is larger than limit
' ============================================
Public Function IsFileTooLarge(ByVal filePath As String, ByVal maxSizeBytes As Long) As Boolean
    On Error Resume Next

    Dim fileSize As Long
    fileSize = FileLen(filePath)

    If Err.Number <> 0 Then
        ' File doesn't exist or can't be accessed
        IsFileTooLarge = False
        Exit Function
    End If

    IsFileTooLarge = (fileSize > maxSizeBytes)
    On Error GoTo 0
End Function

' ============================================
' SafeKill
' Delete file with error handling
'
' Args:
'   filePath: Path to the file
' ============================================
Public Sub SafeKill(ByVal filePath As String)
    On Error Resume Next
    Kill filePath
    On Error GoTo 0
End Sub

' ============================================
' AppendText
' Append text to file
'
' Args:
'   filePath: Path to the file
'   text: Text to append
' ============================================
Public Sub AppendText(ByVal filePath As String, ByVal text As String)
    On Error Resume Next

    Dim fNum As Integer
    fNum = FreeFile

    Open filePath For Append As #fNum
    Print #fNum, text;
    Close #fNum

    On Error GoTo 0
End Sub

' ============================================
' Timestamp
' Get current timestamp string
'
' Returns:
'   Timestamp in format "YYYY-MM-DD HH:NN:SS"
' ============================================
Public Function Timestamp() As String
    Timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function

' ============================================
' CurrentUser
' Get current Windows user name
'
' Returns:
'   User name
' ============================================
Public Function CurrentUser() As String
    CurrentUser = Environ("USERNAME")
End Function

' ============================================
' HostName
' Get computer name
'
' Returns:
'   Computer name
' ============================================
Public Function HostName() As String
    HostName = Environ("COMPUTERNAME")
End Function

' ============================================
' ThisWorkbookPathSafe
' Get ThisWorkbook path with error handling
'
' Returns:
'   Workbook full path, or "(not saved)" if not saved
' ============================================
Public Function ThisWorkbookPathSafe() As String
    On Error Resume Next

    If Len(ThisWorkbook.Path) = 0 Then
        ThisWorkbookPathSafe = "(not saved)"
    Else
        ThisWorkbookPathSafe = ThisWorkbook.FullName
    End If

    On Error GoTo 0
End Function

' ============================================
' FormatElapsed
' Format elapsed time between two dates
'
' Args:
'   startTime: Start time
'   endTime: End time
'
' Returns:
'   Formatted string like "00:05:23.456"
' ============================================
Public Function FormatElapsed(ByVal startTime As Date, ByVal endTime As Date) As String
    Dim totalSeconds As Double
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long
    Dim milliseconds As Long

    totalSeconds = (endTime - startTime) * 86400 ' 24 * 60 * 60

    hours = Int(totalSeconds / 3600)
    totalSeconds = totalSeconds - hours * 3600

    minutes = Int(totalSeconds / 60)
    totalSeconds = totalSeconds - minutes * 60

    seconds = Int(totalSeconds)
    milliseconds = Int((totalSeconds - seconds) * 1000)

    FormatElapsed = Format(hours, "00") & ":" & _
                    Format(minutes, "00") & ":" & _
                    Format(seconds, "00") & "." & _
                    Format(milliseconds, "000")
End Function

' ============================================
' WriteLog
' Simple log writer for tool modules
'
' Args:
'   toolName: Name of the tool
'   level: Log level (INFO, DEBUG, WARN, ERROR)
'   message: Log message
' ============================================
Public Sub WriteLog(ByVal toolName As String, ByVal Level As String, ByVal message As String)
    On Error Resume Next

    Dim logDir As String
    Dim logPath As String
    Dim logLine As String

    logDir = ThisWorkbook.Path & "\logs"

    ' Create directory if not exists
    MkDir logDir

    logPath = logDir & "\vba_" & LCase(toolName) & "_" & Format(Now, "yyyymmdd") & ".log"
    logLine = Timestamp() & " [" & Level & "] " & message & vbCrLf

    AppendText logPath, logLine

    ' Also output to Immediate window for debugging
    Debug.Print logLine

    On Error GoTo 0
End Sub

' ============================================
' LogInfo / LogDebug / LogWarn / LogError
' Convenience wrappers for WriteLog
' ============================================
Public Sub LogInfo(ByVal toolName As String, ByVal message As String)
    WriteLog toolName, "INFO ", message
End Sub

Public Sub LogDebug(ByVal toolName As String, ByVal message As String)
    WriteLog toolName, "DEBUG", message
End Sub

Public Sub LogWarn(ByVal toolName As String, ByVal message As String)
    WriteLog toolName, "WARN ", message
End Sub

Public Sub LogError(ByVal toolName As String, ByVal message As String)
    WriteLog toolName, "ERROR", message
End Sub
