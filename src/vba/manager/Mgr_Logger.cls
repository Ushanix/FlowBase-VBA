VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mgr_Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================
' クラス名   : Mgr_Logger
' 分類       : Manager（調停/WF）
' 概要       : ログ出力処理
' 主要責務   :  ログ出力処理（即時ウィンドウ、ファイル出力、レベル管理、サイズ監視など）
' 主な入出力 : 入力：ログテキスト/出力：.log
' 関連規約   : UModel_命名規約.md / UModel_構造ガイド.md
' 作成日     : 2025-08-14
' 作成者     : Ushas
' 履歴       :
'   2025-08-14  作成
' ============================================

'===== Mgr_Logger.cls =====
Option Explicit

' 依存は「ILoggerConfig」のみ（IConfigやCfg_*に直接依存しない）
Private mLogCfg As ILoggerConfig

'--- ログレベル
Public Enum LogLevel
    LOG_DEBUG = 10
    LOG_INFO = 20
    LOG_WARN = 30
    LOG_ERROR = 40
End Enum

'--- 設定/状態
Private mToolName As String
Private mVersion As String
Private mLogFilePath As String
Private mMinLevel As LogLevel
Private mUseIndent As Boolean
Private mIndentLevel As Long
Private mIndentStep As Long
Private mSessionStart As Date
Private mMaxFileSizeBytes As Long

'=== 初期化 ===
Public Sub Init(ByVal logCfg As ILoggerConfig, _
                Optional ByVal ToolName As String = "", _
                Optional ByVal Version As String = "", _
                Optional ByVal MinLevel As LogLevel = LOG_INFO, _
                Optional ByVal UseIndent As Boolean = False, _
                Optional ByVal indentStep As Long = 4, _
                Optional ByVal MaxFileSizeBytes As Long = 2& * 1024& * 1024&)
    ' --- 表示メタ（値渡しが優先、未指定ならlogCfgから補完） ---
    mToolName = IIf(LenB(ToolName) > 0, ToolName, logCfg.ToolNameForLog)
    mVersion = IIf(LenB(Version) > 0, Version, logCfg.VersionForLog)
    
    mMinLevel = MinLevel
    mUseIndent = UseIndent
    mIndentStep = IIf(indentStep > 0, indentStep, 4)
    mIndentLevel = 0
    mMaxFileSizeBytes = MaxFileSizeBytes
    
    Dim basePath As String
    basePath = Utl_Logger.GetUModelTempPath(mToolName)
    mLogFilePath = basePath & LCase$(mToolName) & "_" & Format(Now, "yyyymmdd") & ".log"
    
    Debug.Print mLogFilePath
    ' 過大なら削除して作り直す
    Utl_Logger.DeleteFileIfTooLarge mLogFilePath, mMaxFileSizeBytes
    
    mSessionStart = Now
    WriteSessionHeader
End Sub

'=== セッションヘッダ/フッタ ===
Private Sub WriteSessionHeader()
    Dim hdr As String
    hdr = String$(80, "=") & vbCrLf & _
          "[Session Start] " & Utl_Logger.Timestamp() & vbCrLf & _
          "Tool        : " & mToolName & vbCrLf & _
          "Version     : " & mVersion & vbCrLf & _
          "User        : " & Utl_Logger.CurrentUser() & vbCrLf & _
          "Host        : " & Utl_Logger.HostName() & vbCrLf & _
          "Workbook    : " & Utl_Logger.ThisWorkbookPathSafe() & vbCrLf & _
          String$(80, "=") & vbCrLf & vbCrLf
    Utl_Logger.AppendText mLogFilePath, hdr
End Sub

Private Sub WriteSessionFooter(ByVal errorCount As Long, ByVal warnCount As Long)
    Dim elapsed As String
    elapsed = Utl_Logger.FormatElapsed(mSessionStart, Now)
    
    Dim ftr As String
    ftr = vbCrLf & String$(80, "=") & vbCrLf & _
          "[Session End]   " & Utl_Logger.Timestamp() & vbCrLf & _
          "Elapsed Time    : " & elapsed & vbCrLf & _
          "Errors          : " & CStr(errorCount) & vbCrLf & _
          "Warnings        : " & CStr(warnCount) & vbCrLf & _
          String$(80, "=") & vbCrLf
    Utl_Logger.AppendText mLogFilePath, ftr
End Sub

'=== ログ出力 ===
Public Sub Log(ByVal Level As LogLevel, ByVal message As String)
    If Level < mMinLevel Then Exit Sub
    
    ' サイズ監視：超過していたら削除し、ヘッダを書き直す
    If Utl_Logger.IsFileTooLarge(mLogFilePath, mMaxFileSizeBytes) Then
        Utl_Logger.SafeKill mLogFilePath
        WriteSessionHeader
        ' 再作成後の先頭に通知
        Utl_Logger.AppendText mLogFilePath, Utl_Logger.Timestamp() & " [INFO ] (Logger) Log file was oversized and recreated." & vbCrLf
    End If
    
    Dim line As String
    Dim lvl As String: lvl = PadLevel(Level)
    Dim indent As String
    If mUseIndent And mIndentLevel > 0 Then
        indent = String$(mIndentLevel, " ")
    Else
        indent = ""
    End If
    
    line = Utl_Logger.Timestamp() & " " & lvl & " " & indent & message & vbCrLf
    Utl_Logger.AppendText mLogFilePath, line
End Sub

Public Sub LogDebug(ByVal message As String)
    Log LOG_DEBUG, message
End Sub

Public Sub LogInfo(ByVal message As String)
    Log LOG_INFO, message
End Sub

Public Sub LogWarn(ByVal message As String)
    Log LOG_WARN, message
End Sub

Public Sub LogError(ByVal message As String)
    Log LOG_ERROR, message
End Sub


Private Function PadLevel(ByVal Level As LogLevel) As String
    Select Case Level
        Case LOG_DEBUG: PadLevel = "[DEBUG]"
        Case LOG_INFO:  PadLevel = "[INFO ]"
        Case LOG_WARN:  PadLevel = "[WARN ]"
        Case LOG_ERROR: PadLevel = "[ERROR]"
        Case Else:      PadLevel = "[INFO ]"
    End Select
End Function

'=== インデント制御（必要に応じて使用） ===
' Enter/Exitは使う場面だけ呼べばOK。使わなければフラットなログのまま。
Public Sub EnterMethod(ByVal methodName As String)
    If mUseIndent Then
        Log LOG_DEBUG, "Enter: " & methodName
        mIndentLevel = mIndentLevel + mIndentStep
    Else
        Log LOG_DEBUG, "Enter: " & methodName
    End If
End Sub

Public Sub ExitMethod(ByVal methodName As String, Optional ByVal elapsedSeconds As Double = -1)
    If mUseIndent Then
        mIndentLevel = Application.WorksheetFunction.Max(0, mIndentLevel - mIndentStep)
    End If
    
    If elapsedSeconds >= 0 Then
        Log LOG_DEBUG, "Exit : " & methodName & " (Elapsed " & Format(elapsedSeconds, "0.000") & "s)"
    Else
        Log LOG_DEBUG, "Exit : " & methodName
    End If
End Sub

' ログをゼロから作り直す（既存ファイルを削除→ヘッダ再出力）
Public Sub ResetLog()
    Debug.Print mLogFilePath
    Utl_Logger.SafeKill mLogFilePath
    WriteSessionHeader
End Sub

' ローテーション（現在のログをリネームして退避→新規作成）
Public Sub RotateLog(Optional ByVal suffix As String = "")
    Dim bak As String
    bak = mLogFilePath & "." & IIf(Len(suffix) > 0, suffix, Format(Now, "yyyymmdd_hhnnss"))
    On Error Resume Next
    Name mLogFilePath As bak
    On Error GoTo 0
    WriteSessionHeader
End Sub

' ディレクトリ内の古いログを一括削除（テスト前の掃除用）
Public Sub PurgeLogs(Optional ByVal days As Long = 0)
    Dim dirPath As String, f As String
    dirPath = Left$(mLogFilePath, InStrRev(mLogFilePath, "\"))
    f = Dir$(dirPath & "*.log")
    Do While Len(f) > 0
        If days <= 0 Then
            Utl_Logger.SafeKill dirPath & f
        Else
            If FileDateTime(dirPath & f) < (Now - days) Then Utl_Logger.SafeKill dirPath & f
        End If
        f = Dir$
    Loop
    ' 新規ログを先頭から開始
    WriteSessionHeader
End Sub


Public Property Let UseIndent(ByVal v As Boolean): mUseIndent = v: End Property
Public Property Get UseIndent() As Boolean: UseIndent = mUseIndent: End Property

Public Property Let MinLevel(ByVal v As LogLevel): mMinLevel = v: End Property
Public Property Get MinLevel() As LogLevel: MinLevel = mMinLevel: End Property

Public Property Let MaxFileSizeBytes(ByVal v As Long): mMaxFileSizeBytes = v: End Property
Public Property Get MaxFileSizeBytes() As Long: MaxFileSizeBytes = mMaxFileSizeBytes: End Property

Public Property Get LogFilePath() As String: LogFilePath = mLogFilePath: End Property

'=== セッション終了（任意で件数サマリを渡せる） ===
Public Sub CloseLogger(Optional ByVal errorCount As Long = 0, Optional ByVal warnCount As Long = 0)
    WriteSessionFooter errorCount, warnCount
End Sub




