Attribute VB_Name = "SQLite3_Logger"
'==============================================================================
' SQLite3_Logger.bas  -  Structured logging subsystem (64-bit only)
'
' Log levels (lowest to highest severity):
'   LOG_DEBUG   = 0   verbose tracing (cache hits, SQL text, open/close)
'   LOG_INFO    = 1   normal operational events (transactions, checkpoints)
'   LOG_WARN    = 2   recoverable problems (SQLITE_BUSY, retry, degraded mode)
'   LOG_ERROR   = 3   hard failures that raise VBA errors
'   LOG_NONE    = 4   suppress all output
'
' Quick start:
'   ' Log INFO and above to the Immediate window
'   Logger_Configure LOG_INFO
'
'   ' Log DEBUG and above to both Immediate window and a file
'   Logger_Configure LOG_DEBUG, True, True, "C:\sqlite\driver.log"
'
'   ' From any module:
'   Logger_Info  "MyModule", "Connection opened"
'   Logger_Debug "MyModule", "Cache hit for: SELECT ..."
'   Logger_Warn  "MyModule", "SQLITE_BUSY, retrying"
'   Logger_Error "MyModule", "Fatal: " & Err.Description
'
'   ' Cheap guard -- skip string building when level is below threshold:
'   If Logger_IsEnabled(LOG_DEBUG) Then
'       Logger_Debug "MyModule", "rs=" & rs.RecordCount & " rows"
'   End If
'
'   ' Flush and close the file sink (call before workbook close):
'   Logger_Close
'
' Output format:
'   [2026-03-12 14:23:01.456] [DEBUG] [Source          ] Message
'
' Version : 0.1.5
'
' Version History:
'   0.1.5 - Initial release.
'
'
'    Copyright (C) 2026  Bryan Mark (bryan.mark@gmail.com)
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'==============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Log level constants  (public so callers can reference by name)
' ---------------------------------------------------------------------------
Public Const LOG_DEBUG As Long = 0
Public Const LOG_INFO  As Long = 1
Public Const LOG_WARN  As Long = 2
Public Const LOG_ERROR As Long = 3
Public Const LOG_NONE  As Long = 4

' ---------------------------------------------------------------------------
' Module-level state
' ---------------------------------------------------------------------------
Private m_level       As Long      ' minimum level that produces output
Private m_toImmediate As Boolean   ' write to Immediate window (Debug.Print)
Private m_toFile      As Boolean   ' write to log file
Private m_filePath    As String    ' path to the log file
Private m_fileNum     As Integer   ' VBA file number (0 = not open)
Private m_isOpen      As Boolean   ' True once Configure has been called

' ---------------------------------------------------------------------------
' Logger_Configure
' Call once at startup (or after changing settings) to initialise the logger.
'
' level       : minimum severity to emit.  Events below this are discarded.
' toImmediate : True to echo each line to the Immediate window (default True).
' toFile      : True to append each line to filePath (default False).
' filePath    : path of the log file.  Ignored when toFile is False.
'               The file is opened for append (existing content is preserved).
'               Pass "" to disable the file sink.
' ---------------------------------------------------------------------------
Public Sub Logger_Configure(Optional ByVal level As Long = LOG_INFO, _
                              Optional ByVal toImmediate As Boolean = True, _
                              Optional ByVal toFile As Boolean = False, _
                              Optional ByVal filePath As String = "")
    ' Close any previously open file sink before reconfiguring
    Logger_Close

    m_level       = level
    m_toImmediate = toImmediate
    m_toFile      = toFile And Len(filePath) > 0
    m_filePath    = filePath
    m_isOpen      = True

    If m_toFile Then
        On Error GoTo FileOpenFail
        m_fileNum = FreeFile()
        Open m_filePath For Append As #m_fileNum
        On Error GoTo 0
    End If

    ' Emit the banner so the log file has a visible session boundary
    Dim banner As String
    banner = String(70, "-")
    WriteRaw banner
    WriteRaw "[" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "] " & _
             "SQLite3 Logger started  level=" & LevelName(m_level)
    WriteRaw banner
    Exit Sub

FileOpenFail:
    m_toFile  = False
    m_fileNum = 0
    On Error GoTo 0
    Debug.Print "SQLite3_Logger WARNING: could not open log file: " & filePath
End Sub

' ---------------------------------------------------------------------------
' Logger_SetLevel
' Change the minimum log level without reconfiguring sinks.
' ---------------------------------------------------------------------------
Public Sub Logger_SetLevel(ByVal level As Long)
    If Not m_isOpen Then Logger_Configure level: Exit Sub
    m_level = level
End Sub

' ---------------------------------------------------------------------------
' Logger_IsEnabled
' Returns True if the given level will produce output.
' Use as a cheap guard to avoid building expensive message strings:
'   If Logger_IsEnabled(LOG_DEBUG) Then Logger_Debug "Mod", "rs=" & ...
' ---------------------------------------------------------------------------
Public Function Logger_IsEnabled(ByVal level As Long) As Boolean
    ' LOG_NONE is a suppression sentinel, not a real level -- never "enabled"
    If level >= LOG_NONE Then Logger_IsEnabled = False: Exit Function
    Logger_IsEnabled = m_isOpen And (level >= m_level)
End Function

' ---------------------------------------------------------------------------
' Named-level convenience wrappers
' ---------------------------------------------------------------------------
Public Sub Logger_Debug(ByVal source As String, ByVal msg As String)
    Logger_Log LOG_DEBUG, source, msg
End Sub

Public Sub Logger_Info(ByVal source As String, ByVal msg As String)
    Logger_Log LOG_INFO, source, msg
End Sub

Public Sub Logger_Warn(ByVal source As String, ByVal msg As String)
    Logger_Log LOG_WARN, source, msg
End Sub

Public Sub Logger_Error(ByVal source As String, ByVal msg As String)
    Logger_Log LOG_ERROR, source, msg
End Sub

' ---------------------------------------------------------------------------
' Logger_Log  -- central dispatch
' Formats one log line and routes it to enabled sinks.
' ---------------------------------------------------------------------------
Public Sub Logger_Log(ByVal level As Long, _
                       ByVal source As String, _
                       ByVal msg As String)
    If Not m_isOpen Then Exit Sub
    If level < m_level Then Exit Sub

    ' Build timestamp with milliseconds
    ' Timer() gives seconds since midnight; the fractional part gives ms.
    Dim t As Double: t = Timer()
    Dim ms As Long:  ms = CLng((t - Int(t)) * 1000)
    Dim ts As String
    ts = Format(Now(), "yyyy-mm-dd hh:mm:ss") & "." & Format(ms, "000")

    ' Pad/truncate source to a fixed width for columnar alignment
    Const SRC_WIDTH As Long = 24
    Dim src As String
    If Len(source) >= SRC_WIDTH Then
        src = Left(source, SRC_WIDTH)
    Else
        src = source & Space(SRC_WIDTH - Len(source))
    End If

    Dim line As String
    line = "[" & ts & "] [" & LevelName(level) & "] [" & src & "] " & msg

    WriteRaw line
End Sub

' ---------------------------------------------------------------------------
' Logger_Close
' Flush and close the file sink.  The Immediate-window sink needs no cleanup.
' Safe to call multiple times.
' ---------------------------------------------------------------------------
Public Sub Logger_Close()
    If m_fileNum <> 0 Then
        On Error Resume Next
        WriteRaw "[" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "] Logger closed"
        Close #m_fileNum
        On Error GoTo 0
        m_fileNum = 0
    End If
    m_isOpen = False
    m_toFile = False
End Sub

' ---------------------------------------------------------------------------
' Logger_GetLevel  -- read the current minimum level
' ---------------------------------------------------------------------------
Public Function Logger_GetLevel() As Long
    Logger_GetLevel = m_level
End Function

' ---------------------------------------------------------------------------
' Logger_GetFilePath  -- read the current file path (empty if none)
' ---------------------------------------------------------------------------
Public Function Logger_GetFilePath() As String
    Logger_GetFilePath = m_filePath
End Function

' ===========================================================================
' Private helpers
' ===========================================================================

' Write a raw (pre-formatted) line to all enabled sinks.
Private Sub WriteRaw(ByVal line As String)
    If m_toImmediate Then Debug.Print line
    If m_toFile And m_fileNum <> 0 Then
        On Error Resume Next
        Print #m_fileNum, line
        On Error GoTo 0
    End If
End Sub

' Return the fixed-width display name for a level constant.
Private Function LevelName(ByVal level As Long) As String
    Select Case level
        Case LOG_DEBUG: LevelName = "DEBUG"
        Case LOG_INFO:  LevelName = "INFO "
        Case LOG_WARN:  LevelName = "WARN "
        Case LOG_ERROR: LevelName = "ERROR"
        Case LOG_NONE:  LevelName = "NONE "
        Case Else:      LevelName = "?????"
    End Select
End Function
