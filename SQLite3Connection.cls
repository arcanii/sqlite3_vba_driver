VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Connection.cls  -  HFT-Grade Connection Object (64-bit only)
' Features : WAL auto-enable, mmap, busy timeout, transaction helpers,
'            64-slot LRU statement cache, last_insert_rowid, error surface
'
' Version : 0.1.7
'
' Version History:
'   0.1.0 - Initial release. OpenDatabase, CloseConnection, WAL auto-enable,
'            mmap, 64-slot LRU statement cache, Begin/Commit/Rollback,
'            LastInsertRowID, ChangesCount, errmsg surface.
'   0.1.1 - CloseConnection and ClearStatementCache guard SQLite_IsLoaded().
'            Fixed CLngPtr(0) for null args in sqlite3_open_v2 call.
'   0.1.2 - No functional changes. Version stamp updated.
'   0.1.3 - Added Savepoint, ReleaseSavepoint, RollbackToSavepoint,
'            SavepointDepth, Interrupt. Added m_savepointDepth state.
'   0.1.4 - Added DllPath property (needed by SQLite3_Serialize.InMemoryClone).
'   0.1.5 - LRU timestamps: Timer() replaced with QueryPerformanceCounter.
'            Added openReadOnly parameter to OpenDatabase; IsReadOnly property.
'            Added Checkpoint() -- wraps sqlite3_wal_checkpoint_v2.
'            Integrated SQLite3_Logger: key events emitted at DEBUG/INFO/WARN/
'            ERROR level. Logger calls are zero-cost when logger is unconfigured
'            or the event level is below the current threshold.
'   0.1.6 - Added Tag property: optional label included in all logger output,
'            useful for identifying connections in multi-connection workbooks.
'            Added ExecScriptFile(path): reads a .sql file and passes the entire
'            content to sqlite3_exec, which handles multiple semicolon-delimited
'            statements and SQL comments natively.
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

' QPC for LRU timestamps -- ~100 ns resolution, no midnight wrap.
' Private so it does not collide with the Public declaration in SQLite3_Tests.
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As LongLong) As Long

' -- Statement cache entry ---------------------------------------------------
Private Type StmtCacheEntry
    hStmt    As LongPtr
    sqlText  As String
    lastUsed As LongLong   ' QPC ticks: monotonic, no midnight wrap
End Type

Private Const STMT_CACHE_SIZE As Long = 64

' -- Instance state -----------------------------------------------------------
Private m_pDb            As LongPtr
Private m_dllPath        As String
Private m_dbPath         As String
Private m_tag            As String     ' optional label shown in all log output
Private m_busyMs         As Long
Private m_inTx           As Boolean
Private m_isReadOnly     As Boolean
Private m_savepointDepth As Long
Private m_cache()        As StmtCacheEntry
Private m_cacheCount     As Long
Private m_lastError      As String
Private m_isOpen         As Boolean

'==============================================================================
' LogSrc -- dynamic log-source tag
' Returns "SQLite3Connection[<tag>]" when Tag is set, else "SQLite3Connection".
' Used as the source argument in every Logger_* call in this class.
'==============================================================================
Private Function LogSrc() As String
    If Len(m_tag) > 0 Then
        LogSrc = "SQLite3Connection[" & m_tag & "]"
    Else
        LogSrc = "SQLite3Connection"
    End If
End Function

'==============================================================================
' QPC helper -- monotonic tick count as LongLong
'==============================================================================
Private Function TickNow() As LongLong
    Dim t As LongLong
    QueryPerformanceCounter t
    TickNow = t
End Function

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Initialize()
    m_busyMs     = 5000
    m_cacheCount = 0
    ReDim m_cache(STMT_CACHE_SIZE - 1)
    m_isOpen = False
End Sub

Private Sub Class_Terminate()
    CloseConnection
End Sub

'==============================================================================
' OpenDatabase
'==============================================================================
Public Sub OpenDatabase(ByVal dbPath As String, _
                        Optional ByVal dllPath As String = "", _
                        Optional ByVal busyTimeoutMs As Long = 5000, _
                        Optional ByVal enableWAL As Boolean = True, _
                        Optional ByVal mmapSize As Long = 0, _
                        Optional ByVal openReadOnly As Boolean = False)

    If m_isOpen Then CloseConnection

    If Len(dllPath) > 0 Then m_dllPath = dllPath
    If Not SQLite3_CoreAPI.SQLite_IsLoaded() Then
        If Len(m_dllPath) = 0 Then m_dllPath = ThisWorkbook.Path & "\sqlite3.dll"
        SQLite3_CoreAPI.SQLite_Load m_dllPath
    End If

    m_dbPath     = dbPath
    m_busyMs     = busyTimeoutMs
    m_isReadOnly = openReadOnly

    Dim pathBytes() As Byte
    pathBytes = SQLite3_CoreAPI.ToUTF8(dbPath)

    Dim flags As Long
    If openReadOnly Then
        flags = SQLITE_OPEN_READONLY Or SQLITE_OPEN_FULLMUTEX
    Else
        flags = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE Or SQLITE_OPEN_FULLMUTEX
    End If

    Dim rc As Long
    rc = sqlite3_open_v2(VarPtr(pathBytes(0)), m_pDb, flags, CLngPtr(0))
    CheckRC rc, "sqlite3_open_v2"

    sqlite3_busy_timeout m_pDb, m_busyMs
    m_isOpen = True

    ' WAL and autocheckpoint require write access.
    If enableWAL And Not openReadOnly Then
        ExecSQL "PRAGMA journal_mode=WAL;"
        ExecSQL "PRAGMA wal_autocheckpoint=1000;"
    End If

    If mmapSize > 0 Then ExecSQL "PRAGMA mmap_size=" & CStr(mmapSize) & ";"

    ExecSQL "PRAGMA synchronous=NORMAL;"
    ExecSQL "PRAGMA cache_size=-65536;"
    ExecSQL "PRAGMA temp_store=MEMORY;"
    If Not openReadOnly Then ExecSQL "PRAGMA locking_mode=NORMAL;"

    Logger_Info LogSrc(), "Opened " & IIf(openReadOnly, "[READ-ONLY] ", "") & _
                          "db=" & dbPath
End Sub

'==============================================================================
' CloseConnection
'==============================================================================
Public Sub CloseConnection()
    If Not m_isOpen Then Exit Sub
    m_isOpen = False
    If Not SQLite3_CoreAPI.SQLite_IsLoaded() Then
        m_pDb = 0
        m_cacheCount = 0
        Exit Sub
    End If
    ClearStatementCache
    If m_pDb <> 0 Then
        sqlite3_close m_pDb
        m_pDb = 0
    End If
    Logger_Info LogSrc(), "Closed db=" & m_dbPath
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get IsOpen() As Boolean
    IsOpen = m_isOpen
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = m_isReadOnly
End Property

Public Property Get Tag() As String
    Tag = m_tag
End Property

Public Property Let Tag(ByVal value As String)
    m_tag = value
End Property

Public Property Get Handle() As LongPtr
    Handle = m_pDb
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property

Public Property Get InTransaction() As Boolean
    InTransaction = m_inTx
End Property

Public Property Get SavepointDepth() As Long
    SavepointDepth = m_savepointDepth
End Property

'==============================================================================
' Statement cache  (64-slot LRU, QPC timestamps)
'==============================================================================
Friend Function GetCachedStatement(ByVal sql As String) As LongPtr
    Dim i As Long
    Dim oldest As Long
    ' The ^ suffix is mandatory: without it VBA treats &H7FFFFFFFFFFFFFFF as a
    ' 32-bit Long literal, which overflows and produces a compile error.
    Dim oldestTick As LongLong: oldestTick = &H7FFFFFFFFFFFFFFF^

    For i = 0 To m_cacheCount - 1
        If m_cache(i).sqlText = sql Then
            m_cache(i).lastUsed = TickNow()
            sqlite3_reset m_cache(i).hStmt
            sqlite3_clear_bindings m_cache(i).hStmt
            If Logger_IsEnabled(LOG_DEBUG) Then
                Logger_Debug LogSrc(), "Cache HIT  slot=" & i & _
                             " sql=" & Left(sql, 80)
            End If
            GetCachedStatement = m_cache(i).hStmt
            Exit Function
        End If
        If m_cache(i).lastUsed < oldestTick Then
            oldestTick = m_cache(i).lastUsed
            oldest = i
        End If
    Next i

    ' Cache miss -- compile and store a new prepared statement
    Dim pStmt As LongPtr, pTail As LongPtr
    Dim sqlBytes() As Byte
    sqlBytes = SQLite3_CoreAPI.ToUTF8(sql)
    Dim rc As Long
    rc = sqlite3_prepare_v2(m_pDb, VarPtr(sqlBytes(0)), -1, pStmt, pTail)
    CheckRC rc, "sqlite3_prepare_v2"

    Dim slot As Long
    If m_cacheCount < STMT_CACHE_SIZE Then
        slot = m_cacheCount
        m_cacheCount = m_cacheCount + 1
    Else
        sqlite3_finalize m_cache(oldest).hStmt
        slot = oldest
        If Logger_IsEnabled(LOG_DEBUG) Then
            Logger_Debug LogSrc(), "Cache EVICT slot=" & slot & _
                         " evicted=" & Left(m_cache(slot).sqlText, 60)
        End If
    End If

    m_cache(slot).sqlText  = sql
    m_cache(slot).hStmt    = pStmt
    m_cache(slot).lastUsed = TickNow()

    If Logger_IsEnabled(LOG_DEBUG) Then
        Logger_Debug LogSrc(), "Cache MISS slot=" & slot & _
                     " sql=" & Left(sql, 80)
    End If
    GetCachedStatement = pStmt
End Function

Private Sub ClearStatementCache()
    If Not SQLite3_CoreAPI.SQLite_IsLoaded() Then
        m_cacheCount = 0
        Exit Sub
    End If
    Dim i As Long
    For i = 0 To m_cacheCount - 1
        If m_cache(i).hStmt <> 0 Then
            sqlite3_finalize m_cache(i).hStmt
            m_cache(i).hStmt = 0
        End If
    Next i
    m_cacheCount = 0
    Logger_Debug LogSrc(), "Statement cache cleared"
End Sub

'==============================================================================
' ExecScriptFile  -  execute a .sql file containing one or more statements
'
' Reads the entire file at filePath and passes it to sqlite3_exec in one call.
' sqlite3_exec handles multiple semicolon-delimited statements and SQL comments
' (-- and /* */) natively, so no client-side parsing is needed.
'
' The file must be UTF-8 encoded (SQLite's native encoding). Windows ANSI and
' UTF-16 files will import incorrectly -- convert before calling if needed.
'==============================================================================
Public Sub ExecScriptFile(ByVal filePath As String)
    AssertOpen

    ' Read the file into a string
    Dim fNum As Integer: fNum = FreeFile
    Dim script As String
    Open filePath For Input As #fNum
    Dim line As String
    Do While Not EOF(fNum)
        Line Input #fNum, line
        script = script & line & vbLf
    Loop
    Close #fNum

    If Len(Trim(script)) = 0 Then Exit Sub

    ' Pass the whole script to sqlite3_exec -- it handles multiple statements
    Dim sqlBytes() As Byte: sqlBytes = SQLite3_CoreAPI.ToUTF8(script)
    Dim pErr As LongPtr
    Dim rc As Long
    rc = sqlite3_exec(m_pDb, VarPtr(sqlBytes(0)), CLngPtr(0), CLngPtr(0), pErr)
    If rc <> SQLITE_OK Then
        m_lastError = sqlite3_errmsg_str(m_pDb)
        Logger_Error LogSrc(), "ExecScriptFile rc=" & rc & _
                               " file=" & filePath & " err=" & m_lastError
        Err.Raise vbObjectError + 100, "SQLite3Connection.ExecScriptFile", _
                  "SQLite error " & rc & ": " & m_lastError & " - File: " & filePath
    End If

    Logger_Info LogSrc(), "ExecScriptFile OK file=" & filePath
End Sub

'==============================================================================
' ExecSQL  -  fire-and-forget, no result set
'==============================================================================
Public Sub ExecSQL(ByVal sql As String)
    AssertOpen
    Dim sqlBytes() As Byte
    sqlBytes = SQLite3_CoreAPI.ToUTF8(sql)
    Dim pErr As LongPtr
    Dim rc As Long
    rc = sqlite3_exec(m_pDb, VarPtr(sqlBytes(0)), CLngPtr(0), CLngPtr(0), pErr)
    If rc <> SQLITE_OK Then
        m_lastError = sqlite3_errmsg_str(m_pDb)
        Logger_Error LogSrc(), "ExecSQL rc=" & rc & " err=" & m_lastError & _
                               " sql=" & sql
        Err.Raise vbObjectError + 100, "SQLite3Connection.ExecSQL", _
                  "SQLite error " & rc & ": " & m_lastError & " - SQL: " & sql
    End If
End Sub

'==============================================================================
' Transaction helpers
'==============================================================================
Public Sub BeginTransaction()
    ExecSQL "BEGIN;"
    m_inTx = True
    Logger_Debug LogSrc(), "BEGIN"
End Sub

Public Sub BeginImmediate()
    ExecSQL "BEGIN IMMEDIATE;"
    m_inTx = True
    Logger_Debug LogSrc(), "BEGIN IMMEDIATE"
End Sub

Public Sub CommitTransaction()
    ExecSQL "COMMIT;"
    m_inTx = False
    Logger_Debug LogSrc(), "COMMIT"
End Sub

Public Sub RollbackTransaction()
    On Error Resume Next
    ExecSQL "ROLLBACK;"
    m_inTx = False
    On Error GoTo 0
    Logger_Warn LogSrc(), "ROLLBACK"
End Sub

'==============================================================================
' Savepoints  (nested transactions -- SQL:2003)
'==============================================================================
Public Sub Savepoint(ByVal spName As String)
    AssertOpen
    ExecSQL "SAVEPOINT [" & spName & "];"
    m_savepointDepth = m_savepointDepth + 1
    Logger_Debug LogSrc(), "SAVEPOINT [" & spName & "] depth=" & m_savepointDepth
End Sub

Public Sub ReleaseSavepoint(ByVal spName As String)
    AssertOpen
    ExecSQL "RELEASE SAVEPOINT [" & spName & "];"
    If m_savepointDepth > 0 Then m_savepointDepth = m_savepointDepth - 1
    Logger_Debug LogSrc(), "RELEASE [" & spName & "] depth=" & m_savepointDepth
End Sub

Public Sub RollbackToSavepoint(ByVal spName As String)
    AssertOpen
    ExecSQL "ROLLBACK TO SAVEPOINT [" & spName & "];"
    Logger_Warn LogSrc(), "ROLLBACK TO SAVEPOINT [" & spName & "]"
End Sub

'==============================================================================
' Interrupt
'==============================================================================
Public Sub Interrupt()
    If m_isOpen Then
        sqlite3_interrupt m_pDb
        Logger_Warn LogSrc(), "sqlite3_interrupt signalled"
    End If
End Sub

'==============================================================================
' Checkpoint
'==============================================================================
Public Function Checkpoint(Optional ByVal mode As String = "PASSIVE", _
                            Optional ByVal schema As String = "main") As Variant
    AssertOpen
    Dim eMode As Long
    Select Case UCase(Trim(mode))
        Case "FULL":     eMode = SQLITE_CHECKPOINT_FULL
        Case "RESTART":  eMode = SQLITE_CHECKPOINT_RESTART
        Case "TRUNCATE": eMode = SQLITE_CHECKPOINT_TRUNCATE
        Case Else:       eMode = SQLITE_CHECKPOINT_PASSIVE
    End Select

    Dim schBytes() As Byte: schBytes = SQLite3_CoreAPI.ToUTF8(schema)
    Dim nLog As Long, nCkpt As Long
    Dim rc As Long
    rc = sqlite3_wal_checkpoint_v2(m_pDb, VarPtr(schBytes(0)), eMode, _
                                    VarPtr(nLog), VarPtr(nCkpt))

    If rc <> SQLITE_OK And rc <> SQLITE_BUSY Then
        m_lastError = sqlite3_errmsg_str(m_pDb)
        Logger_Error LogSrc(), "Checkpoint(" & mode & ") rc=" & rc & _
                               " err=" & m_lastError
        Err.Raise vbObjectError + 110, "SQLite3Connection.Checkpoint", _
                  "sqlite3_wal_checkpoint_v2 failed (rc=" & rc & "): " & m_lastError
    End If

    Logger_Info LogSrc(), "Checkpoint(" & mode & ") pagesWritten=" & nLog & _
                          " pagesRemaining=" & nCkpt & _
                          IIf(rc = SQLITE_BUSY, " [BUSY]", "")
    Checkpoint = Array(nLog, nCkpt)
End Function

'==============================================================================
' OpenRecordset
'==============================================================================
Public Function OpenRecordset(ByVal sql As String) As SQLite3Recordset
    AssertOpen
    Dim rs As New SQLite3Recordset
    rs.OpenInternal Me, sql
    Set OpenRecordset = rs
End Function

'==============================================================================
' Scalar / utility
'==============================================================================
Public Function LastInsertRowID() As LongLong
    AssertOpen
    LastInsertRowID = sqlite3_last_insert_rowid(m_pDb)
End Function

Public Function ChangesCount() As Long
    AssertOpen
    ChangesCount = sqlite3_changes(m_pDb)
End Function

Public Function DbPath() As String
    DbPath = m_dbPath
End Function

Public Function DllPath() As String
    DllPath = m_dllPath
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Sub AssertOpen()
    If Not m_isOpen Then
        Err.Raise vbObjectError + 200, "SQLite3Connection", "Database is not open."
    End If
End Sub

Private Sub CheckRC(ByVal rc As Long, ByVal context As String)
    If rc <> SQLITE_OK Then
        m_lastError = sqlite3_errmsg_str(m_pDb)
        Logger_Error LogSrc(), context & " rc=" & rc & " err=" & m_lastError
        Err.Raise vbObjectError + 100, "SQLite3Connection." & context, _
                  "SQLite error " & rc & ": " & m_lastError
    End If
End Sub
