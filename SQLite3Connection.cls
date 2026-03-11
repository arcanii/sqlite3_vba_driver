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
' Version : 0.1.2
'
' Version History:
'   0.1.0 - Initial release. OpenDatabase, CloseConnection, WAL auto-enable,
'            mmap, 64-slot LRU statement cache, Begin/Commit/Rollback,
'            LastInsertRowID, ChangesCount, errmsg surface.
'   0.1.1 - CloseConnection and ClearStatementCache guard SQLite_IsLoaded().
'            Fixed CLngPtr(0) for null args in sqlite3_open_v2 call.
'   0.1.2 - No functional changes. Version stamp updated.
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

' -- Statement cache entry type -----------------------------------------------
Private Type StmtCacheEntry
    hStmt    As LongPtr
    sqlText  As String
    lastUsed As Double
End Type

Private Const STMT_CACHE_SIZE As Long = 64

' -- Instance state -----------------------------------------------------------
Private m_pDb        As LongPtr
Private m_dllPath    As String
Private m_dbPath     As String
Private m_busyMs     As Long
Private m_inTx       As Boolean
Private m_cache()    As StmtCacheEntry
Private m_cacheCount As Long
Private m_lastError  As String
Private m_isOpen     As Boolean

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

Public Sub OpenDatabase(ByVal dbPath As String, _
                        Optional ByVal dllPath As String = "", _
                        Optional ByVal busyTimeoutMs As Long = 5000, _
                        Optional ByVal enableWAL As Boolean = True, _
                        Optional ByVal mmapSize As Long = 0)

    If m_isOpen Then CloseConnection

    If Len(dllPath) > 0 Then m_dllPath = dllPath
    If Not SQLite3_API.SQLite_IsLoaded() Then
        If Len(m_dllPath) = 0 Then m_dllPath = ThisWorkbook.Path & "\sqlite3.dll"
        SQLite3_API.SQLite_Load m_dllPath
    End If

    m_dbPath = dbPath
    m_busyMs = busyTimeoutMs

    Dim pathBytes() As Byte
    pathBytes = SQLite3_API.ToUTF8(dbPath)

    Dim flags As Long
    flags = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE Or SQLITE_OPEN_FULLMUTEX

    Dim rc As Long
    rc = sqlite3_open_v2(VarPtr(pathBytes(0)), m_pDb, flags, CLngPtr(0))
    CheckRC rc, "sqlite3_open_v2"

    sqlite3_busy_timeout m_pDb, m_busyMs
    m_isOpen = True

    If enableWAL Then
        ExecSQL "PRAGMA journal_mode=WAL;"
        ExecSQL "PRAGMA wal_autocheckpoint=1000;"
    End If

    If mmapSize > 0 Then ExecSQL "PRAGMA mmap_size=" & CStr(mmapSize) & ";"

    ExecSQL "PRAGMA synchronous=NORMAL;"
    ExecSQL "PRAGMA cache_size=-65536;"
    ExecSQL "PRAGMA temp_store=MEMORY;"
    ExecSQL "PRAGMA locking_mode=NORMAL;"
End Sub

Public Sub CloseConnection()
    If Not m_isOpen Then Exit Sub
    m_isOpen = False
    If Not SQLite3_API.SQLite_IsLoaded() Then
        m_pDb = 0
        m_cacheCount = 0
        Exit Sub
    End If
    ClearStatementCache
    If m_pDb <> 0 Then
        sqlite3_close m_pDb
        m_pDb = 0
    End If
End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = m_isOpen
End Property

Public Property Get Handle() As LongPtr
    Handle = m_pDb
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property

'==============================================================================
' Statement cache  (LRU eviction)
'==============================================================================
Friend Function GetCachedStatement(ByVal sql As String) As LongPtr
    Dim i As Long, oldest As Long, oldestTick As Double
    oldestTick = 1E+308
    For i = 0 To m_cacheCount - 1
        If m_cache(i).sqlText = sql Then
            m_cache(i).lastUsed = Timer()
            sqlite3_reset m_cache(i).hStmt
            sqlite3_clear_bindings m_cache(i).hStmt
            GetCachedStatement = m_cache(i).hStmt
            Exit Function
        End If
        If m_cache(i).lastUsed < oldestTick Then
            oldestTick = m_cache(i).lastUsed
            oldest = i
        End If
    Next i

    Dim pStmt As LongPtr, pTail As LongPtr
    Dim sqlBytes() As Byte
    sqlBytes = SQLite3_API.ToUTF8(sql)
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
    End If
    m_cache(slot).sqlText  = sql
    m_cache(slot).hStmt    = pStmt
    m_cache(slot).lastUsed = Timer()
    GetCachedStatement = pStmt
End Function

Private Sub ClearStatementCache()
    If Not SQLite3_API.SQLite_IsLoaded() Then
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
End Sub

'==============================================================================
' Execute (fire-and-forget SQL, no result set)
'==============================================================================
Public Sub ExecSQL(ByVal sql As String)
    AssertOpen
    Dim sqlBytes() As Byte
    sqlBytes = SQLite3_API.ToUTF8(sql)
    Dim pErr As LongPtr
    Dim rc As Long
    rc = sqlite3_exec(m_pDb, VarPtr(sqlBytes(0)), CLngPtr(0), CLngPtr(0), pErr)
    If rc <> SQLITE_OK Then
        m_lastError = sqlite3_errmsg_str(m_pDb)
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
End Sub

Public Sub BeginImmediate()
    ExecSQL "BEGIN IMMEDIATE;"
    m_inTx = True
End Sub

Public Sub CommitTransaction()
    ExecSQL "COMMIT;"
    m_inTx = False
End Sub

Public Sub RollbackTransaction()
    On Error Resume Next
    ExecSQL "ROLLBACK;"
    m_inTx = False
    On Error GoTo 0
End Sub

Public Property Get InTransaction() As Boolean
    InTransaction = m_inTx
End Property

'==============================================================================
' Query helper
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
        Err.Raise vbObjectError + 100, "SQLite3Connection." & context, _
                  "SQLite error " & rc & ": " & m_lastError
    End If
End Sub
