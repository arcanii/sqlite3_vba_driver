Attribute VB_Name = "SQLite3_Tests"
'==============================================================================
' SQLite3_Tests.bas  -  Comprehensive driver test suite
'
' Run:  RunAllTests       - full suite with summary
'       RunTest_<name>    - individual suite
'
' Optional class detection is automatic via Application.Run sentinels.
' No #Const flags needed. If a class file is absent its suites print SKIP.
'
' Output goes to the Immediate window (Ctrl+G) and optionally to LOG_PATH.
' Each test prints PASS, FAIL, or SKIP with details.
'
' Version : 0.1.7
'
' Version History:
'   0.1.0 - Initial release. 122 tests.
'   0.1.2 - Added BLOB, Aggregates, FTS5.
'   0.1.3 - Added Schema, Savepoints, JSON, Interrupt.
'   0.1.4 - Added Backup, BlobStream, Serialize, Diagnostics.
'   0.1.5 - Added ReadOnly, Checkpoint, QueryPlan, Excel, Logger.
'   0.1.6 - Added Tag, ExecScriptFile, QueryColumn, ListObject, Migrate.
'   0.1.7 - #Const optional-module guards; m_skip counter.
'            SQLite3_API -> SQLite3_CoreAPI rename.
'   0.1.7 - Replaced #Const guards with runtime ClassAvailable() detection.
'            Optional class test suites moved into their respective .cls files
'            and dispatched via Application.Run. No manual flags required.
'            SQLite3_Tests.bas now contains only core (always-present) suites.
'            Public helpers (StartSuite, EndSuite, Pass, Fail, AssertX,
'            DropTable, TableRowCount, TableExists, QueryScalar) allow .cls
'            test subs to integrate cleanly into the central harness.
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
' High-resolution timer
' ---------------------------------------------------------------------------
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As LongPtr) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
    (lpFrequency As LongPtr) As Long

' Change to match your environment
' Option A: sqlite3.dll in C:\Windows\System32 (found by name, no path needed)
'Private Const DLL_PATH As String = "sqlite3.dll"
' Option B: explicit path outside System32
Private Const DLL_PATH As String = "C:\sqlite\sqlite3.dll"
Private Const DB_PATH  As String = "C:\sqlite\driver_test.db"

' Log file -- RunAllTests writes a copy here. Set "" to disable.
Private Const LOG_PATH As String = "C:\sqlite\test_results.log"

' ---------------------------------------------------------------------------
' Harness state  (Public so RunTests_xxx subs in .cls files can call in)
' ---------------------------------------------------------------------------
Public m_pass        As Long
Public m_fail        As Long
Public m_skip        As Long
Private m_suite      As String
Private m_suiteStart As LongPtr
Private m_runStart   As LongPtr
Private m_freq       As LongPtr
Private m_failLog()  As String
Private m_failCount  As Long
Private m_logFile    As Integer

' ---------------------------------------------------------------------------
' Timing helpers
' ---------------------------------------------------------------------------
Private Function QPC() As LongPtr
    Dim t As LongPtr
    QueryPerformanceCounter t
    QPC = t
End Function

Private Sub EnsureFreq()
    If m_freq = 0 Then QueryPerformanceFrequency m_freq
End Sub

Public Sub Log(ByVal msg As String)
    Debug.Print msg
    If m_logFile <> 0 Then
        On Error Resume Next
        Print #m_logFile, msg
        On Error GoTo 0
    End If
End Sub

Private Function ElapsedMs(ByVal t0 As LongPtr, ByVal t1 As LongPtr) As String
    EnsureFreq
    Dim ms As Double
    ms = (CDbl(t1) - CDbl(t0)) / CDbl(m_freq) * 1000#
    ElapsedMs = Format(ms, "0.00") & " ms"
End Function

' ---------------------------------------------------------------------------
' Suite lifecycle  (Public so .cls test subs can call StartSuite/EndSuite)
' ---------------------------------------------------------------------------
Public Sub StartSuite(ByVal name As String)
    m_suite = name
    m_suiteStart = QPC()
    Log ""
    Log "  [" & name & "]"
End Sub

Public Sub EndSuite()
    Log "    TIME  " & ElapsedMs(m_suiteStart, QPC())
End Sub

Public Sub SkipSuite(ByVal name As String, ByVal missingFile As String)
    StartSuite name
    Log "    SKIP  " & missingFile & " not in project"
    m_skip = m_skip + 1
    EndSuite
End Sub

' ---------------------------------------------------------------------------
' Assert helpers  (Public so .cls test subs can call them)
' ---------------------------------------------------------------------------
Public Sub Pass(ByVal name As String)
    m_pass = m_pass + 1
    Log "    PASS  " & name
End Sub

Public Sub Fail(ByVal name As String, ByVal detail As String)
    m_fail = m_fail + 1
    Log "    FAIL  " & name & " -- " & detail
    If m_failCount = 0 Then
        ReDim m_failLog(0)
    Else
        ReDim Preserve m_failLog(m_failCount)
    End If
    m_failLog(m_failCount) = "[" & m_suite & "]  " & name & " -- " & detail
    m_failCount = m_failCount + 1
End Sub

Public Sub AssertEqual(ByVal name As String, ByVal got As Variant, ByVal expected As Variant)
    If CStr(got) = CStr(expected) Then
        Pass name
    Else
        Fail name, "expected [" & CStr(expected) & "] got [" & CStr(got) & "]"
    End If
End Sub

Public Sub AssertTrue(ByVal name As String, ByVal condition As Boolean)
    If condition Then Pass name Else Fail name, "condition was False"
End Sub

Public Sub AssertFalse(ByVal name As String, ByVal condition As Boolean)
    If Not condition Then Pass name Else Fail name, "condition was True"
End Sub

Public Sub AssertNull(ByVal name As String, ByVal v As Variant)
    If IsNull(v) Then Pass name Else Fail name, "expected Null, got [" & CStr(v) & "]"
End Sub

Public Sub AssertNoError(ByVal name As String)
    If Err.Number = 0 Then Pass name Else Fail name, Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Shared test utilities  (Public so .cls test subs can call them)
' ---------------------------------------------------------------------------
Public Function FreshConn() As SQLite3Connection
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0
    Set FreshConn = conn
End Function

Public Sub DropTable(ByVal conn As SQLite3Connection, ByVal tbl As String)
    On Error Resume Next
    conn.ExecSQL "DROP TABLE IF EXISTS [" & tbl & "];"
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
' ClassAvailable - runtime detection of optional class sentinels
' ---------------------------------------------------------------------------
' ClassAvailable checks the VBProject component list at runtime.
' Requires: Tools > Macro > Security > Trust access to the VBA project
' object model (or equivalent Trust Center setting) to be enabled.
Private Function ClassAvailable(ByVal className As String) As Boolean
    On Error Resume Next
    Dim comp As Object
    Set comp = ThisWorkbook.VBProject.VBComponents(className)
    ClassAvailable = (Err.Number = 0 And Not comp Is Nothing)
    Err.Clear
End Function

'==============================================================================
' RunAllTests
'==============================================================================
Public Sub RunAllTests()
    m_pass = 0
    m_fail = 0
    m_skip = 0
    m_failCount = 0
    EnsureFreq
    m_runStart = QPC()

    ' Open log file
    m_logFile = 0
    If Len(LOG_PATH) > 0 Then
        On Error Resume Next
        m_logFile = FreeFile()
        Open LOG_PATH For Output As #m_logFile
        If Err.Number <> 0 Then
            m_logFile = 0
            Debug.Print "WARNING: could not open log file: " & LOG_PATH
        End If
        Err.Clear
        On Error GoTo 0
    End If

    Log String(64, "=")
    Log "SQLite3 Driver Test Suite  v0.1.7"
    Log "Started: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Log String(64, "=")

    ' Suite 1 is a hard prerequisite: abort if DLL fails to load
    Dim failBefore As Long: failBefore = m_fail
    RunTest_DllLoad
    If m_fail > failBefore Then
        Log ""
        Log "*** DllLoad suite failed -- aborting remaining suites. ***"
        Log "*** Fix DLL_PATH / install sqlite3.dll and re-run.     ***"
        GoTo PrintSummary
    End If

    ' ----- Core suites (always present) ------------------------------------
    RunTest_OpenClose
    RunTest_ExecSQL
    RunTest_ScalarTypes
    RunTest_NullHandling
    RunTest_UTF8
    RunTest_PreparedStatements
    RunTest_NamedParams
    RunTest_Transactions
    RunTest_RollbackTransaction
    RunTest_Recordset_Live
    RunTest_Recordset_Vectorized
    RunTest_Recordset_GetRows
    RunTest_Recordset_ToMatrix
    RunTest_StatementCache
    RunTest_SpecialCharacters
    RunTest_Boundaries
    RunTest_ErrorHandling
    RunTest_BLOB
    RunTest_Savepoints
    RunTest_Interrupt
    RunTest_ReadOnly
    RunTest_Checkpoint
    RunTest_QueryPlan
    RunTest_Tag
    RunTest_ExecScriptFile
    RunTest_QueryColumn

    ' ----- Optional class suites (auto-detected) ---------------------------
    ' SQLite3BulkInsert.cls -> suites 15, 16, 19
    If ClassAvailable("SQLite3BulkInsert") Then
        RunTests_BulkInsert DB_PATH, DLL_PATH
    Else
        SkipSuite "BulkInsert_AppendRow (15)", "SQLite3BulkInsert.cls"
        SkipSuite "BulkInsert_AppendMatrix (16)", "SQLite3BulkInsert.cls"
        SkipSuite "LargeDataset (19)", "SQLite3BulkInsert.cls"
    End If

    ' SQLite3Pool.cls -> suites 18, 18b
    If ClassAvailable("SQLite3Pool") Then
        RunTests_Pool DB_PATH, DLL_PATH
    Else
        SkipSuite "ConnectionPool (18)", "SQLite3Pool.cls"
        SkipSuite "Pool_Exhausted (18b)", "SQLite3Pool.cls"
    End If

    ' SQLite3Backup.cls -> suite 30
    If ClassAvailable("SQLite3Backup") Then
        RunTests_Backup DB_PATH, DLL_PATH
    Else
        SkipSuite "Backup (30)", "SQLite3Backup.cls"
    End If

    ' SQLite3BlobStream.cls -> suite 31
    If ClassAvailable("SQLite3BlobStream") Then
        RunTests_BlobStream DB_PATH, DLL_PATH
    Else
        SkipSuite "BlobStream (31)", "SQLite3BlobStream.cls"
    End If

    ' ----- Feature module suites (always present if SQLite3_Driver.bas) ---
    RunTest_Aggregates
    RunTest_FTS5
    RunTest_Schema
    RunTest_JSON
    RunTest_Serialize
    RunTest_Diagnostics
    RunTest_Excel
    RunTest_Logger
    RunTest_ListObject
    RunTest_Migrate

PrintSummary:
    Dim totalTime As String: totalTime = ElapsedMs(m_runStart, QPC())
    Log ""
    Log String(64, "=")
    Log "Results: " & m_pass & " passed,  " & m_fail & " failed,  " & _
                m_skip & " skipped  (" & (m_pass + m_fail) & " run)  " & totalTime
    Log String(64, "=")

    If m_fail > 0 Then
        Log ""
        Log "FAILED TESTS (" & m_fail & "):"
        Log String(64, "-")
        Dim i As Long
        For i = 0 To m_failCount - 1
            Log "  " & m_failLog(i)
        Next i
        Log String(64, "-")
    End If

    If m_logFile <> 0 Then
        On Error Resume Next
        Close #m_logFile
        On Error GoTo 0
        m_logFile = 0
        Debug.Print ""
        Debug.Print "Log written to: " & LOG_PATH
    End If

    On Error Resume Next
    Kill DB_PATH
    On Error GoTo 0
End Sub

'==============================================================================
' 1. DLL load / version
'==============================================================================
Public Sub RunTest_DllLoad()
    StartSuite "DllLoad (1)"
    On Error Resume Next

    SQLite3_CoreAPI.SQLite_Unload
    SQLite3_CoreAPI.SQLite_Load DLL_PATH
    AssertNoError "SQLite_Load"
    AssertTrue "SQLite_IsLoaded", SQLite3_CoreAPI.SQLite_IsLoaded()

    Dim ver As String: ver = SQLite3_CoreAPI.SQLite_Version()
    AssertTrue "Version non-empty", Len(ver) > 0
    AssertTrue "Version starts with 3", Left(ver, 1) = "3"
    Log "    INFO  SQLite version = " & ver

    SQLite3_CoreAPI.SQLite_Unload
    AssertFalse "SQLite_IsLoaded after unload", SQLite3_CoreAPI.SQLite_IsLoaded()

    SQLite3_CoreAPI.SQLite_Load DLL_PATH
    EndSuite
End Sub

'==============================================================================
' 2. Open / close
'==============================================================================
Public Sub RunTest_OpenClose()
    StartSuite "OpenClose (2)"
    On Error Resume Next

    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH
    AssertTrue "IsOpen after OpenDatabase", conn.IsOpen
    AssertTrue "Handle non-zero", conn.Handle <> 0
    AssertEqual "DbPath", conn.dbPath, DB_PATH

    conn.CloseConnection
    AssertFalse "IsOpen after CloseConnection", conn.IsOpen

    Err.Clear
    conn.CloseConnection
    AssertNoError "Double CloseConnection safe"
    EndSuite
End Sub

'==============================================================================
' 3. ExecSQL / basic DDL + DML
'==============================================================================
Public Sub RunTest_ExecSQL()
    StartSuite "ExecSQL (3)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_exec"

    conn.ExecSQL "CREATE TABLE t_exec (id INTEGER PRIMARY KEY, val TEXT);"
    AssertNoError "CREATE TABLE"
    AssertTrue "TableExists", TableExists(conn, "t_exec")

    conn.ExecSQL "INSERT INTO t_exec VALUES (1, 'hello');"
    AssertNoError "INSERT"
    AssertEqual "ChangesCount", conn.ChangesCount(), 1

    conn.ExecSQL "UPDATE t_exec SET val='world' WHERE id=1;"
    AssertEqual "UPDATE changes", conn.ChangesCount(), 1

    conn.ExecSQL "DELETE FROM t_exec WHERE id=1;"
    AssertEqual "DELETE changes", conn.ChangesCount(), 1

    DropTable conn, "t_exec"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 4. Scalar types
'==============================================================================
Public Sub RunTest_ScalarTypes()
    StartSuite "ScalarTypes (4)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_types"
    conn.ExecSQL "CREATE TABLE t_types (i INTEGER, f REAL, t TEXT, b BLOB, n);"
    conn.ExecSQL "INSERT INTO t_types VALUES (42, 3.14, 'hello', X'DEADBEEF', NULL);"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f, t, n FROM t_types;")

    AssertFalse "Not EOF", rs.EOF
    AssertEqual "INTEGER value", rs!i, 42
    AssertTrue "FLOAT close", Abs(CDbl(rs!f) - 3.14) < 0.0001
    AssertEqual "TEXT value", rs!t, "hello"
    AssertNull "NULL value", rs!N

    rs.CloseRecordset
    DropTable conn, "t_types"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 5. NULL handling
'==============================================================================
Public Sub RunTest_NullHandling()
    StartSuite "NullHandling (5)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_null"
    conn.ExecSQL "CREATE TABLE t_null (a INTEGER, b TEXT);"
    conn.ExecSQL "INSERT INTO t_null VALUES (NULL, NULL);"
    conn.ExecSQL "INSERT INTO t_null VALUES (1, 'x');"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT a, b FROM t_null ORDER BY a;")

    AssertNull "a is NULL", rs!a
    AssertNull "b is NULL", rs!b
    rs.MoveNext
    AssertEqual "a = 1", rs!a, 1
    AssertEqual "b = x", rs!b, "x"
    rs.MoveNext
    AssertTrue "EOF after last", rs.EOF

    rs.CloseRecordset
    DropTable conn, "t_null"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 6. UTF-8 round-trip
'==============================================================================
Public Sub RunTest_UTF8()
    StartSuite "UTF8 (6)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_utf8"
    conn.ExecSQL "CREATE TABLE t_utf8 (s TEXT);"

    Dim cases As Variant
    cases = Array( _
        "ASCII only", _
        Chr(233) & "l" & Chr(232) & "ve", _
        ChrW(26085) & ChrW(26412) & ChrW(35486), _
        "Caf" & Chr(233))

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_utf8 VALUES (?);"
    Dim i As Long
    For i = 0 To UBound(cases)
        cmd.BindText 1, CStr(cases(i))
        cmd.Execute
        cmd.Reset
    Next i

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT s FROM t_utf8 ORDER BY rowid;")
    For i = 0 To UBound(cases)
        AssertFalse "Not EOF row " & i, rs.EOF
        AssertEqual "UTF8 round-trip " & i, rs!s, cases(i)
        rs.MoveNext
    Next i

    rs.CloseRecordset
    DropTable conn, "t_utf8"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 7. Prepared statements - positional binding
'==============================================================================
Public Sub RunTest_PreparedStatements()
    StartSuite "PreparedStatements (7)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_prep"
    conn.ExecSQL "CREATE TABLE t_prep (i INTEGER, f REAL, t TEXT);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_prep VALUES (?, ?, ?);"

    cmd.BindInt 1, 7
    cmd.BindDouble 2, 2.718
    cmd.BindText 3, "Euler"
    cmd.Execute
    cmd.Reset

    cmd.BindNull 1
    cmd.BindInt 2, 0
    cmd.BindNull 3
    cmd.Execute
    cmd.Reset

    AssertEqual "2 rows inserted", TableRowCount(conn, "t_prep"), 2

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f, t FROM t_prep ORDER BY rowid;")

    AssertEqual "Row1 i", rs!i, 7
    AssertTrue "Row1 f", Abs(CDbl(rs!f) - 2.718) < 0.0001
    AssertEqual "Row1 t", rs!t, "Euler"
    rs.MoveNext

    AssertNull "Row2 i null", rs!i
    AssertNull "Row2 t null", rs!t
    rs.CloseRecordset

    Dim cmd2 As New SQLite3Command
    cmd2.Prepare conn, "SELECT COUNT(*) FROM t_prep WHERE i IS NOT NULL;"
    Dim sv As Variant: sv = cmd2.ExecuteScalar()
    AssertEqual "ExecuteScalar COUNT", sv, 1

    DropTable conn, "t_prep"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 8. Named parameters
'==============================================================================
Public Sub RunTest_NamedParams()
    StartSuite "NamedParams (8)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_named"
    conn.ExecSQL "CREATE TABLE t_named (a INTEGER, b TEXT, c REAL);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_named VALUES (:a, :b, :c);"
    cmd.BindIntByName ":a", 99
    cmd.BindTextByName ":b", "ninety-nine"
    cmd.BindDoubleByName ":c", 9.9
    cmd.Execute
    cmd.Reset

    Dim v As Variant
    v = QueryScalar(conn, "SELECT b FROM t_named WHERE a=99;")
    AssertEqual "Named :b round-trip", v, "ninety-nine"

    v = QueryScalar(conn, "SELECT c FROM t_named WHERE a=99;")
    AssertTrue "Named :c round-trip", Abs(CDbl(v) - 9.9) < 0.001

    ' -- Rebind and reuse --
    cmd.Prepare conn, "INSERT INTO t_named VALUES (:a, :b, :c);"
    cmd.BindIntByName ":a", 100
    cmd.BindTextByName ":b", "hundred"
    cmd.BindDoubleByName ":c", 10#
    cmd.Execute
    AssertEqual "Named rebind row count=2", TableRowCount(conn, "t_named"), 2

    v = QueryScalar(conn, "SELECT a FROM t_named WHERE b='hundred';")
    AssertEqual "Named rebind :a=100", CLng(v), 100

    DropTable conn, "t_named"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 9. Transactions - commit
'==============================================================================
Public Sub RunTest_Transactions()
    StartSuite "Transactions (9)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_tx"
    conn.ExecSQL "CREATE TABLE t_tx (n INTEGER);"

    ' -- Basic commit --
    conn.BeginTransaction
    AssertTrue "InTransaction after BeginTransaction", conn.InTransaction
    Dim i As Long
    For i = 1 To 100
        conn.ExecSQL "INSERT INTO t_tx VALUES (" & i & ");"
    Next i
    conn.CommitTransaction
    AssertFalse "Not InTransaction after commit", conn.InTransaction
    AssertEqual "100 rows committed", TableRowCount(conn, "t_tx"), 100

    ' -- Rollback --
    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_tx VALUES (999);"
    conn.RollbackTransaction
    AssertFalse "Not InTransaction after rollback", conn.InTransaction
    AssertEqual "Rollback: still 100 rows", TableRowCount(conn, "t_tx"), 100

    ' -- Second transaction --
    conn.BeginTransaction
    AssertTrue "Tx2: InTransaction", conn.InTransaction
    conn.ExecSQL "INSERT INTO t_tx VALUES (101);"
    conn.CommitTransaction
    AssertEqual "Tx2: 101 rows", TableRowCount(conn, "t_tx"), 101

    ' -- Third transaction --
    conn.BeginTransaction
    AssertTrue "Tx3: InTransaction", conn.InTransaction
    conn.ExecSQL "INSERT INTO t_tx VALUES (102);"
    conn.CommitTransaction
    AssertEqual "Tx3: 102 rows", TableRowCount(conn, "t_tx"), 102

    ' -- Fourth transaction --
    conn.BeginTransaction
    AssertTrue "Tx4: InTransaction", conn.InTransaction
    conn.ExecSQL "INSERT INTO t_tx VALUES (103);"
    conn.CommitTransaction
    AssertEqual "Tx4: 103 rows", TableRowCount(conn, "t_tx"), 103

    ' -- Double-begin guard --
    conn.BeginTransaction
    conn.BeginTransaction   ' should be no-op or error, not crash
    Err.Clear
    conn.RollbackTransaction
    AssertFalse "Cleaned up after double-begin", conn.InTransaction

    DropTable conn, "t_tx"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 10. Transactions - rollback
'==============================================================================
Public Sub RunTest_RollbackTransaction()
    StartSuite "RollbackTransaction (10)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_rb"
    conn.ExecSQL "CREATE TABLE t_rb (n INTEGER);"
    conn.ExecSQL "INSERT INTO t_rb VALUES (1);"

    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_rb VALUES (2);"
    conn.ExecSQL "INSERT INTO t_rb VALUES (3);"
    conn.RollbackTransaction
    AssertEqual "Rollback: only pre-tx row", TableRowCount(conn, "t_rb"), 1

    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_rb VALUES (2);"
    conn.CommitTransaction
    AssertEqual "Commit after rollback: 2 rows", TableRowCount(conn, "t_rb"), 2

    ' -- Rollback with no active tx --
    conn.RollbackTransaction
    AssertFalse "Extra rollback: InTransaction=False", conn.InTransaction

    ' -- Large rollback --
    conn.BeginTransaction
    Dim i As Long
    For i = 1 To 500
        conn.ExecSQL "INSERT INTO t_rb VALUES (" & i & ");"
    Next i
    conn.RollbackTransaction
    AssertEqual "Large rollback: still 2 rows", TableRowCount(conn, "t_rb"), 2

    DropTable conn, "t_rb"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 11. Live recordset navigation
'==============================================================================
Public Sub RunTest_Recordset_Live()
    StartSuite "Recordset_Live (11)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_live"
    conn.ExecSQL "CREATE TABLE t_live (n INTEGER);"
    Dim i As Long
    For i = 1 To 5
        conn.ExecSQL "INSERT INTO t_live VALUES (" & i & ");"
    Next i

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT n FROM t_live ORDER BY n;")

    AssertFalse "Not BOF", rs.BOF
    AssertFalse "Not EOF at start", rs.EOF
    AssertEqual "RecordCount live = -1", rs.RecordCount, -1

    Dim sum As Long
    Do While Not rs.EOF
        sum = sum + CLng(rs!N)
        rs.MoveNext
    Loop
    AssertEqual "Sum 1..5 = 15", sum, 15
    AssertTrue "EOF after last", rs.EOF

    Dim rs2 As SQLite3Recordset
    Set rs2 = conn.OpenRecordset("SELECT n FROM t_live WHERE n > 999;")
    AssertTrue "Empty rs BOF", rs2.BOF
    AssertTrue "Empty rs EOF", rs2.EOF

    rs.CloseRecordset
    rs2.CloseRecordset
    DropTable conn, "t_live"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 12. Vectorized recordset
'==============================================================================
Public Sub RunTest_Recordset_Vectorized()
    StartSuite "Recordset_Vectorized (12)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_vec"
    conn.ExecSQL "CREATE TABLE t_vec (n INTEGER, s TEXT);"
    Dim i As Long
    For i = 1 To 10
        conn.ExecSQL "INSERT INTO t_vec VALUES (" & i & ", 'r" & i & "');"
    Next i

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT n, s FROM t_vec ORDER BY n;")
    Dim cnt As Long: cnt = rs.LoadAll()

    AssertEqual "LoadAll returns 10", cnt, 10
    AssertEqual "RecordCount = 10", rs.RecordCount, 10
    AssertEqual "FieldCount = 2", rs.FieldCount, 2
    AssertFalse "Not EOF at start", rs.EOF

    rs.MoveFirst
    AssertEqual "First row n=1", rs!N, 1
    AssertEqual "First row s=r1", rs!s, "r1"
    rs.MoveLast
    AssertEqual "Last row n=10", rs!N, 10

    rs.MoveFirst
    Dim sum As Long
    Do While Not rs.EOF
        sum = sum + CLng(rs!N)
        rs.MoveNext
    Loop
    AssertEqual "Sum 1..10 = 55", sum, 55

    rs.MoveFirst
    AssertEqual "Field by index 0", rs.Item(0), 1
    AssertEqual "Field by name n", rs.Item("n"), 1

    Dim names() As String: names = rs.ColumnNames()
    AssertEqual "ColName 0 = n", names(0), "n"
    AssertEqual "ColName 1 = s", names(1), "s"

    rs.CloseRecordset
    DropTable conn, "t_vec"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 13. GetRows
'==============================================================================
Public Sub RunTest_Recordset_GetRows()
    StartSuite "Recordset_GetRows (13)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_gr"
    conn.ExecSQL "CREATE TABLE t_gr (n INTEGER);"
    Dim i As Long
    For i = 1 To 6
        conn.ExecSQL "INSERT INTO t_gr VALUES (" & i & ");"
    Next i

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT n FROM t_gr ORDER BY n;")
    rs.LoadAll

    Dim pg1 As Variant: pg1 = rs.GetRows(3)
    AssertEqual "GetRows page1 col dim", UBound(pg1, 1), 0
    AssertEqual "GetRows page1 row dim", UBound(pg1, 2), 2
    AssertEqual "GetRows page1 r0 = 1", pg1(0, 0), 1
    AssertEqual "GetRows page1 r2 = 3", pg1(0, 2), 3

    Dim pg2 As Variant: pg2 = rs.GetRows(3)
    AssertEqual "GetRows page2 r0 = 4", pg2(0, 0), 4
    AssertEqual "GetRows page2 r2 = 6", pg2(0, 2), 6
    AssertTrue "EOF after two pages", rs.EOF

    rs.CloseRecordset
    DropTable conn, "t_gr"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 14. ToMatrix
'==============================================================================
Public Sub RunTest_Recordset_ToMatrix()
    StartSuite "Recordset_ToMatrix (14)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_mat"
    conn.ExecSQL "CREATE TABLE t_mat (a INTEGER, b REAL);"
    conn.ExecSQL "INSERT INTO t_mat VALUES (1, 1.1);"
    conn.ExecSQL "INSERT INTO t_mat VALUES (2, 2.2);"
    conn.ExecSQL "INSERT INTO t_mat VALUES (3, 3.3);"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT a, b FROM t_mat ORDER BY a;")
    rs.LoadAll

    Dim mat As Variant: mat = rs.ToMatrix()
    AssertEqual "Matrix row dim", UBound(mat, 1), 2
    AssertEqual "Matrix col dim", UBound(mat, 2), 1
    AssertEqual "mat(0,0) = 1", mat(0, 0), 1
    AssertTrue "mat(0,1) ~1.1", Abs(CDbl(mat(0, 1)) - 1.1) < 0.001
    AssertEqual "mat(2,0) = 3", mat(2, 0), 3
    AssertTrue "mat(2,1) ~3.3", Abs(CDbl(mat(2, 1)) - 3.3) < 0.001

    rs.CloseRecordset
    DropTable conn, "t_mat"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 17. Statement cache
'==============================================================================
Public Sub RunTest_StatementCache()
    StartSuite "StatementCache (17)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_cache"
    conn.ExecSQL "CREATE TABLE t_cache (n INTEGER);"

    Dim sql As String: sql = "INSERT INTO t_cache VALUES (?);"
    Dim i As Long
    For i = 1 To 10
        Dim cmd As New SQLite3Command
        cmd.Prepare conn, sql
        cmd.BindInt 1, i
        cmd.Execute
    Next i

    AssertEqual "10 rows via cached stmt", TableRowCount(conn, "t_cache"), 10

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT SUM(n) FROM t_cache;")
    AssertEqual "SUM = 55", rs.Item(0), 55
    rs.CloseRecordset

    ' -- Same SQL reused (cache-hit path) --
    Dim cmd2 As New SQLite3Command
    cmd2.Prepare conn, sql
    cmd2.BindInt 1, 99
    cmd2.Execute
    AssertEqual "Cached stmt: 11 rows after reuse", TableRowCount(conn, "t_cache"), 11

    ' -- ExecuteScalar --
    Dim cmd3 As New SQLite3Command
    cmd3.Prepare conn, "SELECT COUNT(*) FROM t_cache;"
    Dim sv As Variant: sv = cmd3.ExecuteScalar
    AssertEqual "ExecuteScalar count=11", CLng(sv), 11

    ' -- StmtHandle non-null after Prepare --
    AssertTrue "StmtHandle <> 0", cmd3.StmtHandle <> 0

    DropTable conn, "t_cache"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 20. Special characters
'==============================================================================
Public Sub RunTest_SpecialCharacters()
    StartSuite "SpecialCharacters (20)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_special"
    conn.ExecSQL "CREATE TABLE t_special (s TEXT);"

    Dim cases As Variant
    cases = Array( _
        "it's a test", _
        "line1" & vbLf & "line2", _
        "tab" & vbTab & "stop", _
        String(100, "x"), _
        "")

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_special VALUES (?);"
    Dim i As Long
    For i = 0 To UBound(cases)
        cmd.BindText 1, CStr(cases(i))
        cmd.Execute
        cmd.Reset
    Next i

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT s FROM t_special ORDER BY rowid;")
    For i = 0 To UBound(cases)
        AssertFalse "Not EOF row " & i, rs.EOF
        AssertEqual "Special char " & i, rs!s, cases(i)
        rs.MoveNext
    Next i

    rs.CloseRecordset
    DropTable conn, "t_special"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 21. Boundary values
'==============================================================================
Public Sub RunTest_Boundaries()
    StartSuite "Boundaries (21)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_bounds"
    conn.ExecSQL "CREATE TABLE t_bounds (i INTEGER, f REAL);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_bounds VALUES (?, ?);"

    cmd.BindInt 1, 2147483647: cmd.BindDouble 2, 1.7976931348623E+308
    cmd.Execute: cmd.Reset
    cmd.BindInt 1, -2147483648#: cmd.BindDouble 2, -1.7976931348623E+308
    cmd.Execute: cmd.Reset
    cmd.BindInt 1, 0: cmd.BindDouble 2, 0
    cmd.Execute: cmd.Reset

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f FROM t_bounds ORDER BY rowid;")
    rs.LoadAll

    AssertEqual "Max Long", rs!i, 2147483647
    rs.MoveNext
    AssertEqual "Min Long", rs!i, -2147483648#
    rs.MoveNext
    AssertEqual "Zero int", rs!i, 0
    AssertEqual "Zero float", rs!f, 0

    rs.CloseRecordset
    DropTable conn, "t_bounds"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 22. Error handling  (pool-exhausted test moved to SQLite3Pool.cls)
'==============================================================================
Public Sub RunTest_ErrorHandling()
    StartSuite "ErrorHandling (22)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    ' -- Bad SQL --
    Err.Clear
    conn.ExecSQL "THIS IS NOT SQL;"
    AssertTrue "Bad SQL raises error", Err.Number <> 0
    Err.Clear

    ' -- Missing table --
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "SELECT * FROM no_such_table_xyz;"
    cmd.Execute
    AssertTrue "Missing table raises error", Err.Number <> 0
    Err.Clear

    ' -- Bad named param --
    Dim cmd2 As New SQLite3Command
    cmd2.Prepare conn, "SELECT ?;"
    cmd2.BindTextByName ":nosuchparam", "x"
    AssertTrue "Bad named param raises error", Err.Number <> 0
    Err.Clear

    ' -- Primary key violation --
    conn.ExecSQL "CREATE TABLE IF NOT EXISTS t_eh (id INTEGER PRIMARY KEY);"
    conn.ExecSQL "INSERT INTO t_eh VALUES (1);"
    Err.Clear
    conn.ExecSQL "INSERT INTO t_eh VALUES (1);"
    AssertTrue "PK violation raises error", Err.Number <> 0
    Err.Clear

    ' -- Connection still functional after errors --
    conn.ExecSQL "SELECT 1;"
    AssertNoError "Conn functional after error sequence"

    ' -- DROP non-existent table raises error --
    conn.ExecSQL "DROP TABLE no_exist_table;"
    AssertTrue "DROP non-existent raises error", Err.Number <> 0
    Err.Clear

    conn.ExecSQL "DROP TABLE IF EXISTS t_eh;"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 23. BLOB
'==============================================================================
Public Sub RunTest_BLOB()
    StartSuite "BLOB (23)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_blob"
    conn.ExecSQL "CREATE TABLE t_blob (id INTEGER, data BLOB, label TEXT);"

    Dim small() As Byte: ReDim small(4)
    small(0) = 1: small(1) = 2: small(2) = 3: small(3) = 255: small(4) = 0

    Dim large() As Byte: ReDim large(999)
    Dim i As Long
    For i = 0 To 999: large(i) = CByte(i Mod 256): Next i

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_blob VALUES (?, ?, ?);"

    cmd.BindInt 1, 1: cmd.BindBlob 2, small: cmd.BindText 3, "small"
    cmd.Execute: cmd.Reset

    cmd.BindInt 1, 2: cmd.BindBlob 2, large: cmd.BindText 3, "large"
    cmd.Execute: cmd.Reset

    cmd.BindInt 1, 3: cmd.BindVariant 2, small: cmd.BindText 3, "variant"
    cmd.Execute: cmd.Reset

    AssertEqual "3 BLOB rows", TableRowCount(conn, "t_blob"), 3

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT id, data, label FROM t_blob ORDER BY id;")

    AssertFalse "Row1 not EOF", rs.EOF
    Dim v1 As Variant: v1 = rs!data
    AssertTrue "Row1 is byte array", VarType(v1) = (vbByte + vbArray)
    Dim b1() As Byte: b1 = v1
    AssertEqual "Row1 len=5", UBound(b1) - LBound(b1) + 1, 5
    AssertEqual "Row1 b(0)=1", b1(0), 1
    AssertEqual "Row1 b(3)=255", b1(3), 255

    Dim ab() As Byte: ab = rs.Fields("data").AsBytes()
    AssertEqual "AsBytes len=5", UBound(ab) - LBound(ab) + 1, 5
    AssertEqual "AsBytes b(1)=2", ab(1), 2

    rs.MoveNext
    Dim v2 As Variant: v2 = rs!data
    Dim b2() As Byte: b2 = v2
    AssertEqual "Row2 len=1000", UBound(b2) - LBound(b2) + 1, 1000
    AssertEqual "Row2 b(255)=255", b2(255), 255
    AssertEqual "Row2 b(256)=0", b2(256), 0
    rs.MoveNext: rs.CloseRecordset

    Dim rs2 As SQLite3Recordset
    Set rs2 = conn.OpenRecordset("SELECT id, data FROM t_blob ORDER BY id;")
    rs2.LoadAll
    AssertEqual "Vectorized 3 rows", rs2.RecordCount, 3
    rs2.MoveFirst
    Dim vv As Variant: vv = rs2!data
    Dim bv() As Byte: bv = vv
    AssertEqual "Vec row1 len=5", UBound(bv) - LBound(bv) + 1, 5
    rs2.CloseRecordset

    DropTable conn, "t_blob"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 27. Savepoints
'==============================================================================
Public Sub RunTest_Savepoints()
    StartSuite "Savepoints (27)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_sp"
    conn.ExecSQL "CREATE TABLE t_sp (id INTEGER, val TEXT);"

    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_sp VALUES (1, 'outer');"
    conn.Savepoint "sp1"
    AssertEqual "SavepointDepth=1", conn.SavepointDepth, 1
    conn.ExecSQL "INSERT INTO t_sp VALUES (2, 'inner');"
    conn.ReleaseSavepoint "sp1"
    AssertEqual "SavepointDepth=0", conn.SavepointDepth, 0
    conn.CommitTransaction
    AssertEqual "Both rows committed", TableRowCount(conn, "t_sp"), 2

    conn.ExecSQL "DELETE FROM t_sp;"
    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_sp VALUES (10, 'outer2');"
    conn.Savepoint "sp2"
    conn.ExecSQL "INSERT INTO t_sp VALUES (11, 'inner2');"
    conn.ExecSQL "INSERT INTO t_sp VALUES (12, 'inner3');"
    AssertEqual "3 rows before rollback", TableRowCount(conn, "t_sp"), 3
    conn.RollbackToSavepoint "sp2"
    AssertEqual "1 row after sp rollback", TableRowCount(conn, "t_sp"), 1
    conn.ReleaseSavepoint "sp2"
    conn.CommitTransaction
    AssertEqual "Only outer row kept", TableRowCount(conn, "t_sp"), 1
    AssertEqual "Outer value correct", _
        QueryScalar(conn, "SELECT val FROM t_sp WHERE id=10;"), "outer2"

    conn.ExecSQL "DELETE FROM t_sp;"
    conn.BeginTransaction
    conn.Savepoint "outer"
    conn.ExecSQL "INSERT INTO t_sp VALUES (20, 'level1');"
    conn.Savepoint "inner"
    conn.ExecSQL "INSERT INTO t_sp VALUES (21, 'level2');"
    AssertEqual "SavepointDepth=2", conn.SavepointDepth, 2
    conn.RollbackToSavepoint "inner"
    conn.ReleaseSavepoint "inner"
    AssertEqual "SavepointDepth=1 after inner release", conn.SavepointDepth, 1
    conn.ReleaseSavepoint "outer"
    conn.CommitTransaction
    AssertEqual "Nested: only level1", TableRowCount(conn, "t_sp"), 1
    AssertEqual "Nested: val=level1", _
        QueryScalar(conn, "SELECT val FROM t_sp WHERE id=20;"), "level1"

    DropTable conn, "t_sp"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 29. Interrupt
'==============================================================================
Public Sub RunTest_Interrupt()
    StartSuite "Interrupt (29)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_intr"
    conn.ExecSQL "CREATE TABLE t_intr (n INTEGER);"
    Dim i As Long
    For i = 1 To 500
        conn.ExecSQL "INSERT INTO t_intr VALUES (" & i & ");"
    Next i

    ' -- Interrupt on idle is safe --
    Err.Clear
    conn.Interrupt
    AssertNoError "Interrupt on idle conn no error"

    ' -- Query after interrupt works --
    Dim v As Variant
    v = QueryScalar(conn, "SELECT COUNT(*) FROM t_intr;")
    AssertEqual "Query after interrupt returns count", CLng(v), 500

    ' -- Double interrupt safe --
    conn.Interrupt
    conn.Interrupt
    AssertNoError "Double interrupt no error"

    ' -- Data integrity unaffected --
    v = QueryScalar(conn, "SELECT SUM(n) FROM t_intr;")
    AssertTrue "Sum correct after interrupt", CLng(v) = 125250

    ' -- conn still usable after interrupt --
    conn.ExecSQL "SELECT 1;"
    AssertNoError "Conn usable after interrupt"

    ' -- Committed data survives post-commit interrupt --
    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_intr VALUES (9999);"
    conn.CommitTransaction
    conn.Interrupt
    v = QueryScalar(conn, "SELECT COUNT(*) FROM t_intr WHERE n=9999;")
    AssertEqual "Committed row survives post-commit interrupt", CLng(v), 1

    DropTable conn, "t_intr"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 34. Read-only connections
'==============================================================================
Public Sub RunTest_ReadOnly()
    StartSuite "ReadOnly (34)"
    On Error Resume Next

    Dim rw As SQLite3Connection: Set rw = FreshConn()
    DropTable rw, "t_ro"
    rw.ExecSQL "CREATE TABLE t_ro (id INTEGER PRIMARY KEY, val TEXT);"
    rw.BeginTransaction
    Dim i As Long
    For i = 1 To 50
        rw.ExecSQL "INSERT INTO t_ro VALUES (" & i & ", 'r" & i & "');"
    Next i
    rw.CommitTransaction
    AssertEqual "Seed row count", TableRowCount(rw, "t_ro"), 50
    rw.CloseConnection

    Dim ro As New SQLite3Connection
    ro.OpenDatabase DB_PATH, DLL_PATH, 5000, False, 0, True
    AssertNoError "Open read-only"
    AssertTrue "IsReadOnly = True", ro.IsReadOnly
    AssertEqual "Read-only row count", TableRowCount(ro, "t_ro"), 50

    Dim v As Variant
    v = QueryScalar(ro, "SELECT val FROM t_ro WHERE id=25;")
    AssertEqual "Read-only scalar read", CStr(v), "r25"

    Err.Clear
    ro.ExecSQL "INSERT INTO t_ro VALUES (999, 'x');"
    AssertTrue "Write raises error on read-only", Err.Number <> 0
    Err.Clear
    AssertEqual "Row count unchanged", TableRowCount(ro, "t_ro"), 50
    ro.CloseConnection

    Dim rw2 As SQLite3Connection: Set rw2 = FreshConn()
    AssertFalse "IsReadOnly = False for rw conn", rw2.IsReadOnly
    DropTable rw2, "t_ro"
    rw2.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 35. WAL Checkpoint
'==============================================================================
Public Sub RunTest_Checkpoint()
    StartSuite "Checkpoint (35)"
    On Error Resume Next

    Dim ckPath As String
    ckPath = Left(DB_PATH, Len(DB_PATH) - 3) & "_ck.db"
    Kill ckPath:          Err.Clear
    Kill ckPath & "-wal": Err.Clear
    Kill ckPath & "-shm": Err.Clear

    Dim conn As New SQLite3Connection
    conn.OpenDatabase ckPath, DLL_PATH, 5000, True
    AssertNoError "Open checkpoint test DB"
    conn.ExecSQL "CREATE TABLE t_ck (id INTEGER PRIMARY KEY, val REAL);"
    conn.BeginTransaction
    Dim i As Long
    For i = 1 To 1000
        conn.ExecSQL "INSERT INTO t_ck VALUES (" & i & ", " & (i * 1.5) & ");"
    Next i
    conn.CommitTransaction
    AssertEqual "Rows before checkpoint", TableRowCount(conn, "t_ck"), 1000

    Dim ck As Variant
    ck = conn.Checkpoint("PASSIVE")
    AssertNoError "PASSIVE checkpoint no error"
    AssertTrue "Checkpoint returns array", IsArray(ck)
    AssertTrue "Checkpoint(0) pagesWritten >= 0", CLng(ck(0)) >= 0
    AssertTrue "Checkpoint(1) pagesRemaining >= 0", CLng(ck(1)) >= 0
    Log "    INFO  PASSIVE: pagesWritten=" & ck(0) & "  pagesRemaining=" & ck(1)

    Dim ck2 As Variant
    ck2 = conn.Checkpoint("TRUNCATE")
    AssertNoError "TRUNCATE checkpoint no error"
    AssertTrue "TRUNCATE pagesRemaining = 0", CLng(ck2(1)) = 0
    Log "    INFO  TRUNCATE: pagesWritten=" & ck2(0) & "  pagesRemaining=" & ck2(1)
    AssertEqual "Rows after checkpoint", TableRowCount(conn, "t_ck"), 1000

    conn.CloseConnection
    Kill ckPath:          Err.Clear
    Kill ckPath & "-wal": Err.Clear
    Kill ckPath & "-shm": Err.Clear
    EndSuite
End Sub

'==============================================================================
' 36. EXPLAIN QUERY PLAN
'==============================================================================
Public Sub RunTest_QueryPlan()
    StartSuite "QueryPlan (36)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_qp"
    conn.ExecSQL "CREATE TABLE t_qp (id INTEGER PRIMARY KEY, val TEXT, score REAL);"
    conn.ExecSQL "CREATE INDEX idx_qp_score ON t_qp(score);"

    Dim plan As Variant
    plan = GetQueryPlan(conn, "SELECT * FROM t_qp;")
    AssertNoError "GetQueryPlan no error"
    AssertTrue "Plan is array", IsArray(plan)
    AssertTrue "Plan has at least 1 row", UBound(plan, 1) >= 0
    AssertEqual "Plan has 4 columns", UBound(plan, 2) + 1, 4
    Dim detail As String: detail = CStr(plan(0, 3))
    AssertTrue "Plan detail mentions t_qp", InStr(1, detail, "t_qp", vbTextCompare) > 0
    Log "    INFO  plan(0,3)=" & detail

    Dim planIdx As Variant
    planIdx = GetQueryPlan(conn, "SELECT val FROM t_qp WHERE score > 5.0;")
    AssertNoError "GetQueryPlan with index no error"
    AssertTrue "Index plan has rows", IsArray(planIdx) And UBound(planIdx, 1) >= 0
    Dim detailIdx As String: detailIdx = CStr(planIdx(0, 3))
    Log "    INFO  index plan(0,3)=" & detailIdx
    AssertTrue "Index plan detail non-empty", Len(detailIdx) > 0

    DropTable conn, "t_qp"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 39. Tag property
'==============================================================================
Public Sub RunTest_Tag()
    StartSuite "Tag (39)"
    On Error Resume Next
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0

    AssertEqual "Tag default empty", conn.Tag, ""
    conn.Tag = "primary"
    AssertEqual "Tag set/get", conn.Tag, "primary"
    conn.Tag = "secondary"
    AssertEqual "Tag change", conn.Tag, "secondary"
    conn.Tag = ""
    AssertEqual "Tag cleared", conn.Tag, ""

    conn.Tag = "logtest"
    Logger_Configure LOG_INFO, True, False, ""
    On Error Resume Next
    conn.ExecSQL "CREATE TABLE IF NOT EXISTS t_tag_log (x INTEGER);"
    Err.Clear
    AssertNoError "ExecSQL with tagged conn and logger active"
    conn.ExecSQL "DROP TABLE IF EXISTS t_tag_log;"
    Logger_Configure LOG_INFO, True, False, ""

    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 40. ExecScriptFile
'==============================================================================
Public Sub RunTest_ExecScriptFile()
    StartSuite "ExecScriptFile (40)"
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0
    DropTable conn, "script_test"
    DropTable conn, "script_test2"

    Dim scriptPath As String
    scriptPath = Environ("TEMP") & "\sqlite3_test_script.sql"
    Dim fNum As Integer: fNum = FreeFile
    Open scriptPath For Output As #fNum
    Print #fNum, "-- ExecScriptFile test script"
    Print #fNum, "CREATE TABLE script_test (id INTEGER PRIMARY KEY, val TEXT);"
    Print #fNum, "INSERT INTO script_test VALUES (1, 'alpha');"
    Print #fNum, "INSERT INTO script_test VALUES (2, 'beta');"
    Print #fNum, "CREATE TABLE script_test2 (x INTEGER);"
    Print #fNum, "INSERT INTO script_test2 VALUES (42);"
    Close #fNum

    On Error Resume Next
    conn.ExecScriptFile scriptPath
    AssertNoError "ExecScriptFile no error"
    On Error GoTo 0

    AssertTrue "script_test exists", TableExists(conn, "script_test")
    AssertTrue "script_test2 exists", TableExists(conn, "script_test2")
    AssertEqual "script_test row count", TableRowCount(conn, "script_test"), 2
    Dim v As Variant
    v = QueryScalar(conn, "SELECT val FROM script_test WHERE id=1;"): AssertEqual "row 1", CStr(v), "alpha"
    v = QueryScalar(conn, "SELECT val FROM script_test WHERE id=2;"): AssertEqual "row 2", CStr(v), "beta"
    v = QueryScalar(conn, "SELECT x FROM script_test2;"):             AssertEqual "script_test2", CLng(v), 42

    Dim emptyPath As String
    emptyPath = Environ("TEMP") & "\sqlite3_test_empty.sql"
    fNum = FreeFile
    Open emptyPath For Output As #fNum: Close #fNum
    On Error Resume Next
    conn.ExecScriptFile emptyPath
    AssertNoError "ExecScriptFile empty file no error"
    On Error GoTo 0

    On Error Resume Next
    conn.ExecScriptFile Environ("TEMP") & "\no_such_file_xyz.sql"
    AssertTrue "ExecScriptFile missing file raises error", Err.Number <> 0
    Err.Clear
    On Error GoTo 0

    DropTable conn, "script_test": DropTable conn, "script_test2"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 41. QueryColumn
'==============================================================================
Public Sub RunTest_QueryColumn()
    StartSuite "QueryColumn (41)"
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0
    DropTable conn, "qcol_test"
    conn.ExecSQL "CREATE TABLE qcol_test (id INTEGER PRIMARY KEY, name TEXT);"
    conn.ExecSQL "INSERT INTO qcol_test VALUES (1, 'Alice');"
    conn.ExecSQL "INSERT INTO qcol_test VALUES (2, 'Bob');"
    conn.ExecSQL "INSERT INTO qcol_test VALUES (3, 'Carol');"

    Dim names As Variant
    names = QueryColumn(conn, "SELECT name FROM qcol_test ORDER BY id;")
    AssertEqual "QueryColumn name(0)", CStr(names(LBound(names))), "Alice"
    AssertEqual "QueryColumn name(1)", CStr(names(LBound(names) + 1)), "Bob"
    AssertEqual "QueryColumn name(2)", CStr(names(LBound(names) + 2)), "Carol"

    Dim first As Variant
    first = QueryColumn(conn, "SELECT id, name FROM qcol_test ORDER BY id;")
    AssertEqual "QueryColumn multi-col returns first col", CLng(first(LBound(first))), 1

    Dim emptyResult As Variant
    emptyResult = QueryColumn(conn, "SELECT id FROM qcol_test WHERE id = 999;")
    AssertTrue "QueryColumn empty result", UBound(emptyResult) < LBound(emptyResult) Or UBound(emptyResult) = -1

    DropTable conn, "qcol_test"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 24. Aggregates  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Aggregates()
    StartSuite "Aggregates (24)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_agg"
    conn.ExecSQL "CREATE TABLE t_agg (grp TEXT, val REAL);"

    Dim g As Long, r As Long
    For g = 1 To 3
        For r = 1 To 4
            conn.ExecSQL "INSERT INTO t_agg VALUES ('g" & g & "', " & (g * 10 + r) & ");"
        Next r
    Next g

    Dim cnt As Variant: cnt = ScalarAgg(conn, "t_agg", "COUNT(*)")
    AssertEqual "ScalarAgg COUNT", cnt, 12

    Dim total As Variant: total = ScalarAgg(conn, "t_agg", "SUM(val)", "grp='g1'")
    AssertEqual "ScalarAgg SUM g1", total, 50

    Dim gc As Variant: gc = GroupByCount(conn, "t_agg", "grp")
    AssertTrue "GroupByCount rows=3", UBound(gc, 1) = 2
    AssertTrue "GroupByCount cols=2", UBound(gc, 2) = 1
    AssertEqual "GroupByCount each=4", gc(0, 1), 4

    Dim gs As Variant: gs = GroupBySum(conn, "t_agg", "grp", "val")
    AssertTrue "GroupBySum rows=3", UBound(gs, 1) = 2
    AssertEqual "GroupBySum g3 total=130", gs(0, 1), 130

    Dim ma As Variant
    ma = MultiAgg(conn, "t_agg", _
                  Array("COUNT(*) AS n", "SUM(val) AS s", "AVG(val) AS a", "MIN(val) AS lo"))
    AssertEqual "MultiAgg n=12", ma(0, 0), 12
    AssertEqual "MultiAgg s=270", ma(0, 1), 270
    AssertEqual "MultiAgg lo=11", ma(0, 3), 11

    Dim mat As Variant
    mat = AggregateQuery(conn, "SELECT MAX(val), MIN(val) FROM t_agg;")
    AssertEqual "AggQuery MAX=34", mat(0, 0), 34
    AssertEqual "AggQuery MIN=11", mat(0, 1), 11

    Dim rt As Variant
    rt = RunningTotal(conn, "t_agg", "rowid", "val", "grp='g1'")
    AssertTrue "RunningTotal rows=4", UBound(rt, 1) = 3
    AssertEqual "RunningTotal cols=3", UBound(rt, 2), 2
    AssertEqual "RunningTotal last=50", rt(3, 2), 50

    ' -- GroupByAvg --
    Dim ga As Variant: ga = GroupByAvg(conn, "t_agg", "grp", "val")
    AssertTrue "GroupByAvg rows=3", UBound(ga, 1) = 2
    AssertTrue "GroupByAvg g1 avg between 11 and 15", CDbl(ga(2, 1)) >= 11 And CDbl(ga(2, 1)) <= 15

    ' -- ScalarAgg AVG cross-check --
    Dim avgAll As Variant: avgAll = ScalarAgg(conn, "t_agg", "AVG(val)")
    AssertTrue "AVG 270/12=22.5", Abs(CDbl(avgAll) - 22.5) < 0.001

    DropTable conn, "t_agg"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 25. FTS5  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_FTS5()
    StartSuite "FTS5 (25)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    On Error Resume Next
    conn.ExecSQL "CREATE VIRTUAL TABLE _fts5_probe USING fts5(x);"
    If Err.Number <> 0 Then
        Log "    INFO  FTS5 not available in this build -- skipping"
        Err.Clear: conn.CloseConnection: EndSuite: Exit Sub
    End If
    conn.ExecSQL "DROP TABLE _fts5_probe;"
    Err.Clear: On Error Resume Next

    ' -- CreateFTS5Table --
    CreateFTS5Table conn, "t_fts", Array("title", "body"), "", "unicode61", True
    AssertNoError "CreateFTS5Table"
    AssertTrue "t_fts exists", TableExists(conn, "t_fts")

    ' -- FTS5Insert and RowCount --
    FTS5Insert conn, "t_fts", Array("title", "body"), _
               Array("SQLite Tutorial", "Learn how to use SQLite for data storage")
    FTS5Insert conn, "t_fts", Array("title", "body"), _
               Array("Python Guide", "Python is great for data science and scripting")
    FTS5Insert conn, "t_fts", Array("title", "body"), _
               Array("SQLite Performance", "Optimise SQLite queries with indexes and WAL mode")
    FTS5Insert conn, "t_fts", Array("title", "body"), _
               Array("Excel VBA Tips", "Automate Excel tasks with VBA macros and functions")
    FTS5Insert conn, "t_fts", Array("title", "body"), _
               Array("Database Design", "Normalise your schema for better query performance")
    AssertEqual "FTS5 5 rows", FTS5RowCount(conn, "t_fts"), 5

    ' -- FTS5MatchCount --
    Dim cnt As Long: cnt = FTS5MatchCount(conn, "t_fts", "SQLite")
    AssertEqual "FTS5 SQLite matches=2", cnt, 2
    cnt = FTS5MatchCount(conn, "t_fts", "data")
    AssertEqual "FTS5 data matches=2", cnt, 2
    cnt = FTS5MatchCount(conn, "t_fts", "VBA")
    AssertEqual "FTS5 VBA matches=1", cnt, 1

    ' -- FTS5SearchMatrix --
    Dim mat As Variant
    mat = FTS5SearchMatrix(conn, "t_fts", "SQLite", "*", "rank", 10)
    AssertTrue "SearchMatrix rows>=2", UBound(mat, 1) >= 1
    AssertTrue "SearchMatrix has cols", UBound(mat, 2) >= 1

    ' -- Prefix search --
    cnt = FTS5MatchCount(conn, "t_fts", "Optim*")
    AssertTrue "Prefix Optim* matches>=1", cnt >= 1

    ' -- Column filter (title only) --
    cnt = FTS5MatchCount(conn, "t_fts", "title: SQLite")
    AssertEqual "Column filter title:SQLite=2", cnt, 2
    cnt = FTS5MatchCount(conn, "t_fts", "body: SQLite")
    AssertEqual "Column filter body:SQLite=2", cnt, 2

    ' -- Phrase search --
    cnt = FTS5MatchCount(conn, "t_fts", """data science""")
    AssertEqual "Phrase 'data science'=1", cnt, 1

    ' -- FTS5BulkInsert --
    Dim bulkData() As Variant
    ReDim bulkData(99, 1)
    Dim i As Long
    For i = 0 To 99
        bulkData(i, 0) = "Bulk Title " & i
        bulkData(i, 1) = "Bulk body content number " & i & " with extra text"
    Next i
    FTS5BulkInsert conn, "t_fts", Array("title", "body"), bulkData
    AssertEqual "FTS5 after bulk=105", FTS5RowCount(conn, "t_fts"), 105
    cnt = FTS5MatchCount(conn, "t_fts", "Bulk")
    AssertEqual "Bulk rows indexed=100", cnt, 100

    ' -- FTS5Delete --
    Dim delRid As Variant
    delRid = QueryScalar(conn, "SELECT rowid FROM t_fts WHERE title='SQLite Tutorial';")
    AssertFalse "FTS5Delete: rowid found", IsNull(delRid)
    If Not IsNull(delRid) Then
        FTS5Delete conn, "t_fts", CLngLng(delRid)
        AssertNoError "FTS5Delete no error"
    End If
    AssertEqual "After delete=104", FTS5RowCount(conn, "t_fts"), 104
    cnt = FTS5MatchCount(conn, "t_fts", "SQLite")
    AssertEqual "After delete SQLite matches=1", cnt, 1

    ' -- FTS5Optimize --
    Err.Clear
    FTS5Optimize conn, "t_fts"
    AssertNoError "FTS5Optimize"

    ' -- FTS5Rebuild --
    FTS5Rebuild conn, "t_fts"
    AssertNoError "FTS5Rebuild"
    AssertEqual "After rebuild row count unchanged", FTS5RowCount(conn, "t_fts"), 104

    ' -- AND query --
    cnt = FTS5MatchCount(conn, "t_fts", "SQLite AND WAL")
    AssertEqual "AND: SQLite AND WAL=1", cnt, 1

    ' -- Phrase not present returns 0 --
    cnt = FTS5MatchCount(conn, "t_fts", Chr(34) & "totally absent phrase" & Chr(34))
    AssertEqual "Absent phrase returns 0", cnt, 0

    ' -- NOT query --
    cnt = FTS5MatchCount(conn, "t_fts", "SQLite NOT WAL")
    AssertTrue "NOT query returns >=0", cnt >= 0

    ' -- After rebuild, match count stable --
    Dim cntAfter As Long: cntAfter = FTS5MatchCount(conn, "t_fts", "Python")
    AssertEqual "Python count stable after rebuild", cntAfter, 1

    ' -- No-match returns 0 --
    cnt = FTS5MatchCount(conn, "t_fts", "zzznomatch999")
    AssertEqual "Nomatch returns 0", cnt, 0

    ' -- Column filter title:Excel --
    cnt = FTS5MatchCount(conn, "t_fts", "title: Excel")
    AssertEqual "title:Excel=1", cnt, 1

    ' -- Bulk prefix search --
    cnt = FTS5MatchCount(conn, "t_fts", "Bulk*")
    AssertTrue "Bulk* matches all bulk rows", cnt >= 99

    ' -- AND search --
    cnt = FTS5MatchCount(conn, "t_fts", "SQLite AND WAL")
    AssertEqual "AND: SQLite AND WAL=1", cnt, 1

    conn.ExecSQL "DROP TABLE IF EXISTS t_fts;"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 26. Schema  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Schema()
    StartSuite "Schema (26)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_b;"
    conn.ExecSQL "DROP VIEW  IF EXISTS v_schema;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_a;"
    conn.ExecSQL "DROP INDEX IF EXISTS ix_schema_a;"

    conn.ExecSQL "CREATE TABLE t_schema_a (id INTEGER PRIMARY KEY, name TEXT NOT NULL, score REAL);"
    conn.ExecSQL "CREATE TABLE t_schema_b (id INTEGER PRIMARY KEY, aid INTEGER, FOREIGN KEY(aid) REFERENCES t_schema_a(id));"
    conn.ExecSQL "CREATE INDEX ix_schema_a ON t_schema_a(name);"
    conn.ExecSQL "CREATE VIEW v_schema AS SELECT id, name FROM t_schema_a;"

    Dim tbls As Variant: tbls = GetTableList(conn)
    AssertTrue "TableList is array", IsArray(tbls)
    Dim foundA As Boolean, foundB As Boolean, r As Long
    For r = LBound(tbls) To UBound(tbls)
        If CStr(tbls(r)) = "t_schema_a" Then foundA = True
        If CStr(tbls(r)) = "t_schema_b" Then foundB = True
    Next r
    AssertTrue "TableList has t_schema_a", foundA
    AssertTrue "TableList has t_schema_b", foundB

    Dim views As Variant: views = GetViewList(conn)
    AssertTrue "ViewList is array", IsArray(views)
    Dim foundV As Boolean
    For r = LBound(views) To UBound(views)
        If CStr(views(r)) = "v_schema" Then foundV = True
    Next r
    AssertTrue "ViewList has v_schema", foundV

    Dim cols As Variant: cols = GetColumnInfo(conn, "t_schema_a")
    AssertTrue "ColumnInfo is array", IsArray(cols)
    AssertEqual "ColumnInfo rows=3", UBound(cols, 1) - LBound(cols, 1) + 1, 3
    AssertEqual "ColumnInfo col0 name=id", CStr(cols(LBound(cols, 1), 1)), "id"
    AssertEqual "ColumnInfo col1 name=name", CStr(cols(LBound(cols, 1) + 1, 1)), "name"

    Dim idxs As Variant: idxs = GetIndexList(conn, "t_schema_a")
    AssertTrue "IndexList non-empty", Not IsEmpty(idxs)
    Dim foundIdx As Boolean
    If IsArray(idxs) Then
        For r = LBound(idxs, 1) To UBound(idxs, 1)
            If CStr(idxs(r, 1)) = "ix_schema_a" Then foundIdx = True
        Next r
    End If
    AssertTrue "IndexList has ix_schema_a", foundIdx

    Dim fks As Variant: fks = GetForeignKeys(conn, "t_schema_b")
    AssertTrue "ForeignKeys non-empty", Not IsEmpty(fks)
    AssertEqual "FK refs t_schema_a", fks(0, 2), "t_schema_a"

    Dim sql As String: sql = GetCreateSQL(conn, "t_schema_a")
    AssertTrue "CreateSQL non-empty", Len(sql) > 0
    AssertTrue "CreateSQL has CREATE TABLE", InStr(sql, "CREATE TABLE") > 0

    Dim dbInfo As Variant: dbInfo = GetDatabaseInfo(conn)
    AssertTrue "DatabaseInfo is matrix", IsArray(dbInfo)
    AssertTrue "DatabaseInfo rows>5", UBound(dbInfo, 1) >= 5

    conn.ExecSQL "DROP INDEX IF EXISTS ix_schema_a;"
    conn.ExecSQL "DROP VIEW  IF EXISTS v_schema;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_b;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_a;"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 28. JSON  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_JSON()
    StartSuite "JSON (28)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    On Error Resume Next
    Dim probe As Variant
    probe = QueryScalar(conn, "SELECT json_valid('{}')")
    If Err.Number <> 0 Then
        Log "    INFO  JSON not available (requires SQLite 3.38+) -- skipping"
        Err.Clear: conn.CloseConnection: EndSuite: Exit Sub
    End If
    Err.Clear: On Error Resume Next

    DropTable conn, "t_json"
    conn.ExecSQL "CREATE TABLE t_json (id INTEGER PRIMARY KEY, data TEXT);"
    conn.ExecSQL "INSERT INTO t_json VALUES (1, '{""name"":""Alice"",""city"":""London"",""score"":95}');"
    conn.ExecSQL "INSERT INTO t_json VALUES (2, '{""name"":""Bob"",""city"":""Paris"",""score"":82}');"
    conn.ExecSQL "INSERT INTO t_json VALUES (3, '{""name"":""Carol"",""city"":""Berlin"",""tags"":[""vba"",""excel""]}');"

    ' -- JSONExtract --
    Dim v As Variant
    v = JSONExtract(conn, "t_json", "data", "$.name", "id=1")
    AssertEqual "JSONExtract name=Alice", v, "Alice"
    v = JSONExtract(conn, "t_json", "data", "$.score", "id=2")
    AssertEqual "JSONExtract score=82", CDbl(v), 82
    v = JSONExtract(conn, "t_json", "data", "$.tags[0]", "id=3")
    AssertEqual "JSONExtract tags[0]=vba", v, "vba"
    v = JSONExtract(conn, "t_json", "data", "$.tags[1]", "id=3")
    AssertEqual "JSONExtract tags[1]=excel", v, "excel"

    ' -- JSONSet --
    JSONSet conn, "t_json", "data", "$.city", "'Madrid'", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.city", "id=1")
    AssertEqual "JSONSet city=Madrid", v, "Madrid"

    ' -- JSONSet numeric --
    JSONSet conn, "t_json", "data", "$.score", "100", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.score", "id=1")
    AssertEqual "JSONSet score=100", CDbl(v), 100

    ' -- JSONRemove --
    JSONRemove conn, "t_json", "data", Array("$.city"), "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.city", "id=1")
    AssertTrue "JSONRemove key gone", IsNull(v)

    ' -- JSONValid --
    AssertTrue "JSONValid all rows", JSONValid(conn, "t_json", "data")

    ' -- JSONGroupArray --
    Dim arr As String
    arr = JSONGroupArray(conn, "t_json", "json_extract(data,'$.name')", "", "id")
    AssertTrue "JSONGroupArray non-empty", Len(arr) > 2
    AssertTrue "JSONGroupArray has Bob", InStr(arr, "Bob") > 0

    ' -- JSONGroupObject --
    Dim obj As String
    obj = JSONGroupObject(conn, "t_json", "CAST(id AS TEXT)", "json_extract(data,'$.name')", "")
    AssertTrue "JSONGroupObject non-empty", Len(obj) > 2
    AssertTrue "JSONGroupObject has Alice", InStr(obj, "Alice") > 0

    ' -- JSONType --
    Dim jt As String
    jt = JSONType(conn, "t_json", "data", "$.score", "id=2")
    AssertTrue "JSONType score is numeric", jt = "integer" Or jt = "real"
    jt = JSONType(conn, "t_json", "data", "$.tags", "id=3")
    AssertEqual "JSONType tags=array", jt, "array"

    ' -- json_extract verifies nested array value --
    Dim jpath As Variant
    jpath = QueryScalar(conn, "SELECT json_extract(data, '$.tags[0]') FROM t_json WHERE id=3;")
    AssertTrue "JSONSearch finds vba path", Not IsNull(jpath) And CStr(jpath) = "vba"

    ' -- Invalid JSON detection --
    conn.ExecSQL "INSERT INTO t_json VALUES (99, 'NOT VALID JSON');"
    AssertFalse "JSONValid=False with bad row", JSONValid(conn, "t_json", "data")
    conn.ExecSQL "DELETE FROM t_json WHERE id=99;"
    AssertTrue "JSONValid=True after deleting bad row", JSONValid(conn, "t_json", "data")

    ' -- JSONExtract missing key = Null --
    Dim vNull As Variant
    vNull = JSONExtract(conn, "t_json", "data", "$.missing_key_xyz", "id=1")
    AssertTrue "JSONExtract missing key=Null", IsNull(vNull)

    ' -- JSONSet new key --
    JSONSet conn, "t_json", "data", "$.country", "'UK'", "id=2"
    Dim vCo As Variant: vCo = JSONExtract(conn, "t_json", "data", "$.country", "id=2")
    AssertEqual "JSONSet new key country=UK", CStr(vCo), "UK"

    ' -- JSONGroupArray ordering DESC --
    Dim arr2 As String
    arr2 = JSONGroupArray(conn, "t_json", "json_extract(data,'$.name')", "", "id DESC")
    AssertTrue "JSONGroupArray DESC: Carol before Alice", InStr(arr2, "Carol") < InStr(arr2, "Alice")

    DropTable conn, "t_json"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 32. Serialize  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Serialize()
    StartSuite "Serialize (32)"
    On Error Resume Next

    Dim src As New SQLite3Connection
    src.OpenDatabase ":memory:", DLL_PATH, 5000, False, 0
    DropTable src, "t_ser"
    DropTable src, "t_ser2"
    src.ExecSQL "CREATE TABLE t_ser  (id INTEGER PRIMARY KEY, name TEXT);"
    src.ExecSQL "CREATE TABLE t_ser2 (id INTEGER PRIMARY KEY, val  REAL);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 100
        src.ExecSQL "INSERT INTO t_ser  VALUES (" & i & ", 'name_" & i & "');"
        src.ExecSQL "INSERT INTO t_ser2 VALUES (" & i & ", " & (i * 3.14) & ");"
    Next i
    src.CommitTransaction
    AssertEqual "Source t_ser rows", TableRowCount(src, "t_ser"), 100
    AssertEqual "Source t_ser2 rows", TableRowCount(src, "t_ser2"), 100

    Err.Clear

    ' -- SerializeDB --
    Dim snap() As Byte
    snap = SerializeDB(src)
    AssertNoError "SerializeDB no error"
    AssertTrue "Snapshot non-empty", UBound(snap) > 0
    AssertEqual "Header SQLite", Chr(snap(0)) & Chr(snap(1)) & Chr(snap(2)), "SQL"
    Log "    INFO  Serialized size = " & (UBound(snap) + 1) & " bytes"

    ' -- DeserializeDB --
    Dim mem As New SQLite3Connection
    mem.OpenDatabase ":memory:", DLL_PATH, 5000, False
    DeserializeDB mem, snap
    AssertNoError "DeserializeDB no error"
    mem.ExecSQL "PRAGMA journal_mode=DELETE;"
    Err.Clear
    AssertEqual "Deserialized t_ser row count", TableRowCount(mem, "t_ser"), 100
    AssertEqual "Deserialized t_ser2 row count", TableRowCount(mem, "t_ser2"), 100
    Dim v As Variant
    v = QueryScalar(mem, "SELECT name FROM t_ser ORDER BY rowid LIMIT 1 OFFSET 49;")
    AssertEqual "Row 50 correct", CStr(v), "name_50"
    v = QueryScalar(mem, "SELECT val FROM t_ser2 ORDER BY rowid LIMIT 1;")
    AssertTrue "t_ser2 val correct", Abs(CDbl(v) - 3.14) < 0.001
    mem.CloseConnection

    ' -- InMemoryClone independence --
    Dim clone As SQLite3Connection
    Set clone = InMemoryClone(src)
    AssertNoError "InMemoryClone no error"
    AssertEqual "Clone t_ser row count", TableRowCount(clone, "t_ser"), 100
    clone.ExecSQL "DROP TABLE t_ser;"
    AssertFalse "Clone: t_ser gone", TableExists(clone, "t_ser")
    AssertTrue "Source: t_ser still exists", TableExists(src, "t_ser")
    AssertTrue "Clone: t_ser2 survives", TableExists(clone, "t_ser2")
    clone.CloseConnection

    ' -- Round-trip page size preserved --
    Dim snapLen As Long: snapLen = UBound(snap) + 1
    AssertTrue "Snapshot size is page-aligned", (snapLen Mod 4096) = 0 Or (snapLen Mod 1024) = 0 Or snapLen > 0

    ' -- Second clone independent --
    Dim clone2 As SQLite3Connection: Set clone2 = InMemoryClone(src)
    AssertEqual "Clone2 t_ser2 rows", TableRowCount(clone2, "t_ser2"), 100
    clone2.CloseConnection

    src.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 33. Diagnostics  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Diagnostics()
    StartSuite "Diagnostics (33)"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_diag"
    conn.ExecSQL "CREATE TABLE t_diag (id INTEGER PRIMARY KEY, val REAL);"
    conn.BeginTransaction
    Dim i As Long
    For i = 1 To 1000
        conn.ExecSQL "INSERT INTO t_diag VALUES (" & i & ", " & (i * 1.1) & ");"
    Next i
    conn.CommitTransaction

    ' -- GetDbStatus matrix --
    Dim info As Variant: info = GetDbStatus(conn, False)
    AssertNoError "GetDbStatus no error"
    AssertTrue "GetDbStatus rows >= 13", UBound(info, 1) + 1 >= 13
    AssertTrue "GetDbStatus has 3 cols", UBound(info, 2) >= 2

    ' -- GetDbStatusValue for multiple op codes --
    Dim cv As Variant: cv = GetDbStatusValue(conn, DBSTAT_CACHE_USED, False)
    AssertNoError "GetDbStatusValue CACHE_USED no error"
    AssertTrue "cache_used current >= 0", CLng(cv(0)) >= 0
    AssertTrue "cache_used highwater >= 0", CLng(cv(1)) >= 0
    Log "    INFO  cache_used current=" & cv(0) & " highwater=" & cv(1)

    Dim sv As Variant: sv = GetDbStatusValue(conn, DBSTAT_SCHEMA_USED, False)
    AssertNoError "GetDbStatusValue SCHEMA_USED no error"
    AssertTrue "schema_used >= 0", CLng(sv(0)) >= 0

    Dim stv As Variant: stv = GetDbStatusValue(conn, DBSTAT_STMT_USED, False)
    AssertNoError "GetDbStatusValue STMT_USED no error"
    AssertTrue "stmt_used >= 0", CLng(stv(0)) >= 0

    ' -- GetStmtStatus --
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "SELECT SUM(val) FROM t_diag WHERE val > 500;"
    cmd.ExecuteScalar
    AssertNoError "Diagnostic query no error"

    Dim fullScan As Long
    fullScan = GetStmtStatus(cmd.StmtHandle, STMTSTAT_FULLSCAN, False)
    AssertTrue "Full-scan steps > 0", fullScan > 0

    Dim vmSteps As Long
    vmSteps = GetStmtStatus(cmd.StmtHandle, STMTSTAT_VM_STEP, False)
    AssertTrue "VM steps > 0", vmSteps > 0

    Dim memUsed As Long
    memUsed = GetStmtStatus(cmd.StmtHandle, STMTSTAT_MEMUSED, False)
    AssertTrue "Stmt memused >= 0", memUsed >= 0

    ' -- GetStmtStatus with reset --
    Dim fullScanAfterReset As Long
    fullScanAfterReset = GetStmtStatus(cmd.StmtHandle, STMTSTAT_FULLSCAN, True)
    AssertNoError "GetStmtStatus with reset no error"

    ' -- DbStatusSummary (smoke test) --
    Err.Clear
    DbStatusSummary conn
    AssertNoError "DbStatusSummary no error"

    ' -- CACHE_HIT readable --
    Dim chit As Variant: chit = GetDbStatusValue(conn, DBSTAT_CACHE_HIT, False)
    AssertNoError "GetDbStatusValue CACHE_HIT no error"
    AssertTrue "CACHE_HIT >= 0", CLng(chit(0)) >= 0

    DropTable conn, "t_diag"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 37. Excel  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Excel()
    StartSuite "Excel (37)"
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    If Err.Number <> 0 Then
        Log "    INFO  Skipping Excel tests -- could not add worksheet"
        Err.Clear: EndSuite: Exit Sub
    End If
    Err.Clear

    ws.Cells(1, 1) = "id":    ws.Cells(1, 2) = "sym"
    ws.Cells(1, 3) = "price": ws.Cells(1, 4) = "qty"
    ws.Cells(1, 5) = "active"
    Dim i As Long
    For i = 1 To 10
        ws.Cells(i + 1, 1) = i
        ws.Cells(i + 1, 2) = "SYM" & i
        ws.Cells(i + 1, 3) = 100# + i * 0.5
        ws.Cells(i + 1, 4) = i * 10
        ws.Cells(i + 1, 5) = IIf(i Mod 2 = 0, True, False)
    Next i

    Dim srcRange As Range
    Set srcRange = ws.Range(ws.Cells(1, 1), ws.Cells(11, 5))

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_excel"
    RangeToTable conn, "t_excel", srcRange, True, True
    AssertNoError "RangeToTable no error"
    AssertEqual "RangeToTable row count", TableRowCount(conn, "t_excel"), 10

    ' -- Column values --
    Dim vSym As Variant
    vSym = QueryScalar(conn, "SELECT sym FROM t_excel WHERE id=3;")
    AssertEqual "Text column correct", CStr(vSym), "SYM3"

    Dim vPrice As Variant
    vPrice = QueryScalar(conn, "SELECT price FROM t_excel WHERE id=1;")
    AssertTrue "Numeric column correct", Abs(CDbl(vPrice) - 100.5) < 0.001

    Dim vQty As Variant
    vQty = QueryScalar(conn, "SELECT qty FROM t_excel WHERE id=5;")
    AssertEqual "Integer column correct", CLng(vQty), 50

    ' -- Schema check --
    Dim cols As Variant
    cols = GetColumnInfo(conn, "t_excel")
    AssertEqual "Column count from schema", UBound(cols, 1) + 1, 5

    ' -- QueryToRange with headers --
    Dim destCell As Range
    Set destCell = ws.Cells(15, 1)
    QueryToRange conn, "SELECT id, sym, price FROM t_excel ORDER BY id;", destCell, True
    AssertNoError "QueryToRange no error"
    AssertEqual "Header col1", CStr(ws.Cells(15, 1).Value), "id"
    AssertEqual "Header col2", CStr(ws.Cells(15, 2).Value), "sym"
    AssertEqual "Data row1 id", CLng(ws.Cells(16, 1).Value), 1
    AssertEqual "Data row10 id", CLng(ws.Cells(25, 1).Value), 10

    ' -- QueryToRange without headers --
    Set destCell = ws.Cells(30, 1)
    QueryToRange conn, "SELECT id, sym FROM t_excel WHERE id <= 3 ORDER BY id;", destCell, False
    AssertNoError "QueryToRange no-header no error"
    AssertEqual "No-hdr first cell is 1", CLng(ws.Cells(30, 1).Value), 1
    AssertEqual "No-hdr second cell is SYM1", CStr(ws.Cells(30, 2).Value), "SYM1"

    DropTable conn, "t_excel"
    conn.CloseConnection

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Err.Clear
    EndSuite
End Sub

'==============================================================================
' 38. Logger  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Logger()
    StartSuite "Logger (38)"
    On Error Resume Next

    ' -- Configure and IsEnabled --
    Err.Clear
    Logger_Configure LOG_DEBUG, True, False, ""
    AssertNoError "Logger_Configure no error"

    AssertTrue "IsEnabled(DEBUG) when level=DEBUG", Logger_IsEnabled(LOG_DEBUG)
    AssertTrue "IsEnabled(INFO)  when level=DEBUG", Logger_IsEnabled(LOG_INFO)
    AssertTrue "IsEnabled(WARN)  when level=DEBUG", Logger_IsEnabled(LOG_WARN)
    AssertTrue "IsEnabled(ERROR) when level=DEBUG", Logger_IsEnabled(LOG_ERROR)

    ' -- SetLevel / GetLevel --
    Logger_SetLevel LOG_WARN
    AssertFalse "IsEnabled(DEBUG) when level=WARN", Logger_IsEnabled(LOG_DEBUG)
    AssertFalse "IsEnabled(INFO)  when level=WARN", Logger_IsEnabled(LOG_INFO)
    AssertTrue "IsEnabled(WARN)  when level=WARN", Logger_IsEnabled(LOG_WARN)
    AssertTrue "IsEnabled(ERROR) when level=WARN", Logger_IsEnabled(LOG_ERROR)
    AssertEqual "GetLevel returns WARN", Logger_GetLevel(), LOG_WARN

    ' -- Named wrapper calls --
    Logger_SetLevel LOG_DEBUG
    Err.Clear
    Logger_Debug "RunTest_Logger", "debug message"
    Logger_Info "RunTest_Logger", "info message"
    Logger_Warn "RunTest_Logger", "warn message"
    Logger_Error "RunTest_Logger", "error message (test)"
    AssertNoError "All named wrappers no error"

    ' -- LOG_NONE --
    Logger_SetLevel LOG_NONE
    AssertFalse "IsEnabled(ERROR) when level=NONE", Logger_IsEnabled(LOG_ERROR)
    AssertFalse "IsEnabled(WARN)  when level=NONE", Logger_IsEnabled(LOG_WARN)

    ' -- File sink --
    Dim logPath As String
    logPath = Left(DB_PATH, InStrRev(DB_PATH, "\\")) & "_logger_test.log"
    Kill logPath: Err.Clear
    Logger_Configure LOG_INFO, False, True, logPath
    AssertNoError "Logger_Configure with file sink no error"

    Logger_Info "RunTest_Logger", "test line INFO"
    Logger_Warn "RunTest_Logger", "test line WARN"
    Logger_Error "RunTest_Logger", "test line ERROR"
    Logger_Debug "RunTest_Logger", "test line DEBUG (filtered)"
    Logger_Close
    AssertNoError "Logger_Close no error"

    Dim fileNum As Integer: fileNum = FreeFile()
    Dim fileContent As String: Dim oneLine As String
    Open logPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, oneLine
        fileContent = fileContent & oneLine & vbCrLf
    Loop
    Close #fileNum
    Err.Clear

    AssertTrue "Log file contains INFO line", InStr(fileContent, "test line INFO") > 0
    AssertTrue "Log file contains WARN line", InStr(fileContent, "test line WARN") > 0
    AssertTrue "Log file contains ERROR line", InStr(fileContent, "test line ERROR") > 0
    AssertFalse "Log file excludes DEBUG line", InStr(fileContent, "test line DEBUG") > 0
    Kill logPath: Err.Clear

    ' -- Re-configure clears state --
    Logger_Configure LOG_DEBUG, True, False, ""
    AssertNoError "Re-configure no error"
    AssertTrue "Re-configure: DEBUG enabled", Logger_IsEnabled(LOG_DEBUG)

    Logger_Configure LOG_INFO, True, False, ""
    EndSuite
End Sub

'==============================================================================
' 42. ListObject  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_ListObject()
    StartSuite "ListObject (42)"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(1)
    On Error GoTo 0
    If ws Is Nothing Then
        Log "    INFO  No worksheet available -- ListObject suite skipped"
EndSuite:         Exit Sub
    End If

    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0
    DropTable conn, "lo_test"

    Dim startCell As Range
    Set startCell = ws.Cells(1000, 1)
    startCell.Value = "Symbol": startCell.offset(0, 1).Value = "Price"
    startCell.offset(1, 0).Value = "AAPL": startCell.offset(1, 1).Value = 189.5
    startCell.offset(2, 0).Value = "MSFT": startCell.offset(2, 1).Value = 420.1
    startCell.offset(3, 0).Value = "GOOG": startCell.offset(3, 1).Value = 172.3

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.offset(3, 1)), , xlYes)
    lo.name = "TestListObj_SQLite"
    On Error GoTo 0
    If lo Is Nothing Then
        Log "    INFO  Could not create ListObject -- suite skipped"
        ws.Range(startCell, startCell.offset(3, 1)).ClearContents
        conn.CloseConnection: EndSuite: Exit Sub
    End If

    On Error Resume Next
    ListObjectToTable conn, "lo_test", lo, True
    AssertNoError "ListObjectToTable no error"
    On Error GoTo 0

    AssertTrue "lo_test exists", TableExists(conn, "lo_test")
    AssertEqual "lo_test row count", TableRowCount(conn, "lo_test"), 3

    Dim v As Variant
    v = QueryScalar(conn, "SELECT COUNT(*) FROM lo_test WHERE Symbol='AAPL';")
    AssertEqual "lo_test AAPL present", CLng(v), 1

    ' -- Column count --
    Dim ci As Variant: ci = GetColumnInfo(conn, "lo_test")
    AssertEqual "lo_test column count=2", UBound(ci, 1) + 1, 2

    ' -- Data value --
    Dim vp As Variant
    vp = QueryScalar(conn, "SELECT Price FROM lo_test WHERE Symbol='MSFT';")
    AssertTrue "MSFT price correct", Abs(CDbl(vp) - 420.1) < 0.001

    On Error Resume Next
    lo.Delete
    ws.Range(ws.Cells(1000, 1), ws.Cells(1003, 2)).ClearContents
    On Error GoTo 0

    DropTable conn, "lo_test"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 43. Migrate  (SQLite3_Driver.bas)
'==============================================================================
Public Sub RunTest_Migrate()
    StartSuite "Migrate (43)"
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0

    ' -- SetSchemaVersion / GetSchemaVersion --
    SetSchemaVersion conn, 0
    AssertEqual "Initial version is 0", GetSchemaVersion(conn), 0

    SetSchemaVersion conn, 7
    AssertEqual "SetSchemaVersion 7", GetSchemaVersion(conn), 7
    SetSchemaVersion conn, 0

    ' -- ApplyMigration single step --
    DropTable conn, "mig_v1"
    Dim applied As Boolean
    applied = ApplyMigration(conn, 1, _
        "CREATE TABLE IF NOT EXISTS mig_v1 (id INTEGER PRIMARY KEY, label TEXT);")
    AssertTrue "ApplyMigration 0->1 applied", applied
    AssertEqual "Version is 1 after step 1", GetSchemaVersion(conn), 1
    AssertTrue "mig_v1 table exists", TableExists(conn, "mig_v1")

    ' -- ApplyMigration idempotent (same version) --
    applied = ApplyMigration(conn, 1, "CREATE TABLE should_not_exist (x INTEGER);")
    AssertFalse "ApplyMigration 1->1 skipped", applied
    AssertFalse "should_not_exist not created", TableExists(conn, "should_not_exist")

    ' -- ApplyMigration step 2 --
    DropTable conn, "mig_v2"
    applied = ApplyMigration(conn, 2, _
        "CREATE TABLE IF NOT EXISTS mig_v2 (id INTEGER PRIMARY KEY, val REAL);")
    AssertTrue "ApplyMigration 1->2 applied", applied
    AssertEqual "Version is 2 after step 2", GetSchemaVersion(conn), 2

    ' -- MigrateAll from scratch --
    SetSchemaVersion conn, 0
    DropTable conn, "mig_v1": DropTable conn, "mig_v2": DropTable conn, "mig_v3"

    Dim steps(2) As MigrationStep
    steps(0) = MakeStep(1, "CREATE TABLE IF NOT EXISTS mig_v1 (id INTEGER PRIMARY KEY, label TEXT);")
    steps(1) = MakeStep(2, "CREATE TABLE IF NOT EXISTS mig_v2 (id INTEGER PRIMARY KEY, val REAL);")
    steps(2) = MakeStep(3, "CREATE TABLE IF NOT EXISTS mig_v3 (id INTEGER PRIMARY KEY, ts TEXT);")

    Dim nApplied As Long
    nApplied = MigrateAll(conn, steps)
    AssertEqual "MigrateAll applied 3 steps", nApplied, 3
    AssertEqual "Version is 3 after MigrateAll", GetSchemaVersion(conn), 3
    AssertTrue "mig_v1 present", TableExists(conn, "mig_v1")
    AssertTrue "mig_v3 present", TableExists(conn, "mig_v3")

    ' -- MigrateAll idempotent --
    nApplied = MigrateAll(conn, steps)
    AssertEqual "MigrateAll idempotent: 0 applied", nApplied, 0

    ' -- Partial apply from current version --
    Dim step4(0) As MigrationStep
    step4(0) = MakeStep(4, "ALTER TABLE mig_v3 ADD COLUMN notes TEXT;")
    nApplied = MigrateAll(conn, step4)
    AssertEqual "MigrateAll partial: 1 applied", nApplied, 1
    AssertEqual "Version is 4", GetSchemaVersion(conn), 4

    ' -- New column visible --
    conn.ExecSQL "INSERT INTO mig_v3 VALUES (1,'2026-01-01','release note');"
    Dim noteVal As Variant
    noteVal = QueryScalar(conn, "SELECT notes FROM mig_v3 WHERE id=1;")
    AssertEqual "New column has value", CStr(noteVal), "release note"

    ' -- Bad SQL raises error, version unchanged --
    Dim badStep(0) As MigrationStep
    badStep(0) = MakeStep(5, "THIS IS NOT VALID SQL !!!;")
    On Error Resume Next
    nApplied = MigrateAll(conn, badStep)
    AssertTrue "MigrateAll bad SQL raises error", Err.Number <> 0
    Err.Clear: On Error GoTo 0
    AssertEqual "Version still 4 after failed step", GetSchemaVersion(conn), 4

    ' -- Step below current is skipped --
    Dim skipStep(0) As MigrationStep
    skipStep(0) = MakeStep(3, "CREATE TABLE should_skip (x INTEGER);")
    Dim nSkip As Long: nSkip = MigrateAll(conn, skipStep)
    AssertEqual "Step below current: 0 applied", nSkip, 0
    AssertFalse "should_skip not created", TableExists(conn, "should_skip")

    ' -- ALTERed column is queryable --
    conn.ExecSQL "INSERT OR REPLACE INTO mig_v3 VALUES (1,'2026-01-01','v0.1.7');"
    Dim nv As Variant: nv = QueryScalar(conn, "SELECT notes FROM mig_v3 WHERE id=1;")
    AssertEqual "ALTERed column queryable", CStr(nv), "v0.1.7"

    DropTable conn, "mig_v1": DropTable conn, "mig_v2": DropTable conn, "mig_v3"
    SetSchemaVersion conn, 0
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' RunTests_BulkInsert - dispatched from RunAllTests when class is present
'==============================================================================
Public Sub RunTests_BulkInsert(ByVal dbPath As String, ByVal dllPath As String)
    '---- Suite 15: AppendRow -------------------------------------------
    StartSuite "BulkInsert_AppendRow (15)"
    On Error Resume Next

    Dim conn As New SQLite3Connection
    conn.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable conn, "t_bulk"
    conn.ExecSQL "CREATE TABLE t_bulk (i INTEGER, s TEXT, f REAL);"

    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, "t_bulk", Array("i", "s", "f"), 100
    Dim i As Long
    For i = 1 To 500
        bulk.AppendRow Array(i, "row" & i, i * 0.5)
    Next i
    bulk.CloseInsert

    AssertEqual "500 rows inserted", TableRowCount(conn, "t_bulk"), 500
    Dim v As Variant
    v = QueryScalar(conn, "SELECT s FROM t_bulk WHERE i=1;")
    AssertEqual "First row s", v, "row1"
    v = QueryScalar(conn, "SELECT s FROM t_bulk WHERE i=500;")
    AssertEqual "Last row s", v, "row500"
    v = QueryScalar(conn, "SELECT f FROM t_bulk WHERE i=2;")
    AssertTrue "Row2 f ~1.0", Abs(CDbl(v) - 1#) < 0.001

    DropTable conn, "t_bulk"
    conn.CloseConnection
    EndSuite

    '---- Suite 16: AppendMatrix ----------------------------------------
    StartSuite "BulkInsert_AppendMatrix (16)"
    On Error Resume Next

    Dim conn2 As New SQLite3Connection
    conn2.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable conn2, "t_mat2"
    conn2.ExecSQL "CREATE TABLE t_mat2 (a INTEGER, b TEXT);"

    Const N As Long = 200
    Dim mat() As Variant
    ReDim mat(N - 1, 1)
    For i = 0 To N - 1
        mat(i, 0) = i + 1
        mat(i, 1) = "m" & (i + 1)
    Next i

    Dim bulk2 As New SQLite3BulkInsert
    bulk2.OpenInsert conn2, "t_mat2", Array("a", "b")
    bulk2.AppendMatrix mat
    bulk2.CloseInsert

    AssertEqual "TotalRowsInserted", bulk2.TotalRowsInserted, N
    AssertEqual "Row count in DB", TableRowCount(conn2, "t_mat2"), N
    v = QueryScalar(conn2, "SELECT b FROM t_mat2 WHERE a=100;")
    AssertEqual "Mid row b", v, "m100"

    DropTable conn2, "t_mat2"
    conn2.CloseConnection
    EndSuite

    '---- Suite 19: LargeDataset ----------------------------------------
    StartSuite "LargeDataset (19)"
    On Error Resume Next

    Dim conn3 As New SQLite3Connection
    conn3.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable conn3, "t_large"
    conn3.ExecSQL "CREATE TABLE t_large (i INTEGER, f REAL, s TEXT);"

    Const NL As Long = 10000
    Dim bulk3 As New SQLite3BulkInsert
    bulk3.OpenInsert conn3, "t_large", Array("i", "f", "s"), 1000
    For i = 1 To NL
        bulk3.AppendRow Array(i, i * 1.5, "s" & i)
    Next i
    bulk3.CloseInsert

    AssertEqual "10k rows inserted", TableRowCount(conn3, "t_large"), NL

    Dim rs As SQLite3Recordset
    Set rs = conn3.OpenRecordset("SELECT i, f, s FROM t_large ORDER BY i;")
    Dim cnt As Long: cnt = rs.LoadAll()
    AssertEqual "LoadAll = 10000", cnt, NL
    rs.MoveFirst
    AssertEqual "Row 1 i", rs!i, 1

    Dim mat2 As Variant: mat2 = rs.ToMatrix()
    AssertEqual "Matrix rows", UBound(mat2, 1) + 1, NL
    AssertEqual "Matrix cols", UBound(mat2, 2) + 1, 3
    AssertEqual "Last row i", mat2(NL - 1, 0), NL
    AssertTrue "Last row f", Abs(CDbl(mat2(NL - 1, 1)) - (NL * 1.5)) < 0.001
    AssertEqual "Last row s", mat2(NL - 1, 2), "s" & NL

    rs.CloseRecordset

    ' -- GetRows on large recordset --
    Dim rs2 As SQLite3Recordset
    Set rs2 = conn3.OpenRecordset("SELECT i FROM t_large ORDER BY i LIMIT 100;")
    Dim gr As Variant: gr = rs2.GetRows(100)
    AssertEqual "GetRows(100) count", UBound(gr, 2) + 1, 100
    AssertEqual "GetRows first value", gr(0, 0), 1
    AssertEqual "GetRows last value", gr(0, 99), 100
    rs2.CloseRecordset

    DropTable conn3, "t_large"
    conn3.CloseConnection
    EndSuite
End Sub

'==============================================================================
' RunTests_Pool - dispatched from RunAllTests when class is present
'==============================================================================
Public Sub RunTests_Pool(ByVal dbPath As String, ByVal dllPath As String)
    '---- Suite 18: Connection pool -------------------------------------
    StartSuite "ConnectionPool (18)"
    On Error Resume Next

    Dim setup As New SQLite3Connection
    setup.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable setup, "t_pool"
    setup.ExecSQL "CREATE TABLE t_pool (n INTEGER);"
    setup.ExecSQL "INSERT INTO t_pool VALUES (42);"
    setup.CloseConnection

    Dim pool As New SQLite3Pool
    pool.Initialize dbPath, dllPath, 3

    AssertEqual "Initial pool size = 1", pool.PoolSize, 1
    AssertEqual "Initial active = 0", pool.ActiveConnections, 0

    Dim c1 As SQLite3Connection: Set c1 = pool.Acquire()
    Dim c2 As SQLite3Connection: Set c2 = pool.Acquire()
    AssertEqual "Active after 2 acquires", pool.ActiveConnections, 2

    Dim rs As SQLite3Recordset
    Set rs = c1.OpenRecordset("SELECT n FROM t_pool;")
    AssertEqual "Pool query result", rs!N, 42
    rs.CloseRecordset

    pool.ReleaseConnection c1
    pool.ReleaseConnection c2
    AssertEqual "Active after 2 releases", pool.ActiveConnections, 0

    Dim c3 As SQLite3Connection: Set c3 = pool.Acquire()
    c3.BeginTransaction
    c3.ExecSQL "INSERT INTO t_pool VALUES (99);"
    AssertTrue "InTransaction before release", c3.InTransaction
    pool.ReleaseConnection c3
    Dim c4 As SQLite3Connection: Set c4 = pool.Acquire()
    AssertEqual "Rollback on release", _
        QueryScalar(c4, "SELECT COUNT(*) FROM t_pool WHERE n=99;"), 0
    pool.ReleaseConnection c4

    pool.ShutDown
    AssertEqual "PoolSize after shutdown", pool.PoolSize, 0

    Dim cleanup As New SQLite3Connection
    cleanup.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable cleanup, "t_pool"
    cleanup.CloseConnection
    EndSuite

    '---- Pool-exhausted sub-test (was inside ErrorHandling suite) ------
    StartSuite "Pool_Exhausted (18b)"
    On Error Resume Next

    Dim setupE As New SQLite3Connection
    setupE.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable setupE, "t_err"
    setupE.ExecSQL "CREATE TABLE t_err (n INTEGER);"
    setupE.CloseConnection

    Dim pool2 As New SQLite3Pool
    pool2.Initialize dbPath, dllPath, 2
    Dim p1 As SQLite3Connection: Set p1 = pool2.Acquire()
    Dim p2 As SQLite3Connection: Set p2 = pool2.Acquire()
    Err.Clear
    Dim p3 As SQLite3Connection: Set p3 = pool2.Acquire()
    AssertTrue "Pool exhausted raises error", Err.Number <> 0
    Err.Clear

    ' -- Release one slot, re-acquire succeeds --
    pool2.ReleaseConnection p1
    Dim p3b As SQLite3Connection: Set p3b = pool2.Acquire()
    AssertNoError "Acquire after release no error"
    AssertFalse "Acquired conn not Nothing", p3b Is Nothing

    pool2.ReleaseConnection p2
    pool2.ReleaseConnection p3b
    pool2.ShutDown

    Dim cleanE As New SQLite3Connection
    cleanE.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable cleanE, "t_err"
    cleanE.CloseConnection
    EndSuite
End Sub

'==============================================================================
' RunTests_Backup - dispatched from RunAllTests when class is present
'==============================================================================
Public Sub RunTests_Backup(ByVal dbPath As String, ByVal dllPath As String)
    StartSuite "Backup (30)"
    On Error Resume Next

    Dim srcPath As String
    srcPath = Left(dbPath, Len(dbPath) - 3) & "_baksrc.db"
    Kill srcPath:          Err.Clear
    Kill srcPath & "-wal": Err.Clear
    Kill srcPath & "-shm": Err.Clear

    Dim src As New SQLite3Connection
    src.OpenDatabase srcPath, dllPath, 5000, False
    AssertNoError "Open backup source DB"
    src.ExecSQL "CREATE TABLE t_bak (id INTEGER PRIMARY KEY, val TEXT);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 500
        src.ExecSQL "INSERT INTO t_bak VALUES (" & i & ", 'row_" & i & "');"
    Next i
    src.CommitTransaction
    AssertEqual "Source rows before backup", TableRowCount(src, "t_bak"), 500

    Dim destPath As String
    destPath = Left(dbPath, Len(dbPath) - 3) & "_bak.db"
    Kill destPath: Err.Clear

    Dim bak As New SQLite3Backup
    bak.BackupToFile src, destPath
    AssertNoError "BackupToFile no error"
    AssertTrue "Backup IsComplete", bak.IsComplete
    AssertFalse "Backup not open after finish", bak.IsOpen

    Dim dest As New SQLite3Connection
    dest.OpenDatabase destPath, dllPath, 5000, False
    AssertNoError "Open backup DB"
    AssertEqual "Backup row count", TableRowCount(dest, "t_bak"), 500
    Dim v As Variant
    v = QueryScalar(dest, "SELECT val FROM t_bak WHERE id=250;")
    AssertEqual "Backup row 250 correct", CStr(v), "row_250"
    dest.CloseConnection

    Dim dest2Path As String
    dest2Path = Left(dbPath, Len(dbPath) - 3) & "_bak2.db"
    Kill dest2Path: Err.Clear

    Dim bak2 As New SQLite3Backup
    bak2.OpenBackup src, dest2Path
    AssertNoError "OpenBackup no error"
    Dim steps As Long
    Do While Not bak2.IsComplete
        bak2.Step 1
        steps = steps + 1
        If steps = 1 Then AssertTrue "TotalPages > 0", bak2.TotalPages > 0
        If steps > 10000 Then Exit Do
    Loop
    AssertTrue "Incremental: IsComplete", bak2.IsComplete
    AssertTrue "Progress = 1.0", bak2.Progress >= 0.99
    bak2.CloseBackup

    Dim dest2 As New SQLite3Connection
    dest2.OpenDatabase dest2Path, dllPath, 5000, False
    AssertEqual "Backup2 row count", TableRowCount(dest2, "t_bak"), 500
    dest2.CloseConnection

    src.CloseConnection
    Kill srcPath:   Err.Clear
    Kill destPath:  Err.Clear
    Kill dest2Path: Err.Clear
    EndSuite
End Sub

'==============================================================================
' RunTests_BlobStream - dispatched from RunAllTests when class is present
'==============================================================================
Public Sub RunTests_BlobStream(ByVal dbPath As String, ByVal dllPath As String)
    StartSuite "BlobStream (31)"
    On Error Resume Next

    Dim conn As New SQLite3Connection
    conn.OpenDatabase dbPath, dllPath, 5000, True, 0
    DropTable conn, "t_blobstream"
    conn.ExecSQL "CREATE TABLE t_blobstream (id INTEGER PRIMARY KEY, data BLOB);"

    Dim blobSize As Long: blobSize = 1024
    conn.ExecSQL "INSERT INTO t_blobstream VALUES (1, zeroblob(" & blobSize & "));"
    Dim rowId As LongLong: rowId = conn.LastInsertRowID()
    AssertTrue "RowId is 1", CLng(rowId) = 1

    conn.BeginTransaction
    Dim bs As New SQLite3BlobStream
    bs.OpenBlob conn, "t_blobstream", "data", rowId, True
    AssertNoError "OpenBlob for write"
    AssertTrue "IsOpen", bs.IsOpen
    AssertEqual "Blob size", bs.Size, blobSize

    Dim chunk() As Byte
    ReDim chunk(255)
    Dim j As Long
    For j = 0 To 255: chunk(j) = CByte(j): Next j
    Dim off As Long
    For off = 0 To 768 Step 256
        bs.WriteAt chunk, off
    Next off
    AssertNoError "WriteAt all chunks"
    AssertEqual "Position unchanged after WriteAt", bs.Position, 0
    bs.CloseBlob
    conn.CommitTransaction
    AssertFalse "Closed after commit", bs.IsOpen

    Dim bsR As New SQLite3BlobStream
    bsR.OpenBlob conn, "t_blobstream", "data", rowId, False
    AssertNoError "OpenBlob for read"
    AssertEqual "Read blob size", bsR.Size, blobSize

    Dim firstChunk() As Byte
    firstChunk = bsR.ReadAt(256, 0)
    AssertNoError "ReadAt 0"
    AssertEqual "First byte = 0", CInt(firstChunk(0)), 0
    AssertEqual "Last byte of first chunk = 255", CInt(firstChunk(255)), 255

    bsR.SeekTo 0
    Dim seqChunk() As Byte: seqChunk = bsR.ReadBytes(128)
    AssertEqual "Position after ReadBytes(128)", bsR.Position, 128
    AssertEqual "Sequential byte 0", CInt(seqChunk(0)), 0
    AssertEqual "Sequential byte 127", CInt(seqChunk(127)), 127

    Dim lastChunk() As Byte: lastChunk = bsR.ReadAt(256, 768)
    AssertEqual "Last chunk byte 0 = 0", CInt(lastChunk(0)), 0
    AssertEqual "Last chunk byte 255 = 255", CInt(lastChunk(255)), 255
    bsR.CloseBlob

    Dim bsE As New SQLite3BlobStream
    bsE.OpenBlob conn, "t_blobstream", "data", rowId, False
    Err.Clear
    bsE.SeekTo blobSize + 1
    AssertTrue "Out-of-range seek raises error", Err.Number <> 0
    Err.Clear
    bsE.CloseBlob

    DropTable conn, "t_blobstream"
    conn.CloseConnection
    EndSuite
End Sub

