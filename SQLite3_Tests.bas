Attribute VB_Name = "SQLite3_Tests"
'==============================================================================
' SQLite3_Tests.bas  -  Comprehensive driver test suite
'
' Run:  RunAllTests          - full suite with summary
'       RunTest_<name>       - individual test
'
' Output goes to the Immediate window (Ctrl+G).
' Each test prints PASS or FAIL with details on failure.
'
' Version : 0.1.4
'
' Version History:
'   0.1.0 - Initial release. 122 tests across 22 suites.
'   0.1.1 - Added QueryPerformanceCounter/Frequency high-resolution timing.
'            Added EndSuite() per-suite elapsed time reporting.
'            Fixed inline comment after line-continuation in Array() calls.
'   0.1.2 - Added RunTest_BLOB (23), RunTest_Aggregates (24),
'            RunTest_FTS5 (25). Total: 171 tests across 25 suites.
'   0.1.3 - Added RunTest_Schema (26), RunTest_Savepoints (27),
'            RunTest_JSON (28), RunTest_Interrupt (29).
'            Total: 240 tests across 29 suites.
'   0.1.4 - Added RunTest_Backup (30), RunTest_BlobStream (31),
'            RunTest_Serialize (32), RunTest_Diagnostics (33).
'            Added LOG_PATH file logging -- RunAllTests writes a full copy
'            of all output to LOG_PATH (set to "" to disable).
'            Fixed SQLite3Backup.cls: IsComplete returned True before the
'            first Step call (sqlite3_backup_remaining=0 before first step);
'            BackupToFile copied zero pages and produced an empty file.
'            Total: 305 tests across 33 suites.
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
' High-resolution timer (kernel32)
' ---------------------------------------------------------------------------
Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" _
    (lpPerformanceCount As LongPtr) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" _
    (lpFrequency As LongPtr) As Long

' Change these to match your environment
' Option A: place sqlite3.dll in C:\Windows\System32 (recommended)
'   - No Defender scanning overhead, found by name alone
Private Const DLL_PATH As String = "sqlite3.dll"

' Option B: explicit path outside System32
' Private Const DLL_PATH As String = "C:\sqlite\sqlite3.dll"
Private Const DB_PATH  As String = "C:\sqlite\driver_test.db"

' Log file path -- RunAllTests writes a copy of all output here.
' Set to "" to disable file logging.
Private Const LOG_PATH As String = "C:\sqlite\test_results.log"

' ---------------------------------------------------------------------------
' Test harness state
' ---------------------------------------------------------------------------
Private m_pass       As Long
Private m_fail       As Long
Private m_suite      As String
Private m_suiteStart As LongPtr   ' QPC ticks at suite start
Private m_runStart   As LongPtr   ' QPC ticks at RunAllTests start
Private m_freq       As LongPtr   ' QPC frequency (ticks per second)
Private m_failLog()  As String    ' accumulated failure messages for end summary
Private m_failCount  As Long      ' number of entries in m_failLog
Private m_logFile    As Integer   ' VBA file handle; 0 = not open

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

' Write a line to both the Immediate window and the log file (if open).
Private Sub Log(ByVal msg As String)
    Debug.Print msg
    If m_logFile <> 0 Then
        On Error Resume Next
        Print #m_logFile, msg
        On Error GoTo 0
    End If
End Sub

' Elapsed milliseconds between two QPC readings, formatted to 2 decimal places
Private Function ElapsedMs(ByVal t0 As LongPtr, ByVal t1 As LongPtr) As String
    EnsureFreq
    Dim ms As Double
    ms = (CDbl(t1) - CDbl(t0)) / CDbl(m_freq) * 1000#
    ElapsedMs = Format(ms, "0.00") & " ms"
End Function

' ---------------------------------------------------------------------------
' Suite helpers
' ---------------------------------------------------------------------------
Private Sub StartSuite(ByVal name As String)
    m_suite      = name
    m_suiteStart = QPC()
    Log ""
    Log "  [" & name & "]"
End Sub

Private Sub EndSuite()
    Dim elapsed As String: elapsed = ElapsedMs(m_suiteStart, QPC())
    Log "    TIME  " & elapsed
End Sub

Private Sub Pass(ByVal name As String)
    m_pass = m_pass + 1
    Log "    PASS  " & name
End Sub

Private Sub Fail(ByVal name As String, ByVal detail As String)
    m_fail = m_fail + 1
    Log "    FAIL  " & name & " -- " & detail
    ' Append to failure log for end-of-run summary
    If m_failCount = 0 Then
        ReDim m_failLog(0)
    Else
        ReDim Preserve m_failLog(m_failCount)
    End If
    m_failLog(m_failCount) = "[" & m_suite & "]  " & name & " -- " & detail
    m_failCount = m_failCount + 1
End Sub

Private Sub AssertEqual(ByVal name As String, ByVal got As Variant, ByVal expected As Variant)
    If CStr(got) = CStr(expected) Then
        Pass name
    Else
        Fail name, "expected [" & CStr(expected) & "] got [" & CStr(got) & "]"
    End If
End Sub

Private Sub AssertTrue(ByVal name As String, ByVal condition As Boolean)
    If condition Then Pass name Else Fail name, "condition was False"
End Sub

Private Sub AssertFalse(ByVal name As String, ByVal condition As Boolean)
    If Not condition Then Pass name Else Fail name, "condition was True"
End Sub

Private Sub AssertNull(ByVal name As String, ByVal v As Variant)
    If IsNull(v) Then Pass name Else Fail name, "expected Null, got [" & CStr(v) & "]"
End Sub

Private Sub AssertNoError(ByVal name As String)
    If Err.Number = 0 Then Pass name Else Fail name, Err.Description
    Err.Clear
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------
Private Function FreshConn() As SQLite3Connection
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 0
    Set FreshConn = conn
End Function

Private Sub DropTable(ByVal conn As SQLite3Connection, ByVal tbl As String)
    On Error Resume Next
    conn.ExecSQL "DROP TABLE IF EXISTS [" & tbl & "];"
    On Error GoTo 0
End Sub

'==============================================================================
' RunAllTests
'==============================================================================
Public Sub RunAllTests()
    m_pass      = 0
    m_fail      = 0
    m_failCount = 0
    EnsureFreq
    m_runStart = QPC()

    ' Open log file if LOG_PATH is set
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
    Log "SQLite3 Driver Test Suite"
    Log "Started: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    Log String(64, "=")

    RunTest_DllLoad
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
    RunTest_BulkInsert_AppendRow
    RunTest_BulkInsert_AppendMatrix
    RunTest_StatementCache
    RunTest_ConnectionPool
    RunTest_LargeDataset
    RunTest_SpecialCharacters
    RunTest_Boundaries
    RunTest_ErrorHandling
    RunTest_BLOB
    RunTest_Aggregates
    RunTest_FTS5
    RunTest_Schema
    RunTest_Savepoints
    RunTest_JSON
    RunTest_Interrupt
    RunTest_Backup
    RunTest_BlobStream
    RunTest_Serialize
    RunTest_Diagnostics

    Dim totalTime As String: totalTime = ElapsedMs(m_runStart, QPC())
    Log ""
    Log String(64, "=")
    Log "Results: " & m_pass & " passed,  " & m_fail & " failed  " & _
                "(" & (m_pass + m_fail) & " total)  " & totalTime
    Log String(64, "=")

    ' Failure summary -- only printed when there are failures
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

    ' Close log file
    If m_logFile <> 0 Then
        On Error Resume Next
        Close #m_logFile
        On Error GoTo 0
        m_logFile = 0
        Debug.Print ""
        Debug.Print "Log written to: " & LOG_PATH
    End If

    ' Final cleanup
    On Error Resume Next
    Kill DB_PATH
    On Error GoTo 0
End Sub

'==============================================================================
' 1. DLL load / version
'==============================================================================
Public Sub RunTest_DllLoad()
    StartSuite "DllLoad"
    On Error Resume Next

    SQLite3_API.SQLite_Unload
    SQLite3_API.SQLite_Load DLL_PATH
    AssertNoError "SQLite_Load"
    AssertTrue    "SQLite_IsLoaded", SQLite3_API.SQLite_IsLoaded()

    Dim ver As String: ver = SQLite3_API.SQLite_Version()
    AssertTrue "Version non-empty", Len(ver) > 0
    AssertTrue "Version starts with 3", Left(ver, 1) = "3"
    Log "    INFO  SQLite version = " & ver

    SQLite3_API.SQLite_Unload
    AssertFalse "SQLite_IsLoaded after unload", SQLite3_API.SQLite_IsLoaded()

    ' Reload for remaining tests
    SQLite3_API.SQLite_Load DLL_PATH
    EndSuite
End Sub

'==============================================================================
' 2. Open / close
'==============================================================================
Public Sub RunTest_OpenClose()
    StartSuite "OpenClose"
    On Error Resume Next

    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH
    AssertTrue  "IsOpen after OpenDatabase", conn.IsOpen
    AssertTrue  "Handle non-zero", conn.Handle <> 0
    AssertEqual "DbPath", conn.DbPath, DB_PATH

    conn.CloseConnection
    AssertFalse "IsOpen after CloseConnection", conn.IsOpen

    ' Double close must not crash
    Err.Clear
    conn.CloseConnection
    AssertNoError "Double CloseConnection safe"
    EndSuite
End Sub

'==============================================================================
' 3. ExecSQL / basic DDL + DML
'==============================================================================
Public Sub RunTest_ExecSQL()
    StartSuite "ExecSQL"
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
' 4. Scalar types returned correctly
'==============================================================================
Public Sub RunTest_ScalarTypes()
    StartSuite "ScalarTypes"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_types"
    conn.ExecSQL "CREATE TABLE t_types (i INTEGER, f REAL, t TEXT, b BLOB, n);"
    conn.ExecSQL "INSERT INTO t_types VALUES (42, 3.14, 'hello', X'DEADBEEF', NULL);"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f, t, n FROM t_types;")

    AssertFalse "Not EOF", rs.EOF
    AssertEqual "INTEGER value",  rs!i, 42
    AssertTrue  "FLOAT close",    Abs(CDbl(rs!f) - 3.14) < 0.0001
    AssertEqual "TEXT value",     rs!t, "hello"
    AssertNull  "NULL value",     rs!n

    rs.CloseRecordset
    DropTable conn, "t_types"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 5. NULL handling
'==============================================================================
Public Sub RunTest_NullHandling()
    StartSuite "NullHandling"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_null"
    conn.ExecSQL "CREATE TABLE t_null (a INTEGER, b TEXT);"
    conn.ExecSQL "INSERT INTO t_null VALUES (NULL, NULL);"
    conn.ExecSQL "INSERT INTO t_null VALUES (1, 'x');"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT a, b FROM t_null ORDER BY a;")

    ' First row: nulls
    AssertNull "a is NULL", rs!a
    AssertNull "b is NULL", rs!b
    rs.MoveNext

    ' Second row: values
    AssertEqual "a = 1",  rs!a, 1
    AssertEqual "b = x",  rs!b, "x"
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
    StartSuite "UTF8"
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
    StartSuite "PreparedStatements"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_prep"
    conn.ExecSQL "CREATE TABLE t_prep (i INTEGER, f REAL, t TEXT);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_prep VALUES (?, ?, ?);"

    cmd.BindInt    1, 7
    cmd.BindDouble 2, 2.718
    cmd.BindText   3, "Euler"
    cmd.Execute
    cmd.Reset

    cmd.BindNull 1
    cmd.BindInt  2, 0
    cmd.BindNull 3
    cmd.Execute
    cmd.Reset

    AssertEqual "2 rows inserted", TableRowCount(conn, "t_prep"), 2

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f, t FROM t_prep ORDER BY rowid;")

    AssertEqual "Row1 i",  rs!i, 7
    AssertTrue  "Row1 f",  Abs(CDbl(rs!f) - 2.718) < 0.0001
    AssertEqual "Row1 t",  rs!t, "Euler"
    rs.MoveNext

    AssertNull  "Row2 i null", rs!i
    AssertNull  "Row2 t null", rs!t
    rs.CloseRecordset

    ' ExecuteScalar
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
    StartSuite "NamedParams"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_named"
    conn.ExecSQL "CREATE TABLE t_named (a INTEGER, b TEXT);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_named VALUES (:a, :b);"
    cmd.BindIntByName  ":a", 99
    cmd.BindTextByName ":b", "ninety-nine"
    cmd.Execute
    cmd.Reset

    Dim v As Variant
    v = QueryScalar(conn, "SELECT b FROM t_named WHERE a=99;")
    AssertEqual "Named param round-trip", v, "ninety-nine"

    DropTable conn, "t_named"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 9. Transactions - commit
'==============================================================================
Public Sub RunTest_Transactions()
    StartSuite "Transactions"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_tx"
    conn.ExecSQL "CREATE TABLE t_tx (n INTEGER);"

    conn.BeginTransaction
    AssertTrue "InTransaction", conn.InTransaction
    Dim i As Long
    For i = 1 To 100
        conn.ExecSQL "INSERT INTO t_tx VALUES (" & i & ");"
    Next i
    conn.CommitTransaction
    AssertFalse "Not InTransaction after commit", conn.InTransaction

    AssertEqual "100 rows committed", TableRowCount(conn, "t_tx"), 100

    DropTable conn, "t_tx"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 10. Transactions - rollback
'==============================================================================
Public Sub RunTest_RollbackTransaction()
    StartSuite "RollbackTransaction"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_rb"
    conn.ExecSQL "CREATE TABLE t_rb (n INTEGER);"
    conn.ExecSQL "INSERT INTO t_rb VALUES (1);"

    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_rb VALUES (2);"
    conn.ExecSQL "INSERT INTO t_rb VALUES (3);"
    conn.RollbackTransaction
    AssertFalse "Not InTransaction after rollback", conn.InTransaction

    AssertEqual "Only 1 row survives rollback", TableRowCount(conn, "t_rb"), 1

    DropTable conn, "t_rb"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 11. Live recordset navigation
'==============================================================================
Public Sub RunTest_Recordset_Live()
    StartSuite "Recordset_Live"
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
        sum = sum + CLng(rs!n)
        rs.MoveNext
    Loop
    AssertEqual "Sum 1..5 = 15", sum, 15
    AssertTrue  "EOF after last", rs.EOF

    ' Empty query
    Dim rs2 As SQLite3Recordset
    Set rs2 = conn.OpenRecordset("SELECT n FROM t_live WHERE n > 999;")
    AssertTrue  "Empty rs BOF",  rs2.BOF
    AssertTrue  "Empty rs EOF",  rs2.EOF

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
    StartSuite "Recordset_Vectorized"
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

    AssertEqual "LoadAll returns 10",  cnt, 10
    AssertEqual "RecordCount = 10",    rs.RecordCount, 10
    AssertEqual "FieldCount = 2",      rs.FieldCount, 2
    AssertFalse "Not EOF at start",    rs.EOF

    ' MoveFirst / navigation
    rs.MoveFirst
    AssertEqual "First row n=1",  rs!n, 1
    AssertEqual "First row s=r1", rs!s, "r1"

    rs.MoveLast
    AssertEqual "Last row n=10", rs!n, 10

    ' MoveNext exhaustion
    rs.MoveFirst
    Dim sum As Long
    Do While Not rs.EOF
        sum = sum + CLng(rs!n)
        rs.MoveNext
    Loop
    AssertEqual "Sum 1..10 = 55", sum, 55

    ' Index access
    rs.MoveFirst
    AssertEqual "Field by index 0", rs.Item(0), 1
    AssertEqual "Field by name n",  rs.Item("n"), 1

    ' Column names
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
    StartSuite "Recordset_GetRows"
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

    ' First page of 3
    Dim pg1 As Variant: pg1 = rs.GetRows(3)
    AssertEqual "GetRows page1 col dim",  UBound(pg1, 1), 0   ' 1 col, 0-based
    AssertEqual "GetRows page1 row dim",  UBound(pg1, 2), 2   ' 3 rows, 0-based
    AssertEqual "GetRows page1 r0 = 1",   pg1(0, 0), 1
    AssertEqual "GetRows page1 r2 = 3",   pg1(0, 2), 3

    ' Second page of 3
    Dim pg2 As Variant: pg2 = rs.GetRows(3)
    AssertEqual "GetRows page2 r0 = 4",   pg2(0, 0), 4
    AssertEqual "GetRows page2 r2 = 6",   pg2(0, 2), 6
    AssertTrue  "EOF after two pages",     rs.EOF

    rs.CloseRecordset
    DropTable conn, "t_gr"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 14. ToMatrix - shape and values
'==============================================================================
Public Sub RunTest_Recordset_ToMatrix()
    StartSuite "Recordset_ToMatrix"
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

    ' Shape: (rows, cols) for Excel
    AssertEqual "Matrix row dim", UBound(mat, 1), 2   ' 3 rows, 0-based
    AssertEqual "Matrix col dim", UBound(mat, 2), 1   ' 2 cols, 0-based

    AssertEqual "mat(0,0) = 1",   mat(0, 0), 1
    AssertTrue  "mat(0,1) ~1.1",  Abs(CDbl(mat(0, 1)) - 1.1) < 0.001
    AssertEqual "mat(2,0) = 3",   mat(2, 0), 3
    AssertTrue  "mat(2,1) ~3.3",  Abs(CDbl(mat(2, 1)) - 3.3) < 0.001

    rs.CloseRecordset
    DropTable conn, "t_mat"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 15. BulkInsert - AppendRow
'==============================================================================
Public Sub RunTest_BulkInsert_AppendRow()
    StartSuite "BulkInsert_AppendRow"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
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

    ' Spot-check first and last
    Dim v As Variant
    v = QueryScalar(conn, "SELECT s FROM t_bulk WHERE i=1;")
    AssertEqual "First row s", v, "row1"
    v = QueryScalar(conn, "SELECT s FROM t_bulk WHERE i=500;")
    AssertEqual "Last row s", v, "row500"
    v = QueryScalar(conn, "SELECT f FROM t_bulk WHERE i=2;")
    AssertTrue "Row2 f ~1.0", Abs(CDbl(v) - 1.0) < 0.001

    DropTable conn, "t_bulk"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 16. BulkInsert - AppendMatrix
'==============================================================================
Public Sub RunTest_BulkInsert_AppendMatrix()
    StartSuite "BulkInsert_AppendMatrix"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_mat2"
    conn.ExecSQL "CREATE TABLE t_mat2 (a INTEGER, b TEXT);"

    Const N As Long = 200
    Dim mat() As Variant
    ReDim mat(N - 1, 1)
    Dim i As Long
    For i = 0 To N - 1
        mat(i, 0) = i + 1
        mat(i, 1) = "m" & (i + 1)
    Next i

    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, "t_mat2", Array("a", "b")
    bulk.AppendMatrix mat
    bulk.CloseInsert

    AssertEqual "TotalRowsInserted", bulk.TotalRowsInserted, N
    AssertEqual "Row count in DB",   TableRowCount(conn, "t_mat2"), N

    Dim v As Variant
    v = QueryScalar(conn, "SELECT b FROM t_mat2 WHERE a=100;")
    AssertEqual "Mid row b", v, "m100"

    DropTable conn, "t_mat2"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 17. Statement cache - cache hit reuse
'==============================================================================
Public Sub RunTest_StatementCache()
    StartSuite "StatementCache"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_cache"
    conn.ExecSQL "CREATE TABLE t_cache (n INTEGER);"

    ' Same SQL string prepared many times should reuse the cached stmt
    Dim sql As String: sql = "INSERT INTO t_cache VALUES (?);"
    Dim i As Long
    For i = 1 To 10
        Dim cmd As New SQLite3Command
        cmd.Prepare conn, sql
        cmd.BindInt 1, i
        cmd.Execute
    Next i

    AssertEqual "10 rows via cached stmt", TableRowCount(conn, "t_cache"), 10

    ' Verify values are correct
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT SUM(n) FROM t_cache;")
    AssertEqual "SUM = 55", rs.Item(0), 55
    rs.CloseRecordset

    DropTable conn, "t_cache"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 18. Connection pool
'==============================================================================
Public Sub RunTest_ConnectionPool()
    StartSuite "ConnectionPool"
    On Error Resume Next

    ' Need at least the table to exist
    Dim setup As SQLite3Connection: Set setup = FreshConn()
    DropTable setup, "t_pool"
    setup.ExecSQL "CREATE TABLE t_pool (n INTEGER);"
    setup.ExecSQL "INSERT INTO t_pool VALUES (42);"
    setup.CloseConnection

    Dim pool As New SQLite3Pool
    pool.Initialize DB_PATH, DLL_PATH, 3

    AssertEqual "Initial pool size = 1", pool.PoolSize, 1
    AssertEqual "Initial active = 0",    pool.ActiveConnections, 0

    Dim c1 As SQLite3Connection: Set c1 = pool.Acquire()
    Dim c2 As SQLite3Connection: Set c2 = pool.Acquire()
    AssertEqual "Active after 2 acquires", pool.ActiveConnections, 2

    ' Query through pooled connection
    Dim rs As SQLite3Recordset
    Set rs = c1.OpenRecordset("SELECT n FROM t_pool;")
    AssertEqual "Pool query result", rs!n, 42
    rs.CloseRecordset

    pool.ReleaseConnection c1
    pool.ReleaseConnection c2
    AssertEqual "Active after 2 releases", pool.ActiveConnections, 0

    ' Verify auto-rollback on release
    Dim c3 As SQLite3Connection: Set c3 = pool.Acquire()
    c3.BeginTransaction
    c3.ExecSQL "INSERT INTO t_pool VALUES (99);"
    AssertTrue "InTransaction before release", c3.InTransaction
    pool.ReleaseConnection c3
    ' After release the transaction should have been rolled back
    Dim c4 As SQLite3Connection: Set c4 = pool.Acquire()
    AssertEqual "Rollback on release", _
        QueryScalar(c4, "SELECT COUNT(*) FROM t_pool WHERE n=99;"), 0
    pool.ReleaseConnection c4

    pool.ShutDown
    AssertEqual "PoolSize after shutdown", pool.PoolSize, 0

    Dim cleanup As SQLite3Connection: Set cleanup = FreshConn()
    DropTable cleanup, "t_pool"
    cleanup.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 19. Large dataset - 10k rows vectorized
'==============================================================================
Public Sub RunTest_LargeDataset()
    StartSuite "LargeDataset"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_large"
    conn.ExecSQL "CREATE TABLE t_large (i INTEGER, f REAL, s TEXT);"

    Const N As Long = 10000
    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, "t_large", Array("i", "f", "s"), 1000
    Dim i As Long
    For i = 1 To N
        bulk.AppendRow Array(i, i * 1.5, "s" & i)
    Next i
    bulk.CloseInsert

    AssertEqual "10k rows inserted", TableRowCount(conn, "t_large"), N

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f, s FROM t_large ORDER BY i;")
    Dim cnt As Long: cnt = rs.LoadAll()
    AssertEqual "LoadAll = 10000", cnt, N

    ' Spot-check a few rows
    rs.MoveFirst
    AssertEqual "Row 1 i", rs!i, 1

    Dim mat As Variant: mat = rs.ToMatrix()
    AssertEqual "Matrix rows",    UBound(mat, 1) + 1, N
    AssertEqual "Matrix cols",    UBound(mat, 2) + 1, 3
    AssertEqual "Last row i",     mat(N - 1, 0), N
    AssertTrue  "Last row f",     Abs(CDbl(mat(N - 1, 1)) - (N * 1.5)) < 0.001
    AssertEqual "Last row s",     mat(N - 1, 2), "s" & N

    rs.CloseRecordset
    DropTable conn, "t_large"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 20. Special characters in strings
'==============================================================================
Public Sub RunTest_SpecialCharacters()
    StartSuite "SpecialCharacters"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_special"
    conn.ExecSQL "CREATE TABLE t_special (s TEXT);"

    Dim cases As Variant
    ' single quote, newline, tab, long string, empty string
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
    StartSuite "Boundaries"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_bounds"
    conn.ExecSQL "CREATE TABLE t_bounds (i INTEGER, f REAL);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_bounds VALUES (?, ?);"

    ' Max Long
    cmd.BindInt 1, 2147483647: cmd.BindDouble 2, 1.7976931348623E+308
    cmd.Execute: cmd.Reset
    ' Min Long
    cmd.BindInt 1, -2147483648: cmd.BindDouble 2, -1.7976931348623E+308
    cmd.Execute: cmd.Reset
    ' Zero
    cmd.BindInt 1, 0: cmd.BindDouble 2, 0
    cmd.Execute: cmd.Reset

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT i, f FROM t_bounds ORDER BY rowid;")
    rs.LoadAll

    AssertEqual "Max Long",   rs!i, 2147483647
    rs.MoveNext
    AssertEqual "Min Long",   rs!i, -2147483648
    rs.MoveNext
    AssertEqual "Zero int",   rs!i, 0
    AssertEqual "Zero float",  rs!f, 0

    rs.CloseRecordset
    DropTable conn, "t_bounds"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 22. Error handling - bad SQL, missing table
'==============================================================================
Public Sub RunTest_ErrorHandling()
    StartSuite "ErrorHandling"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    ' Bad SQL must raise an error
    Err.Clear
    conn.ExecSQL "THIS IS NOT SQL;"
    AssertTrue "Bad SQL raises error", Err.Number <> 0
    Err.Clear

    ' Query non-existent table
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "SELECT * FROM no_such_table_xyz;"
    cmd.Execute
    AssertTrue "Missing table raises error", Err.Number <> 0
    Err.Clear

    ' Named param not found
    Dim cmd2 As New SQLite3Command
    cmd2.Prepare conn, "SELECT ?;"
    cmd2.BindTextByName ":nosuchparam", "x"
    AssertTrue "Bad named param raises error", Err.Number <> 0
    Err.Clear

    ' Pool exhausted
    Dim setup As SQLite3Connection: Set setup = FreshConn()
    DropTable setup, "t_err"
    setup.ExecSQL "CREATE TABLE t_err (n INTEGER);"
    setup.CloseConnection

    Dim pool As New SQLite3Pool
    pool.Initialize DB_PATH, DLL_PATH, 2
    Dim p1 As SQLite3Connection: Set p1 = pool.Acquire()
    Dim p2 As SQLite3Connection: Set p2 = pool.Acquire()
    Err.Clear
    Dim p3 As SQLite3Connection: Set p3 = pool.Acquire()
    AssertTrue "Pool exhausted raises error", Err.Number <> 0
    Err.Clear
    pool.ReleaseConnection p1
    pool.ReleaseConnection p2
    pool.ShutDown

    Dim cleanup As SQLite3Connection: Set cleanup = FreshConn()
    DropTable cleanup, "t_err"
    cleanup.CloseConnection

    conn.CloseConnection
    EndSuite
End Sub
'==============================================================================
' 23. BLOB - bind, read live, read vectorized
'==============================================================================
Public Sub RunTest_BLOB()
    StartSuite "BLOB"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_blob"
    conn.ExecSQL "CREATE TABLE t_blob (id INTEGER, data BLOB, label TEXT);"

    ' Build test payloads
    Dim small() As Byte: ReDim small(4)
    small(0)=1: small(1)=2: small(2)=3: small(3)=255: small(4)=0

    Dim large() As Byte: ReDim large(999)
    Dim i As Long
    For i = 0 To 999: large(i) = CByte(i Mod 256): Next i

    ' Insert via BindBlob
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO t_blob VALUES (?, ?, ?);"

    cmd.BindInt  1, 1
    cmd.BindBlob 2, small
    cmd.BindText 3, "small"
    cmd.Execute: cmd.Reset

    cmd.BindInt  1, 2
    cmd.BindBlob 2, large
    cmd.BindText 3, "large"
    cmd.Execute: cmd.Reset

    ' Insert via BindVariant (byte array path)
    cmd.BindInt     1, 3
    cmd.BindVariant 2, small
    cmd.BindText    3, "variant"
    cmd.Execute: cmd.Reset

    AssertEqual "3 BLOB rows", TableRowCount(conn, "t_blob"), 3

    ' Read back via live recordset
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT id, data, label FROM t_blob ORDER BY id;")

    ' Row 1 - small blob via Value
    AssertFalse "Row1 not EOF", rs.EOF
    Dim v1 As Variant: v1 = rs!data
    AssertTrue  "Row1 is byte array", VarType(v1) = (vbByte + vbArray)
    Dim b1() As Byte: b1 = v1
    AssertEqual "Row1 len=5",    UBound(b1) - LBound(b1) + 1, 5
    AssertEqual "Row1 b(0)=1",   b1(0), 1
    AssertEqual "Row1 b(3)=255", b1(3), 255

    ' Read via AsBytes typed accessor
    Dim ab() As Byte: ab = rs.Fields("data").AsBytes()
    AssertEqual "AsBytes len=5",  UBound(ab) - LBound(ab) + 1, 5
    AssertEqual "AsBytes b(1)=2", ab(1), 2

    rs.MoveNext

    ' Row 2 - large blob
    Dim v2 As Variant: v2 = rs!data
    Dim b2() As Byte: b2 = v2
    AssertEqual "Row2 len=1000",   UBound(b2) - LBound(b2) + 1, 1000
    AssertEqual "Row2 b(255)=255", b2(255), 255
    AssertEqual "Row2 b(256)=0",   b2(256), 0

    rs.MoveNext
    rs.CloseRecordset

    ' Vectorized BLOB load
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
' 24. Aggregates - helpers
'==============================================================================
Public Sub RunTest_Aggregates()
    StartSuite "Aggregates"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_agg"
    conn.ExecSQL "CREATE TABLE t_agg (grp TEXT, val REAL);"

    ' Insert test data: 3 groups, 4 rows each
    Dim g As Long, r As Long
    For g = 1 To 3
        For r = 1 To 4
            conn.ExecSQL "INSERT INTO t_agg VALUES ('g" & g & "', " & _
                         (g * 10 + r) & ");"
        Next r
    Next g
    ' g1: 11,12,13,14  g2: 21,22,23,24  g3: 31,32,33,34

    ' ScalarAgg
    Dim cnt As Variant: cnt = ScalarAgg(conn, "t_agg", "COUNT(*)")
    AssertEqual "ScalarAgg COUNT", cnt, 12

    Dim total As Variant: total = ScalarAgg(conn, "t_agg", "SUM(val)", "grp='g1'")
    AssertEqual "ScalarAgg SUM g1", total, 50   ' 11+12+13+14

    ' GroupByCount
    Dim gc As Variant: gc = GroupByCount(conn, "t_agg", "grp")
    AssertTrue "GroupByCount rows=3",  UBound(gc, 1) = 2   ' 3 rows 0-based
    AssertTrue "GroupByCount cols=2",  UBound(gc, 2) = 1
    AssertEqual "GroupByCount each=4", gc(0, 1), 4

    ' GroupBySum
    Dim gs As Variant: gs = GroupBySum(conn, "t_agg", "grp", "val")
    AssertTrue "GroupBySum rows=3", UBound(gs, 1) = 2
    ' highest sum first: g3=130, g2=90, g1=50
    AssertEqual "GroupBySum g3 total=130", gs(0, 1), 130

    ' GroupByAvg
    Dim ga As Variant: ga = GroupByAvg(conn, "t_agg", "grp", "val")
    AssertTrue "GroupByAvg rows=3",   UBound(ga, 1) = 2
    AssertEqual "GroupByAvg cols=3",  UBound(ga, 2), 2

    ' MultiAgg
    Dim ma As Variant
    ma = MultiAgg(conn, "t_agg", _
                  Array("COUNT(*) AS n", "SUM(val) AS s", "AVG(val) AS a", "MIN(val) AS lo"))
    AssertEqual "MultiAgg n=12",   ma(0, 0), 12
    AssertEqual "MultiAgg s=270",  ma(0, 1), 270  ' sum all = 50+90+130
    AssertEqual "MultiAgg lo=11",  ma(0, 3), 11

    ' AggregateQuery raw
    Dim mat As Variant
    mat = AggregateQuery(conn, "SELECT MAX(val), MIN(val) FROM t_agg;")
    AssertEqual "AggQuery MAX=34", mat(0, 0), 34
    AssertEqual "AggQuery MIN=11", mat(0, 1), 11

    ' RunningTotal (window function)
    Dim rt As Variant
    rt = RunningTotal(conn, "t_agg", "rowid", "val", "grp='g1'")
    AssertTrue "RunningTotal rows=4",   UBound(rt, 1) = 3
    AssertEqual "RunningTotal cols=3",  UBound(rt, 2), 2
    ' running total after 4th g1 row = 11+12+13+14 = 50
    AssertEqual "RunningTotal last=50", rt(3, 2), 50

    DropTable conn, "t_agg"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 25. FTS5 - create, insert, search, snippet, highlight, BM25
'==============================================================================
Public Sub RunTest_FTS5()
    StartSuite "FTS5"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    ' Check FTS5 is available
    On Error Resume Next
    conn.ExecSQL "CREATE VIRTUAL TABLE _fts5_probe USING fts5(x);"
    If Err.Number <> 0 Then
        Log "    SKIP  FTS5 not available in this sqlite3.dll build"
        Err.Clear
        conn.CloseConnection
        EndSuite
        Exit Sub
    End If
    conn.ExecSQL "DROP TABLE _fts5_probe;"
    Err.Clear
    On Error Resume Next

    ' Create FTS5 table
    CreateFTS5Table conn, "t_fts", Array("title", "body"), "", "unicode61", True
    AssertNoError "CreateFTS5Table"

    ' Insert rows
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

    ' Basic search
    Dim cnt As Long: cnt = FTS5MatchCount(conn, "t_fts", "SQLite")
    AssertEqual "FTS5 SQLite matches=2", cnt, 2

    cnt = FTS5MatchCount(conn, "t_fts", "data")
    AssertEqual "FTS5 data matches=2", cnt, 2

    ' Prefix search
    cnt = FTS5MatchCount(conn, "t_fts", "Optimi*")
    AssertEqual "FTS5 prefix Optimi*=1", cnt, 1

    ' Column-scoped search
    cnt = FTS5MatchCount(conn, "t_fts", "title : SQLite")
    AssertEqual "FTS5 title:SQLite=2", cnt, 2

    cnt = FTS5MatchCount(conn, "t_fts", "title : Python")
    AssertEqual "FTS5 title:Python=1", cnt, 1

    ' Boolean AND
    cnt = FTS5MatchCount(conn, "t_fts", "SQLite AND performance")
    AssertEqual "FTS5 AND=1", cnt, 1

    ' FTS5SearchMatrix
    Dim mat As Variant
    mat = FTS5SearchMatrix(conn, "t_fts", "SQLite", "*", "rank", 10)
    AssertTrue "SearchMatrix rows>=2", UBound(mat, 1) >= 1

    ' Snippet
    Dim snip As Variant
    snip = FTS5Snippet(conn, "t_fts", "SQLite", 0, "[", "]", "...", 8, 5)
    AssertTrue  "Snippet non-empty", Not IsEmpty(snip)
    ' snippet_text is last column; check it contains our markers
    Dim snippetText As String
    snippetText = CStr(snip(0, UBound(snip, 2)))
    AssertTrue "Snippet contains marker", InStr(snippetText, "[") > 0

    ' Highlight
    Dim hl As Variant
    hl = FTS5Highlight(conn, "t_fts", "SQLite", 0, "**", "**", 5)
    AssertTrue "Highlight non-empty", Not IsEmpty(hl)

    ' BM25 search
    Dim bm As Variant
    bm = FTS5BM25Search(conn, "t_fts", "SQLite data storage")
    AssertTrue  "BM25 non-empty",     Not IsEmpty(bm)
    ' score column (last) should be negative (bm25 returns negative in SQLite)
    AssertTrue "BM25 score negative", CDbl(bm(0, UBound(bm, 2))) < 0

    ' Bulk insert
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
    AssertEqual "FTS5 Bulk matches=100", cnt, 100

    ' Optimize
    Err.Clear
    FTS5Optimize conn, "t_fts"
    AssertNoError "FTS5Optimize"

    ' Delete
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT rowid FROM t_fts WHERE t_fts MATCH 'Python';")
    AssertFalse "Python row exists", rs.EOF
    Dim rid As LongLong: rid = CLngLng(rs.Item(0))
    rs.CloseRecordset
    FTS5Delete conn, "t_fts", rid
    cnt = FTS5MatchCount(conn, "t_fts", "title : Python")
    AssertEqual "After delete Python=0", cnt, 0

    conn.ExecSQL "DROP TABLE IF EXISTS t_fts;"
    conn.CloseConnection
    EndSuite
End Sub
'==============================================================================
' 26. Schema introspection
'==============================================================================
Public Sub RunTest_Schema()
    StartSuite "Schema"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    ' Setup: create two tables, a view, an index
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_a;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_b;"
    conn.ExecSQL "DROP VIEW  IF EXISTS v_schema;"
    conn.ExecSQL "DROP INDEX IF EXISTS ix_schema_a;"

    conn.ExecSQL _
        "CREATE TABLE t_schema_a (" & _
        "  id    INTEGER PRIMARY KEY, " & _
        "  name  TEXT NOT NULL, " & _
        "  score REAL DEFAULT 0.0);"
    conn.ExecSQL _
        "CREATE TABLE t_schema_b (" & _
        "  id    INTEGER PRIMARY KEY, " & _
        "  aid   INTEGER REFERENCES t_schema_a(id), " & _
        "  value BLOB);"
    conn.ExecSQL "CREATE INDEX ix_schema_a ON t_schema_a(name);"
    conn.ExecSQL "CREATE VIEW  v_schema AS SELECT id, name FROM t_schema_a;"

    ' GetTableList
    Dim tbls As Variant: tbls = GetTableList(conn)
    AssertTrue "TableList is array",  IsArray(tbls)
    Dim found As Boolean, i As Long
    For i = LBound(tbls) To UBound(tbls)
        If tbls(i) = "t_schema_a" Then found = True
    Next i
    AssertTrue "TableList has t_schema_a", found

    ' GetTableList excludes views by default
    Dim foundView As Boolean
    For i = LBound(tbls) To UBound(tbls)
        If tbls(i) = "v_schema" Then foundView = True
    Next i
    AssertFalse "TableList excludes view", foundView

    ' GetTableList with views included
    Dim all As Variant: all = GetTableList(conn, True)
    Dim foundV As Boolean
    For i = LBound(all) To UBound(all)
        If all(i) = "v_schema" Then foundV = True
    Next i
    AssertTrue "TableList(includeViews) has view", foundV

    ' GetViewList
    Dim views As Variant: views = GetViewList(conn)
    AssertTrue "ViewList non-empty", IsArray(views)
    AssertEqual "ViewList(0)=v_schema", views(0), "v_schema"

    ' TableExists / ViewExists / IndexExists
    AssertTrue  "TableExists t_schema_a", TableExists(conn, "t_schema_a")
    AssertFalse "TableExists nosuchtable", TableExists(conn, "nosuchtable")
    AssertTrue  "ViewExists v_schema",    ViewExists(conn, "v_schema")
    AssertTrue  "IndexExists ix_schema_a", IndexExists(conn, "ix_schema_a")
    AssertFalse "IndexExists nosuchindex", IndexExists(conn, "nosuchindex")

    ' GetColumnInfo
    Dim cols As Variant: cols = GetColumnInfo(conn, "t_schema_a")
    ' PRAGMA table_info returns: cid, name, type, notnull, dflt_value, pk
    AssertEqual "ColumnInfo 3 cols", UBound(cols, 1) - LBound(cols, 1) + 1, 3
    AssertEqual "Col0 name=id",   cols(0, 1), "id"
    AssertEqual "Col1 name=name", cols(1, 1), "name"
    AssertEqual "Col1 notnull=1", cols(1, 3), 1
    AssertEqual "Col2 default=0.0", CStr(cols(2, 4)), "0.0"
    AssertEqual "Col0 pk=1",      cols(0, 5), 1   ' id is PK

    ' GetIndexList
    Dim idxs As Variant: idxs = GetIndexList(conn, "t_schema_a")
    AssertTrue "IndexList non-empty", Not IsEmpty(idxs)
    ' ix_schema_a should be in the list
    Dim foundIdx As Boolean
    Dim r As Long
    If IsArray(idxs) Then
        For r = LBound(idxs, 1) To UBound(idxs, 1)
            If CStr(idxs(r, 1)) = "ix_schema_a" Then foundIdx = True
        Next r
    End If
    AssertTrue "IndexList has ix_schema_a", foundIdx

    ' GetIndexColumns
    Dim ixCols As Variant: ixCols = GetIndexColumns(conn, "ix_schema_a")
    AssertTrue  "IndexColumns non-empty",  Not IsEmpty(ixCols)
    AssertEqual "IndexColumns col=name",   ixCols(0, 2), "name"

    ' GetForeignKeys
    Dim fks As Variant: fks = GetForeignKeys(conn, "t_schema_b")
    AssertTrue  "ForeignKeys non-empty",          Not IsEmpty(fks)
    AssertEqual "FK refs t_schema_a",  fks(0, 2), "t_schema_a"
    AssertEqual "FK from=aid",         fks(0, 3), "aid"

    ' GetCreateSQL
    Dim sql As String: sql = GetCreateSQL(conn, "t_schema_a")
    AssertTrue  "CreateSQL non-empty",      Len(sql) > 0
    AssertTrue  "CreateSQL has CREATE TABLE", InStr(sql, "CREATE TABLE") > 0

    ' GetDatabaseInfo
    Dim dbInfo As Variant: dbInfo = GetDatabaseInfo(conn)
    AssertTrue "DatabaseInfo is matrix", IsArray(dbInfo)
    AssertTrue "DatabaseInfo rows>5",    UBound(dbInfo, 1) >= 5
    AssertEqual "DatabaseInfo col0=page_count", dbInfo(0, 0), "page_count"
    AssertTrue  "DatabaseInfo page_size>0", CLng(dbInfo(1, 1)) > 0

    ' Cleanup
    conn.ExecSQL "DROP INDEX IF EXISTS ix_schema_a;"
    conn.ExecSQL "DROP VIEW  IF EXISTS v_schema;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_b;"
    conn.ExecSQL "DROP TABLE IF EXISTS t_schema_a;"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 27. Savepoints
'==============================================================================
Public Sub RunTest_Savepoints()
    StartSuite "Savepoints"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_sp"
    conn.ExecSQL "CREATE TABLE t_sp (id INTEGER, val TEXT);"

    ' Basic savepoint: release (commit)
    conn.BeginTransaction
    conn.ExecSQL "INSERT INTO t_sp VALUES (1, 'outer');"
    conn.Savepoint "sp1"
    AssertEqual "SavepointDepth=1", conn.SavepointDepth, 1
    conn.ExecSQL "INSERT INTO t_sp VALUES (2, 'inner');"
    conn.ReleaseSavepoint "sp1"
    AssertEqual "SavepointDepth=0 after release", conn.SavepointDepth, 0
    conn.CommitTransaction
    AssertEqual "Both rows committed", TableRowCount(conn, "t_sp"), 2

    ' Savepoint rollback: undo inner, keep outer
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

    ' Nested savepoints
    conn.ExecSQL "DELETE FROM t_sp;"
    conn.BeginTransaction
    conn.Savepoint "outer"
    conn.ExecSQL "INSERT INTO t_sp VALUES (20, 'level1');"
    conn.Savepoint "inner"
    conn.ExecSQL "INSERT INTO t_sp VALUES (21, 'level2');"
    AssertEqual "SavepointDepth=2", conn.SavepointDepth, 2
    conn.RollbackToSavepoint "inner"
    AssertEqual "SavepointDepth still 2", conn.SavepointDepth, 2
    conn.ReleaseSavepoint "inner"
    AssertEqual "SavepointDepth=1 after inner release", conn.SavepointDepth, 1
    conn.ReleaseSavepoint "outer"
    conn.CommitTransaction
    ' Only level1 row should exist
    AssertEqual "Nested: only level1", TableRowCount(conn, "t_sp"), 1
    AssertEqual "Nested: val=level1", _
        QueryScalar(conn, "SELECT val FROM t_sp WHERE id=20;"), "level1"

    DropTable conn, "t_sp"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 28. JSON functions
'==============================================================================
Public Sub RunTest_JSON()
    StartSuite "JSON"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()

    ' Check JSON functions available (3.38+)
    On Error Resume Next
    Dim probe As Variant
    probe = QueryScalar(conn, "SELECT json_valid('{""a"":1}');")
    If Err.Number <> 0 Then
        Log "    SKIP  JSON functions not available (requires SQLite 3.38+)"
        Err.Clear
        conn.CloseConnection
        EndSuite
        Exit Sub
    End If
    Err.Clear
    On Error Resume Next

    DropTable conn, "t_json"
    conn.ExecSQL _
        "CREATE TABLE t_json (id INTEGER PRIMARY KEY, data TEXT);"

    ' Insert rows with JSON data
    conn.ExecSQL "INSERT INTO t_json VALUES (1, '{""name"":""Alice"",""city"":""London"",""score"":95}');"
    conn.ExecSQL "INSERT INTO t_json VALUES (2, '{""name"":""Bob"",""city"":""Paris"",""score"":82}');"
    conn.ExecSQL "INSERT INTO t_json VALUES (3, '{""name"":""Carol"",""city"":""Berlin"",""tags"":[""vba"",""excel""]}');"

    ' JSONExtract -- single value
    Dim v As Variant
    v = JSONExtract(conn, "t_json", "data", "$.name", "id=1")
    AssertEqual "JSONExtract name=Alice", v, "Alice"

    v = JSONExtract(conn, "t_json", "data", "$.score", "id=2")
    AssertEqual "JSONExtract score=82", CDbl(v), 82

    ' Nested path
    v = JSONExtract(conn, "t_json", "data", "$.tags[0]", "id=3")
    AssertEqual "JSONExtract tags[0]=vba", v, "vba"

    ' JSONExtractColumn
    Dim mat As Variant
    mat = JSONExtractColumn(conn, "t_json", "data", "$.city")
    AssertEqual "JSONExtractColumn rows=3", UBound(mat,1) - LBound(mat,1) + 1, 3

    ' JSONSearch
    Dim res As Variant
    res = JSONSearch(conn, "t_json", "data", "$.city", "'London'")
    AssertEqual "JSONSearch London rows=1", UBound(res,1) - LBound(res,1) + 1, 1

    ' JSONSet -- update existing key
    JSONSet conn, "t_json", "data", "$.city", "'Madrid'", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.city", "id=1")
    AssertEqual "JSONSet city=Madrid", v, "Madrid"

    ' JSONSet -- add new key
    JSONSet conn, "t_json", "data", "$.active", "1", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.active", "id=1")
    AssertEqual "JSONSet new key=1", CLng(v), 1

    ' JSONInsert -- does nothing if key exists
    JSONInsert conn, "t_json", "data", "$.city", "'Tokyo'", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.city", "id=1")
    AssertEqual "JSONInsert keeps Madrid", v, "Madrid"

    ' JSONReplace -- updates existing key
    JSONReplace conn, "t_json", "data", "$.score", "100", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.score", "id=1")
    AssertEqual "JSONReplace score=100", CDbl(v), 100

    ' JSONRemove
    JSONRemove conn, "t_json", "data", Array("$.active"), "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.active", "id=1")
    AssertTrue "JSONRemove key gone", IsNull(v)

    ' JSONPatch
    JSONPatch conn, "t_json", "data", "'{""country"":""Spain""}'", "id=1"
    v = JSONExtract(conn, "t_json", "data", "$.country", "id=1")
    AssertEqual "JSONPatch country=Spain", v, "Spain"

    ' JSONValid
    AssertTrue "JSONValid all rows", JSONValid(conn, "t_json", "data")
    conn.ExecSQL "INSERT INTO t_json VALUES (99, 'not json at all');"
    AssertFalse "JSONValid bad row", JSONValid(conn, "t_json", "data")
    conn.ExecSQL "DELETE FROM t_json WHERE id=99;"

    ' JSONGroupArray
    Dim arr As String
    arr = JSONGroupArray(conn, "t_json", "json_extract(data,'$.name')", "", "id")
    AssertTrue  "JSONGroupArray non-empty", Len(arr) > 2
    AssertTrue  "JSONGroupArray has Alice", InStr(arr, "Alice") > 0
    AssertTrue  "JSONGroupArray has Bob",   InStr(arr, "Bob") > 0

    ' JSONType
    Dim types As Variant
    types = JSONType(conn, "t_json", "data", "$.score")
    AssertTrue "JSONType rows=3", UBound(types,1) - LBound(types,1) + 1 = 3
    ' score is stored as integer or real depending on value
    AssertTrue "JSONType score is numeric", _
        CStr(types(0, 1)) = "integer" Or CStr(types(0, 1)) = "real"

    ' JSONEach -- expand tags array for row 3
    Dim elems As Variant
    elems = JSONEach(conn, "t_json", "data", "t.id=3")
    AssertTrue "JSONEach rows>0", Not IsEmpty(elems)

    ' JSONBuildObject
    Dim obj As String
    obj = JSONBuildObject(conn, Array("key", "val"), Array("'hello'", "42"))
    AssertTrue "JSONBuildObject has key", InStr(obj, "key") > 0
    AssertTrue "JSONBuildObject has 42",  InStr(obj, "42") > 0

    ' JSONBuildArray
    Dim jarr As String
    jarr = JSONBuildArray(conn, Array("1", "2", "'three'"))
    AssertTrue "JSONBuildArray has 1",     InStr(jarr, "1") > 0
    AssertTrue "JSONBuildArray has three", InStr(jarr, "three") > 0

    DropTable conn, "t_json"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 29. Interrupt
'==============================================================================
Public Sub RunTest_Interrupt()
    StartSuite "Interrupt"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_interrupt"
    conn.ExecSQL "CREATE TABLE t_interrupt (x INTEGER);"

    ' Insert enough rows to make a query measurable
    Dim i As Long
    conn.BeginTransaction
    For i = 1 To 50000
        conn.ExecSQL "INSERT INTO t_interrupt VALUES (" & i & ");"
    Next i
    conn.CommitTransaction
    AssertEqual "50k rows inserted", TableRowCount(conn, "t_interrupt"), 50000

    ' Interrupt an idle connection: should be a no-op and not crash
    Err.Clear
    conn.Interrupt
    AssertNoError "Interrupt on idle connection"

    ' Run a legitimate query after interrupt -- should work normally
    Dim n As Variant
    n = QueryScalar(conn, "SELECT COUNT(*) FROM t_interrupt;")
    AssertEqual "Query after interrupt OK", CLng(n), 50000

    ' Interrupt immediately before a query:
    ' SQLite clears the interrupt flag once checked, so if we call Interrupt
    ' before opening a statement the subsequent query may or may not be
    ' interrupted depending on timing.  We just verify no crash either way.
    conn.Interrupt
    Err.Clear
    n = QueryScalar(conn, "SELECT SUM(x) FROM t_interrupt;")
    ' Accept either a numeric result or SQLITE_INTERRUPT (rc 9 surfaces as error)
    AssertTrue "No crash after interrupt+query", _
        (Err.Number = 0 And Not IsNull(n)) Or Err.Number <> 0
    Err.Clear

    DropTable conn, "t_interrupt"
    conn.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 30. Online Backup API
'==============================================================================
Public Sub RunTest_Backup()
    StartSuite "Backup"
    On Error Resume Next

    ' Use a dedicated source file opened WITHOUT WAL from creation.
    ' FreshConn() opens driver_test.db in WAL mode; after 29 prior suites its
    ' statement cache holds open read snapshots that prevent wal_checkpoint(TRUNCATE)
    ' from resetting the WAL.  A fresh rollback-journal-only file has a clean
    ' page 1 from the start -- no checkpoint tricks needed.
    Dim srcPath As String
    srcPath = Left(DB_PATH, Len(DB_PATH) - 3) & "_baksrc.db"
    Kill srcPath:          Err.Clear
    Kill srcPath & "-wal": Err.Clear
    Kill srcPath & "-shm": Err.Clear

    Dim src As New SQLite3Connection
    src.OpenDatabase srcPath, DLL_PATH, 5000, False   ' enableWAL=False
    AssertNoError "Open backup source DB"
    src.ExecSQL "CREATE TABLE t_bak (id INTEGER PRIMARY KEY, val TEXT);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 500
        src.ExecSQL "INSERT INTO t_bak VALUES (" & i & ", 'row_" & i & "');"
    Next i
    src.CommitTransaction
    AssertEqual "Source rows before backup", TableRowCount(src, "t_bak"), 500

    ' Full one-shot backup
    Dim destPath As String
    destPath = Left(DB_PATH, Len(DB_PATH) - 3) & "_bak.db"
    Kill destPath: Err.Clear

    Dim bak As New SQLite3Backup
    bak.BackupToFile src, destPath
    AssertNoError "BackupToFile no error"
    AssertTrue "Backup IsComplete", bak.IsComplete
    AssertFalse "Backup not open after finish", bak.IsOpen

    ' Verify backup
    Dim dest As New SQLite3Connection
    dest.OpenDatabase destPath, DLL_PATH, 5000, False
    AssertNoError "Open backup DB"
    AssertEqual "Backup row count", TableRowCount(dest, "t_bak"), 500
    Dim v As Variant
    v = QueryScalar(dest, "SELECT val FROM t_bak WHERE id=250;")
    AssertEqual "Backup row 250 correct", CStr(v), "row_250"
    dest.CloseConnection

    ' Test incremental backup with progress tracking
    Dim dest2Path As String
    dest2Path = Left(DB_PATH, Len(DB_PATH) - 3) & "_bak2.db"
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

    ' Verify second backup
    Dim dest2 As New SQLite3Connection
    dest2.OpenDatabase dest2Path, DLL_PATH, 5000, False
    AssertEqual "Backup2 row count", TableRowCount(dest2, "t_bak"), 500
    dest2.CloseConnection

    src.CloseConnection
    Kill srcPath:   Err.Clear
    Kill destPath:  Err.Clear
    Kill dest2Path: Err.Clear
    EndSuite
End Sub

'==============================================================================
' 31. Incremental BLOB Stream
'==============================================================================
Public Sub RunTest_BlobStream()
    StartSuite "BlobStream"
    On Error Resume Next

    Dim conn As SQLite3Connection: Set conn = FreshConn()
    DropTable conn, "t_blobstream"
    conn.ExecSQL "CREATE TABLE t_blobstream (id INTEGER PRIMARY KEY, data BLOB);"

    ' Insert a zeroblob placeholder (1024 bytes)
    Dim blobSize As Long: blobSize = 1024
    conn.ExecSQL "INSERT INTO t_blobstream VALUES (1, zeroblob(" & blobSize & "));"
    Dim rowId As LongLong: rowId = conn.LastInsertRowID()
    AssertTrue "RowId is 1", CLng(rowId) = 1

    ' Open blob for writing
    conn.BeginTransaction
    Dim bs As New SQLite3BlobStream
    bs.OpenBlob conn, "t_blobstream", "data", rowId, True
    AssertNoError "OpenBlob for write"
    AssertTrue "IsOpen", bs.IsOpen
    AssertEqual "Blob size", bs.Size, blobSize

    ' Write a pattern: fill with sequential byte values
    Dim chunk() As Byte
    ReDim chunk(255)
    Dim j As Long
    For j = 0 To 255
        chunk(j) = CByte(j)
    Next j
    ' Write four 256-byte chunks to fill the 1024-byte blob
    Dim off As Long
    For off = 0 To 768 Step 256
        bs.WriteAt chunk, off
    Next off
    AssertNoError "WriteAt all chunks"
    AssertEqual "Position unchanged after WriteAt", bs.Position, 0  ' WriteAt doesn't move position

    bs.CloseBlob
    conn.CommitTransaction
    AssertFalse "Closed after commit", bs.IsOpen

    ' Reopen for reading and verify pattern
    Dim bsR As New SQLite3BlobStream
    bsR.OpenBlob conn, "t_blobstream", "data", rowId, False
    AssertNoError "OpenBlob for read"
    AssertEqual "Read blob size", bsR.Size, blobSize

    ' Read first 256 bytes and verify
    Dim firstChunk() As Byte
    firstChunk = bsR.ReadAt(256, 0)
    AssertNoError "ReadAt 0"
    AssertEqual "First byte = 0", CInt(firstChunk(0)), 0
    AssertEqual "Last byte of first chunk = 255", CInt(firstChunk(255)), 255

    ' Read using sequential ReadBytes (advances position)
    bsR.SeekTo 0
    Dim seqChunk() As Byte: seqChunk = bsR.ReadBytes(128)
    AssertEqual "Position after ReadBytes(128)", bsR.Position, 128
    AssertEqual "Sequential byte 0", CInt(seqChunk(0)), 0
    AssertEqual "Sequential byte 127", CInt(seqChunk(127)), 127

    ' Read last 256 bytes
    Dim lastChunk() As Byte: lastChunk = bsR.ReadAt(256, 768)
    AssertEqual "Last chunk byte 0 = 0", CInt(lastChunk(0)), 0
    AssertEqual "Last chunk byte 255 = 255", CInt(lastChunk(255)), 255

    bsR.CloseBlob

    ' Seek out-of-range should error
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

'==============================================================================
' 32. Serialize / Deserialize
'==============================================================================
Public Sub RunTest_Serialize()
    StartSuite "Serialize"
    On Error Resume Next

    ' Create a source DB with known content
    Dim src As SQLite3Connection: Set src = FreshConn()
    DropTable src, "t_ser"
    src.ExecSQL "CREATE TABLE t_ser (id INTEGER PRIMARY KEY, name TEXT);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 100
        src.ExecSQL "INSERT INTO t_ser VALUES (" & i & ", 'name_" & i & "');"
    Next i
    src.CommitTransaction
    AssertEqual "Source rows", TableRowCount(src, "t_ser"), 100

    ' Checkpoint all WAL frames into the main file, then switch to rollback journal.
    ' sqlite3_serialize captures exactly the bytes SQLite would write to disk; if the
    ' connection is still in WAL mode those bytes include a WAL-format page 1 that
    ' causes text reads to fail when deserialized into a :memory: connection that
    ' cannot support WAL.  Switching to DELETE first gives a clean rollback snapshot.
    src.ExecSQL "PRAGMA wal_checkpoint(TRUNCATE);"
    src.ExecSQL "PRAGMA journal_mode=DELETE;"
    Err.Clear

    ' Serialize to bytes
    Dim snap() As Byte
    Err.Clear
    snap = SerializeDB(src)
    AssertNoError "SerializeDB no error"
    AssertTrue "Snapshot non-empty", UBound(snap) > 0
    ' SQLite DB files start with the magic header "SQLite format 3\000"
    ' S(0) Q(1) L(2) i(3) t(4) e(5) space(6) ...
    AssertEqual "Header byte 0 = S", Chr(snap(0)), "S"
    AssertEqual "Header byte 1 = Q", Chr(snap(1)), "Q"
    AssertEqual "Header byte 2 = L", Chr(snap(2)), "L"
    Log "    INFO  Serialized size = " & (UBound(snap) + 1) & " bytes"

    ' Deserialize into a fresh in-memory connection.
    Dim mem As New SQLite3Connection
    mem.OpenDatabase ":memory:", DLL_PATH, 5000, False
    AssertNoError "Open :memory: for deserialize"
    DeserializeDB mem, snap
    AssertNoError "DeserializeDB no error"
    AssertEqual "Deserialized row count", TableRowCount(mem, "t_ser"), 100
    Dim v As Variant
    v = QueryScalar(mem, "SELECT name FROM t_ser WHERE id=50;")
    AssertEqual "Row 50 correct", CStr(v), "name_50"

    ' Mutations to mem do NOT affect src
    mem.ExecSQL "DELETE FROM t_ser WHERE id <= 10;"
    AssertEqual "mem after delete", TableRowCount(mem, "t_ser"), 90
    AssertEqual "src unaffected", TableRowCount(src, "t_ser"), 100
    mem.CloseConnection

    ' InMemoryClone
    Dim clone As SQLite3Connection
    Set clone = InMemoryClone(src)
    AssertNoError "InMemoryClone no error"
    AssertEqual "Clone row count", TableRowCount(clone, "t_ser"), 100
    v = QueryScalar(clone, "SELECT name FROM t_ser WHERE id=99;")
    AssertEqual "Clone row 99", CStr(v), "name_99"

    ' Verify clone is independent
    clone.ExecSQL "DROP TABLE t_ser;"
    AssertFalse "Clone: t_ser gone", TableExists(clone, "t_ser")
    AssertTrue  "Source: t_ser still exists", TableExists(src, "t_ser")
    clone.CloseConnection

    src.CloseConnection
    EndSuite
End Sub

'==============================================================================
' 33. Diagnostics (db_status / stmt_status)
'==============================================================================
Public Sub RunTest_Diagnostics()
    StartSuite "Diagnostics"
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

    ' GetDbStatus returns a matrix
    Dim info As Variant: info = GetDbStatus(conn, False)
    AssertNoError "GetDbStatus no error"
    AssertTrue "GetDbStatus rows >= 13", UBound(info, 1) + 1 >= 13
    AssertEqual "GetDbStatus cols", UBound(info, 2) + 1, 3
    AssertEqual "First counter name", CStr(info(0, 0)), "lookaside_used"
    AssertTrue "cache_used >= 0", CLng(info(1, 1)) >= 0

    ' GetDbStatusValue returns (current, highwater) array
    Dim cv As Variant: cv = GetDbStatusValue(conn, DBSTAT_CACHE_USED, False)
    AssertNoError "GetDbStatusValue no error"
    AssertTrue "cache_used current >= 0", CLng(cv(0)) >= 0
    AssertTrue "cache_used highwater >= 0", CLng(cv(1)) >= 0
    Log "    INFO  cache_used current=" & cv(0) & " highwater=" & cv(1)

    ' Run a full-scan query and check stmt_status
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "SELECT SUM(val) FROM t_diag WHERE val > 500;"
    cmd.ExecuteScalar
    AssertNoError "Diagnostic query no error"

    ' Full-scan step count should be > 0 (no index on val)
    Dim fullScan As Long
    fullScan = GetStmtStatus(cmd.StmtHandle, STMTSTAT_FULLSCAN, False)
    AssertTrue "Full-scan steps > 0", fullScan > 0
    Log "    INFO  fullscan_step=" & fullScan

    ' VM step count should also be > 0
    Dim vmSteps As Long
    vmSteps = GetStmtStatus(cmd.StmtHandle, STMTSTAT_VM_STEP, False)
    AssertTrue "VM steps > 0", vmSteps > 0

    ' GetAllStmtStatus matrix
    Dim stmtInfo As Variant: stmtInfo = GetAllStmtStatus(cmd.StmtHandle, False)
    AssertNoError "GetAllStmtStatus no error"
    AssertTrue "GetAllStmtStatus rows >= 9", UBound(stmtInfo, 1) + 1 >= 9
    AssertEqual "GetAllStmtStatus cols", UBound(stmtInfo, 2) + 1, 2
    AssertEqual "First stmt stat name", CStr(stmtInfo(0, 0)), "fullscan_step"

    ' ResetDbStatus: highwater should be zeroed after reset
    ResetDbStatus conn, DBSTAT_CACHE_HIT
    Dim cv2 As Variant: cv2 = GetDbStatusValue(conn, DBSTAT_CACHE_HIT, False)
    AssertEqual "Highwater zeroed after reset", CLng(cv2(1)), 0

    ' DbStatusSummary should not raise
    Err.Clear
    DbStatusSummary conn
    AssertNoError "DbStatusSummary no error"

    DropTable conn, "t_diag"
    conn.CloseConnection
    EndSuite
End Sub

