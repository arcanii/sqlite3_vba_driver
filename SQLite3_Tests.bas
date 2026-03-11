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
' Version : 0.1.2
'
' Version History:
'   0.1.0 - Initial release. 122 tests across 22 suites.
'   0.1.1 - Added QueryPerformanceCounter/Frequency high-resolution timing.
'            Added EndSuite() per-suite elapsed time reporting.
'            Fixed inline comment after line-continuation in Array() calls.
'   0.1.2 - Added RunTest_BLOB (23), RunTest_Aggregates (24),
'            RunTest_FTS5 (25). Total: 171 tests across 25 suites.
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
Private Const DB_PATH As String = "C:\sqlite\driver_test.db"

' ---------------------------------------------------------------------------
' Test harness state
' ---------------------------------------------------------------------------
Private m_pass       As Long
Private m_fail       As Long
Private m_suite      As String
Private m_suiteStart As LongPtr   ' QPC ticks at suite start
Private m_runStart   As LongPtr   ' QPC ticks at RunAllTests start
Private m_freq       As LongPtr   ' QPC frequency (ticks per second)

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
    Debug.Print ""
    Debug.Print "  [" & name & "]"
End Sub

Private Sub EndSuite()
    Dim elapsed As String: elapsed = ElapsedMs(m_suiteStart, QPC())
    Debug.Print "    TIME  " & elapsed
End Sub

Private Sub Pass(ByVal name As String)
    m_pass = m_pass + 1
    Debug.Print "    PASS  " & name
End Sub

Private Sub Fail(ByVal name As String, ByVal detail As String)
    m_fail = m_fail + 1
    Debug.Print "    FAIL  " & name & " -- " & detail
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
    m_pass = 0
    m_fail = 0
    EnsureFreq
    m_runStart = QPC()

    Debug.Print String(64, "=")
    Debug.Print "SQLite3 Driver Test Suite"
    Debug.Print String(64, "=")

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

    Dim totalTime As String: totalTime = ElapsedMs(m_runStart, QPC())
    Debug.Print ""
    Debug.Print String(64, "=")
    Debug.Print "Results: " & m_pass & " passed,  " & m_fail & " failed  " & _
                "(" & (m_pass + m_fail) & " total)  " & totalTime
    Debug.Print String(64, "=")

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
    Debug.Print "    INFO  SQLite version = " & ver

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
        Debug.Print "    SKIP  FTS5 not available in this sqlite3.dll build"
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
