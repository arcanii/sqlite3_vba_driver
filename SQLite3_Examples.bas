Attribute VB_Name = "SQLite3_Examples"
'==============================================================================
' SQLite3_Examples.bas  -  Usage examples and integration tests (64-bit only)
' Run TestAll() to validate the complete setup.
'
' Version : 0.1.4
'
' Version History:
'   0.1.0 - Initial release. BasicCRUD, VectorizedQuery, BulkInsert,
'            ConnectionPool, NamedParams, QuantDBTemplate examples.
'   0.1.1 - Fixed Integer overflow in large literal multiplications (256&).
'            Added Diagnose() step-by-step debug helper.
'            Updated DLL_PATH to support System32 placement.
'   0.1.2 - No functional changes. Version stamp updated.
'   0.1.3 - No functional changes. Version stamp updated.
'   0.1.4 - Added Example_Backup, Example_BlobStream, Example_Serialize,
'            Example_Diagnostics. Added all four to TestAll().
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

' Option A: place sqlite3.dll in C:\Windows\System32 (recommended)
'   - No Defender scanning overhead, found by name alone
Private Const DLL_PATH As String = "sqlite3.dll"

' Option B: explicit path outside System32
' Private Const DLL_PATH As String = "C:\sqlite\sqlite3.dll"
Private Const DB_PATH  As String = "C:\sqlite\test_quant.db"

'==============================================================================
' Example 1: Basic CRUD with ADO-style recordset
'==============================================================================
Public Sub Example_BasicCRUD()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "CREATE TABLE IF NOT EXISTS instruments (" & _
                 "  id     INTEGER PRIMARY KEY AUTOINCREMENT," & _
                 "  symbol TEXT    NOT NULL UNIQUE," & _
                 "  name   TEXT," & _
                 "  sector TEXT," & _
                 "  weight REAL    DEFAULT 0.0" & _
                 ");"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT OR REPLACE INTO instruments " & _
                      "(symbol, name, sector, weight) VALUES (?,?,?,?)"

    Dim tickers As Variant
    tickers = Array( _
        Array("AAPL", "Apple Inc",      "Technology", 0.065), _
        Array("MSFT", "Microsoft Corp", "Technology", 0.058), _
        Array("NVDA", "NVIDIA Corp",    "Technology", 0.042), _
        Array("JPM",  "JPMorgan Chase", "Financials", 0.031), _
        Array("GS",   "Goldman Sachs",  "Financials", 0.018))

    conn.BeginTransaction
    Dim i As Long
    For i = 0 To UBound(tickers)
        Dim row As Variant: row = tickers(i)
        cmd.BindText   1, row(0)
        cmd.BindText   2, row(1)
        cmd.BindText   3, row(2)
        cmd.BindDouble 4, row(3)
        cmd.Execute
        cmd.Reset
    Next i
    conn.CommitTransaction

    Debug.Print "Inserted. Last rowid: " & conn.LastInsertRowID()

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset( _
        "SELECT symbol, name, weight FROM instruments ORDER BY weight DESC;")

    Debug.Print String(50, "-")
    Debug.Print "Symbol  | Name                  | Weight"
    Debug.Print String(50, "-")
    Do While Not rs.EOF
        Debug.Print rs!symbol & " | " & rs!name & " | " & Format(rs!weight, "0.000")
        rs.MoveNext
    Loop
    rs.CloseRecordset
    conn.CloseConnection
    Debug.Print "Example_BasicCRUD complete."
End Sub

'==============================================================================
' Example 2: Vectorized bulk load (50x faster for large result sets)
'==============================================================================
Public Sub Example_VectorizedQuery()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 256& * 1024 * 1024

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT * FROM instruments;")
    Dim rowCount As Long: rowCount = rs.LoadAll()

    Debug.Print "Loaded " & rowCount & " rows vectorized."

    rs.MoveFirst
    Do While Not rs.EOF
        Debug.Print rs!symbol & " weight=" & rs!weight
        rs.MoveNext
    Loop

    If rowCount > 0 Then
        RecordsetToRange rs, ActiveSheet.Range("A1"), True
        Debug.Print "Data written to sheet."
    End If

    rs.CloseRecordset
    conn.CloseConnection
End Sub

'==============================================================================
' Example 3: Bulk insert - 100k rows/sec
'==============================================================================
Public Sub Example_BulkInsert()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "CREATE TABLE IF NOT EXISTS tick_data (" & _
                 "  id     INTEGER PRIMARY KEY AUTOINCREMENT," & _
                 "  symbol TEXT, ts REAL, price REAL, volume INTEGER);"

    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, "tick_data", Array("symbol", "ts", "price", "volume"), 5000

    Dim t0 As Double:    t0   = Timer()
    Dim syms As Variant: syms = Array("AAPL", "MSFT", "NVDA", "JPM", "GS")
    Dim i As Long
    For i = 1 To 50000
        bulk.AppendRow Array(syms(i Mod 5), CDbl(Now()) + (i / 86400), _
                             100 + (Rnd() * 50), CLng(Rnd() * 10000))
    Next i
    bulk.CloseInsert

    Dim elapsed As Double: elapsed = Timer() - t0
    Dim rate As Long:      rate    = CLng(50000 / IIf(elapsed = 0, 0.001, elapsed))
    Debug.Print "Inserted 50,000 rows in " & Format(elapsed, "0.00") & _
                "s  (" & rate & " rows/sec)"
    conn.CloseConnection
End Sub

'==============================================================================
' Example 4: AppendMatrix - feed a pre-built 2-D array
'==============================================================================
Public Sub Example_MatrixInsert()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "CREATE TABLE IF NOT EXISTS signals " & _
                 "(date_str TEXT, symbol TEXT, signal REAL, weight REAL);"

    Const N As Long = 1000
    Dim mat() As Variant
    ReDim mat(N - 1, 3)
    Dim r As Long
    For r = 0 To N - 1
        mat(r, 0) = Format(Date - r, "YYYY-MM-DD")
        mat(r, 1) = "AAPL"
        mat(r, 2) = (Rnd() - 0.5) * 2
        mat(r, 3) = Rnd()
    Next r

    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, "signals", Array("date_str", "symbol", "signal", "weight")
    bulk.AppendMatrix mat
    bulk.CloseInsert

    Debug.Print "Matrix insert: " & bulk.TotalRowsInserted & " rows"
    conn.CloseConnection
End Sub

'==============================================================================
' Example 5: Connection pool
'==============================================================================
Public Sub Example_ConnectionPool()
    Dim pool As New SQLite3Pool
    pool.Initialize DB_PATH, DLL_PATH, 4, 5000, True, 64& * 1024 * 1024

    Debug.Print "Pool size: " & pool.PoolSize & " | Active: " & pool.ActiveConnections

    Dim conn As SQLite3Connection
    Set conn = pool.Acquire()

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("SELECT COUNT(*) AS cnt FROM instruments;")
    Debug.Print "Instruments count = " & rs!cnt
    rs.CloseRecordset

    pool.ReleaseConnection conn
    Debug.Print "After release - Active: " & pool.ActiveConnections
    pool.ShutDown
End Sub

'==============================================================================
' Example 6: Named parameters
'==============================================================================
Public Sub Example_NamedParams()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "CREATE TABLE IF NOT EXISTS orders (" & _
                 "  id INTEGER PRIMARY KEY AUTOINCREMENT," & _
                 "  symbol TEXT, side TEXT, qty INTEGER, price REAL);"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO orders (symbol, side, qty, price) " & _
                      "VALUES (:symbol, :side, :qty, :price)"

    conn.BeginTransaction
    cmd.BindTextByName   ":symbol", "NVDA"
    cmd.BindTextByName   ":side",   "BUY"
    cmd.BindIntByName    ":qty",    500
    cmd.BindDoubleByName ":price",  875.25
    cmd.Execute
    cmd.Reset

    cmd.BindTextByName   ":symbol", "AAPL"
    cmd.BindTextByName   ":side",   "SELL"
    cmd.BindIntByName    ":qty",    200
    cmd.BindDoubleByName ":price",  182.1
    cmd.Execute
    cmd.Reset
    conn.CommitTransaction

    Debug.Print "Last rowid: " & conn.LastInsertRowID()
    conn.CloseConnection
End Sub

'==============================================================================
' Example 7: WAL + mmap quant database template
'==============================================================================
Public Sub Example_QuantDBTemplate()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 512& * 1024 * 1024

    Debug.Print "SQLite version: " & SQLite3_API.SQLite_Version()

    conn.ExecSQL "CREATE TABLE IF NOT EXISTS price_history (" & _
                 "  symbol   TEXT NOT NULL," & _
                 "  date_val TEXT NOT NULL," & _
                 "  open_px  REAL, high_px REAL, low_px REAL," & _
                 "  close_px REAL NOT NULL," & _
                 "  volume   INTEGER, adj_close REAL," & _
                 "  PRIMARY KEY (symbol, date_val)" & _
                 ") WITHOUT ROWID;"

    conn.ExecSQL "CREATE INDEX IF NOT EXISTS idx_ph_date " & _
                 "ON price_history(date_val);"

    Debug.Print "Table exists: " & CStr(TableExists(conn, "price_history"))
    Debug.Print "Row count:    " & TableRowCount(conn, "price_history")

    conn.CloseConnection
    Debug.Print "Quant DB template created."
End Sub

'==============================================================================
' Test runner
'==============================================================================
Public Sub TestAll()
    On Error GoTo ErrHandler
    Debug.Print String(60, "=")
    Debug.Print "SQLite3 HFT VBA Class Suite - Integration Test"
    Debug.Print String(60, "=")
    Example_BasicCRUD
    Example_VectorizedQuery
    Example_BulkInsert
    Example_MatrixInsert
    Example_ConnectionPool
    Example_NamedParams
    Example_QuantDBTemplate
    Example_Backup
    Example_BlobStream
    Example_Serialize
    Example_Diagnostics
    Debug.Print String(60, "=")
    Debug.Print "All tests passed."
    Exit Sub
ErrHandler:
    Debug.Print "ERROR: " & Err.Description
    Debug.Print "Source: " & Err.Source
End Sub

'==============================================================================
' Example 8: Online Backup API
' Hot-backup a live database to a file without interrupting readers/writers.
'==============================================================================
Public Sub Example_Backup()
    ' Source: a fresh non-WAL file so the backup destination is a clean copy.
    Dim srcPath  As String: srcPath  = Left(DB_PATH, Len(DB_PATH) - 3) & "_bak_src.db"
    Dim destPath As String: destPath = Left(DB_PATH, Len(DB_PATH) - 3) & "_bak_dst.db"
    On Error Resume Next: Kill srcPath: Kill destPath: Err.Clear: On Error GoTo 0

    Dim src As New SQLite3Connection
    src.OpenDatabase srcPath, DLL_PATH, 5000, False   ' no WAL for clean backup
    src.ExecSQL "CREATE TABLE prices (sym TEXT, price REAL, ts TEXT);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 1000
        src.ExecSQL "INSERT INTO prices VALUES ('AAPL', " & (150 + i * 0.01) & ", '" & Now() & "');"
    Next i
    src.CommitTransaction
    Debug.Print "  Source rows: " & TableRowCount(src, "prices")

    ' One-shot backup (blocks until complete; fine for end-of-day snapshots)
    Dim bak As New SQLite3Backup
    bak.BackupToFile src, destPath
    Debug.Print "  BackupToFile: " & bak.TotalPages & " pages, complete=" & bak.IsComplete

    ' Verify
    Dim dest As New SQLite3Connection
    dest.OpenDatabase destPath, DLL_PATH, 5000, False
    Debug.Print "  Backup rows : " & TableRowCount(dest, "prices")
    dest.CloseConnection

    ' Incremental backup with progress (use for large DBs or progress bars)
    Dim dest2Path As String: dest2Path = Left(DB_PATH, Len(DB_PATH) - 3) & "_bak_inc.db"
    On Error Resume Next: Kill dest2Path: Err.Clear: On Error GoTo 0

    Dim bak2 As New SQLite3Backup
    bak2.OpenBackup src, dest2Path
    Do
        bak2.Step 10   ' 10 pages per step -- yield between steps for large DBs
        Debug.Print "  Progress: " & Format(bak2.Progress * 100, "0.0") & "%"
    Loop Until bak2.IsComplete
    bak2.CloseBackup
    Debug.Print "  Incremental backup complete."

    src.CloseConnection
    On Error Resume Next: Kill srcPath: Kill destPath: Kill dest2Path: Err.Clear
    On Error GoTo 0
End Sub

'==============================================================================
' Example 9: Incremental BLOB I/O (SQLite3BlobStream)
' Read and write arbitrary byte ranges of a BLOB without loading it fully.
'==============================================================================
Public Sub Example_BlobStream()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "DROP TABLE IF EXISTS blobs;"
    conn.ExecSQL "CREATE TABLE blobs (id INTEGER PRIMARY KEY, data BLOB);"

    ' Insert a zeroblob placeholder (64 KB)
    Dim blobSize As Long: blobSize = 65536
    conn.ExecSQL "INSERT INTO blobs VALUES (1, zeroblob(" & blobSize & "));"
    Dim rowId As LongLong: rowId = conn.LastInsertRowID()

    ' Open the BLOB for writing
    conn.BeginTransaction
    Dim bs As New SQLite3BlobStream
    bs.OpenBlob conn, "main", "blobs", "data", rowId, True  ' True = write mode

    ' Write a pattern into three separate regions
    Dim hdr(7) As Byte: Dim j As Long
    For j = 0 To 7: hdr(j) = j: Next j          ' bytes 0-7: 0,1,2,...,7
    bs.WriteAt hdr, 0

    Dim middle(9) As Byte
    For j = 0 To 9: middle(j) = 255 - j: Next j  ' bytes 1000-1009
    bs.WriteAt middle, 1000
    bs.CloseBlob
    conn.CommitTransaction

    ' Re-open for reading and verify
    bs.OpenBlob conn, "main", "blobs", "data", rowId, False
    Debug.Print "  BLOB size: " & bs.Size & " bytes"

    Dim rHdr() As Byte: rHdr = bs.ReadAt(8, 0)
    Debug.Print "  Header[0]: " & rHdr(0) & "  Header[7]: " & rHdr(7)

    Dim rMid() As Byte: rMid = bs.ReadAt(10, 1000)
    Debug.Print "  Middle[0]: " & rMid(0) & "  Middle[9]: " & rMid(9)
    bs.CloseBlob

    conn.ExecSQL "DROP TABLE IF EXISTS blobs;"
    conn.CloseConnection
End Sub

'==============================================================================
' Example 10: Serialize / Deserialize and InMemoryClone
' Snapshot a live DB to a Byte array; restore to :memory: for fast in-process work.
'==============================================================================
Public Sub Example_Serialize()
    ' Build a source DB (non-WAL so the snapshot has a clean page 1)
    Dim srcPath As String: srcPath = Left(DB_PATH, Len(DB_PATH) - 3) & "_ser_src.db"
    On Error Resume Next: Kill srcPath: Err.Clear: On Error GoTo 0

    Dim src As New SQLite3Connection
    src.OpenDatabase srcPath, DLL_PATH, 5000, False
    src.ExecSQL "CREATE TABLE ticks (sym TEXT, price REAL, ts INTEGER);"
    src.BeginTransaction
    Dim i As Long
    For i = 1 To 500
        src.ExecSQL "INSERT INTO ticks VALUES ('MSFT', " & (300 + i * 0.05) & ", " & i & ");"
    Next i
    src.CommitTransaction
    Debug.Print "  Source rows: " & TableRowCount(src, "ticks")

    ' Snapshot to a Byte array (in-process; no file I/O)
    Dim snap() As Byte
    snap = SerializeDB(src)
    Debug.Print "  Snapshot:    " & (UBound(snap) + 1) & " bytes"

    ' Restore into :memory: for isolated analysis
    Dim mem As New SQLite3Connection
    mem.OpenDatabase ":memory:", DLL_PATH, 5000, False
    DeserializeDB mem, snap
    Debug.Print "  Deserialized rows: " & TableRowCount(mem, "ticks")
    mem.ExecSQL "DELETE FROM ticks WHERE ts > 250;"
    Debug.Print "  After delete in mem: " & TableRowCount(mem, "ticks") & "  (src unaffected)"
    Debug.Print "  src still has: " & TableRowCount(src, "ticks")
    mem.CloseConnection

    ' InMemoryClone: independent :memory: copy via the backup API
    Dim clone As SQLite3Connection
    Set clone = InMemoryClone(src)
    Debug.Print "  Clone rows:  " & TableRowCount(clone, "ticks")
    clone.ExecSQL "DROP TABLE ticks;"
    Debug.Print "  After DROP in clone -- src still has: " & TableRowCount(src, "ticks")
    clone.CloseConnection

    src.CloseConnection
    On Error Resume Next: Kill srcPath: Err.Clear: On Error GoTo 0
End Sub

'==============================================================================
' Example 11: Diagnostics (db_status / stmt_status)
' Inspect memory usage and statement-level performance counters at runtime.
'==============================================================================
Public Sub Example_Diagnostics()
    Dim conn As New SQLite3Connection
    conn.OpenDatabase DB_PATH, DLL_PATH

    conn.ExecSQL "DROP TABLE IF EXISTS diag_tbl;"
    conn.ExecSQL "CREATE TABLE diag_tbl (id INTEGER PRIMARY KEY, val TEXT);"
    conn.BeginTransaction
    Dim i As Long
    For i = 1 To 1000
        conn.ExecSQL "INSERT INTO diag_tbl VALUES (" & i & ", 'v" & i & "');"
    Next i
    conn.CommitTransaction

    ' Run a query to populate statement stats
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "SELECT SUM(id) FROM diag_tbl WHERE val LIKE 'v%';"
    Dim total As Variant: total = cmd.ExecuteScalar()
    Debug.Print "  Query result: " & total

    ' Per-statement counters
    Dim stmtInfo As Variant
    stmtInfo = GetAllStmtStatus(cmd.StmtHandle)
    Debug.Print "  VM steps:     " & stmtInfo(3, 1)   ' STMTSTAT_VM_STEP row
    Debug.Print "  Full scans:   " & stmtInfo(0, 1)   ' STMTSTAT_FULLSCAN row

    ' Per-connection memory counters (DbStatusSummary prints directly)
    DbStatusSummary conn

    ' Individual counters
    Debug.Print "  Page cache bytes in use: " & GetDbStatusValue(conn, DBSTAT_CACHE_USED)
    Debug.Print "  Schema memory:           " & GetDbStatusValue(conn, DBSTAT_SCHEMA_USED)

    conn.ExecSQL "DROP TABLE IF EXISTS diag_tbl;"
    conn.CloseConnection
End Sub
'==============================================================================
' Diagnose  -  run this first to find exactly where OpenDatabase fails.
' Check the Immediate window (Ctrl+G) for output.
'==============================================================================
Public Sub Diagnose()
    Dim ok As Boolean
    Debug.Print String(60, "=")
    Debug.Print "SQLite3 Diagnostic"
    Debug.Print String(60, "=")

    ' ---- Step 1: DLL file accessible? --------------------------------------
    Debug.Print ""
    Debug.Print "Step 1: DLL file exists?"
    If Len(DLL_PATH) = 0 Then
        Debug.Print "  FAIL - DLL_PATH is empty. Set DLL_PATH at the top of SQLite3_Examples.bas"
        Exit Sub
    End If
    If Dir(DLL_PATH) = "" Then
        Debug.Print "  FAIL - File not found: " & DLL_PATH
        Debug.Print "  Fix : place sqlite3.dll at that path, or update DLL_PATH."
        Exit Sub
    End If
    Debug.Print "  OK   - Found: " & DLL_PATH

    ' ---- Step 2: DB directory writable? ------------------------------------
    Debug.Print ""
    Debug.Print "Step 2: DB parent folder exists?"
    Dim dbDir As String
    dbDir = Left(DB_PATH, InStrRev(DB_PATH, "\"))
    If Len(dbDir) > 0 And Dir(dbDir, vbDirectory) = "" Then
        Debug.Print "  FAIL - Directory not found: " & dbDir
        Debug.Print "  Fix : create the folder, or update DB_PATH."
        Exit Sub
    End If
    Debug.Print "  OK   - " & IIf(Len(dbDir) > 0, dbDir, "(current dir)")

    ' ---- Step 3: Load the DLL ----------------------------------------------
    Debug.Print ""
    Debug.Print "Step 3: LoadLibrary?"
    On Error Resume Next
    SQLite3_API.SQLite_Unload          ' start clean
    SQLite3_API.SQLite_Load DLL_PATH
    If Err.Number <> 0 Then
        Debug.Print "  FAIL - " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "  OK   - DLL loaded. Version = " & SQLite3_API.SQLite_Version()

    ' ---- Step 4: Open the database -----------------------------------------
    Debug.Print ""
    Debug.Print "Step 4: sqlite3_open_v2?"
    Dim conn As New SQLite3Connection
    On Error Resume Next
    conn.OpenDatabase DB_PATH, DLL_PATH, 5000, False, 0   ' WAL off for this test
    If Err.Number <> 0 Then
        Debug.Print "  FAIL - " & Err.Description
        Debug.Print "  Source: " & Err.Source
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "  OK   - DB open: " & DB_PATH

    ' ---- Step 5: Simple round-trip query -----------------------------------
    Debug.Print ""
    Debug.Print "Step 5: Basic query?"
    On Error Resume Next
    Dim v As Variant
    v = QueryScalar(conn, "SELECT sqlite_version();")
    If Err.Number <> 0 Then
        Debug.Print "  FAIL - " & Err.Description
        On Error GoTo 0
        conn.CloseConnection
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "  OK   - sqlite_version() = " & v

    ' ---- Step 6: WAL pragma ------------------------------------------------
    Debug.Print ""
    Debug.Print "Step 6: PRAGMA journal_mode=WAL?"
    On Error Resume Next
    conn.ExecSQL "PRAGMA journal_mode=WAL;"
    If Err.Number <> 0 Then
        Debug.Print "  FAIL - " & Err.Description
        On Error GoTo 0
        conn.CloseConnection
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print "  OK"

    conn.CloseConnection
    Debug.Print ""
    Debug.Print "All diagnostic steps passed."
    Debug.Print String(60, "=")
End Sub
