SQLite3 VBA Driver
==================
Skip down to TLDR section to get it working a.s.a.p. Please look at the Security section to make sure you don't hit a slowdown caused by Windows/Excel security settings.

**What is this?**<br/>
This is a VBA SQLite3 "driver" so no registered COM objects, no third-party dependencies beyond the SQLite DLL itself. It is also very flexible to support a user defined 'sqlite3.dll' location, but this can cause "security issues" (refer to security issues section below).

TLDR Section (for those who have no attention span)
=======
1. Get the sqlite3.dll from [sqlite.org/download](https://sqlite.org/download.html).
2. You need to have Microsoft Excel (64-bit) installed.
3. Make a directory (e.g. "C:\sqlite\") : this can go anywhere you want, but if you change this you also need to change the excel file in step 7 (WARNING: see security issues)
4. Put the sqlite3.dll in the directory you made in step 3 ("C:\sqlite\" if you want to avoid future work in step 7).
5. Download the Excel file "Test-SQLite3-VBA-Driver.xlsm" from this GitHub repo, open it and turn on macros.
6. Do a (`Alt+F11`) to open the Visual Basic editor.
7. ONLY if you changed "C:\sqlite\" to something else ... look for, and make changes as below, if you made no changes in step 3, skip right to 8.
   - find the "SQLite3_Tests.bas" in the 'Project - VBAProject' explorer window, at the top of the file look for the file locations.
   - Change this to your DLL location: `Private Const DLL_PATH  As String = "C:\sqlite\sqlite3.dll"`. Note: if you are using the System32 option from security, it is just `"sqlite3.dll"`
   - Change this to where you want the DB location: `Private Const DB_PATH   As String = "C:\sqlite\driver_test.db"`.
8. In the 'Immediate window', type `RunAllTests` and hit enter. Reminder: to show the Immediate window do a (`Ctrl+G`).
9. If everything is ok, the tests should run and produce a report (before the universe ends).

Versions
=======
| Version | Date           | Tests | Pass | Highlights |
|---------|----------------|-------|------|------------|
| 0.1.3 | 11 March, 2026 |  240  |  240 | Schema introspection, savepoints, JSON functions, interrupt, failed-test summary |
| 0.1.2 | 11 March, 2026 |  171  |  171 | BLOB support, aggregate helpers, FTS5 full-text search, GPLv3 license |
| 0.1.1 | 11 March, 2026 |  122  |  122 | DispCallFunc ABI fixes, UTF-8 fixes, QPC benchmarking, security notes |
| 0.1.0 | 09 March, 2026 |  n/a  |  n/a | Initial release. Core driver, recordset, bulk insert, pool. My inaugural github check-in! |

Releases Page
=======
https://github.com/arcanii/sqlite3_vba_driver/releases

Security Issues (important)
=======
1. Microsoft Excel has default security settings that slow down this driver 100x (it feels like the universe may end before the 240 tests can run).
2. In Excel File->Options->Trust Center, click "Trust Center Settings..." button.
3. Macro Settings tab: `Enable VBA Macros.`
4. You may also need add the file location in the "Trusted Location" tab. Simply click the "Add new location..." button and add the directory location you want to trust to run macros.
   
There are 2 alternatives to doing this : (these may require administrative privileges)
1. Option A: place `sqlite3.dll` in `C:\Windows\System32` (recommended, but it may not be allowed in your environment).
No Defender scanning overhead, found by name alone.
2. Option B: explicit path outside System32 as the original, but you need to make sure your AV program has it excluded. Windows Defender example provided below.

### (for example) Windows Defender folder exclusion
1. Open **Windows Security** (search "Windows Security" in the Start menu)
2. Go to **Virus & threat protection -> Manage settings**
3. Scroll to **Exclusions -> Add or remove exclusions**
4. Add a **Folder** exclusion for the folder containing `sqlite3.dll`
   (e.g. `C:\sqlite\`)

The DLL path in your VBA constants stays unchanged. Defender will skip
real-time scanning for that folder entirely.

---

## What's New in 0.1.3

- **`SQLite3_Schema.bas`** — new module: schema introspection via `PRAGMA`. `GetTableList`, `GetViewList`, `GetColumnInfo`, `GetIndexList`, `GetIndexColumns`, `GetForeignKeys`, `GetCreateSQL`, `GetDatabaseInfo`, `GetTriggerList`
- **`SQLite3_JSON.bas`** — new module: 15 wrappers for SQLite's built-in JSON functions (requires SQLite 3.38+). `JSONExtract`, `JSONSet`, `JSONInsert`, `JSONReplace`, `JSONRemove`, `JSONPatch`, `JSONValid`, `JSONSearch`, `JSONGroupArray`, `JSONGroupObject`, `JSONEach`, `JSONType`, `JSONBuildObject`, `JSONBuildArray`
- **Savepoints** — nested transactions added to `SQLite3Connection`: `Savepoint`, `ReleaseSavepoint`, `RollbackToSavepoint`, `SavepointDepth`
- **Interrupt** — `conn.Interrupt` calls `sqlite3_interrupt` to cancel a running query
- **Failed-test summary** — `RunAllTests` now prints a consolidated list of every failure at the end so nothing gets lost in the output
- **`ViewExists`, `IndexExists`** — added to `SQLite3_Helpers.bas` alongside the existing `TableExists`
- 4 new test suites (26-29), 69 new tests — total 240/240

---

## Features

| Feature | Detail |
|---------|--------|
| Dynamic DLL loading | `LoadLibrary` at runtime — no hard `Declare` required |
| UTF-8 marshalling | Full round-trip via `WideCharToMultiByte` / `MultiByteToWideChar` |
| Prepared statements | Positional (`?`) and named (`:param`) binding |
| BLOB support | `BindBlob`, `AsBytes()`, vectorized BLOB load, `BlobToBytes` |
| Statement cache | 64-slot LRU per connection — cache hit = reset only, no re-prepare |
| ADO-style recordset | `BOF`/`EOF`/`MoveNext`/`MoveLast`, `rs!FieldName` syntax |
| Vectorized load | `LoadAll()` pulls entire result into a Variant matrix (~50x faster than live iteration) |
| `ToMatrix()` | Returns `(row, col)` Variant array ready for direct Excel range assignment |
| Bulk insert | Single reused prepared `INSERT`, transaction-batched commits (~100k rows/sec) |
| Connection pool | LRU idle reaping, auto-rollback on release, configurable max size |
| Savepoints | Nested transactions: `Savepoint` / `ReleaseSavepoint` / `RollbackToSavepoint` |
| Interrupt | `conn.Interrupt` cancels a running query via `sqlite3_interrupt` |
| Schema introspection | Tables, views, columns, indexes, FKs, triggers, CREATE SQL, PRAGMA diagnostics |
| Aggregate helpers | `GroupBy*`, `ScalarAgg`, `MultiAgg`, `RunningTotal`, `Histogram`, window functions |
| JSON functions | `JSONExtract`, `JSONSet`, `JSONPatch`, `JSONSearch`, `JSONGroupArray` and more |
| FTS5 full-text search | Create, insert, search, snippet, highlight, BM25 ranking, optimize |
| WAL mode | Enabled by default on `OpenDatabase` |
| QPC benchmarking | `QueryPerformanceCounter` timing in every test suite |
| Failed-test summary | All failures reprinted at the end of `RunAllTests` — none get lost |
| 64-bit only | All handles are `LongPtr` / `LongLong` — requires 64-bit Excel |

---

## File Reference

| File | Role |
|------|------|
| `SQLite3_API.bas` | DLL loader, 31 cached proc addresses, all SQLite wrappers via `DispCallFunc` |
| `SQLite3_API_Ext.bas` | Auxiliary dispatch bridge; secondary DLL handle copy |
| `SQLite3_Helpers.bas` | `QueryScalar`, `TableExists`, `ViewExists`, `IndexExists`, `TableRowCount`, `RecordsetToRange` |
| `SQLite3Connection.cls` | Open/close, WAL, mmap, 64-slot LRU statement cache, transactions, savepoints, interrupt |
| `SQLite3Recordset.cls` | Live and vectorized recordset, `GetRows()`, `ToMatrix()`, `rs!Field` |
| `SQLite3Fields.cls` | Case-insensitive field collection, `For Each` enumerator |
| `SQLite3Field.cls` | Zero-copy value reads; `Value`, `AsString`, `AsBytes`, `AsInt64` |
| `SQLite3Command.cls` | Positional and named binding, `BindBlob`, `BindVariant`, `ExecuteScalar` |
| `SQLite3BulkInsert.cls` | High-speed batch insert, `AppendRow`, `AppendMatrix` |
| `SQLite3Pool.cls` | Connection pool, LRU reaping, auto-rollback on release, pre-warm |
| `SQLite3_Schema.bas` | Schema introspection — tables, columns, indexes, FKs, triggers, PRAGMA info |
| `SQLite3_Aggregates.bas` | SQL aggregate and window function helpers |
| `SQLite3_FTS5.bas` | FTS5 full-text search helpers |
| `SQLite3_JSON.bas` | SQLite built-in JSON function wrappers (requires SQLite 3.38+) |
| `SQLite3_Examples.bas` | Annotated usage examples for every feature |
| `SQLite3_Tests.bas` | 240-test automated suite with QPC timing and failure summary |
| `Test-SQLite3-VBA-Driver.xlsm` | Template Excel workbook with all VBA pre-loaded. You just need this to try it out. |

---

## Requirements
- **64-bit Excel** (Excel 2016 or later, 64-bit install)
- **`sqlite3.dll`** (64-bit) — download from [sqlite.org/download](https://sqlite.org/download.html)
  - Under *Precompiled Binaries for Windows*, grab **`sqlite-dll-win-x64-*.zip`**
  - Extract `sqlite3.dll` to a location (e.g. `C:\sqlite\sqlite3.dll`)
  - FTS5 and JSON functions are enabled in all official precompiled binaries
- **Microsoft Scripting Runtime** reference (for `Dictionary` used in `SQLite3Fields`)

---

## Installation

### 1. Place the DLL

Copy `sqlite3.dll` (64-bit) to a stable path. The recommended location is
alongside the `.xlsm` workbook, or a dedicated folder such as `C:\sqlite\`.
This filepath is set programmatically in VBA, so it can be part of settings (e.g. dynamic).

### 2. Import the VBA modules

Open the Visual Basic Editor (`Alt+F11`), then for each file below choose
**File -> Import File**, or drag and drop them to the project explorer (this is faster):

```
1.  SQLite3_API.bas           <- must be first
2.  SQLite3_API_Ext.bas
3.  SQLite3_Helpers.bas
4.  SQLite3_Schema.bas
5.  SQLite3_Aggregates.bas
6.  SQLite3_FTS5.bas
7.  SQLite3_JSON.bas
8.  SQLite3Field.cls
9.  SQLite3Fields.cls
10. SQLite3Command.cls
11. SQLite3Recordset.cls
12. SQLite3Connection.cls
13. SQLite3BulkInsert.cls
14. SQLite3Pool.cls
15. SQLite3_Examples.bas      <- optional, for learning only
16. SQLite3_Tests.bas         <- optional, for testing only
```

### 3. Add the Scripting Runtime reference

In the VBA Editor: **Tools -> References -> check "Microsoft Scripting Runtime"**

---

## Running the Tests

1. Set `DLL_PATH` and `DB_PATH` at the top of `SQLite3_Tests.bas` to match
   your environment:

```vba
Private Const DLL_PATH As String = "C:\sqlite\sqlite3.dll"
Private Const DB_PATH  As String = "C:\sqlite\driver_test.db"
```

2. Open the Immediate window (`Ctrl+G`), then type:

```
RunAllTests
```

Expected output (abridged):

```
================================================================
SQLite3 Driver Test Suite
================================================================

  [DllLoad]
    PASS  SQLite_Load
    PASS  SQLite_IsLoaded
    INFO  SQLite version = 3.47.0
    TIME  1.23 ms

  [Schema]
    PASS  TableList has t_schema_a
    PASS  TableExists t_schema_a
    PASS  ColumnInfo 3 cols
    PASS  FK refs t_schema_a
    TIME  18.55 ms

  [JSON]
    PASS  JSONExtract name=Alice
    PASS  JSONSet city=Madrid
    PASS  JSONPatch country=Spain
    PASS  JSONGroupArray has Alice
    TIME  24.31 ms

  ...

================================================================
Results: 240 passed,  0 failed  (240 total)  1137.46 ms
================================================================
```

If any tests fail, a consolidated summary is printed at the very end so nothing gets lost:

```
FAILED TESTS (2):
----------------------------------------------------------------
  [Schema]  TableExists t_schema_a -- condition was False
  [JSON]    JSONExtract name=Alice -- expected [Alice] got []
----------------------------------------------------------------
```

The test database is deleted automatically at the end of the suite. Individual
tests can be run in isolation — each is a standalone `Public Sub`:

```
RunTest_Transactions
RunTest_BulkInsert_AppendRow
RunTest_ConnectionPool
RunTest_BLOB
RunTest_Aggregates
RunTest_FTS5
RunTest_Schema
RunTest_Savepoints
RunTest_JSON
RunTest_Interrupt
```

---

## Quick Start

### Open a database and query

```vba
Dim conn As New SQLite3Connection
conn.OpenDatabase "C:\data\mydb.db", "sqlite3.dll"

Dim rs As SQLite3Recordset
Set rs = conn.OpenRecordset("SELECT id, name FROM customers ORDER BY name;")
Do While Not rs.EOF
    Debug.Print rs!id, rs!name
    rs.MoveNext
Loop
rs.CloseRecordset
conn.CloseConnection
```

### Dump a query directly to a worksheet

```vba
Dim conn As New SQLite3Connection
conn.OpenDatabase DB_PATH, DLL_PATH, 5000, True, 256& * 1024 * 1024

Dim rs As SQLite3Recordset
Set rs = conn.OpenRecordset("SELECT * FROM prices;")
Dim n As Long: n = rs.LoadAll()

Sheet1.Range("A1").Resize(n, rs.FieldCount).Value = rs.ToMatrix()
rs.CloseRecordset
conn.CloseConnection
```

### Prepared statement — named parameters

```vba
Dim cmd As New SQLite3Command
cmd.Prepare conn, "INSERT INTO orders VALUES (:id, :product, :qty);"
cmd.BindIntByName  ":id",      42
cmd.BindTextByName ":product", "Widget"
cmd.BindIntByName  ":qty",     100
cmd.Execute
```

### BLOB — store and retrieve binary data

```vba
' Write
Dim img() As Byte: img = ReadFileToBytesArray("C:\logo.png")
Dim cmd As New SQLite3Command
cmd.Prepare conn, "INSERT INTO assets (name, data) VALUES (?, ?);"
cmd.BindText 1, "logo.png"
cmd.BindBlob 2, img
cmd.Execute

' Read - live recordset
Dim rs As SQLite3Recordset
Set rs = conn.OpenRecordset("SELECT data FROM assets WHERE name='logo.png';")
Dim blob() As Byte: blob = rs.Fields("data").AsBytes()

' Read - vectorized (LoadAll stores BLOBs as Byte() in the matrix)
rs.LoadAll
Dim v As Variant: v = rs!data
Dim b() As Byte: b = v
```

### Schema introspection

```vba
' List all tables
Dim tables As Variant: tables = GetTableList(conn)

' Column details: name, type, NOT NULL, default, PK position
Dim cols As Variant: cols = GetColumnInfo(conn, "orders")

' Foreign keys on a table
Dim fks As Variant: fks = GetForeignKeys(conn, "order_lines")

' Indexes on a table
Dim idxs As Variant: idxs = GetIndexList(conn, "orders")

' Retrieve the original CREATE TABLE statement
Dim ddl As String: ddl = GetCreateSQL(conn, "orders")

' Key PRAGMA diagnostics as a (name, value) matrix
Dim info As Variant: info = GetDatabaseInfo(conn)

' Existence checks
If TableExists(conn, "orders")  Then ...
If ViewExists(conn,  "v_open")  Then ...
If IndexExists(conn, "ix_date") Then ...
```

### Savepoints — nested transactions

```vba
conn.BeginTransaction
conn.ExecSQL "INSERT INTO orders VALUES (1, 'outer');"

conn.Savepoint "sp1"
conn.ExecSQL "INSERT INTO orders VALUES (2, 'inner');"

' Something went wrong - undo just the inner work
conn.RollbackToSavepoint "sp1"
conn.ReleaseSavepoint "sp1"

conn.CommitTransaction   ' only the outer INSERT is kept
```

### JSON functions (requires SQLite 3.38+)

```vba
' Table with a JSON TEXT column
conn.ExecSQL "CREATE TABLE users (id INTEGER, profile TEXT);"
conn.ExecSQL "INSERT INTO users VALUES (1, '{""name"":""Alice"",""city"":""London""}');"

' Extract a value by JSONPath
Dim city As Variant
city = JSONExtract(conn, "users", "profile", "$.city", "id=1")
' -> "London"

' Update a key in place
JSONSet conn, "users", "profile", "$.city", "'Paris'", "id=1"

' Apply an RFC 7396 merge-patch
JSONPatch conn, "users", "profile", "'{""country"":""France""}'", "id=1"

' Remove a key
JSONRemove conn, "users", "profile", Array("$.city"), "id=1"

' Search rows by JSON value
Dim results As Variant
results = JSONSearch(conn, "users", "profile", "$.country", "'France'")

' Aggregate all names into a JSON array
Dim arr As String
arr = JSONGroupArray(conn, "users", "json_extract(profile,'$.name')")
' -> '["Alice","Bob"]'

' Validate all rows contain valid JSON
If Not JSONValid(conn, "users", "profile") Then Debug.Print "Bad JSON found"
```

### Aggregate helpers

```vba
' Count rows per group, top 10
Dim mat As Variant
mat = GroupByCount(conn, "sales", "region", "", 10)
Sheet1.Range("A1").Resize(UBound(mat,1)+1, 2).Value = mat

' Multiple aggregates in one query
mat = MultiAgg(conn, "trades", _
               Array("COUNT(*) AS n", "SUM(qty) AS vol", "AVG(price) AS avg_px"))

' Running total window function
mat = RunningTotal(conn, "trades", "trade_date", "pnl")

' Histogram: bucket price into 20 bins
mat = Histogram(conn, "trades", "price", 20)
```

### FTS5 full-text search

```vba
' Create FTS5 table with Porter stemmer
CreateFTS5Table conn, "docs", Array("title", "body"), "", "porter unicode61", True

' Insert documents
FTS5Insert conn, "docs", Array("title", "body"), _
           Array("SQLite Guide", "How to use SQLite for fast data storage")

' Bulk insert from a matrix
FTS5BulkInsert conn, "docs", Array("title", "body"), myDataMatrix

' Search - returns (row, col) matrix sorted by relevance
Dim results As Variant
results = FTS5SearchMatrix(conn, "docs", "sqlite storage")

' Search with highlighted snippets
results = FTS5Snippet(conn, "docs", "sqlite", 0, "<b>", "</b>", "...", 16, 20)

' Prefix search, column-scoped search, boolean operators
results = FTS5SearchMatrix(conn, "docs", "fast*")
results = FTS5SearchMatrix(conn, "docs", "title : SQLite")
results = FTS5SearchMatrix(conn, "docs", "sqlite AND performance")

' BM25 explicit scoring
results = FTS5BM25Search(conn, "docs", "fast data storage", "*", 10)

' Maintenance
FTS5Optimize conn, "docs"    ' merge b-tree segments
FTS5Rebuild  conn, "docs"    ' rebuild index from content table
```

### Connection pool

```vba
Dim pool As New SQLite3Pool
pool.Initialize DB_PATH, DLL_PATH, 4

Dim conn As SQLite3Connection
Set conn = pool.Acquire()
' ... use conn ...
pool.ReleaseConnection conn    ' auto-rolls back any open transaction
pool.ShutDown
```

### Bulk insert — 100k rows

```vba
Dim bulk As New SQLite3BulkInsert
bulk.OpenInsert conn, "prices", Array("symbol", "price", "ts"), 10000
Dim i As Long
For i = 1 To 100000
    bulk.AppendRow Array("AAPL", 189.5 + i * 0.001, Now())
Next i
bulk.CloseInsert
Debug.Print bulk.TotalRowsInserted & " rows inserted"
```

---

## Performance Under Macro Security Restrictions

If Excel's Trust Center is set to **"Disable Macros with Notification"** and
the DLL is loaded from an untrusted path, the driver may run 100x slower than
normal. This is not an Excel Trust Center issue — it is Windows Defender
scanning `sqlite3.dll`'s code pages on every `DispCallFunc` call.

Two fixes, neither requiring changes to Trust Center:

### Fix A — Place `sqlite3.dll` in `C:\Windows\System32` (recommended, but may not be allowed in your environment)

```
copy sqlite3.dll C:\Windows\System32\sqlite3.dll
```

Then use just the filename in your constants:

```vba
Private Const DLL_PATH As String = "sqlite3.dll"
```

System32 DLLs receive "known DLL" treatment — Defender does not repeatedly
scan them at runtime.

### Fix B — Add a Windows Defender folder exclusion (requires Administrator permission)

1. Open **Windows Security**
2. Go to **Virus & threat protection -> Manage settings**
3. Scroll to **Exclusions -> Add or remove exclusions**
4. Add a **Folder** exclusion for the folder containing `sqlite3.dll`

The DLL path in your VBA constants is unchanged. Defender skips real-time
scanning for that folder entirely.

---

## Architecture

```
VBA code
  +--> SQLite3Connection / SQLite3Command / SQLite3Recordset / ...
         +--> SQLite3_API.bas
                +-- LoadLibraryW("sqlite3.dll")    <- once at first OpenDatabase
                +-- GetProcAddress x 31            <- cached in m_procs(31)
                +--> DispCallFunc(0, m_procs(n), CC_CDECL, ...)  <- every call
                       +--> sqlite3.dll  (__cdecl ABI)
```

**Why `DispCallFunc` instead of `Declare`?**
`Declare PtrSafe` requires the DLL to be on the system PATH or in the same
folder as the workbook at load time, and the declaration is fixed at compile
time. `DispCallFunc` lets the driver load any path at runtime, fail gracefully
if the DLL is missing, and reload without restarting Excel.

**`DispCallFunc` calling convention:**

```vba
' prgpvarg is VARIANTARG** -- array of pointers to Variants
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance  As LongPtr,  _ ' 0 for plain function
    ByVal oVft        As LongPtr,  _ ' function address
    ByVal cc          As Long,     _ ' CC_CDECL = 1
    ByVal vtReturn    As Integer,  _ ' VT_I4 or VT_I8
    ByVal cActuals    As Long,     _
    ByRef prgvt       As Integer,  _ ' vt(0)
    ByRef prgpvarg    As LongPtr,  _ ' ptrs(0)  <- NOT Variant
    ByRef pvargResult As Variant) As Long

' Each wrapper builds ptrs(i) = VarPtr(args(i)) before calling
```

---

## FTS5 Query Syntax Reference

| Syntax | Meaning |
|--------|---------|
| `word` | Match rows containing `word` |
| `word1 word2` | Both words must appear (implicit AND) |
| `word1 AND word2` | Explicit AND |
| `word1 OR word2` | Either word |
| `NOT word` | Exclude rows containing `word` |
| `"exact phrase"` | Phrase must appear verbatim |
| `prefix*` | Prefix match — `optim*` matches `optimise`, `optimize`, etc. |
| `col : word` | Word must appear in the named column |
| `^word` | Word must be the first token in the column |

FTS5 requires SQLite 3.9.0 or later. JSON functions require SQLite 3.38.0 or later.
Both are included in all official precompiled binaries from sqlite.org. The FTS5 and
JSON test suites probe availability at startup and print `SKIP` gracefully on older builds.

---

## Pragmas applied at OpenDatabase

```sql
PRAGMA journal_mode       = WAL;
PRAGMA wal_autocheckpoint = 1000;
PRAGMA synchronous        = NORMAL;
PRAGMA cache_size         = -65536;   -- 64 MB page cache
PRAGMA temp_store         = MEMORY;
PRAGMA locking_mode       = NORMAL;
PRAGMA mmap_size          = <n>;      -- if mmapSizeBytes > 0
```

---

## Limitations

- **64-bit Excel only.** The driver unconditionally uses `LongPtr` / `LongLong`
  and will not compile in 32-bit VBA.
- **Windows only.** Relies on `kernel32` and `oleaut32`.
- **No async execution.** SQLite itself is synchronous; `busy_timeout` handles
  contention between threads.
- **No custom aggregate functions.** SQLite's `sqlite3_create_function_v2`
  requires C callback pointers that VBA cannot produce without a shim DLL.
- **No streaming BLOBs.** BLOBs are read fully into memory. Very large BLOBs
  (>50 MB) should be stored externally and referenced by path. Streaming BLOB
  I/O is planned for 0.1.4.
- **JSON functions require SQLite 3.38+.** The JSON test suite probes
  availability at startup and prints `SKIP` gracefully on older builds.

---

## License

    Copyright (C) 2026  Bryan Mark (bryan.mark@gmail.com)

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.
