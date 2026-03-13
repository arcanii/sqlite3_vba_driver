SQLite3 VBA Driver
==================
Skip down to the TLDR section to get it working a.s.a.p. Please read the Security section — it explains a 100x slowdown that catches almost everyone the first time.

**What is this?**<br/>
A **very rich** VBA SQLite3 driver for 64-bit Excel. No registered COM objects, no third-party
dependencies beyond the SQLite DLL itself. Every SQLite call goes through
`DispCallFunc`, so the DLL path is configurable at runtime with no compile-time
`Declare` statements. This tries to implement the full sqlite functionality set so you can use it in Excel (not tested with Word or Access, but I guess it should work)
Before you ask, this Project was developed in vibe with Claude Sonnet 4.6 Extended in about 3 days (I am very impressed so far!)

---

TLDR Section (for those who have no attention span)
=======
1. Get `sqlite3.dll` (64-bit) from [sqlite.org/download](https://sqlite.org/download.html) — under *Precompiled Binaries for Windows*, grab **`sqlite-dll-win-x64-*.zip`**.
2. You need 64-bit Microsoft Excel (2016 or later).
3. Make a directory (e.g. `C:\sqlite\`) — this can go anywhere, but if you change it you must also update the constants in step 7.
4. Put `sqlite3.dll` in that directory (`C:\sqlite\` if you want to skip step 7).
5. Download `Test-SQLite3-VBA-Driver.xlsm` from this repo, open it and enable macros.
6. Press `Alt+F11` to open the Visual Basic Editor.
7. **Only if you changed `C:\sqlite\`** — find `SQLite3_Tests.bas` in the Project Explorer. Near the top of the file, update:
   - `Private Const DLL_PATH As String = "C:\sqlite\sqlite3.dll"` (or just `"sqlite3.dll"` if you used System32)
   - `Private Const DB_PATH  As String = "C:\sqlite\driver_test.db"`
   - `Private Const LOG_PATH As String = "C:\sqlite\test_results.log"` (set to `""` to disable)
8. Enable "Trust access to the VBA project object model": `File → Options → Trust Center → Trust Center Settings → Macro Settings` — tick the checkbox. This lets the test runner detect which optional `.cls` files you have imported, so it can skip those suites gracefully. Without it you'll get a permissions error.
9. Open the Immediate window (`Ctrl+G`), type `RunAllTests` and press Enter.
10. All 434 tests should pass. Results are printed to the Immediate window **and** written to `LOG_PATH`.

---

Versions
=======
| Version | Date           | Tests | Pass | Highlights |
|---------|----------------|-------|------|------------|
| 0.1.7 | 14 March, 2026 |  434  |  434 | Module consolidation (23 → 13 files): `SQLite3_CoreAPI.bas` + `SQLite3_Driver.bas`; `ClassAvailable()` via `VBProject.VBComponents`; full test suite restoration and expansion |
| 0.1.6 | 12 March, 2026 |  415  |  415 | conn.Tag, ExecScriptFile, QueryColumn, ListObjectToTable, SQLite3_Migrate |
| 0.1.5 | 12 March, 2026 |  365  |  365 | Excel integration, EXPLAIN QUERY PLAN, read-only connections, Checkpoint(), QPC LRU fix, SerializeDB auto-checkpoint, structured logging |
| 0.1.4 | 12 March, 2026 |  305  |  305 | Online Backup API, incremental BLOB I/O, serialize/deserialize, diagnostics, file log output |
| 0.1.3 | 11 March, 2026 |  240  |  240 | Schema introspection, savepoints, JSON functions, interrupt, failed-test summary |
| 0.1.2 | 11 March, 2026 |  171  |  171 | BLOB support, aggregate helpers, FTS5 full-text search, GPLv3 license |
| 0.1.1 | 11 March, 2026 |  122  |  122 | DispCallFunc ABI fixes, UTF-8 fixes, QPC benchmarking, security notes |
| 0.1.0 | 09 March, 2026 |  n/a  |  n/a | Initial release — core driver, recordset, bulk insert, connection pool |

Releases Page
=======
https://github.com/arcanii/sqlite3_vba_driver/releases

---

Security Issues (important)
=======
Microsoft Excel's default security settings can slow down this driver **100x**. It is
not a Trust Center issue — it is Windows Defender scanning `sqlite3.dll`'s code pages
on every single `DispCallFunc` call.

There are two fixes (either requires Administrator privileges):

**Option A — Place `sqlite3.dll` in `C:\Windows\System32` (recommended)**
```
copy sqlite3.dll C:\Windows\System32\sqlite3.dll
```
Then use just the filename in your constants:
```vba
Private Const DLL_PATH As String = "sqlite3.dll"
```
System32 DLLs get "known DLL" treatment — Defender does not re-scan them at runtime.

**Option B — Add a Windows Defender folder exclusion**
1. Open **Windows Security** (search "Windows Security" in the Start menu)
2. Go to **Virus & threat protection → Manage settings**
3. Scroll to **Exclusions → Add or remove exclusions**
4. Add a **Folder** exclusion for the folder containing `sqlite3.dll` (e.g. `C:\sqlite\`)

The DLL path constant in your VBA is unchanged. Defender skips that folder entirely.

---

## What's New in 0.1.7

### Module consolidation — 23 files → 13 files

The biggest change in this release is housekeeping. The three Core API modules
(`SQLite3_API.bas`, `SQLite3_API_Ext.bas`, `SQLite3_Helpers.bas`) are merged into a single
**`SQLite3_CoreAPI.bas`**. The nine feature modules (`SQLite3_Aggregates.bas`,
`SQLite3_Diagnostics.bas`, `SQLite3_Excel.bas`, `SQLite3_FTS5.bas`, `SQLite3_JSON.bas`,
`SQLite3_Logger.bas`, `SQLite3_Migrate.bas`, `SQLite3_Schema.bas`, `SQLite3_Serialize.bas`)
are merged into a single **`SQLite3_Driver.bas`**. All public function names are
unchanged — existing calling code requires no modifications.

### SQLite3_Tests.bas

- `ClassAvailable()` rewritten to use `VBProject.VBComponents` for robust detection of
  optional class modules. Requires "Trust access to the VBA project object model" (see TLDR step 8).
- `RunTests_BulkInsert`, `RunTests_Pool`, `RunTests_Backup`, `RunTests_BlobStream` dispatch
  subs moved here from the individual feature modules (they no longer exist as separate files).
- 19 new assertions across existing suites: deeper coverage of FTS5, JSON, Diagnostics,
  Logger, Migrate, Transactions, Excel, Serialize, and more.
- Total: **434 / 434**

---

## What's New in 0.1.6

### New module

- **`SQLite3_Migrate.bas`** — Schema versioning via SQLite's built-in `PRAGMA user_version`
  (no extra table required). `GetSchemaVersion` / `SetSchemaVersion` for direct access.
  `ApplyMigration(conn, toVersion, sql)` applies a single DDL/DML block inside a
  transaction and advances the version only on success; skips silently if the DB is
  already at or above `toVersion`. `MigrateAll(conn, steps)` applies an ordered array
  of `MigrationStep` values and returns the count actually applied. Safe to call on
  every workbook open — already-applied steps are no-ops.

### SQLite3Connection.cls

- **`Tag` property** — optional string label attached to a connection instance.
  Included in every logger output line from that connection as
  `[SQLite3Connection[<tag>]]`, making it straightforward to distinguish
  connections in multi-connection workbook designs without changing any other API.
- **`ExecScriptFile(filePath)`** — reads a UTF-8 `.sql` file and passes the entire
  content to `sqlite3_exec` in one call. SQLite's exec handles multiple
  semicolon-delimited statements and `--` / `/* */` comments natively, so no
  client-side parsing is required. Useful for applying migration scripts, loading
  seed data, and running schema dumps produced by external tools.

### SQLite3_Helpers.bas

- **`QueryColumn(conn, sql)`** — executes a query and returns the first column of
  every result row as a zero-based `Variant()` array. Returns an empty array when
  the query produces no rows. Eliminates the boilerplate of opening a recordset,
  iterating rows, and closing — a pattern that appears constantly in practice.

### SQLite3_Excel.bas

- **`ListObjectToTable(conn, tableName, lo, dropIfExists, batchSize)`** — imports an
  Excel `ListObject` (structured Table) into SQLite. Uses the ListObject's own header
  row directly; callers do not supply `hasHeaders`. Guards against zero-row
  ListObjects. Thin wrapper over `RangeToTable` — all type inference and bulk insert
  logic is shared.

### SQLite3_Tests.bas

- 5 new test suites (39–43): `RunTest_Tag`, `RunTest_ExecScriptFile`,
  `RunTest_QueryColumn`, `RunTest_ListObject`, `RunTest_Migrate`
- Total: **415 / 415**

---

## What's New in 0.1.5

### New modules

- **`SQLite3_Logger.bas`** — Structured logging subsystem. Five levels: `LOG_DEBUG`,
  `LOG_INFO`, `LOG_WARN`, `LOG_ERROR`, `LOG_NONE`. Two sinks: Immediate window and
  append-mode file. `Logger_Configure` / `Logger_SetLevel` / `Logger_Close` for
  lifecycle control. `Logger_IsEnabled(level)` cheap boolean guard avoids building
  message strings when the level is filtered. `SQLite3Connection` emits events
  automatically at appropriate levels throughout — open/close, cache hit/miss/evict,
  every transaction verb, checkpoint results, and all error paths.

- **`SQLite3_Excel.bas`** — Excel ↔ SQLite integration. `RangeToTable` imports any
  `Range` or `ListObject` into a SQLite table: reads the entire range in one `.Value`
  call, infers column types from the first data row (Date→TEXT/ISO 8601,
  whole numbers→INTEGER, floats→REAL, else TEXT), sanitizes header strings into valid
  SQL identifiers, creates the table, and bulk-inserts via `SQLite3BulkInsert`.
  `QueryToRange` runs a SQL query and writes the result set back to a worksheet range,
  optionally including column headers.

### SQLite3Connection.cls

- **Read-only connections** — new `openReadOnly` parameter on `OpenDatabase`; opens
  with `SQLITE_OPEN_READONLY`. WAL and `locking_mode` pragmas are skipped. Write
  attempts raise `SQLITE_READONLY` from SQLite. `IsReadOnly` property added.
- **`Checkpoint(mode, schema)`** — wraps `sqlite3_wal_checkpoint_v2`. Modes: `PASSIVE`
  (default), `FULL`, `RESTART`, `TRUNCATE`. Returns `Array(pagesWritten, pagesRemaining)`.
- **QPC LRU timestamps** — statement cache `lastUsed` field changed from `Double`
  (Timer(), 1-second resolution, wraps at midnight) to `LongLong` (QPC ticks,
  ~100 ns resolution, no wrap for 292 years).
- **Logger integration** — key events emitted automatically at DEBUG/INFO/WARN/ERROR.

### SQLite3_API.bas

- `sqlite3_wal_checkpoint_v2` wrapper added (`P_WAL_CKPT = 47`).
- `SQLITE_OPEN_READONLY` and `SQLITE_CHECKPOINT_PASSIVE/FULL/RESTART/TRUNCATE` constants added.
- `PROC_COUNT` bumped from 47 → 48.

### SQLite3_Helpers.bas

- **`GetQueryPlan(conn, sql)`** — runs `EXPLAIN QUERY PLAN` and returns the result as
  a `(n × 4)` Variant matrix: columns `(id, parent, notused, detail)`. One call,
  immediately useful for index tuning.

### SQLite3_Serialize.bas

- **`SerializeDB` auto-checkpoint** — calls `sqlite3_wal_checkpoint_v2` with `TRUNCATE`
  mode before serializing. Outstanding WAL frames are folded into the main file first,
  so the snapshot is always clean regardless of whether the source is WAL-mode. The
  checkpoint error is swallowed silently (non-WAL databases return `SQLITE_OK`
  immediately; a `SQLITE_BUSY` still produces a valid, slightly stale, snapshot).

### SQLite3_Tests.bas

- 5 new test suites (34–38): `RunTest_ReadOnly`, `RunTest_Checkpoint`,
  `RunTest_QueryPlan`, `RunTest_Excel`, `RunTest_Logger`
- Total: **365 / 365**

---

## Features

| Feature | Detail |
|---------|--------|
| Dynamic DLL loading | `LoadLibrary` at runtime — no hard `Declare` required |
| UTF-8 marshalling | Full round-trip via `WideCharToMultiByte` / `MultiByteToWideChar` |
| Prepared statements | Positional (`?`) and named (`:param`) binding |
| BLOB support | `BindBlob`, `AsBytes()`, vectorized BLOB load |
| Incremental BLOB I/O | `SQLite3BlobStream` — read/write byte ranges without full load |
| Statement cache | 64-slot LRU per connection (QPC timestamps) — cache hit = reset only, no re-prepare |
| ADO-style recordset | `BOF`/`EOF`/`MoveNext`/`MoveLast`, `rs!FieldName` syntax |
| Vectorized load | `LoadAll()` pulls entire result into a Variant matrix (~50× faster than live) |
| `ToMatrix()` | Returns `(row, col)` Variant array ready for direct Excel range assignment |
| Bulk insert | Single reused prepared `INSERT`, transaction-batched (~100k rows/sec) |
| Connection pool | LRU idle reaping, auto-rollback on release, configurable max size |
| Savepoints | Nested transactions: `Savepoint` / `ReleaseSavepoint` / `RollbackToSavepoint` |
| Interrupt | `conn.Interrupt` cancels a running query via `sqlite3_interrupt` |
| Online backup | Hot-backup any live DB; one-shot or incremental with progress reporting |
| Serialize / deserialize | Snapshot a DB to a `Byte()` array; restore to any connection |
| In-memory clone | `InMemoryClone` — independent `:memory:` copy via the backup API |
| Diagnostics | `sqlite3_db_status` / `sqlite3_stmt_status` counters per connection and statement |
| Schema introspection | Tables, views, columns, indexes, FKs, triggers, CREATE SQL, PRAGMA info |
| Aggregate helpers | `GroupBy*`, `ScalarAgg`, `MultiAgg`, `RunningTotal`, `Histogram` |
| JSON functions | `JSONExtract`, `JSONSet`, `JSONPatch`, `JSONSearch`, `JSONGroupArray` and more |
| FTS5 full-text search | Create, insert, search, snippet, highlight, BM25 ranking, optimize |
| WAL mode | Enabled by default on `OpenDatabase`; skip with `enableWAL=False` |
| Connection tagging | `conn.Tag` — label included in all log output; identifies connections in multi-connection workbooks |
| Script file execution | `conn.ExecScriptFile(path)` — run a multi-statement UTF-8 `.sql` file in one call |
| QueryColumn | `QueryColumn(conn, sql)` — first column of every result row as a flat `Variant()` array |
| ListObject import | `ListObjectToTable` — import an Excel structured Table into SQLite directly |
| Schema versioning | `MigrateAll` / `ApplyMigration`: `GetSchemaVersion`, `SetSchemaVersion` |
| WAL checkpoint | `conn.Checkpoint(mode)` — PASSIVE / FULL / RESTART / TRUNCATE |
| Excel integration | `RangeToTable` (range→SQLite) and `QueryToRange` (SQL→worksheet) |
| EXPLAIN QUERY PLAN | `GetQueryPlan(conn, sql)` returns plan nodes as a Variant matrix |
| Structured logging | DEBUG/INFO/WARN/ERROR/NONE, Immediate + file sinks |
| QPC benchmarking | `QueryPerformanceCounter` timing in every test suite and LRU cache |
| File log output | `RunAllTests` writes a full copy of results to `LOG_PATH` |
| Failed-test summary | All failures reprinted at the end of `RunAllTests` |
| 64-bit only | All handles are `LongPtr` / `LongLong` — requires 64-bit Excel |

---

## File Reference

| File | Role |
|------|------|
| `SQLite3_CoreAPI.bas` | DLL loader, 48 cached proc addresses, all SQLite wrappers via `DispCallFunc`, UTF-8 marshalling; public helpers: `QueryScalar`, `QueryColumn`, `TableExists`, `TableRowCount`, `GetQueryPlan`, `GetColumnInfo` |
| `SQLite3_Driver.bas` | All feature logic in one module: Aggregates, Diagnostics, Excel, FTS5, JSON, Logger, Migrate, Schema, Serialize |
| `SQLite3Connection.cls` | Open/close, WAL, mmap, 64-slot LRU cache (QPC), transactions, savepoints, interrupt, checkpoint, read-only mode, Tag, ExecScriptFile |
| `SQLite3Recordset.cls` | Live and vectorized recordset, `GetRows()`, `ToMatrix()`, `rs!Field` |
| `SQLite3Fields.cls` | Case-insensitive field collection, `For Each` enumerator |
| `SQLite3Field.cls` | Zero-copy value reads; `Value`, `AsString`, `AsBytes`, `AsInt64` |
| `SQLite3Command.cls` | Positional and named binding, `BindBlob`, `BindVariant`, `ExecuteScalar`, `StmtHandle` |
| `SQLite3BulkInsert.cls` | High-speed batch insert, `AppendRow`, `AppendMatrix` |
| `SQLite3Pool.cls` | Connection pool, LRU reaping, auto-rollback on release, pre-warm |
| `SQLite3Backup.cls` | Online Backup API — `BackupToFile`, `OpenBackup`, `Step`, `CloseBackup`, progress |
| `SQLite3BlobStream.cls` | Incremental BLOB I/O — `OpenBlob`, `ReadAt`, `WriteAt`, `SeekTo` |
| `SQLite3_Tests.bas` | 44 suites, 434 assertions, `RunAllTests` entry point, QPC timing, file logging, failure summary |
| `SQLite3_Examples.bas` | Annotated usage examples for every feature |

---

## Requirements

- **64-bit Excel** (Excel 2016 or later, 64-bit install)
- **`sqlite3.dll`** (64-bit) — download from [sqlite.org/download](https://sqlite.org/download.html)
  - Under *Precompiled Binaries for Windows*, grab **`sqlite-dll-win-x64-*.zip`**
  - FTS5, JSON, backup, and serialize are enabled in all official precompiled binaries
  - Serialize/deserialize requires SQLite 3.23.0 or later (released 2018 — all current builds qualify)
- **Microsoft Scripting Runtime** reference (for `Dictionary` used in `SQLite3Fields`)
- **Trust access to the VBA project object model** enabled (for `ClassAvailable()` in `SQLite3_Tests.bas`)

---

## Installation

### 1. Place the DLL

Copy `sqlite3.dll` (64-bit) to a stable path. The recommended location is
`C:\Windows\System32` (see Security section). Any path works as long as you set
`DLL_PATH` accordingly.

### 2. Import the VBA modules

Open the Visual Basic Editor (`Alt+F11`), then for each file choose
**File → Import File** (or drag-and-drop onto the Project Explorer).
Import `SQLite3_CoreAPI.bas` first. You only need the core files and the classes
for features you want. To run all the tests, import everything.

```
 1.  SQLite3_CoreAPI.bas       ← Core (import first)
 2.  SQLite3_Driver.bas        ← All features
 3.  SQLite3Field.cls          ← Core
 4.  SQLite3Fields.cls         ← Core
 5.  SQLite3Command.cls
 6.  SQLite3Recordset.cls      ← Core
 7.  SQLite3Connection.cls     ← Core
 8.  SQLite3BulkInsert.cls
 9.  SQLite3Pool.cls
10.  SQLite3Backup.cls
11.  SQLite3BlobStream.cls
12.  SQLite3_Examples.bas      ← for learning
13.  SQLite3_Tests.bas         ← for testing
```

### 3. Add the Scripting Runtime reference

In the VBA Editor: **Tools → References → check "Microsoft Scripting Runtime"**

### 4. Trust access to the VBA project object model

`File → Options → Trust Center → Trust Center Settings → Macro Settings`
→ check **Trust access to the VBA project object model**

The test runner uses `VBProject.VBComponents` to detect which optional class
modules you have imported and skips those suites gracefully. Without this setting
you will get a permissions error when `RunAllTests` starts.

---

## Running the Tests

1. Set the constants at the top of `SQLite3_Tests.bas`:

```vba
Private Const DLL_PATH As String = "sqlite3.dll"                        ' or full path
Private Const DB_PATH  As String = "C:\sqlite\driver_test.db"
Private Const LOG_PATH As String = "C:\sqlite\test_results.log"         ' "" to disable
```

2. Open the Immediate window (`Ctrl+G`) and type:

```
RunAllTests
```

Expected output (abridged):

```
================================================================
SQLite3 Driver Test Suite
Started: 2026-03-14 09:41:03
================================================================

  [DllLoad]
    PASS  SQLite_Load
    PASS  SQLite_IsLoaded
    INFO  SQLite version = 3.47.0
    TIME  1.23 ms

  [Backup]
    PASS  Open backup source DB
    PASS  Source rows before backup
    PASS  BackupToFile no error
    PASS  Backup IsComplete
    PASS  Backup row 250 correct
    TIME  18.42 ms

  [Serialize]
    PASS  SerializeDB no error
    PASS  Row 50 correct
    PASS  InMemoryClone no error
    PASS  Clone row 99
    TIME  22.18 ms

  [Diagnostics]
    PASS  GetDbStatus no error
    PASS  First counter name
    PASS  GetStmtStatus no error
    TIME  8.91 ms

  ...

================================================================
Results: 434 passed,  0 failed  (434 total)  1946.71 ms
================================================================

Log written to: C:\sqlite\test_results.log
```

If any tests fail, a consolidated summary is printed at the very end:

```
FAILED TESTS (1):
----------------------------------------------------------------
  [Backup]  Backup row 250 correct -- expected [row_250] got []
----------------------------------------------------------------
```

Individual suites can be run in isolation — each is a standalone `Public Sub`:

```
RunTest_Backup          RunTest_BlobStream      RunTest_Serialize
RunTest_Diagnostics     RunTest_Schema          RunTest_Savepoints
RunTest_JSON            RunTest_Interrupt       RunTest_FTS5
RunTest_Aggregates      RunTest_BLOB            RunTest_ConnectionPool
RunTest_BulkInsert_AppendRow                    RunTest_Transactions
RunTest_Tag             RunTest_ExecScriptFile  RunTest_QueryColumn
RunTest_ListObject      RunTest_Migrate         RunTest_Logger
RunTest_ReadOnly        RunTest_Checkpoint      RunTest_QueryPlan
RunTest_Excel
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

' Read via live recordset
Dim rs As SQLite3Recordset
Set rs = conn.OpenRecordset("SELECT data FROM assets WHERE name='logo.png';")
Dim blob() As Byte: blob = rs.Fields("data").AsBytes()
```

### Online backup

```vba
' One-shot backup (blocks until complete)
Dim bak As New SQLite3Backup
bak.BackupToFile conn, "C:\backups\mydb_backup.db"
Debug.Print bak.TotalPages & " pages copied"

' Incremental backup with progress
bak.OpenBackup conn, "C:\backups\mydb_backup.db"
Do While Not bak.IsComplete
    bak.Step 100
    Debug.Print Format(bak.Progress * 100, "0.0") & "% complete"
Loop
bak.CloseBackup
```

### Incremental BLOB I/O

```vba
' Insert a zeroblob placeholder, then fill it via BlobStream
conn.ExecSQL "INSERT INTO chunks (id, data) VALUES (1, zeroblob(65536));"
Dim rowId As LongLong: rowId = conn.LastInsertRowID()

conn.BeginTransaction
Dim bs As New SQLite3BlobStream
bs.OpenBlob conn, "main", "chunks", "data", rowId, True   ' True = writable
Dim chunk(4095) As Byte
' ... fill chunk ...
bs.WriteAt chunk, 0       ' write at byte offset 0
bs.WriteAt chunk, 4096    ' write at byte offset 4096
bs.CloseBlob
conn.CommitTransaction

' Read back a specific range without loading the full BLOB
bs.OpenBlob conn, "main", "chunks", "data", rowId, False
Dim portion() As Byte: portion = bs.ReadAt(4096, 0)   ' 4096 bytes from offset 0
bs.CloseBlob
```

### Serialize / deserialize

```vba
' Snapshot to a byte array — use :memory: source for cleanest result
' (a file DB opened by earlier code may be in WAL mode; :memory: is always clean)
Dim src As New SQLite3Connection
src.OpenDatabase ":memory:", DLL_PATH
' ... populate src ...
Dim snap() As Byte: snap = SerializeDB(src)

' Restore to a fresh in-memory connection
Dim mem As New SQLite3Connection
mem.OpenDatabase ":memory:", DLL_PATH, 5000, False
DeserializeDB mem, snap
' mem is now a complete independent copy of src at that moment

' Convenience: one-call clone of any open connection
Dim clone As SQLite3Connection
Set clone = InMemoryClone(conn)
' clone is fully independent — mutations to conn don't affect clone, and vice versa
clone.ExecSQL "DELETE FROM sensitive_data;"
clone.CloseConnection
```

### Diagnostics

```vba
' Print a full db_status summary to the Immediate window
DbStatusSummary conn

' Read a specific counter
Dim used As Long, maxVal As Long
GetDbStatus conn, DBSTAT_LOOKASIDE_USED, used, maxVal
Debug.Print "Lookaside in use: " & used & "  peak: " & maxVal

' Statement-level counters
Dim cmd As New SQLite3Command
cmd.Prepare conn, "SELECT * FROM prices WHERE symbol = ?;"
cmd.BindText 1, "AAPL"
' ... execute ...
Debug.Print "VM steps: " & GetStmtStatus(cmd.StmtHandle, STMTSTAT_VM_STEP, False)
```

### Schema introspection

```vba
Dim tables As Variant: tables = GetTableList(conn)
Dim cols   As Variant: cols   = GetColumnInfo(conn, "orders")
Dim fks    As Variant: fks    = GetForeignKeys(conn, "order_lines")
Dim ddl    As String:  ddl    = GetCreateSQL(conn, "orders")
Dim info   As Variant: info   = GetDatabaseInfo(conn)

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
conn.RollbackToSavepoint "sp1"   ' undo just the inner work
conn.ReleaseSavepoint "sp1"

conn.CommitTransaction            ' only the outer INSERT is kept
```

### JSON functions (requires SQLite 3.38+)

```vba
conn.ExecSQL "CREATE TABLE users (id INTEGER, profile TEXT);"
conn.ExecSQL "INSERT INTO users VALUES (1, '{""name"":""Alice"",""city"":""London""}');"

Dim city As Variant
city = JSONExtract(conn, "users", "profile", "$.city", "id=1")   ' -> "London"

JSONSet    conn, "users", "profile", "$.city",             "'Paris'",     "id=1"
JSONPatch  conn, "users", "profile", "'{""country"":""France""}'",        "id=1"
JSONRemove conn, "users", "profile", Array("$.city"),                     "id=1"

Dim arr As String
arr = JSONGroupArray(conn, "users", "json_extract(profile,'$.name')")
' -> '["Alice","Bob"]'
```

### Aggregate helpers

```vba
Dim mat As Variant
mat = GroupByCount(conn, "sales", "region", "", 10)
Sheet1.Range("A1").Resize(UBound(mat,1)+1, 2).Value = mat

mat = MultiAgg(conn, "trades", _
               Array("COUNT(*) AS n", "SUM(qty) AS vol", "AVG(price) AS avg_px"))
mat = RunningTotal(conn, "trades", "trade_date", "pnl")
mat = Histogram(conn, "trades", "price", 20)
```

### FTS5 full-text search

```vba
CreateFTS5Table conn, "docs", Array("title", "body"), "", "porter unicode61", True

FTS5Insert conn, "docs", Array("title", "body"), _
           Array("SQLite Guide", "How to use SQLite for fast data storage")

Dim results As Variant
results = FTS5SearchMatrix(conn, "docs", "sqlite storage")
results = FTS5Snippet(conn, "docs", "sqlite", 0, "<b>", "</b>", "...", 16, 20)
results = FTS5BM25Search(conn, "docs", "fast data storage", "*", 10)

FTS5Optimize conn, "docs"
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

## Architecture

```
VBA code
  +--> SQLite3Connection / SQLite3Command / SQLite3Recordset / ...
         +--> SQLite3_CoreAPI.bas
                +-- LoadLibraryW("sqlite3.dll")    <- once at first OpenDatabase
                +-- GetProcAddress x 48            <- cached in m_procs(47)
                +--> DispCallFunc(0, m_procs(n), CC_CDECL, ...)  <- every call
                       +--> sqlite3.dll  (__cdecl ABI)
```

**Why `DispCallFunc` instead of `Declare`?**
`Declare PtrSafe` requires the DLL on the system PATH or next to the workbook at
load time, and the declaration is fixed at compile time. `DispCallFunc` lets the
driver load any path at runtime, fail gracefully if the DLL is missing, and reload
without restarting Excel.

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

' Each wrapper builds ptrs(i) = VarPtr(args(i)) before calling.
' Pointer args:  CLngLng(ptr), VT_I8
' Integer args:  CLng(val),    VT_I4
' Double args:   val,          VT_R8
' Return ptr:    VT_I8, CLngPtr(ret)
' Return int:    VT_I4, CLng(ret)
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

FTS5 requires SQLite 3.9.0+. JSON requires SQLite 3.38.0+. Both are included in all
official precompiled binaries from sqlite.org. The FTS5 and JSON test suites probe
availability at startup and print `SKIP` gracefully on older builds.

---

## Diagnostics Reference

### db_status constants (`SQLite3_Driver.bas`)

| Constant | Value | Meaning |
|----------|-------|---------|
| `DBSTAT_LOOKASIDE_USED` | 0 | Lookaside slots currently in use |
| `DBSTAT_CACHE_USED` | 1 | Page cache memory in use (bytes) |
| `DBSTAT_SCHEMA_USED` | 2 | Memory used for schema (bytes) |
| `DBSTAT_STMT_USED` | 3 | Memory used for prepared statements (bytes) |
| `DBSTAT_LOOKASIDE_HIT` | 4 | Lookaside allocations that succeeded |
| `DBSTAT_LOOKASIDE_MISS_SIZE` | 5 | Failed due to allocation size |
| `DBSTAT_LOOKASIDE_MISS_FULL` | 6 | Failed due to lookaside buffer full |
| `DBSTAT_CACHE_HIT` | 7 | Page cache hits |
| `DBSTAT_CACHE_MISS` | 8 | Page cache misses |
| `DBSTAT_CACHE_WRITE` | 9 | Page cache write-backs |
| `DBSTAT_DEFERRED_FKS` | 10 | Unresolved deferred foreign keys |
| `DBSTAT_CACHE_USED_SHARED` | 11 | Shared cache memory (bytes) |
| `DBSTAT_CACHE_SPILL` | 12 | Cache spill-to-disk events |

### stmt_status constants

| Constant | Value | Meaning |
|----------|-------|---------|
| `STMTSTAT_FULLSCAN` | 1 | Full table scans performed |
| `STMTSTAT_SORT` | 2 | Sort operations |
| `STMTSTAT_AUTOINDEX` | 3 | Automatic indexes created |
| `STMTSTAT_VM_STEP` | 4 | Virtual machine instructions executed |
| `STMTSTAT_REPREPARE` | 5 | Statement re-preparations |
| `STMTSTAT_RUN` | 6 | Times the statement has been run |
| `STMTSTAT_FILTER_MISS` | 7 | Bloom filter misses |
| `STMTSTAT_FILTER_HIT` | 8 | Bloom filter hits |
| `STMTSTAT_MEMUSED` | 99 | Memory in use by the statement (bytes) |

---

## Pragmas applied at OpenDatabase

```sql
PRAGMA journal_mode       = WAL;          -- write-ahead logging (omitted if enableWAL=False)
PRAGMA wal_autocheckpoint = 1000;
PRAGMA synchronous        = NORMAL;
PRAGMA cache_size         = -65536;       -- 64 MB page cache
PRAGMA temp_store         = MEMORY;
PRAGMA locking_mode       = NORMAL;
PRAGMA mmap_size          = <n>;          -- only if mmapSizeBytes > 0
```

---

## Limitations

- **64-bit Excel only.** Uses `LongPtr` / `LongLong` throughout — will not compile in 32-bit VBA.
- **Windows only.** Relies on `kernel32` and `oleaut32`.
- **No async execution.** SQLite is synchronous; `busy_timeout` handles lock contention.
- **No custom aggregate functions.** `sqlite3_create_function_v2` requires C callback
  pointers that VBA cannot produce without a shim DLL.
- **JSON functions require SQLite 3.38+.** Both JSON and FTS5 test suites probe
  availability at startup and print `SKIP` gracefully on older builds.
- **Serialize requires SQLite 3.23+** (released 2018). All current official binaries qualify.
- **`RunAllTests` requires "Trust access to the VBA project object model"** — needed for
  `ClassAvailable()` to detect optional class modules. Without it you get a permissions
  error at startup; individual `RunTest_*` subs still work fine without it.

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
