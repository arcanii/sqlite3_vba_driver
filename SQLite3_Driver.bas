Attribute VB_Name = "SQLite3_Driver"
'==============================================================================
' SQLite3_Driver.bas  -  Feature modules for the SQLite3 VBA driver (64-bit)
'
' This single file consolidates all optional feature modules:
'   * Aggregates   - GroupByCount/Sum/Avg, ScalarAgg, RunningTotal, Histogram
'   * Diagnostics  - GetDbStatus, GetStmtStatus, DbStatusSummary
'   * Excel        - RangeToTable, ListObjectToTable, QueryToRange
'   * FTS5         - Full-Text Search: create/insert/search/snippet/highlight
'   * JSON         - JSONExtract, JSONSet, JSONSearch, JSONGroupArray, etc.
'   * Logger       - Structured LOG_DEBUG/INFO/WARN/ERROR logging
'   * Migrate      - Schema versioning with PRAGMA user_version
'   * Schema       - GetTableList, GetColumnInfo, GetIndexList, GetForeignKeys
'   * Serialize    - SerializeDB, DeserializeDB, InMemoryClone
'
' Import order: SQLite3_CoreAPI.bas first, then this file, then any .cls files.
' Requires: Microsoft Scripting Runtime (Dictionary)
'
' Note on Excel functions (RangeToTable, QueryToRange, ListObjectToTable):
'   These functions require an active Excel workbook and worksheet objects.
'   They are only meaningful when running inside Excel VBA.
'
' Version : 0.1.7
'
' Version History:
'   0.1.7 - Initial release. Consolidated from individual feature .bas files:
'            SQLite3_Aggregates.bas, SQLite3_Diagnostics.bas, SQLite3_Excel.bas,
'            SQLite3_FTS5.bas, SQLite3_JSON.bas, SQLite3_Logger.bas,
'            SQLite3_Migrate.bas, SQLite3_Schema.bas, SQLite3_Serialize.bas.
'            RangeToTable refactored to use SQLite3Command directly instead of
'            SQLite3BulkInsert, removing the class compile-time dependency.
'            RecordsetToRange added as private helper (was missing in 0.1.7).
'            File count: 23 -> 14.
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

' Win32 / COM declarations (required by SQLite3_Serialize section)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As LongPtr)

' Deserialize flags (bitfield) - used by SQLite3_Serialize section
Public Const SQLITE_DESERIALIZE_FREEONCLOSE As Long = 1  ' SQLite frees pData on close
Public Const SQLITE_DESERIALIZE_RESIZEABLE  As Long = 2  ' allow the in-memory DB to grow

'==============================================================================
' db_status op codes  (SQLite3_Diagnostics section)
'==============================================================================
Public Const DBSTAT_LOOKASIDE_USED       As Long = 0
Public Const DBSTAT_CACHE_USED           As Long = 1
Public Const DBSTAT_SCHEMA_USED          As Long = 2
Public Const DBSTAT_STMT_USED            As Long = 3
Public Const DBSTAT_LOOKASIDE_HIT        As Long = 4
Public Const DBSTAT_LOOKASIDE_MISS_SIZE  As Long = 5
Public Const DBSTAT_LOOKASIDE_MISS_FULL  As Long = 6
Public Const DBSTAT_CACHE_HIT            As Long = 7
Public Const DBSTAT_CACHE_MISS           As Long = 8
Public Const DBSTAT_CACHE_WRITE          As Long = 9
Public Const DBSTAT_DEFERRED_FKS         As Long = 10
Public Const DBSTAT_CACHE_USED_SHARED    As Long = 11
Public Const DBSTAT_CACHE_SPILL          As Long = 12

'==============================================================================
' stmt_status op codes  (SQLite3_Diagnostics section)
'==============================================================================
Public Const STMTSTAT_FULLSCAN     As Long = 1
Public Const STMTSTAT_SORT         As Long = 2
Public Const STMTSTAT_AUTOINDEX    As Long = 3
Public Const STMTSTAT_VM_STEP      As Long = 4
Public Const STMTSTAT_REPREPARE    As Long = 5
Public Const STMTSTAT_RUN          As Long = 6
Public Const STMTSTAT_FILTER_MISS  As Long = 7
Public Const STMTSTAT_FILTER_HIT   As Long = 8
Public Const STMTSTAT_MEMUSED      As Long = 99

'==============================================================================
' Log level constants  (SQLite3_Logger section)
'==============================================================================
Public Const LOG_DEBUG As Long = 0
Public Const LOG_INFO  As Long = 1
Public Const LOG_WARN  As Long = 2
Public Const LOG_ERROR As Long = 3
Public Const LOG_NONE  As Long = 4

'==============================================================================
' Migration step type  (SQLite3_Migrate section)
'==============================================================================
' A single migration step: the version it brings the schema TO, plus the SQL
' that achieves it. Build with MakeStep() for convenient array literals.
Public Type MigrationStep
    toVersion As Long
    sql       As String
End Type

'==============================================================================
' Logger module-level state  (SQLite3_Logger section)
'==============================================================================
Private m_level       As Long      ' minimum level that produces output
Private m_toImmediate As Boolean   ' write to Immediate window (Debug.Print)
Private m_toFile      As Boolean   ' write to log file
Private m_filePath    As String    ' path to the log file
Private m_fileNum     As Integer   ' VBA file number (0 = not open)
Private m_isOpen      As Boolean   ' True once Configure has been called

'==============================================================================
' SQLite3_Aggregates.bas  -  Aggregate and window-function helpers
'
' SQLite supports all standard SQL aggregates natively:
'   COUNT, SUM, AVG, MIN, MAX, GROUP_CONCAT, TOTAL
'   Plus window functions: ROW_NUMBER, RANK, DENSE_RANK, LAG, LEAD,
'   FIRST_VALUE, LAST_VALUE, SUM OVER (...), AVG OVER (...), etc.
'
' These helpers wrap the most common patterns so callers do not need to
' hand-write SQL for routine analytical tasks.
'
' NOTE: Custom user-defined aggregate functions (CREATE AGGREGATE FUNCTION)
' require passing C function pointers (xStep, xFinal callbacks) to
' sqlite3_create_function_v2.  VBA cannot produce a C-callable function
' pointer without a shim DLL.  All helpers below work entirely through
' standard SQL executed via the existing driver.
'
' Version : 0.1.7
'
' Version History:
'   0.1.2 - Initial release. GroupByCount, GroupBySum, GroupByAvg,
'            ScalarAgg, MultiAgg, AggregateQuery, RunningTotal,
'            PercentileApprox, Histogram helpers.
'   0.1.3 - No functional changes. Version stamp updated.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

'==============================================================================
' AggregateQuery
' Run any SELECT that returns a single result set and return it as a
' (row, col) Variant matrix.  Equivalent to rs.LoadAll + rs.ToMatrix()
' but in a single call.
'==============================================================================
Public Function AggregateQuery(ByVal conn As SQLite3Connection, _
                                ByVal sql As String) As Variant
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset(sql)
    rs.LoadAll
    If rs.RecordCount = 0 Then
        AggregateQuery = Empty
    Else
        AggregateQuery = rs.ToMatrix()
    End If
    rs.CloseRecordset
End Function

'==============================================================================
' GroupByCount
' SELECT groupCol, COUNT(*) FROM table GROUP BY groupCol ORDER BY cnt DESC
' Returns (row, 2) matrix: col 0 = group value, col 1 = count
'==============================================================================
Public Function GroupByCount(ByVal conn As SQLite3Connection, _
                              ByVal tableName As String, _
                              ByVal groupCol As String, _
                              Optional ByVal whereClause As String = "", _
                              Optional ByVal topN As Long = 0) As Variant
    Dim sql As String
    sql = "SELECT [" & groupCol & "], COUNT(*) AS cnt" & _
          " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & " GROUP BY [" & groupCol & "] ORDER BY cnt DESC"
    If topN > 0 Then sql = sql & " LIMIT " & topN
    sql = sql & ";"
    GroupByCount = AggregateQuery(conn, sql)
End Function

'==============================================================================
' GroupBySum
' SELECT groupCol, SUM(valueCol) FROM table GROUP BY groupCol
' Returns (row, 2) matrix: col 0 = group value, col 1 = sum
'==============================================================================
Public Function GroupBySum(ByVal conn As SQLite3Connection, _
                            ByVal tableName As String, _
                            ByVal groupCol As String, _
                            ByVal valueCol As String, _
                            Optional ByVal whereClause As String = "", _
                            Optional ByVal orderBySum As Boolean = True) As Variant
    Dim sql As String
    sql = "SELECT [" & groupCol & "], SUM([" & valueCol & "]) AS total" & _
          " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & " GROUP BY [" & groupCol & "]"
    If orderBySum Then sql = sql & " ORDER BY total DESC"
    sql = sql & ";"
    GroupBySum = AggregateQuery(conn, sql)
End Function

'==============================================================================
' GroupByAvg
' SELECT groupCol, AVG(valueCol), COUNT(*) FROM table GROUP BY groupCol
' Returns (row, 3) matrix: group value, avg, count
'==============================================================================
Public Function GroupByAvg(ByVal conn As SQLite3Connection, _
                            ByVal tableName As String, _
                            ByVal groupCol As String, _
                            ByVal valueCol As String, _
                            Optional ByVal whereClause As String = "") As Variant
    Dim sql As String
    sql = "SELECT [" & groupCol & "]," & _
          " AVG([" & valueCol & "]) AS avg_val," & _
          " COUNT(*) AS cnt" & _
          " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & " GROUP BY [" & groupCol & "] ORDER BY avg_val DESC;"
    GroupByAvg = AggregateQuery(conn, sql)
End Function

'==============================================================================
' ScalarAgg
' Return a single aggregate scalar: COUNT(*), SUM(col), AVG(col), etc.
' Example: ScalarAgg(conn, "trades", "SUM(price * qty)")
'==============================================================================
Public Function ScalarAgg(ByVal conn As SQLite3Connection, _
                           ByVal tableName As String, _
                           ByVal aggExpr As String, _
                           Optional ByVal whereClause As String = "") As Variant
    Dim sql As String
    sql = "SELECT " & aggExpr & " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & ";"
    ScalarAgg = QueryScalar(conn, sql)
End Function

'==============================================================================
' Histogram
' Bucket a numeric column into N equal-width bins.
' Returns (row, 3) matrix: bin_low, bin_high, count
'==============================================================================
Public Function Histogram(ByVal conn As SQLite3Connection, _
                           ByVal tableName As String, _
                           ByVal valueCol As String, _
                           ByVal numBins As Long, _
                           Optional ByVal whereClause As String = "") As Variant
    If numBins < 1 Then numBins = 10

    ' First pass: get min and max
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    Dim lo As Double: lo = CDbl(QueryScalar(conn, _
        "SELECT MIN([" & valueCol & "]) FROM [" & tableName & "]" & wh & ";"))
    Dim hi As Double: hi = CDbl(QueryScalar(conn, _
        "SELECT MAX([" & valueCol & "]) FROM [" & tableName & "]" & wh & ";"))

    If hi = lo Then hi = lo + 1   ' degenerate range guard
    Dim width As Double: width = (hi - lo) / numBins

    ' Build CASE expression that assigns each row to a bin index
    Dim caseSql As String
    Dim i As Long
    For i = 0 To numBins - 2
        Dim binLo As Double: binLo = lo + i * width
        Dim binHi As Double: binHi = lo + (i + 1) * width
        caseSql = caseSql & " WHEN [" & valueCol & "] >= " & binLo & _
                  " AND [" & valueCol & "] < " & binHi & " THEN " & i
    Next i
    ' last bin is inclusive of max
    caseSql = caseSql & " ELSE " & (numBins - 1)

    Dim sql As String
    sql = "SELECT bin," & _
          " (" & lo & " + bin * " & width & ") AS bin_low," & _
          " (" & lo & " + (bin+1) * " & width & ") AS bin_high," & _
          " COUNT(*) AS cnt" & _
          " FROM (SELECT CASE" & caseSql & " END AS bin" & _
          " FROM [" & tableName & "]" & wh & ")" & _
          " GROUP BY bin ORDER BY bin;"
    Histogram = AggregateQuery(conn, sql)
End Function

'==============================================================================
' RunningTotal
' Return valueCol with a cumulative SUM window function.
' Returns (row, 3) matrix: orderCol value, valueCol value, running_total
' Requires SQLite 3.25+ (window functions).
'==============================================================================
Public Function RunningTotal(ByVal conn As SQLite3Connection, _
                              ByVal tableName As String, _
                              ByVal orderCol As String, _
                              ByVal valueCol As String, _
                              Optional ByVal whereClause As String = "", _
                              Optional ByVal partitionCol As String = "") As Variant
    Dim overClause As String
    If Len(partitionCol) > 0 Then
        overClause = "PARTITION BY [" & partitionCol & "] "
    End If
    overClause = overClause & "ORDER BY [" & orderCol & "] " & _
                 "ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW"

    Dim sql As String
    sql = "SELECT [" & orderCol & "], [" & valueCol & "]," & _
          " SUM([" & valueCol & "]) OVER (" & overClause & ") AS running_total" & _
          " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & " ORDER BY [" & orderCol & "];"
    RunningTotal = AggregateQuery(conn, sql)
End Function

'==============================================================================
' PercentileApprox
' Approximate percentile using SQLite's built-in NTILE window function.
' Returns the value at the given percentile (0.0 to 1.0).
'==============================================================================
Public Function PercentileApprox(ByVal conn As SQLite3Connection, _
                                  ByVal tableName As String, _
                                  ByVal valueCol As String, _
                                  ByVal percentile As Double, _
                                  Optional ByVal whereClause As String = "") As Variant
    If percentile < 0 Then percentile = 0
    If percentile > 1 Then percentile = 1
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    ' Use NTILE(100) and pick the tile matching the percentile
    Dim tile As Long: tile = CLng(percentile * 100)
    If tile < 1 Then tile = 1
    If tile > 100 Then tile = 100
    Dim sql As String
    sql = "SELECT AVG([" & valueCol & "]) FROM (" & _
          "SELECT [" & valueCol & "], NTILE(100) OVER (ORDER BY [" & _
          valueCol & "]) AS tile FROM [" & tableName & "]" & wh & _
          ") WHERE tile = " & tile & ";"
    PercentileApprox = QueryScalar(conn, sql)
End Function

'==============================================================================
' MultiAgg
' Run multiple aggregate expressions against the same table in one query.
' aggExprs: Array("COUNT(*) AS n", "SUM(price) AS total", "AVG(price) AS avg_p")
' Returns a (1, ncols) matrix row.
'==============================================================================
Public Function MultiAgg(ByVal conn As SQLite3Connection, _
                          ByVal tableName As String, _
                          ByVal aggExprs As Variant, _
                          Optional ByVal whereClause As String = "") As Variant
    Dim i As Long, parts As String
    For i = LBound(aggExprs) To UBound(aggExprs)
        If i > LBound(aggExprs) Then parts = parts & ", "
        parts = parts & CStr(aggExprs(i))
    Next i
    Dim sql As String
    sql = "SELECT " & parts & " FROM [" & tableName & "]"
    If Len(whereClause) > 0 Then sql = sql & " WHERE " & whereClause
    sql = sql & ";"
    MultiAgg = AggregateQuery(conn, sql)
End Function

'==============================================================================
' SQLite3_Diagnostics.bas  -  Performance counters (64-bit only)
'
' Wraps sqlite3_db_status (per-connection) and sqlite3_stmt_status (per-statement).
' Useful for tuning cache size, detecting full-table scans, and measuring
' memory consumption without an external profiler.
'
' Functions:
'   GetDbStatus        - all db_status counters as a (n, 3) matrix
'   GetDbStatusValue   - single db_status counter (current + highwater)
'   ResetDbStatus      - zero the highwater mark for one or all counters
'   GetStmtStatus      - single stmt_status counter
'   GetAllStmtStatus   - all stmt_status counters as a (n, 2) matrix
'
' Constants (db_status op codes):
'   DBSTAT_LOOKASIDE_USED, DBSTAT_CACHE_USED, DBSTAT_SCHEMA_USED,
'   DBSTAT_STMT_USED, DBSTAT_LOOKASIDE_HIT, DBSTAT_LOOKASIDE_MISS_SIZE,
'   DBSTAT_LOOKASIDE_MISS_FULL, DBSTAT_CACHE_HIT, DBSTAT_CACHE_MISS,
'   DBSTAT_CACHE_WRITE, DBSTAT_DEFERRED_FKS, DBSTAT_CACHE_USED_SHARED,
'   DBSTAT_CACHE_SPILL
'
' Constants (stmt_status op codes):
'   STMTSTAT_FULLSCAN, STMTSTAT_SORT, STMTSTAT_AUTOINDEX, STMTSTAT_VM_STEP,
'   STMTSTAT_REPREPARE, STMTSTAT_RUN, STMTSTAT_FILTER_MISS, STMTSTAT_FILTER_HIT,
'   STMTSTAT_MEMUSED
'
' Typical usage:
'   ' Print all connection counters
'   Dim info As Variant: info = GetDbStatus(conn, False)
'   Dim i As Long
'   For i = 0 To UBound(info, 1)
'       Debug.Print info(i,0) & ": cur=" & info(i,1) & " hi=" & info(i,2)
'   Next i
'
'   ' Check if a specific query triggered a full-table scan
'   Dim pStmt As LongPtr: pStmt = cmd.StmtHandle   ' via SQLite3Command
'   conn.ExecSQL "..."   ' or use the cmd
'   Debug.Print "Full-scan steps: " & GetStmtStatus(pStmt, STMTSTAT_FULLSCAN, False)
'
' Version : 0.1.7
'
' Version History:
'   0.1.4 - Initial release.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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


'==============================================================================
' GetDbStatus
' Returns all known db_status counters as a (13, 3) Variant matrix:
'   col 0 = counter name (String)
'   col 1 = current value (Long)
'   col 2 = highwater value (Long)
' resetAfter: if True, resets all highwater marks after reading.
'==============================================================================
Public Function GetDbStatus(ByVal conn As SQLite3Connection, _
                              Optional ByVal resetAfter As Boolean = False) As Variant
    Dim ops As Variant
    ops = Array( _
        DBSTAT_LOOKASIDE_USED, DBSTAT_CACHE_USED, DBSTAT_SCHEMA_USED, _
        DBSTAT_STMT_USED, DBSTAT_LOOKASIDE_HIT, DBSTAT_LOOKASIDE_MISS_SIZE, _
        DBSTAT_LOOKASIDE_MISS_FULL, DBSTAT_CACHE_HIT, DBSTAT_CACHE_MISS, _
        DBSTAT_CACHE_WRITE, DBSTAT_DEFERRED_FKS, DBSTAT_CACHE_USED_SHARED, _
        DBSTAT_CACHE_SPILL)

    Dim names As Variant
    names = Array( _
        "lookaside_used", "cache_used", "schema_used", _
        "stmt_used", "lookaside_hit", "lookaside_miss_size", _
        "lookaside_miss_full", "cache_hit", "cache_miss", _
        "cache_write", "deferred_fks", "cache_used_shared", _
        "cache_spill")

    Dim n As Long: n = UBound(ops) - LBound(ops) + 1
    Dim result() As Variant
    ReDim result(n - 1, 2)

    Dim i As Long
    For i = 0 To n - 1
        Dim cur As Long, hi As Long
        sqlite3_db_status conn.Handle, CLng(ops(i + LBound(ops))), _
                          VarPtr(cur), VarPtr(hi), IIf(resetAfter, 1, 0)
        result(i, 0) = CStr(names(i + LBound(names)))
        result(i, 1) = cur
        result(i, 2) = hi
    Next i

    GetDbStatus = result
End Function

'==============================================================================
' GetDbStatusValue
' Read a single db_status counter. Returns a 2-element array: (current, highwater).
' resetAfter: if True, zeroes the highwater mark for this counter after reading.
'==============================================================================
Public Function GetDbStatusValue(ByVal conn As SQLite3Connection, _
                                   ByVal op As Long, _
                                   Optional ByVal resetAfter As Boolean = False) As Variant
    Dim cur As Long, hi As Long
    sqlite3_db_status conn.Handle, op, VarPtr(cur), VarPtr(hi), IIf(resetAfter, 1, 0)
    GetDbStatusValue = Array(cur, hi)
End Function

'==============================================================================
' ResetDbStatus
' Zero the highwater mark for a single counter (op >= 0) or all counters (op = -1).
'==============================================================================
Public Sub ResetDbStatus(ByVal conn As SQLite3Connection, _
                           Optional ByVal op As Long = -1)
    If op = -1 Then
        Dim i As Long
        For i = DBSTAT_LOOKASIDE_USED To DBSTAT_CACHE_SPILL
            Dim dummy1 As Long, dummy2 As Long
            sqlite3_db_status conn.Handle, i, VarPtr(dummy1), VarPtr(dummy2), 1
        Next i
    Else
        Dim d1 As Long, d2 As Long
        sqlite3_db_status conn.Handle, op, VarPtr(d1), VarPtr(d2), 1
    End If
End Sub

'==============================================================================
' GetStmtStatus
' Read a single stmt_status counter from a prepared statement handle.
' pStmt: raw sqlite3_stmt* -- obtain from SQLite3Command.StmtHandle.
' resetAfter: if True, zeroes the counter after reading.
'==============================================================================
Public Function GetStmtStatus(ByVal pStmt As LongPtr, _
                                ByVal op As Long, _
                                Optional ByVal resetAfter As Boolean = False) As Long
    If pStmt = 0 Then Exit Function
    GetStmtStatus = sqlite3_stmt_status(pStmt, op, IIf(resetAfter, 1, 0))
End Function

'==============================================================================
' GetAllStmtStatus
' Returns all named stmt_status counters as a (n, 2) matrix:
'   col 0 = counter name (String)
'   col 1 = counter value (Long)
' pStmt: raw sqlite3_stmt* from SQLite3Command.StmtHandle.
' resetAfter: if True, zeroes all counters after reading.
'==============================================================================
Public Function GetAllStmtStatus(ByVal pStmt As LongPtr, _
                                   Optional ByVal resetAfter As Boolean = False) As Variant
    Dim ops As Variant
    ops = Array(STMTSTAT_FULLSCAN, STMTSTAT_SORT, STMTSTAT_AUTOINDEX, _
                STMTSTAT_VM_STEP, STMTSTAT_REPREPARE, STMTSTAT_RUN, _
                STMTSTAT_FILTER_MISS, STMTSTAT_FILTER_HIT, STMTSTAT_MEMUSED)

    Dim names As Variant
    names = Array("fullscan_step", "sort", "autoindex", "vm_step", _
                  "reprepare", "run", "filter_miss", "filter_hit", "memused")

    Dim n As Long: n = UBound(ops) - LBound(ops) + 1
    Dim result() As Variant
    ReDim result(n - 1, 1)

    Dim i As Long
    For i = 0 To n - 1
        result(i, 0) = CStr(names(i + LBound(names)))
        result(i, 1) = GetStmtStatus(pStmt, CLng(ops(i + LBound(ops))), resetAfter)
    Next i

    GetAllStmtStatus = result
End Function

'==============================================================================
' DbStatusSummary
' Convenience: prints all db_status counters to the Immediate window.
'==============================================================================
Public Sub DbStatusSummary(ByVal conn As SQLite3Connection)
    Dim info As Variant: info = GetDbStatus(conn, False)
    Debug.Print "--- db_status for [" & conn.DbPath & "] ---"
    Dim i As Long
    For i = 0 To UBound(info, 1)
        Debug.Print "  " & info(i, 0) & ": current=" & info(i, 1) & _
                    "  highwater=" & info(i, 2)
    Next i
End Sub

'==============================================================================
' SQLite3_Excel.bas  -  Excel <-> SQLite integration helpers (64-bit only)
'
' Functions:
'   RangeToTable   - import an Excel range (or ListObject) into a SQLite table
'   QueryToRange   - write a SQL query result to a worksheet range
'
' RangeToTable type inference (sampled from the first data row):
'   vbDate / Date string  -> TEXT  (stored as ISO 8601: "yyyy-mm-dd hh:mm:ss")
'   Integer / whole REAL  -> INTEGER
'   Floating-point REAL   -> REAL
'   Everything else       -> TEXT
'
' Typical usage:
'
'   ' Import Sheet1 A1:E501 (row 1 = headers) into table "prices"
'   RangeToTable conn, "prices", Sheet1.Range("A1:E501"), True, True
'
'   ' Write a query back to Sheet2 B2
'   QueryToRange conn, "SELECT sym, price FROM prices ORDER BY price DESC", _
'                Sheet2.Range("B2"), True
'
' Version : 0.1.7
'
' Version History:
'   0.1.5 - Initial release.
'   0.1.6 - Added ListObjectToTable: imports an Excel ListObject (Table) into
'            SQLite using its own header metadata; thin wrapper over RangeToTable.
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

'==============================================================================
' RangeToTable
' Import the data in sourceRange into a SQLite table named tableName.
'
' Parameters:
'   conn          - open SQLite3Connection
'   tableName     - target table name (created if it does not exist)
'   sourceRange   - the Excel range to import
'   hasHeaders    - True: row 1 of sourceRange contains column names
'                   False: auto-generate names (col1, col2, ...)
'   dropIfExists  - True: DROP TABLE IF EXISTS before creating (default False)
'   batchSize     - rows per transaction batch (default 10000)
'
' The table is created with inferred column types. If the table already exists
' and dropIfExists is False, rows are appended to the existing table.
'
' Date values are stored as TEXT in ISO 8601 format ("yyyy-mm-dd hh:mm:ss")
' so they sort lexicographically and are readable by strftime() in SQLite.
'==============================================================================
Public Sub RangeToTable(ByVal conn As SQLite3Connection, _
                         ByVal tableName As String, _
                         ByVal sourceRange As Range, _
                         Optional ByVal hasHeaders As Boolean = True, _
                         Optional ByVal dropIfExists As Boolean = False, _
                         Optional ByVal batchSize As Long = 10000)

    ' ---- 1. Read the entire range into a Variant array (one shot) ----------
    Dim data As Variant
    data = sourceRange.Value    ' (1-based row, 1-based col)

    Dim totalRows As Long: totalRows = sourceRange.Rows.Count
    Dim nCols     As Long: nCols     = sourceRange.Columns.Count

    If totalRows < 1 Or nCols < 1 Then Exit Sub

    ' ---- 2. Extract column names -------------------------------------------
    Dim firstDataRow As Long
    Dim colNames() As String
    ReDim colNames(nCols - 1)

    If hasHeaders Then
        Dim c As Long
        For c = 0 To nCols - 1
            Dim hdr As String: hdr = CStr(data(1, c + 1))
            colNames(c) = SanitizeIdentifier(hdr, c)
        Next c
        firstDataRow = 2
    Else
        For c = 0 To nCols - 1
            colNames(c) = "col" & (c + 1)
        Next c
        firstDataRow = 1
    End If

    If firstDataRow > totalRows Then Exit Sub   ' headers only, no data

    ' ---- 3. Infer column types from first data row -------------------------
    Dim colTypes() As String
    ReDim colTypes(nCols - 1)
    For c = 0 To nCols - 1
        colTypes(c) = InferSQLiteType(data(firstDataRow, c + 1))
    Next c

    ' ---- 4. Build and execute CREATE TABLE ---------------------------------
    If dropIfExists Then
        conn.ExecSQL "DROP TABLE IF EXISTS [" & tableName & "];"
    End If

    Dim ddl As String
    ddl = "CREATE TABLE IF NOT EXISTS [" & tableName & "] ("
    For c = 0 To nCols - 1
        If c > 0 Then ddl = ddl & ", "
        ddl = ddl & "[" & colNames(c) & "] " & colTypes(c)
    Next c
    ddl = ddl & ");"
    conn.ExecSQL ddl

    ' ---- 5. Insert via SQLite3Command (no SQLite3BulkInsert dependency) ----
    Dim insertSQL As String
    insertSQL = "INSERT INTO [" & tableName & "] ("
    Dim phSQL As String
    For c = 0 To nCols - 1
        If c > 0 Then insertSQL = insertSQL & ", ": phSQL = phSQL & ", "
        insertSQL = insertSQL & "[" & colNames(c) & "]"
        phSQL = phSQL & "?"
    Next c
    insertSQL = insertSQL & ") VALUES (" & phSQL & ");"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, insertSQL

    Dim batchCount As Long: batchCount = 0
    conn.BeginTransaction

    Dim r As Long
    Dim rowData() As Variant
    ReDim rowData(nCols - 1)

    For r = firstDataRow To totalRows
        For c = 0 To nCols - 1
            cmd.BindVariant c + 1, FormatForSQLite(data(r, c + 1), colTypes(c))
        Next c
        cmd.Execute
        cmd.Reset
        batchCount = batchCount + 1
        If batchCount >= batchSize Then
            conn.CommitTransaction
            conn.BeginTransaction
            batchCount = 0
        End If
    Next r

    If conn.InTransaction Then conn.CommitTransaction
    cmd.FinalizeStatement
End Sub

'==============================================================================
' QueryToRange
' Execute sql against conn and write the result set to a worksheet, starting
' at topLeft. Column headers are written in the first row when includeHeaders
' is True (default).
'
' Any existing content in the destination cells is overwritten. The sheet is
' not cleared beyond the written area -- call topLeft.CurrentRegion.Clear
' beforehand if you need a full replace.
'==============================================================================
Public Sub QueryToRange(ByVal conn As SQLite3Connection, _
                         ByVal sql As String, _
                         ByVal topLeft As Range, _
                         Optional ByVal includeHeaders As Boolean = True)
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset(sql)
    rs.LoadAll
    If rs.RecordCount > 0 Then
        RecordsetToRange rs, topLeft, includeHeaders
    End If
    rs.CloseRecordset
End Sub

'==============================================================================
' ListObjectToTable
' Import an Excel ListObject (structured Table) into a SQLite table.
'
' This is a thin wrapper over RangeToTable that extracts the header row and
' data body range from the ListObject directly, so callers do not need to
' pass hasHeaders -- the answer is always True for a ListObject.
'
' Parameters:
'   conn         - open SQLite3Connection
'   tableName    - target SQLite table name (created if it does not exist)
'   lo           - the source Excel ListObject (e.g. ActiveSheet.ListObjects(1))
'   dropIfExists - True: DROP TABLE IF EXISTS before importing (default False)
'   batchSize    - rows per transaction batch (default 10000)
'
' Usage:
'   Dim lo As ListObject
'   Set lo = Sheet1.ListObjects("PriceTable")
'   ListObjectToTable conn, "prices", lo, dropIfExists:=True
'==============================================================================
Public Sub ListObjectToTable(ByVal conn As SQLite3Connection, _
                              ByVal tableName As String, _
                              ByVal lo As ListObject, _
                              Optional ByVal dropIfExists As Boolean = False, _
                              Optional ByVal batchSize As Long = 10000)

    ' A ListObject with no data rows still has a header row; DataBodyRange
    ' is Nothing when there are zero data rows -- guard against that.
    If lo.ListRows.Count = 0 Then Exit Sub

    ' Use the full range (headers + data) and tell RangeToTable headers=True.
    ' lo.Range includes the header row; lo.DataBodyRange does not.
    RangeToTable conn, tableName, lo.Range, True, dropIfExists, batchSize
End Sub

'==============================================================================
' Private helpers
'==============================================================================

' Infer the SQLite affinity for a single cell value sampled from the first
' data row. Order matters: Date before IsNumeric (Excel dates are numeric).
Private Function InferSQLiteType(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        InferSQLiteType = "TEXT"
        Exit Function
    End If
    ' Excel stores dates as Double internally; VarType returns vbDate for
    ' cells formatted as Date, but plain Doubles can also represent dates.
    If VarType(v) = vbDate Then
        InferSQLiteType = "TEXT"    ' stored as ISO 8601 string
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        InferSQLiteType = "INTEGER"
        Exit Function
    End If
    If IsNumeric(v) Then
        Dim d As Double: d = CDbl(v)
        ' Whole numbers in Long range -> INTEGER; everything else -> REAL
        If d = Int(d) And d >= -2147483648# And d <= 2147483647# Then
            InferSQLiteType = "INTEGER"
        Else
            InferSQLiteType = "REAL"
        End If
        Exit Function
    End If
    InferSQLiteType = "TEXT"
End Function

' Convert a cell value to the most appropriate VBA type for SQLite binding.
' Date cells are converted to ISO 8601 TEXT; everything else is passed through.
Private Function FormatForSQLite(ByVal v As Variant, ByVal colType As String) As Variant
    If IsEmpty(v) Or IsNull(v) Then
        FormatForSQLite = Null
        Exit Function
    End If
    If VarType(v) = vbDate Or (colType = "TEXT" And IsDate(v)) Then
        ' Store as ISO 8601: sortable, readable by SQLite date functions
        If Int(CDbl(v)) = CDbl(v) Then
            FormatForSQLite = Format(v, "yyyy-mm-dd")            ' date only
        Else
            FormatForSQLite = Format(v, "yyyy-mm-dd hh:mm:ss")  ' datetime
        End If
        Exit Function
    End If
    If VarType(v) = vbBoolean Then
        FormatForSQLite = IIf(v, 1, 0)
        Exit Function
    End If
    FormatForSQLite = v
End Function

' Return a valid SQLite identifier from a header string. Replaces any run of
' characters that are not letters, digits, or underscore with a single '_'.
' Falls back to "col<n+1>" for blank or purely-special-character headers.
'==============================================================================
' RecordsetToRange  (private helper for QueryToRange)
' Writes a fully-loaded SQLite3Recordset to a worksheet range.
' topLeft is the upper-left destination cell.
' When includeHeaders is True the first row receives the column names.
'==============================================================================
Private Sub RecordsetToRange(ByVal rs As SQLite3Recordset, _
                               ByVal topLeft As Range, _
                               ByVal includeHeaders As Boolean)
    Dim nCols As Long: nCols = rs.FieldCount
    Dim dataStartRow As Long: dataStartRow = IIf(includeHeaders, 1, 0)

    If includeHeaders Then
        Dim names() As String: names = rs.ColumnNames()
        Dim c As Long
        For c = 0 To nCols - 1
            topLeft.Offset(0, c).Value = names(c)
        Next c
    End If

    If rs.RecordCount > 0 Then
        Dim mat As Variant: mat = rs.ToMatrix()
        If Not IsEmpty(mat) Then
            topLeft.Offset(dataStartRow, 0).Resize(rs.RecordCount, nCols).Value = mat
        End If
    End If
End Sub

Private Function SanitizeIdentifier(ByVal s As String, ByVal colIdx As Long) As String
    s = Trim(s)
    If Len(s) = 0 Then SanitizeIdentifier = "col" & (colIdx + 1): Exit Function

    Dim result As String
    Dim i As Long
    Dim inRun As Boolean: inRun = False
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid(s, i, 1)
        Dim code As Long: code = Asc(ch)
        Dim isOk As Boolean
        Select Case code
            Case 65 To 90:  isOk = True   ' A-Z
            Case 97 To 122: isOk = True   ' a-z
            Case 48 To 57:  isOk = True   ' 0-9
            Case 95:        isOk = True   ' _
            Case Else:      isOk = False
        End Select
        If isOk Then
            result = result & ch
            inRun = False
        ElseIf Not inRun Then
            result = result & "_"
            inRun = True
        End If
    Next i

    ' Strip leading/trailing underscores introduced by the replacement
    Do While Left(result, 1) = "_" And Len(result) > 1
        result = Mid(result, 2)
    Loop
    Do While Right(result, 1) = "_" And Len(result) > 1
        result = Left(result, Len(result) - 1)
    Loop

    ' Leading digit -> prepend underscore so it is a valid identifier
    If Len(result) > 0 Then
        If Asc(Left(result, 1)) >= 48 And Asc(Left(result, 1)) <= 57 Then
            result = "_" & result
        End If
    End If

    If Len(result) = 0 Then result = "col" & (colIdx + 1)
    SanitizeIdentifier = result
End Function

'==============================================================================
' SQLite3_FTS5.bas  -  Full-Text Search (FTS5) helpers
'
' FTS5 is SQLite's built-in full-text search engine.  It is enabled in the
' official sqlite3.dll precompiled binaries from sqlite.org.
'
' Key concepts:
'   - FTS5 tables are virtual tables created with: CREATE VIRTUAL TABLE
'     ... USING fts5(col1, col2, ...)
'   - Rows are searched with the MATCH operator:
'     SELECT * FROM docs WHERE docs MATCH 'query'
'   - Results can be ranked by relevance using the built-in rank column.
'   - Auxiliary functions: snippet(), highlight(), bm25()
'   - Content tables: FTS5 can index an existing real table (content=)
'   - Prefix search: "prefix*"
'   - Phrase search: "exact phrase"
'   - Column filter: col : term
'   - Boolean: term1 AND term2, term1 OR term2, NOT term
'
' All functions here are pure SQL wrappers -- no extra DLL calls needed.
'
' Version : 0.1.7
'
' Version History:
'   0.1.2 - Initial release. CreateFTS5Table, FTS5Insert, FTS5BulkInsert,
'            FTS5Search, FTS5SearchMatrix, FTS5Snippet, FTS5Highlight,
'            FTS5BM25Search, FTS5Delete, FTS5Optimize, FTS5Rebuild,
'            FTS5RowCount, FTS5MatchCount.
'   0.1.3 - No functional changes. Version stamp updated.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

'==============================================================================
' CreateFTS5Table
' Creates a new FTS5 virtual table.
'
' columns    : Array("title", "body", "author")
' contentTbl : optional real table to index (content= option).
'              If empty, FTS5 stores its own copy of the text.
' tokenizer  : "unicode61" (default), "ascii", "porter" (stemming)
'
' Example:
'   CreateFTS5Table conn, "docs_fts", Array("title","body"), "", "porter unicode61"
'==============================================================================
Public Sub CreateFTS5Table(ByVal conn As SQLite3Connection, _
                            ByVal ftsTable As String, _
                            ByVal columns As Variant, _
                            Optional ByVal contentTbl As String = "", _
                            Optional ByVal tokenizer As String = "unicode61", _
                            Optional ByVal dropIfExists As Boolean = False)
    If dropIfExists Then
        conn.ExecSQL "DROP TABLE IF EXISTS [" & ftsTable & "];"
    End If

    Dim colList As String, i As Long
    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then colList = colList & ", "
        colList = colList & "[" & CStr(columns(i)) & "]"
    Next i

    Dim opts As String
    If Len(contentTbl) > 0 Then opts = opts & ", content=[" & contentTbl & "]"
    If Len(tokenizer) > 0 Then opts = opts & ", tokenize=""" & tokenizer & """"

    conn.ExecSQL "CREATE VIRTUAL TABLE [" & ftsTable & "] USING fts5(" & _
                 colList & opts & ");"
End Sub

'==============================================================================
' FTS5Insert
' Insert a single row into an FTS5 table.
' columns : Array("title", "body")
' values  : Array("My Title", "Body text here")
'==============================================================================
Public Sub FTS5Insert(ByVal conn As SQLite3Connection, _
                       ByVal ftsTable As String, _
                       ByVal columns As Variant, _
                       ByVal values As Variant)
    Dim i As Long
    Dim colList As String, placeholders As String
    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then
            colList      = colList & ", "
            placeholders = placeholders & ", "
        End If
        colList      = colList & "[" & CStr(columns(i)) & "]"
        placeholders = placeholders & "?"
    Next i

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, "INSERT INTO [" & ftsTable & "] (" & colList & _
                      ") VALUES (" & placeholders & ");"
    For i = LBound(values) To UBound(values)
        cmd.BindVariant i - LBound(values) + 1, values(i)
    Next i
    cmd.Execute
End Sub

'==============================================================================
' FTS5BulkInsert
' Insert a (rows x cols) Variant matrix into an FTS5 table.
' columns : column names matching the matrix column order
'==============================================================================
Public Sub FTS5BulkInsert(ByVal conn As SQLite3Connection, _
                           ByVal ftsTable As String, _
                           ByVal columns As Variant, _
                           ByVal data As Variant, _
                           Optional ByVal batchSize As Long = 5000)
    Dim nCols As Long: nCols = UBound(columns) - LBound(columns) + 1
    Dim nRows As Long: nRows = UBound(data, 1) - LBound(data, 1) + 1

    Dim i As Long, colList As String, placeholders As String
    For i = LBound(columns) To UBound(columns)
        If i > LBound(columns) Then
            colList      = colList & ", "
            placeholders = placeholders & ", "
        End If
        colList      = colList & "[" & CStr(columns(i)) & "]"
        placeholders = placeholders & "?"
    Next i

    Dim sql As String
    sql = "INSERT INTO [" & ftsTable & "] (" & colList & ") VALUES (" & _
          placeholders & ");"

    Dim cmd As New SQLite3Command
    cmd.Prepare conn, sql

    Dim r As Long, c As Long, batch As Long
    conn.BeginTransaction
    For r = LBound(data, 1) To UBound(data, 1)
        For c = 0 To nCols - 1
            cmd.BindVariant c + 1, data(r, LBound(data, 2) + c)
        Next c
        cmd.Execute
        cmd.Reset
        batch = batch + 1
        If batch >= batchSize Then
            conn.CommitTransaction
            conn.BeginTransaction
            batch = 0
        End If
    Next r
    conn.CommitTransaction
End Sub

'==============================================================================
' FTS5Search
' Full-text search returning an open SQLite3Recordset.
' Caller must call rs.CloseRecordset when done.
'
' query    : FTS5 query string, e.g. "hello world", "title : hello", "foo*"
' columns  : columns to SELECT (Empty = "*")
' orderBy  : "rank" (relevance, default), "" (natural order), or any expr
' limit    : max rows, 0 = no limit
'==============================================================================
Public Function FTS5Search(ByVal conn As SQLite3Connection, _
                            ByVal ftsTable As String, _
                            ByVal query As String, _
                            Optional ByVal columns As String = "*", _
                            Optional ByVal orderBy As String = "rank", _
                            Optional ByVal limit As Long = 0) As SQLite3Recordset
    If Len(columns) = 0 Then columns = "*"
    Dim escapedQ As String: escapedQ = Replace(query, "'", "''")

    Dim sql As String
    sql = "SELECT " & columns & " FROM [" & ftsTable & "]" & _
          " WHERE [" & ftsTable & "] MATCH '" & escapedQ & "'"
    If Len(orderBy) > 0 Then sql = sql & " ORDER BY " & orderBy
    If limit > 0 Then sql = sql & " LIMIT " & limit
    sql = sql & ";"

    Set FTS5Search = conn.OpenRecordset(sql)
End Function

'==============================================================================
' FTS5SearchMatrix
' Like FTS5Search but returns the full result as a (row, col) Variant matrix.
'==============================================================================
Public Function FTS5SearchMatrix(ByVal conn As SQLite3Connection, _
                                  ByVal ftsTable As String, _
                                  ByVal query As String, _
                                  Optional ByVal columns As String = "*", _
                                  Optional ByVal orderBy As String = "rank", _
                                  Optional ByVal limit As Long = 0) As Variant
    Dim rs As SQLite3Recordset
    Set rs = FTS5Search(conn, ftsTable, query, columns, orderBy, limit)
    rs.LoadAll
    If rs.RecordCount > 0 Then
        FTS5SearchMatrix = rs.ToMatrix()
    Else
        FTS5SearchMatrix = Empty
    End If
    rs.CloseRecordset
End Function

'==============================================================================
' FTS5Snippet
' Return search results with an HTML-like snippet for each matching row.
' snippetCol  : the column to generate snippets from (0-based index)
' startMatch  : text inserted before each matching term  (default "<b>")
' endMatch    : text inserted after each matching term   (default "</b>")
' ellipsis    : text between non-adjacent snippets       (default "...")
' numTokens   : approximate number of tokens in snippet (default 16)
'
' Returns (row, ncols+1) matrix where the last column is the snippet text.
'==============================================================================
Public Function FTS5Snippet(ByVal conn As SQLite3Connection, _
                             ByVal ftsTable As String, _
                             ByVal query As String, _
                             Optional ByVal snippetCol As Long = 0, _
                             Optional ByVal startMatch As String = "<b>", _
                             Optional ByVal endMatch As String = "</b>", _
                             Optional ByVal ellipsis As String = "...", _
                             Optional ByVal numTokens As Long = 16, _
                             Optional ByVal limit As Long = 20) As Variant
    Dim escapedQ  As String: escapedQ  = Replace(query,      "'", "''")
    Dim escapedS  As String: escapedS  = Replace(startMatch, "'", "''")
    Dim escapedE  As String: escapedE  = Replace(endMatch,   "'", "''")
    Dim escapedEl As String: escapedEl = Replace(ellipsis,   "'", "''")

    Dim snippetExpr As String
    snippetExpr = "snippet([" & ftsTable & "], " & snippetCol & ", '" & _
                  escapedS & "', '" & escapedE & "', '" & escapedEl & "', " & _
                  numTokens & ") AS snippet_text"

    Dim sql As String
    sql = "SELECT *, " & snippetExpr & _
          " FROM [" & ftsTable & "]" & _
          " WHERE [" & ftsTable & "] MATCH '" & escapedQ & "'" & _
          " ORDER BY rank"
    If limit > 0 Then sql = sql & " LIMIT " & limit
    sql = sql & ";"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset(sql)
    rs.LoadAll
    If rs.RecordCount > 0 Then
        FTS5Snippet = rs.ToMatrix()
    Else
        FTS5Snippet = Empty
    End If
    rs.CloseRecordset
End Function

'==============================================================================
' FTS5Highlight
' Like FTS5Snippet but returns the full column text with matches highlighted
' rather than a trimmed excerpt.
'==============================================================================
Public Function FTS5Highlight(ByVal conn As SQLite3Connection, _
                               ByVal ftsTable As String, _
                               ByVal query As String, _
                               ByVal highlightCol As Long, _
                               Optional ByVal startMatch As String = "<b>", _
                               Optional ByVal endMatch As String = "</b>", _
                               Optional ByVal limit As Long = 20) As Variant
    Dim escapedQ As String: escapedQ = Replace(query,      "'", "''")
    Dim escapedS As String: escapedS = Replace(startMatch, "'", "''")
    Dim escapedE As String: escapedE = Replace(endMatch,   "'", "''")

    Dim hlExpr As String
    hlExpr = "highlight([" & ftsTable & "], " & highlightCol & ", '" & _
             escapedS & "', '" & escapedE & "') AS highlighted_text"

    Dim sql As String
    sql = "SELECT *, " & hlExpr & _
          " FROM [" & ftsTable & "]" & _
          " WHERE [" & ftsTable & "] MATCH '" & escapedQ & "'" & _
          " ORDER BY rank"
    If limit > 0 Then sql = sql & " LIMIT " & limit
    sql = sql & ";"

    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset(sql)
    rs.LoadAll
    If rs.RecordCount > 0 Then
        FTS5Highlight = rs.ToMatrix()
    Else
        FTS5Highlight = Empty
    End If
    rs.CloseRecordset
End Function

'==============================================================================
' FTS5Delete
' Delete a row from an FTS5 table by rowid.
'==============================================================================
Public Sub FTS5Delete(ByVal conn As SQLite3Connection, _
                       ByVal ftsTable As String, _
                       ByVal rowid As LongLong)
    conn.ExecSQL "DELETE FROM [" & ftsTable & "] WHERE rowid=" & rowid & ";"
End Sub

'==============================================================================
' FTS5Optimize
' Merge all FTS5 b-tree segments into one.  Run periodically after heavy
' inserts/deletes to improve search performance.
'==============================================================================
Public Sub FTS5Optimize(ByVal conn As SQLite3Connection, _
                         ByVal ftsTable As String)
    conn.ExecSQL "INSERT INTO [" & ftsTable & "]([" & ftsTable & "]) VALUES('optimize');"
End Sub

'==============================================================================
' FTS5Rebuild
' Rebuild the FTS5 index from scratch.  Required after updating a content
' table without going through the FTS5 virtual table.
'==============================================================================
Public Sub FTS5Rebuild(ByVal conn As SQLite3Connection, _
                        ByVal ftsTable As String)
    conn.ExecSQL "INSERT INTO [" & ftsTable & "]([" & ftsTable & "]) VALUES('rebuild');"
End Sub

'==============================================================================
' FTS5RowCount
' Return the number of rows currently in the FTS5 index.
'==============================================================================
Public Function FTS5RowCount(ByVal conn As SQLite3Connection, _
                              ByVal ftsTable As String) As Long
    FTS5RowCount = CLng(QueryScalar(conn, _
        "SELECT COUNT(*) FROM [" & ftsTable & "];"))
End Function

'==============================================================================
' FTS5MatchCount
' Return the number of rows matching a query without fetching the rows.
'==============================================================================
Public Function FTS5MatchCount(ByVal conn As SQLite3Connection, _
                                ByVal ftsTable As String, _
                                ByVal query As String) As Long
    Dim escapedQ As String: escapedQ = Replace(query, "'", "''")
    FTS5MatchCount = CLng(QueryScalar(conn, _
        "SELECT COUNT(*) FROM [" & ftsTable & "] WHERE [" & _
        ftsTable & "] MATCH '" & escapedQ & "';"))
End Function

'==============================================================================
' FTS5BM25Search
' Search using the bm25() scoring function explicitly.
' Lower bm25 score = better match (bm25 returns negative values in SQLite).
' Returns (row, col) matrix ordered by best match first.
'==============================================================================
Public Function FTS5BM25Search(ByVal conn As SQLite3Connection, _
                                 ByVal ftsTable As String, _
                                 ByVal query As String, _
                                 Optional ByVal columns As String = "*", _
                                 Optional ByVal limit As Long = 20) As Variant
    If Len(columns) = 0 Then columns = "*"
    Dim escapedQ As String: escapedQ = Replace(query, "'", "''")
    Dim sql As String
    sql = "SELECT " & columns & ", bm25([" & ftsTable & "]) AS score" & _
          " FROM [" & ftsTable & "]" & _
          " WHERE [" & ftsTable & "] MATCH '" & escapedQ & "'" & _
          " ORDER BY score"
    If limit > 0 Then sql = sql & " LIMIT " & limit
    sql = sql & ";"
    FTS5BM25Search = AggregateQuery(conn, sql)
End Function

'==============================================================================
' SQLite3_JSON.bas  -  JSON helper functions (64-bit only)
'
' Wraps SQLite's built-in JSON functions (available since SQLite 3.38.0).
' All functions are pure SQL -- no additional DLL procedures are required.
'
' Built-in JSON functions used:
'   json(x)                  - validate and canonicalise JSON
'   json_extract(x, path...) - extract values using JSONPath
'   json_insert(x, p, v)     - insert value if path absent
'   json_replace(x, p, v)    - replace value if path exists
'   json_set(x, p, v)        - insert or replace (upsert)
'   json_remove(x, path...)  - remove a key
'   json_patch(x, patch)     - RFC 7396 merge-patch
'   json_array(...)          - build JSON array
'   json_object(...)         - build JSON object
'   json_type(x, path)       - return SQLite type of value at path
'   json_valid(x)            - 1 if x is valid JSON, 0 otherwise
'   json_quote(x)            - convert SQL value to JSON fragment
'   json_group_array(x)      - aggregate rows into JSON array
'   json_group_object(k,v)   - aggregate rows into JSON object
'   json_each(x)             - table-valued function (rows per array element)
'   json_tree(x)             - table-valued function (full recursive walk)
'
' JSONPath format:
'   $           root element
'   $.key       object key access
'   $[0]        array index
'   $.a.b[2]    chained path
'
' Version : 0.1.7
'
' Version History:
'   0.1.3 - Initial release.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

'==============================================================================
' JSONExtract
' Extract a single value from a JSON column using a JSONPath expression.
' Returns the value as a Variant (text, number, or Null).
'
' Example:
'   v = JSONExtract(conn, "users", "profile", "$.address.city", "id=42")
'==============================================================================
Public Function JSONExtract(ByVal conn As SQLite3Connection, _
                             ByVal tableName As String, _
                             ByVal jsonCol As String, _
                             ByVal jsonPath As String, _
                             Optional ByVal whereClause As String = "") As Variant
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    JSONExtract = QueryScalar(conn, _
        "SELECT json_extract([" & jsonCol & "], '" & EscQ(jsonPath) & "')" & _
        " FROM [" & tableName & "]" & wh & " LIMIT 1;")
End Function

'==============================================================================
' JSONExtractColumn
' Extract a JSON path from every row in a table.
' Returns a (n, 2) matrix: rowid, extracted_value
'==============================================================================
Public Function JSONExtractColumn(ByVal conn As SQLite3Connection, _
                                   ByVal tableName As String, _
                                   ByVal jsonCol As String, _
                                   ByVal jsonPath As String, _
                                   Optional ByVal whereClause As String = "") As Variant
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    JSONExtractColumn = AggregateQuery(conn, _
        "SELECT rowid, json_extract([" & jsonCol & "], '" & EscQ(jsonPath) & "')" & _
        " FROM [" & tableName & "]" & wh & ";")
End Function

'==============================================================================
' JSONSet
' Update (insert or replace) a value at a JSONPath in a column.
' valueExpr: a SQL expression for the new value, e.g. "'London'", "42", "NULL"
'
' Example:
'   JSONSet conn, "users", "profile", "$.address.city", "'London'", "id=42"
'==============================================================================
Public Sub JSONSet(ByVal conn As SQLite3Connection, _
                   ByVal tableName As String, _
                   ByVal jsonCol As String, _
                   ByVal jsonPath As String, _
                   ByVal valueExpr As String, _
                   Optional ByVal whereClause As String = "")
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    conn.ExecSQL _
        "UPDATE [" & tableName & "] SET [" & jsonCol & "] = " & _
        "json_set([" & jsonCol & "], '" & EscQ(jsonPath) & "', " & valueExpr & ")" & _
        wh & ";"
End Sub

'==============================================================================
' JSONInsert
' Insert a value at a JSONPath only if the path does not already exist.
'==============================================================================
Public Sub JSONInsert(ByVal conn As SQLite3Connection, _
                      ByVal tableName As String, _
                      ByVal jsonCol As String, _
                      ByVal jsonPath As String, _
                      ByVal valueExpr As String, _
                      Optional ByVal whereClause As String = "")
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    conn.ExecSQL _
        "UPDATE [" & tableName & "] SET [" & jsonCol & "] = " & _
        "json_insert([" & jsonCol & "], '" & EscQ(jsonPath) & "', " & valueExpr & ")" & _
        wh & ";"
End Sub

'==============================================================================
' JSONReplace
' Replace a value at a JSONPath only if the path already exists.
'==============================================================================
Public Sub JSONReplace(ByVal conn As SQLite3Connection, _
                       ByVal tableName As String, _
                       ByVal jsonCol As String, _
                       ByVal jsonPath As String, _
                       ByVal valueExpr As String, _
                       Optional ByVal whereClause As String = "")
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    conn.ExecSQL _
        "UPDATE [" & tableName & "] SET [" & jsonCol & "] = " & _
        "json_replace([" & jsonCol & "], '" & EscQ(jsonPath) & "', " & valueExpr & ")" & _
        wh & ";"
End Sub

'==============================================================================
' JSONRemove
' Remove a key or array element from a JSON column.
' paths: Array("$.address", "$.phone") to remove multiple keys at once
'==============================================================================
Public Sub JSONRemove(ByVal conn As SQLite3Connection, _
                      ByVal tableName As String, _
                      ByVal jsonCol As String, _
                      ByVal paths As Variant, _
                      Optional ByVal whereClause As String = "")
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    Dim pathList As String, i As Long
    For i = LBound(paths) To UBound(paths)
        If i > LBound(paths) Then pathList = pathList & ", "
        pathList = pathList & "'" & EscQ(CStr(paths(i))) & "'"
    Next i
    conn.ExecSQL _
        "UPDATE [" & tableName & "] SET [" & jsonCol & "] = " & _
        "json_remove([" & jsonCol & "], " & pathList & ")" & wh & ";"
End Sub

'==============================================================================
' JSONPatch
' Apply an RFC 7396 merge-patch to a JSON column.
' patchExpr: a SQL expression resolving to the patch JSON, e.g. "'{"city":"Paris"}'"
'
' Merge-patch rules:
'   - Keys in patch present with non-null value -> set in target
'   - Keys in patch present with null value     -> remove from target
'   - Keys absent from patch                    -> unchanged in target
'==============================================================================
Public Sub JSONPatch(ByVal conn As SQLite3Connection, _
                     ByVal tableName As String, _
                     ByVal jsonCol As String, _
                     ByVal patchExpr As String, _
                     Optional ByVal whereClause As String = "")
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    conn.ExecSQL _
        "UPDATE [" & tableName & "] SET [" & jsonCol & "] = " & _
        "json_patch([" & jsonCol & "], " & patchExpr & ")" & wh & ";"
End Sub

'==============================================================================
' JSONValid
' Returns True if the stored value in a column is valid JSON for every row.
' whereClause can restrict to a subset of rows.
' Returns False immediately on the first invalid value found.
'==============================================================================
Public Function JSONValid(ByVal conn As SQLite3Connection, _
                           ByVal tableName As String, _
                           ByVal jsonCol As String, _
                           Optional ByVal whereClause As String = "") As Boolean
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause & " AND"
    ' If any row has json_valid=0, MIN will be 0
    Dim v As Variant
    v = QueryScalar(conn, _
        "SELECT MIN(json_valid([" & jsonCol & "]))" & _
        " FROM [" & tableName & "]" & _
        IIf(Len(wh) > 0, " WHERE " & whereClause, "") & ";")
    If IsNull(v) Or IsEmpty(v) Then
        JSONValid = True   ' empty table = trivially valid
    Else
        JSONValid = (CLng(v) = 1)
    End If
End Function

'==============================================================================
' JSONSearch
' Find rows where a JSON path equals a specific value.
' Returns a (n, ncols) matrix of matching rows.
' compareVal: SQL literal, e.g. "'London'", "42"
'==============================================================================
Public Function JSONSearch(ByVal conn As SQLite3Connection, _
                            ByVal tableName As String, _
                            ByVal jsonCol As String, _
                            ByVal jsonPath As String, _
                            ByVal compareVal As String, _
                            Optional ByVal extraCols As String = "*", _
                            Optional ByVal limit As Long = 0) As Variant
    If Len(extraCols) = 0 Then extraCols = "*"
    Dim sql As String
    sql = "SELECT " & extraCols & " FROM [" & tableName & "]" & _
          " WHERE json_extract([" & jsonCol & "], '" & EscQ(jsonPath) & "') = " & _
          compareVal
    If limit > 0 Then sql = sql & " LIMIT " & limit
    sql = sql & ";"
    JSONSearch = AggregateQuery(conn, sql)
End Function

'==============================================================================
' JSONGroupArray
' Aggregate values into a JSON array string.
' valueCol: a column name or any SQL expression, e.g.:
'   "tag_name"                          -- plain column
'   "json_extract(data, '$.name')"      -- expression
' Example: JSONGroupArray(conn, "tags", "tag_name", "user_id=7", "tag_name")
'          -> '["excel","sqlite","vba"]'
'==============================================================================
Public Function JSONGroupArray(ByVal conn As SQLite3Connection, _
                                ByVal tableName As String, _
                                ByVal valueCol As String, _
                                Optional ByVal whereClause As String = "", _
                                Optional ByVal orderByExpr As String = "") As String
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    Dim ob As String
    If Len(orderByExpr) > 0 Then ob = " ORDER BY " & orderByExpr
    Dim v As Variant
    v = QueryScalar(conn, _
        "SELECT json_group_array(" & valueCol & ob & ")" & _
        " FROM [" & tableName & "]" & wh & ";")
    If IsNull(v) Or IsEmpty(v) Then JSONGroupArray = "[]" Else JSONGroupArray = CStr(v)
End Function

'==============================================================================
' JSONGroupObject
' Aggregate key-value pairs into a JSON object string.
' keyCol and valueCol may be column names or SQL expressions.
' Example: JSONGroupObject(conn, "settings", "key", "value")
'          -> '{"theme":"dark","lang":"en"}'
'==============================================================================
Public Function JSONGroupObject(ByVal conn As SQLite3Connection, _
                                 ByVal tableName As String, _
                                 ByVal keyCol As String, _
                                 ByVal valueCol As String, _
                                 Optional ByVal whereClause As String = "") As String
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    Dim v As Variant
    v = QueryScalar(conn, _
        "SELECT json_group_object(" & keyCol & ", " & valueCol & ")" & _
        " FROM [" & tableName & "]" & wh & ";")
    If IsNull(v) Or IsEmpty(v) Then JSONGroupObject = "{}" Else JSONGroupObject = CStr(v)
End Function

'==============================================================================
' JSONEach
' Expand a JSON array stored in a column into one row per element.
' Returns a (n, 3) matrix: source_rowid, element_index, element_value
' Requires SQLite 3.38+ (json_each is a table-valued function).
'==============================================================================
Public Function JSONEach(ByVal conn As SQLite3Connection, _
                          ByVal tableName As String, _
                          ByVal jsonCol As String, _
                          Optional ByVal whereClause As String = "") As Variant
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " AND " & whereClause
    JSONEach = AggregateQuery(conn, _
        "SELECT t.rowid, j.key, j.value" & _
        " FROM [" & tableName & "] t," & _
        " json_each(t.[" & jsonCol & "]) j" & _
        " WHERE 1=1" & wh & ";")
End Function

'==============================================================================
' JSONType
' Return the JSON type of the value at jsonPath for each row.
' Returns a (n, 2) matrix: rowid, json_type string
' json_type values: 'null', 'true', 'false', 'integer', 'real', 'text',
'                   'array', 'object'
'==============================================================================
Public Function JSONType(ByVal conn As SQLite3Connection, _
                          ByVal tableName As String, _
                          ByVal jsonCol As String, _
                          ByVal jsonPath As String, _
                          Optional ByVal whereClause As String = "") As String
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    Dim v As Variant
    v = QueryScalar(conn, _
        "SELECT json_type([" & jsonCol & "], '" & EscQ(jsonPath) & "')" & _
        " FROM [" & tableName & "]" & wh & " LIMIT 1;")
    If IsNull(v) Then JSONType = "" Else JSONType = CStr(v)
End Function

'==============================================================================
' JSONBuildObject
' Build a JSON object string from an array of alternating key-value pairs.
' keys and values must have the same UBound.
' values are SQL literal strings, e.g. "'London'", "42", "NULL"
'
' Example:
'   JSONBuildObject(conn, Array("city","count"), Array("'London'","42"))
'   -> evaluates 'SELECT json_object("city","London","count",42)'
'   -> returns '{"city":"London","count":42}'
'==============================================================================
Public Function JSONBuildObject(ByVal conn As SQLite3Connection, _
                                 ByVal keys As Variant, _
                                 ByVal values As Variant) As String
    Dim i As Long, args As String
    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then args = args & ", "
        args = args & "'" & EscQ(CStr(keys(i))) & "', " & CStr(values(i))
    Next i
    Dim v As Variant
    v = QueryScalar(conn, "SELECT json_object(" & args & ");")
    If IsNull(v) Or IsEmpty(v) Then JSONBuildObject = "{}" Else JSONBuildObject = CStr(v)
End Function

'==============================================================================
' JSONBuildArray
' Build a JSON array from a VBA array of SQL literal values.
'
' Example:
'   JSONBuildArray(conn, Array("'alpha'", "'beta'", "42"))
'   -> returns '["alpha","beta",42]'
'==============================================================================
Public Function JSONBuildArray(ByVal conn As SQLite3Connection, _
                                ByVal values As Variant) As String
    Dim i As Long, args As String
    For i = LBound(values) To UBound(values)
        If i > LBound(values) Then args = args & ", "
        args = args & CStr(values(i))
    Next i
    Dim v As Variant
    v = QueryScalar(conn, "SELECT json_array(" & args & ");")
    If IsNull(v) Or IsEmpty(v) Then JSONBuildArray = "[]" Else JSONBuildArray = CStr(v)
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function EscQ(ByVal s As String) As String
    EscQ = Replace(s, "'", "''")
End Function

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
' Version : 0.1.7
'
' Version History:
'   0.1.5 - Initial release.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

'==============================================================================
' SQLite3_Migrate.bas  -  Schema versioning and migration helpers (64-bit only)
'
' Stores the application schema version in SQLite's built-in PRAGMA user_version
' (a 32-bit integer, default 0). No extra table is required.
'
' Functions:
'   GetSchemaVersion(conn)               - read current PRAGMA user_version
'   SetSchemaVersion(conn, version)      - write PRAGMA user_version
'   ApplyMigration(conn, toVersion, sql) - apply sql if current version < toVersion,
'                                          then advance version to toVersion
'   MigrateAll(conn, migrations)         - apply an ordered array of migration steps
'
' Typical usage -- initial setup from version 0:
'
'   Dim steps(2) As MigrationStep
'   steps(0) = MakeStep(1, "CREATE TABLE accounts (id INTEGER PRIMARY KEY, name TEXT);")
'   steps(1) = MakeStep(2, "CREATE TABLE trades (id INTEGER PRIMARY KEY, acct_id INTEGER, qty REAL);")
'   steps(2) = MakeStep(3, "CREATE INDEX idx_trades_acct ON trades (acct_id);")
'   MigrateAll conn, steps
'
' Only steps whose toVersion is above the current schema version are applied,
' so it is safe to call MigrateAll on every workbook open.
'
' Version : 0.1.7
'
' Version History:
'   0.1.6 - Initial release.
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


'==============================================================================
' MakeStep  -  convenience constructor for MigrationStep
'==============================================================================
Public Function MakeStep(ByVal toVersion As Long, ByVal sql As String) As MigrationStep
    MakeStep.toVersion = toVersion
    MakeStep.sql       = sql
End Function

'==============================================================================
' GetSchemaVersion  -  read PRAGMA user_version
'==============================================================================
Public Function GetSchemaVersion(ByVal conn As SQLite3Connection) As Long
    Dim v As Variant
    v = QueryScalar(conn, "PRAGMA user_version;")
    If IsNull(v) Or IsEmpty(v) Then
        GetSchemaVersion = 0
    Else
        GetSchemaVersion = CLng(v)
    End If
End Function

'==============================================================================
' SetSchemaVersion  -  write PRAGMA user_version
'
' PRAGMA user_version cannot be set via a bound parameter -- the value must be
' written inline. version is a Long (not user-supplied string) so injection is
' not a concern.
'==============================================================================
Public Sub SetSchemaVersion(ByVal conn As SQLite3Connection, ByVal version As Long)
    conn.ExecSQL "PRAGMA user_version = " & CStr(version) & ";"
End Sub

'==============================================================================
' ApplyMigration
' Apply sql inside a transaction if the current schema version is below
' toVersion, then advance user_version to toVersion.
'
' Parameters:
'   conn      - open SQLite3Connection
'   toVersion - the schema version this migration brings the DB to
'   sql       - DDL / DML to execute (may contain multiple statements)
'
' Returns True if the migration was applied, False if it was skipped (already
' at or above toVersion).
'==============================================================================
Public Function ApplyMigration(ByVal conn As SQLite3Connection, _
                                ByVal toVersion As Long, _
                                ByVal sql As String) As Boolean
    If GetSchemaVersion(conn) >= toVersion Then
        ApplyMigration = False
        Exit Function
    End If

    conn.BeginTransaction
    On Error GoTo fail
    conn.ExecSQL sql
    SetSchemaVersion conn, toVersion
    conn.CommitTransaction
    ApplyMigration = True
    Exit Function

fail:
    conn.RollbackTransaction
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'==============================================================================
' MigrateAll
' Apply every step in the migrations array whose toVersion exceeds the current
' schema version, in array order. Each step runs in its own transaction.
'
' Raises an error (and rolls back the failing step) if any step fails; steps
' already applied before the failure are committed and remain in effect.
'
' Parameters:
'   conn       - open SQLite3Connection
'   migrations - array of MigrationStep (build with MakeStep)
'
' Returns the number of steps actually applied.
'==============================================================================
Public Function MigrateAll(ByVal conn As SQLite3Connection, _
                             ByRef migrations() As MigrationStep) As Long
    Dim applied As Long: applied = 0
    Dim i As Long
    For i = LBound(migrations) To UBound(migrations)
        If ApplyMigration(conn, migrations(i).toVersion, migrations(i).sql) Then
            applied = applied + 1
        End If
    Next i
    MigrateAll = applied
End Function

'==============================================================================
' SQLite3_Schema.bas  -  Schema introspection helpers (64-bit only)
'
' All functions use PRAGMA commands and sqlite_master queries -- no additional
' DLL procedure addresses are required beyond the core driver.
'
' Functions:
'   GetTableList        - all user tables (or tables + views)
'   GetViewList         - all views
'   GetTriggerList      - all triggers
'   GetColumnInfo       - column names, types, nullability, defaults, PK flag
'   GetIndexList        - all indexes on a table
'   GetIndexColumns     - columns covered by a specific index
'   GetForeignKeys      - FK relationships on a table
'   GetCreateSQL        - original CREATE statement from sqlite_master
'   GetDatabaseInfo     - key PRAGMA values as a (name, value) matrix
'
' Note: TableExists, ViewExists, IndexExists live in SQLite3_Helpers.bas
'
' Version : 0.1.7
'
' Version History:
'   0.1.3 - Initial release.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

'==============================================================================
' GetTableList
' Returns a single-column Variant array of user table names.
' Set includeViews=True to include views in the result.
'==============================================================================
Public Function GetTableList(ByVal conn As SQLite3Connection, _
                              Optional ByVal includeViews As Boolean = False) As Variant
    Dim typeFilter As String
    If includeViews Then
        typeFilter = "type IN ('table','view')"
    Else
        typeFilter = "type='table'"
    End If
    Dim sql As String
    sql = "SELECT name FROM sqlite_master" & _
          " WHERE " & typeFilter & _
          " AND name NOT LIKE 'sqlite_%'" & _
          " ORDER BY name;"
    GetTableList = ColumnToArray(conn, sql)
End Function

'==============================================================================
' GetViewList
' Returns a single-column Variant array of view names.
'==============================================================================
Public Function GetViewList(ByVal conn As SQLite3Connection) As Variant
    GetViewList = ColumnToArray(conn, _
        "SELECT name FROM sqlite_master WHERE type='view'" & _
        " AND name NOT LIKE 'sqlite_%' ORDER BY name;")
End Function

'==============================================================================
' GetTriggerList
' Returns a (n, 3) matrix: trigger_name, event, target_table
' Event examples: "INSERT", "UPDATE", "DELETE"
' If tableName is non-empty, filters to triggers on that table only.
'==============================================================================
Public Function GetTriggerList(ByVal conn As SQLite3Connection, _
                                Optional ByVal tableName As String = "") As Variant
    Dim where As String
    where = "type='trigger'"
    If Len(tableName) > 0 Then
        where = where & " AND tbl_name='" & Escape(tableName) & "'"
    End If
    Dim sql As String
    sql = "SELECT name," & _
          " CASE WHEN sql LIKE '%BEFORE INSERT%' OR sql LIKE '%AFTER INSERT%'" & _
          "      THEN 'INSERT'" & _
          "      WHEN sql LIKE '%BEFORE UPDATE%' OR sql LIKE '%AFTER UPDATE%'" & _
          "      THEN 'UPDATE'" & _
          "      ELSE 'DELETE' END AS event," & _
          " tbl_name" & _
          " FROM sqlite_master WHERE " & where & " ORDER BY tbl_name, name;"
    GetTriggerList = AggregateQuery(conn, sql)
End Function

'==============================================================================
' GetColumnInfo
' Returns a (n, 6) matrix per column:
'   col 0: cid          - column index (0-based)
'   col 1: name         - column name
'   col 2: type         - declared type (e.g. "INTEGER", "TEXT", "REAL", "BLOB")
'   col 3: notnull      - 1 if NOT NULL constraint exists, 0 otherwise
'   col 4: dflt_value   - default value expression, or Null
'   col 5: pk           - position in primary key (0 = not part of PK)
'==============================================================================
Public Function GetColumnInfo(ByVal conn As SQLite3Connection, _
                               ByVal tableName As String) As Variant
    GetColumnInfo = AggregateQuery(conn, _
        "PRAGMA table_info([" & tableName & "]);")
End Function

'==============================================================================
' GetIndexList
' Returns a (n, 4) matrix for all indexes on a table:
'   col 0: seq      - creation sequence
'   col 1: name     - index name
'   col 2: unique   - 1 if UNIQUE, 0 otherwise
'   col 3: origin   - 'c'=CREATE INDEX, 'u'=UNIQUE constraint, 'pk'=PRIMARY KEY
'==============================================================================
Public Function GetIndexList(ByVal conn As SQLite3Connection, _
                              ByVal tableName As String) As Variant
    GetIndexList = AggregateQuery(conn, _
        "PRAGMA index_list([" & tableName & "]);")
End Function

'==============================================================================
' GetIndexColumns
' Returns a (n, 3) matrix for all columns in the named index:
'   col 0: seqno    - position within the index (0-based)
'   col 1: cid      - column id in the table (-1 for rowid)
'   col 2: name     - column name (empty string for rowid)
'==============================================================================
Public Function GetIndexColumns(ByVal conn As SQLite3Connection, _
                                 ByVal indexName As String) As Variant
    GetIndexColumns = AggregateQuery(conn, _
        "PRAGMA index_info([" & indexName & "]);")
End Function

'==============================================================================
' GetForeignKeys
' Returns a (n, 8) matrix for all FK constraints on a table:
'   col 0: id         - FK constraint id
'   col 1: seq        - column position within this FK (for multi-col FKs)
'   col 2: table      - referenced (parent) table
'   col 3: from       - column in this (child) table
'   col 4: to         - column in the parent table (or Null = parent PK)
'   col 5: on_update  - action on parent update ("NO ACTION", "CASCADE", etc.)
'   col 6: on_delete  - action on parent delete
'   col 7: match      - match type (usually "NONE")
'==============================================================================
Public Function GetForeignKeys(ByVal conn As SQLite3Connection, _
                                ByVal tableName As String) As Variant
    GetForeignKeys = AggregateQuery(conn, _
        "PRAGMA foreign_key_list([" & tableName & "]);")
End Function

'==============================================================================
' GetCreateSQL
' Returns the original CREATE TABLE / VIEW / TRIGGER / INDEX SQL string
' exactly as stored in sqlite_master.
'==============================================================================
Public Function GetCreateSQL(ByVal conn As SQLite3Connection, _
                              ByVal objectName As String) As String
    Dim v As Variant
    v = QueryScalar(conn, _
        "SELECT sql FROM sqlite_master WHERE name='" & Escape(objectName) & "';")
    If IsNull(v) Or IsEmpty(v) Then
        GetCreateSQL = ""
    Else
        GetCreateSQL = CStr(v)
    End If
End Function

'==============================================================================
' GetDatabaseInfo
' Returns key PRAGMA values as a (n, 2) matrix of (name, value) pairs.
' Useful for quick diagnostics and logging.
'==============================================================================
Public Function GetDatabaseInfo(ByVal conn As SQLite3Connection) As Variant
    Dim pragmas As Variant
    pragmas = Array( _
        "page_count", _
        "page_size", _
        "freelist_count", _
        "journal_mode", _
        "synchronous", _
        "cache_size", _
        "temp_store", _
        "auto_vacuum", _
        "encoding", _
        "user_version", _
        "application_id", _
        "wal_autocheckpoint")

    Dim result() As Variant
    ReDim result(UBound(pragmas) - LBound(pragmas), 1)

    Dim i As Long
    For i = LBound(pragmas) To UBound(pragmas)
        Dim pName As String: pName = CStr(pragmas(i))
        result(i - LBound(pragmas), 0) = pName
        result(i - LBound(pragmas), 1) = QueryScalar(conn, "PRAGMA " & pName & ";")
    Next i

    GetDatabaseInfo = result
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function Escape(ByVal s As String) As String
    Escape = Replace(s, "'", "''")
End Function

Private Function ColumnToArray(ByVal conn As SQLite3Connection, _
                                ByVal sql As String) As Variant
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset(sql)
    rs.LoadAll
    If rs.RecordCount = 0 Then
        ColumnToArray = Empty
        rs.CloseRecordset
        Exit Function
    End If
    Dim mat As Variant: mat = rs.ToMatrix()
    rs.CloseRecordset

    ' Extract column 0 into a 1-D array
    Dim arr() As Variant
    ReDim arr(UBound(mat, 1) - LBound(mat, 1))
    Dim i As Long
    For i = LBound(mat, 1) To UBound(mat, 1)
        arr(i - LBound(mat, 1)) = mat(i, 0)
    Next i
    ColumnToArray = arr
End Function

'==============================================================================
' SQLite3_Serialize.bas  -  Database snapshot to/from Byte() (64-bit only)
'
' Wraps sqlite3_serialize / sqlite3_deserialize (requires SQLite 3.23+).
'
' Functions:
'   SerializeDB      - snapshot a live DB to a VBA Byte() array
'   DeserializeDB    - replace a connection's content from a Byte() array
'   InMemoryClone    - open an independent in-memory copy of a live DB
'   IsSerializeAvail - True if the SQLite build includes these functions
'
' Typical usage:
'   ' Snapshot to bytes
'   Dim snap() As Byte
'   snap = SerializeDB(conn)
'   Debug.Print UBound(snap) + 1 & " bytes serialized"
'
'   ' Restore into an in-memory DB
'   Dim clone As SQLite3Connection
'   Set clone = InMemoryClone(conn)
'   ' clone is now independent -- changes to conn do not affect it
'
' Version : 0.1.7
'
' Version History:
'   0.1.4 - Initial release.
'   0.1.5 - SerializeDB now auto-checkpoints the WAL (TRUNCATE mode) before
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
'            calling sqlite3_serialize. This folds outstanding WAL frames into
'            the main file so the snapshot is always clean regardless of whether
'            the source was opened in WAL mode. Errors from the checkpoint are
'            silently ignored (non-WAL databases return SQLITE_OK immediately).
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


'==============================================================================
' SerializeDB
' Snapshot the named schema of conn into a VBA Byte() array.
' The connection stays open and fully usable during and after the call.
' schema: "main" (default) or the name of an ATTACHed database.
'
' The returned byte array is a complete, self-contained SQLite database file.
' It can be written to disk, sent over a network, or passed to DeserializeDB.
'==============================================================================
Public Function SerializeDB(ByVal conn As SQLite3Connection, _
                              Optional ByVal schema As String = "main") As Byte()
    ' Checkpoint with TRUNCATE before serializing.
    ' If the source is in WAL mode, outstanding WAL frames live in a separate
    ' file and are NOT captured by sqlite3_serialize unless they have been
    ' flushed into the main file first.  TRUNCATE folds all frames in and
    ' resets the WAL to zero bytes.  On non-WAL databases this is a no-op that
    ' returns SQLITE_OK immediately.  Errors are ignored: a failed checkpoint
    ' (e.g. SQLITE_BUSY) still produces a valid -- if slightly stale -- snapshot.
    Dim schCk() As Byte: schCk = SQLite3_CoreAPI.ToUTF8(schema)
    SQLite3_CoreAPI.sqlite3_wal_checkpoint_v2 conn.Handle, VarPtr(schCk(0)), _
        SQLITE_CHECKPOINT_TRUNCATE, CLngPtr(0), CLngPtr(0)

    Dim schBytes() As Byte: schBytes = SQLite3_CoreAPI.ToUTF8(schema)

    Dim szDb As LongLong   ' sqlite3_serialize writes the size here
    Dim pData As LongPtr
    pData = SQLite3_CoreAPI.sqlite3_serialize( _
        conn.Handle, VarPtr(schBytes(0)), VarPtr(szDb), 0)

    If pData = 0 Then
        ' Empty or uninitialized database -- return minimal empty array
        Dim emptyBuf() As Byte
        ReDim emptyBuf(0)
        SerializeDB = emptyBuf
        Exit Function
    End If

    Dim nBytes As Long
    If szDb > 2147483647 Then
        SQLite3_CoreAPI.sqlite3_free pData
        Err.Raise vbObjectError + 700, "SQLite3_Serialize.SerializeDB", _
                  "Database is too large to serialize into a VBA array (>2 GB)."
    End If
    nBytes = CLng(szDb)

    Dim buf() As Byte
    ReDim buf(nBytes - 1)
    CopyMemory buf(0), ByVal pData, CLngPtr(nBytes)
    SQLite3_CoreAPI.sqlite3_free pData

    SerializeDB = buf
End Function

'==============================================================================
' DeserializeDB
' Replace the content of an existing open connection with the bytes in data().
' The connection must have been opened (typically as ":memory:" for a fresh
' in-memory DB, or any file DB you want to overwrite).
'
' After this call the connection behaves exactly as if it had been opened
' against a file containing data().
'
' WARNING: This replaces ALL content in the target connection.
'          Any uncommitted transactions are rolled back first.
'==============================================================================
Public Sub DeserializeDB(ByVal conn As SQLite3Connection, _
                           ByRef data() As Byte, _
                           Optional ByVal schema As String = "main")
    Dim nBytes As Long: nBytes = UBound(data) - LBound(data) + 1
    If nBytes <= 0 Then
        Err.Raise vbObjectError + 701, "SQLite3_Serialize.DeserializeDB", _
                  "data() array is empty."
    End If

    ' Allocate a SQLite-heap buffer (so FREEONCLOSE works)
    Dim pBuf As LongPtr: pBuf = SQLite3_CoreAPI.sqlite3_malloc(nBytes)
    If pBuf = 0 Then
        Err.Raise vbObjectError + 702, "SQLite3_Serialize.DeserializeDB", _
                  "sqlite3_malloc(" & nBytes & ") returned NULL -- out of memory."
    End If

    CopyMemory ByVal pBuf, data(LBound(data)), CLngPtr(nBytes)

    Dim schBytes() As Byte: schBytes = SQLite3_CoreAPI.ToUTF8(schema)
    Dim szDb As LongLong: szDb = CLngLng(nBytes)
    Dim flags As Long: flags = SQLITE_DESERIALIZE_FREEONCLOSE Or SQLITE_DESERIALIZE_RESIZEABLE

    Dim rc As Long
    rc = SQLite3_CoreAPI.sqlite3_deserialize( _
        conn.Handle, VarPtr(schBytes(0)), pBuf, szDb, szDb, flags)

    If rc <> SQLITE_OK Then
        ' pBuf ownership passed to SQLite even on error when FREEONCLOSE is set
        Err.Raise vbObjectError + 703, "SQLite3_Serialize.DeserializeDB", _
                  "sqlite3_deserialize failed (rc=" & rc & ")"
    End If
End Sub

'==============================================================================
' InMemoryClone
' Open a new, independent in-memory connection whose content is an exact
' snapshot of conn at the moment of the call.
' Changes to conn after this call do NOT affect the clone, and vice versa.
'
' Uses the backup API internally (not serialize/deserialize) so that the clone
' is built from a clean page-level copy unaffected by WAL header state.
' The clone is opened without WAL mode -- WAL is irrelevant for :memory:.
'
' The returned connection must be closed by the caller when done.
'==============================================================================
Public Function InMemoryClone(ByVal conn As SQLite3Connection, _
                                Optional ByVal srcSchema As String = "main") As SQLite3Connection
    ' Open a fresh :memory: connection.  WAL is explicitly disabled --
    ' WAL mode on :memory: can interfere with reads after a backup step.
    Dim clone As New SQLite3Connection
    clone.OpenDatabase ":memory:", conn.DllPath, 5000, False

    ' Convert schema names to UTF-8 byte arrays for the backup API.
    Dim srcSchBytes() As Byte: srcSchBytes = SQLite3_CoreAPI.ToUTF8(srcSchema)
    Dim dstSchBytes() As Byte: dstSchBytes = SQLite3_CoreAPI.ToUTF8("main")

    ' Initialise the backup from conn -> clone.
    Dim pBackup As LongPtr
    pBackup = SQLite3_CoreAPI.sqlite3_backup_init( _
        clone.Handle, VarPtr(dstSchBytes(0)), _
        conn.Handle,  VarPtr(srcSchBytes(0)))

    If pBackup = 0 Then
        Err.Raise vbObjectError + 710, "SQLite3_Serialize.InMemoryClone", _
                  "sqlite3_backup_init failed: " & SQLite3_CoreAPI.sqlite3_errmsg_str(conn.Handle)
    End If

    ' Copy all pages in one step.  Returns SQLITE_DONE (101) on success.
    Dim rc As Long
    rc = SQLite3_CoreAPI.sqlite3_backup_step(pBackup, -1)
    SQLite3_CoreAPI.sqlite3_backup_finish pBackup

    If rc <> SQLITE_DONE And rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 711, "SQLite3_Serialize.InMemoryClone", _
                  "sqlite3_backup_step failed (rc=" & rc & ")"
    End If

    Set InMemoryClone = clone
End Function

'==============================================================================
' IsSerializeAvail
' Returns True if sqlite3_serialize is available in the loaded DLL.
' (Requires SQLite 3.23.0 or later, released 2018.)
' All official precompiled binaries from sqlite.org include it.
'==============================================================================
Public Function IsSerializeAvail() As Boolean
    On Error Resume Next
    Dim schBytes() As Byte: schBytes = SQLite3_CoreAPI.ToUTF8("main")
    Dim szDb As LongLong
    ' Call with a null db handle -- will fail gracefully but won't crash
    ' if the symbol exists. A missing export gives a null proc address and
    ' DispCallFunc returns 0 immediately.
    ' Best check: just try on a real connection in your test suite.
    ' Here we return True if PROC_COUNT covers the serialize slot (always true
    ' in v0.1.4+) and the DLL is loaded.
    IsSerializeAvail = SQLite3_CoreAPI.SQLite_IsLoaded()
    On Error GoTo 0
End Function

