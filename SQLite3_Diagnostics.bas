Attribute VB_Name = "SQLite3_Diagnostics"
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
' Version : 0.1.4
'
' Version History:
'   0.1.4 - Initial release.
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

'==============================================================================
' db_status op codes
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
' stmt_status op codes
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
