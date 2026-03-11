Attribute VB_Name = "SQLite3_Schema"
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
' Version : 0.1.4
'
' Version History:
'   0.1.3 - Initial release.
'   0.1.4 - No functional changes. Version stamp updated.
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
