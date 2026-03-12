Attribute VB_Name = "SQLite3_Helpers"
'==============================================================================
' SQLite3_Helpers.bas  -  Public utility helpers (64-bit only)
' No Declare needed here: CopyMemoryByte lives in SQLite3_API.bas,
' and all SQLite wrappers are in SQLite3_API.bas.
'
' Version : 0.1.5
'
' Version History:
'   0.1.0 - Initial release. QueryScalar, TableExists, TableRowCount,
'            RecordsetToRange, BindParamIndex utilities.
'   0.1.1 - No functional changes.
'   0.1.2 - No functional changes. Version stamp updated.
'   0.1.3 - No functional changes. Version stamp updated.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - Added GetQueryPlan -- EXPLAIN QUERY PLAN wrapper returning a matrix.
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
' Named parameter index lookup
' Thin wrapper so callers never touch raw pointer arithmetic.
'==============================================================================
Public Function BindParamIndex(ByVal pStmt As LongPtr, _
                                ByVal pName As LongPtr) As Long
    BindParamIndex = sqlite3_bind_parameter_index(pStmt, pName)
End Function

'==============================================================================
' Excel integration helpers
'==============================================================================

' Dump a vectorized recordset to a worksheet in a single API call.
Public Sub RecordsetToRange(ByVal rs As SQLite3Recordset, _
                             ByVal topLeft As Range, _
                             Optional ByVal includeHeaders As Boolean = True)
    If Not rs.EOF And rs.RecordCount < 0 Then rs.LoadAll
    If rs.RecordCount = 0 Then Exit Sub

    Dim mat As Variant:    mat   = rs.ToMatrix
    Dim names() As String: names = rs.ColumnNames
    Dim nRows As Long:     nRows = rs.RecordCount
    Dim nCols As Long:     nCols = rs.FieldCount

    Dim startRow As Long: startRow = topLeft.Row
    Dim startCol As Long: startCol = topLeft.Column
    Dim ws As Worksheet:  Set ws  = topLeft.Worksheet

    If includeHeaders Then
        Dim hdr() As Variant
        ReDim hdr(0, nCols - 1)
        Dim c As Long
        For c = 0 To nCols - 1
            hdr(0, c) = names(c)
        Next c
        ws.Cells(startRow, startCol).Resize(1, nCols).Value = hdr
        startRow = startRow + 1
    End If

    ws.Cells(startRow, startCol).Resize(nRows, nCols).Value = mat
End Sub

' Execute a scalar query and return the first column of the first row.
Public Function QueryScalar(ByVal conn As SQLite3Connection, _
                             ByVal sql As String) As Variant
    Dim cmd As New SQLite3Command
    cmd.Prepare conn, sql
    QueryScalar = cmd.ExecuteScalar
End Function

' Return the row count for a table.
Public Function TableRowCount(ByVal conn As SQLite3Connection, _
                               ByVal tableName As String) As Long
    Dim v As Variant
    v = QueryScalar(conn, "SELECT COUNT(*) FROM [" & tableName & "];")
    If IsNull(v) Or IsEmpty(v) Then TableRowCount = 0 Else TableRowCount = CLng(v)
End Function

' Return True if a table exists in the schema.
Public Function TableExists(ByVal conn As SQLite3Connection, _
                             ByVal tableName As String) As Boolean
    TableExists = SchemaObjectExists(conn, tableName, "table")
End Function

' Return True if a view exists in the schema.
Public Function ViewExists(ByVal conn As SQLite3Connection, _
                            ByVal viewName As String) As Boolean
    ViewExists = SchemaObjectExists(conn, viewName, "view")
End Function

' Return True if an index exists in the schema.
Public Function IndexExists(ByVal conn As SQLite3Connection, _
                             ByVal indexName As String) As Boolean
    IndexExists = SchemaObjectExists(conn, indexName, "index")
End Function

Private Function SchemaObjectExists(ByVal conn As SQLite3Connection, _
                                     ByVal objName As String, _
                                     ByVal objType As String) As Boolean
    Dim v As Variant
    v = QueryScalar(conn, "SELECT COUNT(*) FROM sqlite_master " & _
                          "WHERE type='" & objType & "' AND name='" & _
                          Replace(objName, "'", "''") & "';")
    SchemaObjectExists = (Not IsNull(v)) And (CLng(v) > 0)
End Function

'==============================================================================
' GetQueryPlan
' Run EXPLAIN QUERY PLAN on sql and return the result as a Variant matrix.
'
' Returns a (nNodes x 4) matrix with columns:
'   (col 0) id        INTEGER  - node id within the plan
'   (col 1) parent    INTEGER  - parent node id (0 = root)
'   (col 2) notused   INTEGER  - always 0 in current SQLite
'   (col 3) detail    TEXT     - human-readable description of the plan step
'
' Returns Empty if the query produces no plan nodes (DDL statements, etc.).
'
' Usage:
'   Dim plan As Variant
'   plan = GetQueryPlan(conn, "SELECT * FROM orders WHERE customer_id = 42")
'   Dim i As Long
'   For i = 0 To UBound(plan, 1)
'       Debug.Print plan(i, 3)   ' detail column
'   Next i
'==============================================================================
Public Function GetQueryPlan(ByVal conn As SQLite3Connection, _
                               ByVal sql As String) As Variant
    Dim rs As SQLite3Recordset
    Set rs = conn.OpenRecordset("EXPLAIN QUERY PLAN " & sql)
    rs.LoadAll
    If rs.RecordCount = 0 Then
        GetQueryPlan = Empty
    Else
        GetQueryPlan = rs.ToMatrix
    End If
    rs.CloseRecordset
End Function
