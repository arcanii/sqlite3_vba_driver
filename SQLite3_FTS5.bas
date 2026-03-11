Attribute VB_Name = "SQLite3_FTS5"
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
' Version : 0.1.3
'
' Version History:
'   0.1.2 - Initial release. CreateFTS5Table, FTS5Insert, FTS5BulkInsert,
'            FTS5Search, FTS5SearchMatrix, FTS5Snippet, FTS5Highlight,
'            FTS5BM25Search, FTS5Delete, FTS5Optimize, FTS5Rebuild,
'            FTS5RowCount, FTS5MatchCount.
'   0.1.3 - No functional changes. Version stamp updated.
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
