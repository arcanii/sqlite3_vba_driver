Attribute VB_Name = "SQLite3_JSON"
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
' Version : 0.1.3
'
' Version History:
'   0.1.3 - Initial release.
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
                          Optional ByVal whereClause As String = "") As Variant
    Dim wh As String
    If Len(whereClause) > 0 Then wh = " WHERE " & whereClause
    JSONType = AggregateQuery(conn, _
        "SELECT rowid, json_type([" & jsonCol & "], '" & EscQ(jsonPath) & "')" & _
        " FROM [" & tableName & "]" & wh & ";")
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
