Attribute VB_Name = "SQLite3_Excel"
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
' Version : 0.1.6
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
Option Explicit

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
    Dim firstDataRow As Long   ' 1-based index into data() of first data row
    Dim colNames() As String
    ReDim colNames(nCols - 1)

    If hasHeaders Then
        Dim c As Long
        For c = 0 To nCols - 1
            Dim hdr As String: hdr = CStr(data(1, c + 1))
            ' Sanitize: replace non-alphanumeric chars with underscore
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

    ' ---- 5. Bulk insert all data rows --------------------------------------
    Dim bulk As New SQLite3BulkInsert
    bulk.OpenInsert conn, tableName, colNames, batchSize

    Dim r As Long
    Dim rowData() As Variant
    ReDim rowData(nCols - 1)

    For r = firstDataRow To totalRows
        For c = 0 To nCols - 1
            rowData(c) = FormatForSQLite(data(r, c + 1), colTypes(c))
        Next c
        bulk.AppendRow rowData
    Next r

    bulk.CloseInsert
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
