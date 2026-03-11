Attribute VB_Name = "SQLite3_Aggregates"
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
' Version : 0.1.3
'
' Version History:
'   0.1.2 - Initial release. GroupByCount, GroupBySum, GroupByAvg,
'            ScalarAgg, MultiAgg, AggregateQuery, RunningTotal,
'            PercentileApprox, Histogram helpers.
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
