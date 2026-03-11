VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3BulkInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3BulkInsert.cls  -  HFT-Grade Bulk Insert Engine (64-bit only)
'
' Architecture:
'   - Single prepared INSERT statement reused for every row
'   - Rows batched in explicit transactions (auto-commits every batchSize rows)
'   - No per-row VBA object allocation
'   - Achieves 100k+ rows/sec for simple schemas
'
' Usage:
'   Dim bulk As New SQLite3BulkInsert
'   bulk.OpenInsert conn, "trades", Array("symbol","price","qty","ts"), 10000
'   bulk.AppendRow Array("AAPL", 182.5, 100, Now())
'   bulk.Flush          ' commit final partial batch
'   bulk.CloseInsert
'
' Or feed a 2-D Variant array directly:
'   bulk.AppendMatrix data2D   ' (rowCount, colCount)
'
' Version : 0.1.4
'
' Version History:
'   0.1.0 - Initial release. AppendRow, AppendMatrix, transaction batching,
'            single reused prepared INSERT, TotalRowsInserted property.
'   0.1.1 - Added SQLite_IsLoaded() guard in CloseInsert.
'            Class_Terminate no longer calls native functions.
'            Fixed missing Set for conn assignment; UDT field copy fixed.
'   0.1.2 - BindVariant now handles byte-array (BLOB) columns.
'   0.1.3 - No functional changes. Version stamp updated.
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

Private m_conn        As SQLite3Connection
Private m_cmd         As SQLite3Command
Private m_tableName   As String
Private m_colNames()  As String
Private m_colCount    As Long
Private m_batchSize   As Long
Private m_rowsPending As Long
Private m_totalRows   As Long
Private m_isOpen      As Boolean
Private m_sql         As String

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Initialize()
    m_batchSize   = 10000
    m_rowsPending = 0
    m_totalRows   = 0
    m_isOpen      = False
End Sub

Private Sub Class_Terminate()
    ' Drop references only - native cleanup must be done by calling CloseInsert
    ' explicitly while the DLL is still loaded.
    Set m_cmd  = Nothing
    Set m_conn = Nothing
    m_isOpen   = False
End Sub

' colNames: 1-D array of column name strings
Public Sub OpenInsert(ByVal conn As SQLite3Connection, _
                      ByVal tableName As String, _
                      ByVal colNames As Variant, _
                      Optional ByVal batchSize As Long = 10000)

    Set m_conn  = conn
    m_tableName = tableName
    m_batchSize = batchSize
    m_colCount  = UBound(colNames) - LBound(colNames) + 1
    ReDim m_colNames(m_colCount - 1)

    Dim i As Long
    For i = 0 To m_colCount - 1
        m_colNames(i) = CStr(colNames(LBound(colNames) + i))
    Next i

    m_sql = BuildInsertSQL()

    Set m_cmd = New SQLite3Command
    m_cmd.Prepare conn, m_sql

    m_conn.BeginImmediate
    m_isOpen = True
End Sub

' Append a single row. rowData is a 1-D array of length colCount.
Public Sub AppendRow(ByVal rowData As Variant)
    AssertOpen
    Dim i As Long
    For i = 0 To m_colCount - 1
        m_cmd.BindVariant i + 1, rowData(LBound(rowData) + i)
    Next i
    m_cmd.Execute
    m_cmd.Reset
    m_rowsPending = m_rowsPending + 1
    m_totalRows   = m_totalRows + 1
    If m_rowsPending >= m_batchSize Then CommitBatch
End Sub

' Append a 2-D matrix (rowIdx, colIdx)
Public Sub AppendMatrix(ByVal data As Variant)
    AssertOpen
    Dim nRows As Long, nCols As Long
    nRows = UBound(data, 1) - LBound(data, 1) + 1
    nCols = UBound(data, 2) - LBound(data, 2) + 1
    If nCols <> m_colCount Then
        Err.Raise vbObjectError + 600, "SQLite3BulkInsert.AppendMatrix", _
                  "Column count mismatch. Expected " & m_colCount & ", got " & nCols
    End If
    Dim r As Long, c As Long
    Dim rBase As Long: rBase = LBound(data, 1)
    Dim cBase As Long: cBase = LBound(data, 2)
    For r = 0 To nRows - 1
        For c = 0 To m_colCount - 1
            m_cmd.BindVariant c + 1, data(rBase + r, cBase + c)
        Next c
        m_cmd.Execute
        m_cmd.Reset
        m_rowsPending = m_rowsPending + 1
        m_totalRows   = m_totalRows + 1
        If m_rowsPending >= m_batchSize Then CommitBatch
    Next r
End Sub

' Flush commits any pending rows without closing
Public Sub Flush()
    AssertOpen
    If m_rowsPending > 0 Then
        CommitBatch
        m_conn.BeginImmediate
    End If
End Sub

' Close commits the final batch and releases resources
Public Sub CloseInsert()
    If Not m_isOpen Then Exit Sub
    m_isOpen = False
    If Not SQLite3_API.SQLite_IsLoaded() Then
        Set m_cmd  = Nothing
        Set m_conn = Nothing
        Exit Sub
    End If
    If m_rowsPending > 0 Or m_conn.InTransaction Then
        m_conn.CommitTransaction
    End If
    If Not m_cmd Is Nothing Then
        m_cmd.FinalizeStatement
        Set m_cmd = Nothing
    End If
    Set m_conn = Nothing
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get TotalRowsInserted() As Long
    TotalRowsInserted = m_totalRows
End Property

Public Property Get BatchSize() As Long
    BatchSize = m_batchSize
End Property

Public Property Let BatchSize(ByVal val As Long)
    m_batchSize = val
End Property

Public Property Get SQL() As String
    SQL = m_sql
End Property

'==============================================================================
' Private helpers
'==============================================================================
Private Function BuildInsertSQL() As String
    Dim cols As String, ph As String
    Dim i As Long
    For i = 0 To m_colCount - 1
        If i > 0 Then cols = cols & ",": ph = ph & ","
        cols = cols & "[" & m_colNames(i) & "]"
        ph   = ph   & "?"
    Next i
    BuildInsertSQL = "INSERT INTO [" & m_tableName & "] (" & cols & ") VALUES (" & ph & ");"
End Function

Private Sub CommitBatch()
    m_conn.CommitTransaction
    m_rowsPending = 0
    m_conn.BeginImmediate
End Sub

Private Sub AssertOpen()
    If Not m_isOpen Then
        Err.Raise vbObjectError + 610, "SQLite3BulkInsert", _
                  "BulkInsert is not open. Call OpenInsert first."
    End If
End Sub
