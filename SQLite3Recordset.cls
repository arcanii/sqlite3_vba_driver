VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Recordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Recordset.cls  -  HFT-Grade ADO-style Recordset (64-bit only)
'
' Modes:
'   Live mode  : MoveNext reads rows one at a time from SQLite (default)
'   Vectorized : LoadAll() pulls all rows into a Variant matrix in one pass
'                (50x faster for large result sets - no per-row dispatch)
'
' ADO surface:
'   .BOF / .EOF                   navigation sentinels
'   .RecordCount                  exact after LoadAll, -1 in live mode
'   .MoveFirst / .MoveLast        vectorized mode only
'   .MoveNext / .MovePrevious     navigation
'   .Fields("name") / .Fields(n)  column access
'   rs!ColumnName                 shortcut via default member
'   .GetRows(n)                   return next n rows as 2-D Variant array
'   .CloseRecordset               release resources
'
' Version : 0.1.4
'
' Version History:
'   0.1.0 - Initial release. Live mode, LoadAll vectorized mode, GetRows,
'            ToMatrix, BOF/EOF/MoveNext/MoveLast, rs!Field default member.
'   0.1.1 - Fixed ReDim Preserve: flipped storage to (col, row) so last
'            dimension is rows. ToMatrix transposes to (row, col) for Excel.
'            LoadAll guards against sqlite3_column_count returning 0 on reset.
'            Class_Terminate no longer calls native functions.
'   0.1.2 - LoadAll stores SQLITE_BLOB columns as Byte() in Variant matrix
'            using sqlite3_column_blob + BlobToBytes.
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

' -- Live-mode state ----------------------------------------------------------
Private m_pStmt      As LongPtr
Private m_conn       As SQLite3Connection
Private m_sql        As String
Private m_stepResult As Long
Private m_colCount   As Long

' -- ADO navigation state -----------------------------------------------------
Private m_bof    As Boolean
Private m_eof    As Boolean
Private m_fields As SQLite3Fields

' -- Vectorized cache ---------------------------------------------------------
Private m_data()       As Variant
Private m_colNames()   As String
Private m_isVectorized As Boolean
Private m_rowCount     As Long
Private m_cursor       As Long
Private m_isOpen       As Boolean

'==============================================================================
' Internal open (called by SQLite3Connection.OpenRecordset)
'==============================================================================
Friend Sub OpenInternal(ByVal conn As SQLite3Connection, ByVal sql As String)
    Set m_conn     = conn
    m_sql          = sql
    m_isVectorized = False

    m_pStmt    = conn.GetCachedStatement(sql)
    m_colCount = sqlite3_column_count(m_pStmt)

    m_stepResult = sqlite3_step(m_pStmt)

    If m_stepResult = SQLITE_ROW Then
        Set m_fields = New SQLite3Fields
        m_fields.BuildFields m_pStmt, m_colCount
        m_bof = False
        m_eof = False
    ElseIf m_stepResult = SQLITE_DONE Then
        m_bof = True
        m_eof = True
    Else
        Err.Raise vbObjectError + 400, "SQLite3Recordset.OpenInternal", _
                  "sqlite3_step error " & m_stepResult & ": " & conn.LastError
    End If
    m_isOpen = True
End Sub

'==============================================================================
' Vectorized bulk load
'==============================================================================
Public Function LoadAll() As Long
    AssertOpen
    If m_isVectorized Then LoadAll = m_rowCount: Exit Function

    sqlite3_reset m_pStmt
    sqlite3_clear_bindings m_pStmt

    ' Use m_colCount already set by OpenInternal (column_count on a reset
    ' statement may return 0 on some SQLite builds before the first step)
    If m_colCount = 0 Then
        m_colCount = sqlite3_column_count(m_pStmt)
    End If
    If m_colCount = 0 Then
        m_isVectorized = True
        m_rowCount     = 0
        m_bof          = True
        m_eof          = True
        LoadAll        = 0
        Exit Function
    End If

    ReDim m_colNames(m_colCount - 1)

    ' Store as m_data(colIdx, rowIdx) so ReDim Preserve can grow/trim the
    ' last dimension (rows) without VBA's "cannot resize first dimension" error.
    Dim bufSize As Long: bufSize = 4096
    ReDim m_data(m_colCount - 1, bufSize - 1)
    m_rowCount = 0

    Dim rc As Long
    Dim c  As Long
    rc = sqlite3_step(m_pStmt)

    ' Capture column names from the first row
    If rc = SQLITE_ROW Then
        For c = 0 To m_colCount - 1
            m_colNames(c) = SQLite3_API.PtrToStringA( _
                                sqlite3_column_name(m_pStmt, c))
        Next c
    End If

    Do While rc = SQLITE_ROW
        If m_rowCount >= bufSize Then
            bufSize = bufSize * 2
            ReDim Preserve m_data(m_colCount - 1, bufSize - 1)
        End If
        Dim ct As Long
        For c = 0 To m_colCount - 1
            ct = sqlite3_column_type(m_pStmt, c)
            Select Case ct
                Case SQLITE_INTEGER
                    m_data(c, m_rowCount) = sqlite3_column_int64(m_pStmt, c)
                Case SQLITE_FLOAT
                    m_data(c, m_rowCount) = sqlite3_column_double(m_pStmt, c)
                Case SQLITE_TEXT
                    m_data(c, m_rowCount) = SQLite3_API.PtrToStringA( _
                                             sqlite3_column_text(m_pStmt, c))
                Case SQLITE_BLOB
                    Dim pBlob  As LongPtr: pBlob  = sqlite3_column_blob(m_pStmt, c)
                    Dim nBytes As Long:    nBytes = sqlite3_column_bytes(m_pStmt, c)
                    m_data(c, m_rowCount) = SQLite3_API.BlobToBytes(pBlob, nBytes)
                Case Else
                    m_data(c, m_rowCount) = Null
            End Select
        Next c
        m_rowCount = m_rowCount + 1
        rc = sqlite3_step(m_pStmt)
    Loop

    If rc <> SQLITE_DONE Then
        Err.Raise vbObjectError + 401, "SQLite3Recordset.LoadAll", _
                  "Step error " & rc
    End If

    ' Trim last dimension to exact row count
    If m_rowCount > 0 Then
        ReDim Preserve m_data(m_colCount - 1, m_rowCount - 1)
    End If

    m_isVectorized = True
    m_cursor       = 0
    m_bof          = (m_rowCount = 0)
    m_eof          = (m_rowCount = 0)
    LoadAll        = m_rowCount
End Function

'==============================================================================
' ADO navigation
'==============================================================================
Public Sub MoveFirst()
    If m_isVectorized Then
        If m_rowCount = 0 Then m_bof = True: m_eof = True: Exit Sub
        m_cursor = 0
        m_bof    = False
        m_eof    = False
    Else
        sqlite3_reset m_pStmt
        sqlite3_clear_bindings m_pStmt
        m_stepResult = sqlite3_step(m_pStmt)
        If m_stepResult = SQLITE_ROW Then
            m_bof = False: m_eof = False
            m_fields.RefreshStmtPointer m_pStmt
        Else
            m_bof = True: m_eof = True
        End If
    End If
End Sub

Public Sub MoveNext()
    AssertOpen
    If m_eof Then Exit Sub
    If m_isVectorized Then
        m_cursor = m_cursor + 1
        If m_cursor >= m_rowCount Then m_eof = True: m_cursor = m_rowCount
    Else
        m_stepResult = sqlite3_step(m_pStmt)
        If m_stepResult = SQLITE_ROW Then
            m_fields.RefreshStmtPointer m_pStmt
        ElseIf m_stepResult = SQLITE_DONE Then
            m_eof = True
        Else
            Err.Raise vbObjectError + 402, "SQLite3Recordset.MoveNext", _
                      "Step error " & m_stepResult
        End If
    End If
End Sub

Public Sub MovePrevious()
    AssertVectorized "MovePrevious"
    If m_cursor > 0 Then
        m_cursor = m_cursor - 1
        m_eof    = False
    Else
        m_bof = True
    End If
End Sub

Public Sub MoveLast()
    AssertVectorized "MoveLast"
    If m_rowCount = 0 Then Exit Sub
    m_cursor = m_rowCount - 1
    m_bof    = False
    m_eof    = False
End Sub

Public Sub Move(ByVal numRows As Long)
    AssertVectorized "Move"
    m_cursor = m_cursor + numRows
    If m_cursor < 0 Then m_cursor = 0: m_bof = True
    If m_cursor >= m_rowCount Then m_cursor = m_rowCount: m_eof = True
End Sub

'==============================================================================
' ADO-style properties
'==============================================================================
Public Property Get BOF() As Boolean
    BOF = m_bof
End Property

Public Property Get EOF() As Boolean
    EOF = m_eof
End Property

Public Property Get RecordCount() As Long
    If m_isVectorized Then RecordCount = m_rowCount Else RecordCount = -1
End Property

Public Property Get FieldCount() As Long
    FieldCount = m_colCount
End Property

Public Property Get Fields() As SQLite3Fields
    Set Fields = m_fields
End Property

' rs!ColumnName shortcut via default member
Public Property Get Item(ByVal nameOrIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    If m_isVectorized Then
        If m_eof Or m_bof Then
            Item = Null
        ElseIf VarType(nameOrIndex) = vbString Then
            Item = m_data(ColumnIndex(CStr(nameOrIndex)), m_cursor)
        Else
            Item = m_data(CLng(nameOrIndex), m_cursor)
        End If
    Else
        If Not m_fields Is Nothing Then
            If VarType(nameOrIndex) = vbString Then
                Item = m_fields.Item(CStr(nameOrIndex)).Value
            Else
                Item = m_fields.Item(nameOrIndex).Value
            End If
        Else
            Item = Null
        End If
    End If
End Property

'==============================================================================
' GetRows  -  return next n rows as a (cols x rows) Variant array (ADO compat)
'==============================================================================
Public Function GetRows(Optional ByVal numRows As Long = -1) As Variant
    If Not m_isVectorized Then LoadAll
    If m_eof Then GetRows = Empty: Exit Function

    Dim startRow As Long: startRow = m_cursor
    Dim endRow As Long
    If numRows < 0 Then
        endRow = m_rowCount - 1
    Else
        endRow = m_cursor + numRows - 1
        If endRow >= m_rowCount Then endRow = m_rowCount - 1
    End If

    Dim nRows As Long: nRows = endRow - startRow + 1
    ' ADO convention: result(col, row)
    Dim result() As Variant
    ReDim result(m_colCount - 1, nRows - 1)
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 0 To m_colCount - 1
            result(c, r) = m_data(c, startRow + r)
        Next c
    Next r
    m_cursor = endRow + 1
    If m_cursor >= m_rowCount Then m_eof = True
    GetRows = result
End Function

' Return data as (row, col) matrix suitable for direct Excel range assignment
Public Function ToMatrix() As Variant
    If Not m_isVectorized Then LoadAll
    If m_rowCount = 0 Then ToMatrix = Empty: Exit Function
    Dim result() As Variant
    ReDim result(m_rowCount - 1, m_colCount - 1)
    Dim r As Long, c As Long
    For r = 0 To m_rowCount - 1
        For c = 0 To m_colCount - 1
            result(r, c) = m_data(c, r)
        Next c
    Next r
    ToMatrix = result
End Function

Public Function ColumnNames() As String()
    ColumnNames = m_colNames
End Function

'==============================================================================
' Close and cleanup
'==============================================================================
Public Sub CloseRecordset()
    If SQLite3_API.SQLite_IsLoaded() Then
        If m_pStmt <> 0 Then
            sqlite3_reset m_pStmt
            sqlite3_clear_bindings m_pStmt
        End If
    End If
    Set m_fields = Nothing
    Set m_conn   = Nothing
    m_pStmt  = 0
    m_isOpen = False
End Sub

Private Sub Class_Terminate()
    ' Stmt is owned by the connection cache - do not call native functions here.
    m_pStmt  = 0
    m_isOpen = False
    Set m_fields = Nothing
    Set m_conn   = Nothing
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Sub AssertOpen()
    If Not m_isOpen Then
        Err.Raise vbObjectError + 410, "SQLite3Recordset", "Recordset is not open."
    End If
End Sub

Private Sub AssertVectorized(ByVal method As String)
    If Not m_isVectorized Then
        Err.Raise vbObjectError + 411, "SQLite3Recordset." & method, _
                  method & " requires LoadAll() to have been called first."
    End If
End Sub

Private Function ColumnIndex(ByVal name As String) As Long
    Dim i As Long
    For i = 0 To m_colCount - 1
        If StrComp(m_colNames(i), name, vbTextCompare) = 0 Then
            ColumnIndex = i: Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 412, "SQLite3Recordset.ColumnIndex", _
              "Column not found: " & name
End Function
