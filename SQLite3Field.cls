VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Field.cls  -  ADO-style Field object (64-bit only)
' Zero-allocation: value reads directly against the live sqlite3_stmt*
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
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.'
'
'
'==============================================================================

Option Explicit

Private m_name    As String
Private m_ordinal As Long
Private m_pStmt   As LongPtr

'==============================================================================
' Internal factory (called by SQLite3Fields.BuildFields)
'==============================================================================
Friend Sub InitField(ByVal colName As String, _
                     ByVal Ordinal As Long, _
                     ByVal pStmt As LongPtr)
    m_name = colName
    m_ordinal = Ordinal
    m_pStmt = pStmt
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get name() As String
    name = m_name
End Property

Public Property Get Ordinal() As Long
    Ordinal = m_ordinal
End Property

' Value - zero-copy reads directly from SQLite memory
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    If m_pStmt = 0 Then Value = Null: Exit Property
    Select Case sqlite3_column_type(m_pStmt, m_ordinal)
        Case SQLITE_INTEGER
            Value = sqlite3_column_int64(m_pStmt, m_ordinal)
        Case SQLITE_FLOAT
            Value = sqlite3_column_double(m_pStmt, m_ordinal)
        Case SQLITE_TEXT
            Value = SQLite3_API.PtrToStringA(sqlite3_column_text(m_pStmt, m_ordinal))
        Case SQLITE_BLOB
            Value = ReadBlob()
        Case Else
            Value = Null
    End Select
End Property

' Typed accessors for hot paths (avoid Variant boxing overhead)
Public Property Get AsString() As String
    Dim p As LongPtr
    p = sqlite3_column_text(m_pStmt, m_ordinal)
    If p <> 0 Then AsString = SQLite3_API.PtrToStringA(p)
End Property

Public Property Get AsLong() As Long
    AsLong = sqlite3_column_int(m_pStmt, m_ordinal)
End Property

Public Property Get AsDouble() As Double
    AsDouble = sqlite3_column_double(m_pStmt, m_ordinal)
End Property

Public Property Get AsInt64() As LongLong
    AsInt64 = sqlite3_column_int64(m_pStmt, m_ordinal)
End Property

Public Property Get IsNull() As Boolean
    IsNull = (sqlite3_column_type(m_pStmt, m_ordinal) = SQLITE_NULL)
End Property

Public Property Get ColumnType() As Long
    ColumnType = sqlite3_column_type(m_pStmt, m_ordinal)
End Property

Public Property Get ByteLength() As Long
    ByteLength = sqlite3_column_bytes(m_pStmt, m_ordinal)
End Property

' Internal: refresh stmt pointer on MoveNext (called by SQLite3Fields)
Friend Sub RefreshStmt(ByVal pStmt As LongPtr)
    m_pStmt = pStmt
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function ReadBlob() As Variant
    Dim cb As Long
    cb = sqlite3_column_bytes(m_pStmt, m_ordinal)
    If cb = 0 Then ReadBlob = Empty: Exit Function
    Dim buf() As Byte
    ReDim buf(cb - 1)
    Dim p As LongPtr
    p = sqlite3_column_text(m_pStmt, m_ordinal)
    If p <> 0 Then
        Dim i As Long, b As Byte
        For i = 0 To cb - 1
            SQLite3_API.CopyMemoryByte b, p + i
            buf(i) = b
        Next i
    End If
    ReadBlob = buf
End Function
