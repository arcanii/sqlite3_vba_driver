VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Field.cls  -  ADO-style Field object (64-bit only)
' Zero-allocation: value reads directly against the live sqlite3_stmt*
'
' Version : 0.1.2
'
' Version History:
'   0.1.0 - Initial release. Name, Ordinal, Value (default member), AsString,
'            AsLong, AsDouble, AsInt64, IsNull, ColumnType, ByteLength.
'   0.1.1 - No functional changes.
'   0.1.2 - Added AsBytes() typed accessor for BLOB columns.
'            ReadBlob() rewritten to use sqlite3_column_blob + BlobToBytes
'            (replaces slow per-byte CopyMemory loop).
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

Private m_name    As String
Private m_ordinal As Long
Private m_pStmt   As LongPtr

'==============================================================================
' Internal factory (called by SQLite3Fields.BuildFields)
'==============================================================================
Friend Sub InitField(ByVal colName As String, _
                     ByVal ordinal As Long, _
                     ByVal pStmt As LongPtr)
    m_name    = colName
    m_ordinal = ordinal
    m_pStmt   = pStmt
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get Name() As String
    Name = m_name
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

Public Property Get AsBytes() As Byte()
    Dim pBlob As LongPtr: pBlob = sqlite3_column_blob(m_pStmt, m_ordinal)
    Dim nBytes As Long:   nBytes = sqlite3_column_bytes(m_pStmt, m_ordinal)
    AsBytes = SQLite3_API.BlobToBytes(pBlob, nBytes)
End Property

' Internal: refresh stmt pointer on MoveNext (called by SQLite3Fields)
Friend Sub RefreshStmt(ByVal pStmt As LongPtr)
    m_pStmt = pStmt
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function ReadBlob() As Variant
    Dim pBlob  As LongPtr: pBlob  = sqlite3_column_blob(m_pStmt, m_ordinal)
    Dim nBytes As Long:    nBytes = sqlite3_column_bytes(m_pStmt, m_ordinal)
    If nBytes = 0 Then ReadBlob = Empty: Exit Function
    ReadBlob = SQLite3_API.BlobToBytes(pBlob, nBytes)
End Function
