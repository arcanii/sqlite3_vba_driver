VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Fields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Fields.cls  -  ADO-style Fields collection (64-bit only)
' Enables: rs.Fields("Name").Value / rs.Fields(0).Value / rs!ColumnName
'
' Version : 0.1.5
'
' Version History:
'   0.1.0 - Initial release. Case-insensitive Dictionary name index,
'            For Each enumerator, Fields("name") and Fields(0) access.
'   0.1.1 - No functional changes.
'   0.1.2 - No functional changes. Version stamp updated.
'   0.1.3 - No functional changes. Version stamp updated.
'   0.1.4 - No functional changes. Version stamp updated.
'   0.1.5 - No functional changes. Version stamp updated.
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

Private m_fields()  As SQLite3Field
Private m_count     As Long
Private m_nameIndex As Object   ' Scripting.Dictionary for O(1) name lookup

'==============================================================================
' Internal builder (called by SQLite3Recordset after first step)
'==============================================================================
Friend Sub BuildFields(ByVal pStmt As LongPtr, ByVal colCount As Long)
    m_count = colCount
    ReDim m_fields(colCount - 1)
    Set m_nameIndex = CreateObject("Scripting.Dictionary")
    m_nameIndex.CompareMode = 1     ' vbTextCompare - case-insensitive

    Dim i As Long
    For i = 0 To colCount - 1
        Set m_fields(i) = New SQLite3Field
        Dim colName As String
        colName = SQLite3_API.PtrToStringA(sqlite3_column_name(pStmt, i))
        m_fields(i).InitField colName, i, pStmt
        m_nameIndex(colName) = i
    Next i
End Sub

' Refresh all field stmt pointers after MoveNext
Friend Sub RefreshStmtPointer(ByVal pStmt As LongPtr)
    Dim i As Long
    For i = 0 To m_count - 1
        m_fields(i).RefreshStmt pStmt
    Next i
End Sub

'==============================================================================
' Public interface
'==============================================================================
Public Property Get Count() As Long
    Count = m_count
End Property

' Default Item property - supports Fields("name") and Fields(index)
Public Property Get Item(ByVal indexOrName As Variant) As SQLite3Field
Attribute Item.VB_UserMemId = 0
    If VarType(indexOrName) = vbString Then
        Dim key As String: key = CStr(indexOrName)
        If m_nameIndex.Exists(key) Then
            Set Item = m_fields(m_nameIndex(key))
        Else
            Err.Raise vbObjectError + 300, "SQLite3Fields.Item", _
                      "Field not found: " & key
        End If
    Else
        Dim idx As Long: idx = CLng(indexOrName)
        If idx < 0 Or idx >= m_count Then
            Err.Raise vbObjectError + 301, "SQLite3Fields.Item", _
                      "Field index out of range: " & idx
        End If
        Set Item = m_fields(idx)
    End If
End Property

' Standard VBA For Each enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim col As New Collection
    Dim i As Long
    For i = 0 To m_count - 1
        col.Add m_fields(i)
    Next i
    Set NewEnum = col.[_NewEnum]
End Function

Public Function Exists(ByVal fieldName As String) As Boolean
    Exists = m_nameIndex.Exists(fieldName)
End Function
