VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Command.cls  -  Prepared statement with full parameter binding (64-bit)
'
' Usage:
'   Dim cmd As New SQLite3Command
'   cmd.Prepare conn, "INSERT INTO tbl (a,b,c) VALUES (?,?,?)"
'   cmd.BindText 1, "hello"
'   cmd.BindInt  2, 42
'   cmd.Execute
'   cmd.Reset
'
' Named params:
'   cmd.Prepare conn, "SELECT * FROM tbl WHERE id=:id AND name=:name"
'   cmd.BindTextByName ":name", "Alice"
'   cmd.BindIntByName  ":id",   7
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

Private m_pStmt      As LongPtr
Private m_conn       As SQLite3Connection
Private m_sql        As String
Private m_isPrepared As Boolean

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Initialize()
    m_isPrepared = False
End Sub

Private Sub Class_Terminate()
    ' Do NOT call FinalizeStatement here - the stmt handle is owned by the
    ' connection's statement cache, which may already have finalized it.
    ' Just drop the references silently.
    m_pStmt = 0
    m_isPrepared = False
    Set m_conn = Nothing
End Sub

Public Sub Prepare(ByVal conn As SQLite3Connection, ByVal sql As String)
    Set m_conn = conn
    m_sql = sql
    m_pStmt = conn.GetCachedStatement(sql)
    m_isPrepared = True
End Sub

Public Sub FinalizeStatement()
    If m_isPrepared And m_pStmt <> 0 Then
        If SQLite3_API.SQLite_IsLoaded() Then
            sqlite3_reset m_pStmt
            sqlite3_clear_bindings m_pStmt
        End If
    End If
    m_pStmt = 0
    m_isPrepared = False
    Set m_conn = Nothing
End Sub

'==============================================================================
' Parameter binding - positional (1-based, matching SQLite convention)
'==============================================================================
Public Sub BindText(ByVal paramIdx As Long, ByVal val As String)
    AssertPrepared
    Dim buf() As Byte
    buf = SQLite3_API.ToUTF8(val)
    Dim rc As Long
    rc = sqlite3_bind_text(m_pStmt, paramIdx, VarPtr(buf(0)), _
                           UBound(buf), SQLITE_TRANSIENT)
    CheckBindRC rc, "BindText", paramIdx
End Sub

Public Sub BindInt(ByVal paramIdx As Long, ByVal val As Long)
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_bind_int(m_pStmt, paramIdx, val)
    CheckBindRC rc, "BindInt", paramIdx
End Sub

Public Sub BindInt64(ByVal paramIdx As Long, ByVal val As LongLong)
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_bind_int64(m_pStmt, paramIdx, val)
    CheckBindRC rc, "BindInt64", paramIdx
End Sub

Public Sub BindDouble(ByVal paramIdx As Long, ByVal val As Double)
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_bind_double(m_pStmt, paramIdx, val)
    CheckBindRC rc, "BindDouble", paramIdx
End Sub

Public Sub BindNull(ByVal paramIdx As Long)
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_bind_null(m_pStmt, paramIdx)
    CheckBindRC rc, "BindNull", paramIdx
End Sub

' Smart binding: infers type from Variant
Public Sub BindVariant(ByVal paramIdx As Long, ByVal val As Variant)
    Select Case VarType(val)
        Case vbNull, vbEmpty:            BindNull paramIdx
        Case vbLong, vbInteger, vbByte:  BindInt paramIdx, CLng(val)
        Case vbDouble, vbSingle:         BindDouble paramIdx, CDbl(val)
        Case vbCurrency:                 BindDouble paramIdx, CDbl(val)
        Case vbDate:                     BindText paramIdx, Format(val, "YYYY-MM-DD HH:MM:SS")
        Case vbBoolean:                  BindInt paramIdx, IIf(val, 1, 0)
        Case Else:                       BindText paramIdx, CStr(val)
    End Select
End Sub

'==============================================================================
' Named parameter binding  (:name, @name, $name)
'==============================================================================
Public Sub BindTextByName(ByVal paramName As String, ByVal val As String)
    BindText ParamIndex(paramName), val
End Sub

Public Sub BindIntByName(ByVal paramName As String, ByVal val As Long)
    BindInt ParamIndex(paramName), val
End Sub

Public Sub BindDoubleByName(ByVal paramName As String, ByVal val As Double)
    BindDouble ParamIndex(paramName), val
End Sub

Public Sub BindNullByName(ByVal paramName As String)
    BindNull ParamIndex(paramName)
End Sub

'==============================================================================
' Execution
'==============================================================================
Public Function Execute() As Long
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_step(m_pStmt)
    If rc <> SQLITE_DONE And rc <> SQLITE_ROW Then
        Dim msg As String: msg = sqlite3_errmsg_str(m_conn.Handle)
        Err.Raise vbObjectError + 500, "SQLite3Command.Execute", _
                  "sqlite3_step error " & rc & ": " & msg
    End If
    Execute = rc
End Function

Public Function ExecuteScalar() As Variant
    AssertPrepared
    Dim rc As Long
    rc = sqlite3_step(m_pStmt)
    If rc = SQLITE_ROW Then
        Select Case sqlite3_column_type(m_pStmt, 0)
            Case SQLITE_INTEGER: ExecuteScalar = sqlite3_column_int(m_pStmt, 0)
            Case SQLITE_FLOAT:   ExecuteScalar = sqlite3_column_double(m_pStmt, 0)
            Case SQLITE_TEXT:    ExecuteScalar = SQLite3_API.PtrToStringA( _
                                                   sqlite3_column_text(m_pStmt, 0))
            Case Else:           ExecuteScalar = Null
        End Select
    ElseIf rc = SQLITE_DONE Then
        ExecuteScalar = Null
    Else
        Err.Raise vbObjectError + 501, "SQLite3Command.ExecuteScalar", _
                  "sqlite3_step error " & rc
    End If
    sqlite3_reset m_pStmt
    sqlite3_clear_bindings m_pStmt
End Function

Public Sub Reset()
    AssertPrepared
    sqlite3_reset m_pStmt
    sqlite3_clear_bindings m_pStmt
End Sub

'==============================================================================
' Properties
'==============================================================================
Public Property Get sql() As String
    sql = m_sql
End Property

Public Property Get IsPrepared() As Boolean
    IsPrepared = m_isPrepared
End Property

'==============================================================================
' Private helpers
'==============================================================================
Private Sub AssertPrepared()
    If Not m_isPrepared Then
        Err.Raise vbObjectError + 510, "SQLite3Command", _
                  "Statement not prepared. Call Prepare() first."
    End If
End Sub

Private Sub CheckBindRC(ByVal rc As Long, ByVal method As String, ByVal idx As Long)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 511, "SQLite3Command." & method, _
                  "Bind error on param " & idx & ": code " & rc
    End If
End Sub

' Resolve a named parameter (:name / @name / $name) to its 1-based index.
Private Function ParamIndex(ByVal paramName As String) As Long
    Dim nameBytes() As Byte
    nameBytes = SQLite3_API.ToUTF8(paramName)
    Dim idx As Long
    idx = sqlite3_bind_parameter_index(m_pStmt, VarPtr(nameBytes(0)))
    If idx = 0 Then
        Err.Raise vbObjectError + 512, "SQLite3Command.ParamIndex", _
                  "Named parameter not found: " & paramName
    End If
    ParamIndex = idx
End Function
