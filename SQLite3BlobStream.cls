VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3BlobStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3BlobStream.cls  -  Incremental BLOB I/O (64-bit only)
'
' Wraps sqlite3_blob_open / _read / _write / _close / _bytes.
' Allows random-access read and write of a single BLOB cell without loading
' the entire value into memory. Useful for BLOBs larger than ~50 MB.
'
' IMPORTANT CONSTRAINTS
'   - The BLOB must already exist in the row; you cannot create a new BLOB
'     row via this API. INSERT first with a zeroblob(), then open for write.
'   - The blob handle is NOT invalidated by other database operations, but IS
'     invalidated if the row is modified by another statement on the same
'     connection while the blob is open (returns SQLITE_ABORT on next access).
'   - Reading does not require an open transaction; writing does.
'
' Typical usage:
'   ' 1. Insert a zeroed BLOB placeholder of the right size
'   conn.ExecSQL "INSERT INTO assets (name, data) VALUES ('file.bin', zeroblob(1048576));"
'   Dim rowId As LongLong: rowId = conn.LastInsertRowID()
'
'   ' 2. Open the blob handle for writing
'   Dim bs As New SQLite3BlobStream
'   bs.OpenBlob conn, "assets", "data", rowId, True
'
'   ' 3. Write in chunks
'   Dim chunk() As Byte
'   ReDim chunk(65535)  ' 64 KB at a time
'   Dim offset As Long
'   Do While offset < bs.Size
'       ' ... fill chunk() from your file ...
'       bs.WriteAt chunk, offset
'       offset = offset + 65536
'   Loop
'   bs.CloseBlob
'
' Version : 0.1.7
'
' Version History:
'   0.1.4 - Initial release.
'   0.1.5 - No functional changes. Version stamp updated.
'   0.1.6 - No functional changes. Version stamp updated.
'   0.1.7 - Module renamed from SQLite3_API/SQLite3_API_Ext/SQLite3_Helpers to
'   0.1.7 - Module reference updated from SQLite3_API to SQLite3_CoreAPI.
'            No functional changes.

'            SQLite3_CoreAPI. No functional changes; all public symbols unchanged.
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

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As LongPtr)

' -- Instance state -----------------------------------------------------------
Private m_pBlob      As LongPtr
Private m_size       As Long
Private m_position   As Long
Private m_isOpen     As Boolean
Private m_readWrite  As Boolean

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Terminate()
    If m_isOpen Then CloseBlob
End Sub

'==============================================================================
' OpenBlob
' Opens a handle to the BLOB at (tableName, columnName) in the row with the
' given rowId. readWrite=False for read-only, True for read/write.
' schema: "main" (default), "temp", or the name of an ATTACH-ed database.
'==============================================================================
Public Sub OpenBlob(ByVal conn As SQLite3Connection, _
                     ByVal tableName As String, _
                     ByVal columnName As String, _
                     ByVal rowId As LongLong, _
                     Optional ByVal readWrite As Boolean = False, _
                     Optional ByVal schema As String = "main")
    If m_isOpen Then CloseBlob

    Dim schBytes()  As Byte: schBytes  = SQLite3_CoreAPI.ToUTF8(schema)
    Dim tblBytes()  As Byte: tblBytes  = SQLite3_CoreAPI.ToUTF8(tableName)
    Dim colBytes()  As Byte: colBytes  = SQLite3_CoreAPI.ToUTF8(columnName)
    Dim flags As Long:       flags     = IIf(readWrite, 1, 0)

    Dim rc As Long
    rc = sqlite3_blob_open( _
        conn.Handle, _
        VarPtr(schBytes(0)), _
        VarPtr(tblBytes(0)), _
        VarPtr(colBytes(0)), _
        rowId, flags, VarPtr(m_pBlob))

    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 600, "SQLite3BlobStream.OpenBlob", _
                  "sqlite3_blob_open failed (rc=" & rc & ") table=[" & _
                  tableName & "] col=[" & columnName & "] row=" & CStr(rowId)
    End If

    m_size      = sqlite3_blob_bytes(m_pBlob)
    m_position  = 0
    m_readWrite = readWrite
    m_isOpen    = True
End Sub

'==============================================================================
' CloseBlob
' Closes the BLOB handle. Must be called when done, even on error paths.
'==============================================================================
Public Sub CloseBlob()
    If Not m_isOpen Then Exit Sub
    m_isOpen = False
    If m_pBlob <> 0 Then
        sqlite3_blob_close m_pBlob
        m_pBlob = 0
    End If
    m_position = 0
    m_size     = 0
End Sub

'==============================================================================
' ReadBytes / ReadAt
' ReadBytes: read nBytes from the current stream position, advance position.
' ReadAt:    read nBytes from a specific byte offset (position unchanged).
' Both return a Byte() array of exactly nBytes.
' Note: "Read" and "Write" are reserved words in VBA -- use ReadBytes/WriteBytes.
'==============================================================================
Public Function ReadBytes(ByVal nBytes As Long) As Byte()
    ReadBytes = ReadAt(nBytes, m_position)
    m_position = m_position + nBytes
End Function

Public Function ReadAt(ByVal nBytes As Long, ByVal offset As Long) As Byte()
    AssertOpen
    If nBytes <= 0 Or offset < 0 Or offset + nBytes > m_size Then
        Err.Raise vbObjectError + 601, "SQLite3BlobStream.ReadAt", _
                  "Read out of range: offset=" & offset & " n=" & nBytes & _
                  " size=" & m_size
    End If
    Dim buf() As Byte
    ReDim buf(nBytes - 1)
    Dim rc As Long
    rc = sqlite3_blob_read(m_pBlob, VarPtr(buf(0)), nBytes, offset)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 602, "SQLite3BlobStream.ReadAt", _
                  "sqlite3_blob_read failed (rc=" & rc & ")"
    End If
    ReadAt = buf
End Function

'==============================================================================
' WriteBytes / WriteAt
' WriteBytes: write data() at the current stream position, advance position.
' WriteAt:    write data() at a specific offset (position unchanged).
' The blob size is FIXED after open -- you cannot grow it here.
' To change blob size, use zeroblob() in a regular UPDATE and re-open.
'==============================================================================
Public Sub WriteBytes(ByRef data() As Byte)
    WriteAt data, m_position
    m_position = m_position + (UBound(data) - LBound(data) + 1)
End Sub

Public Sub WriteAt(ByRef data() As Byte, ByVal offset As Long)
    AssertOpen
    If Not m_readWrite Then
        Err.Raise vbObjectError + 603, "SQLite3BlobStream.WriteAt", _
                  "Blob was opened read-only."
    End If
    Dim nBytes As Long: nBytes = UBound(data) - LBound(data) + 1
    If offset < 0 Or offset + nBytes > m_size Then
        Err.Raise vbObjectError + 604, "SQLite3BlobStream.WriteAt", _
                  "Write out of range: offset=" & offset & " n=" & nBytes & _
                  " size=" & m_size
    End If
    Dim rc As Long
    rc = sqlite3_blob_write(m_pBlob, VarPtr(data(LBound(data))), nBytes, offset)
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 605, "SQLite3BlobStream.WriteAt", _
                  "sqlite3_blob_write failed (rc=" & rc & ")"
    End If
End Sub

'==============================================================================
' Seek / Position / Size
'==============================================================================
Public Sub SeekTo(ByVal pos As Long)
    If pos < 0 Or pos > m_size Then
        Err.Raise vbObjectError + 606, "SQLite3BlobStream.SeekTo", _
                  "Seek position " & pos & " out of range [0, " & m_size & "]"
    End If
    m_position = pos
End Sub

Public Property Get Position() As Long
    Position = m_position
End Property

' Total size of the BLOB in bytes (fixed at open time).
Public Property Get Size() As Long
    Size = m_size
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m_isOpen
End Property

'==============================================================================
' Private helpers
'==============================================================================
Private Sub AssertOpen()
    If Not m_isOpen Then
        Err.Raise vbObjectError + 607, "SQLite3BlobStream", "Blob handle is not open."
    End If
End Sub
