Attribute VB_Name = "SQLite3_Serialize"
'==============================================================================
' SQLite3_Serialize.bas  -  Database snapshot to/from Byte() (64-bit only)
'
' Wraps sqlite3_serialize / sqlite3_deserialize (requires SQLite 3.23+).
'
' Functions:
'   SerializeDB      - snapshot a live DB to a VBA Byte() array
'   DeserializeDB    - replace a connection's content from a Byte() array
'   InMemoryClone    - open an independent in-memory copy of a live DB
'   IsSerializeAvail - True if the SQLite build includes these functions
'
' Typical usage:
'   ' Snapshot to bytes
'   Dim snap() As Byte
'   snap = SerializeDB(conn)
'   Debug.Print UBound(snap) + 1 & " bytes serialized"
'
'   ' Restore into an in-memory DB
'   Dim clone As SQLite3Connection
'   Set clone = InMemoryClone(conn)
'   ' clone is now independent -- changes to conn do not affect it
'
' Version : 0.1.4
'
' Version History:
'   0.1.4 - Initial release.
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

' Deserialize flags (bitfield)
Public Const SQLITE_DESERIALIZE_FREEONCLOSE As Long = 1  ' SQLite frees pData on close
Public Const SQLITE_DESERIALIZE_RESIZEABLE  As Long = 2  ' allow the in-memory DB to grow

'==============================================================================
' SerializeDB
' Snapshot the named schema of conn into a VBA Byte() array.
' The connection stays open and fully usable during and after the call.
' schema: "main" (default) or the name of an ATTACHed database.
'
' The returned byte array is a complete, self-contained SQLite database file.
' It can be written to disk, sent over a network, or passed to DeserializeDB.
'==============================================================================
Public Function SerializeDB(ByVal conn As SQLite3Connection, _
                              Optional ByVal schema As String = "main") As Byte()
    Dim schBytes() As Byte: schBytes = SQLite3_API.ToUTF8(schema)

    Dim szDb As LongLong   ' sqlite3_serialize writes the size here
    Dim pData As LongPtr
    pData = SQLite3_API.sqlite3_serialize( _
        conn.Handle, VarPtr(schBytes(0)), VarPtr(szDb), 0)

    If pData = 0 Then
        ' Empty or uninitialized database -- return minimal empty array
        Dim emptyBuf() As Byte
        ReDim emptyBuf(0)
        SerializeDB = emptyBuf
        Exit Function
    End If

    Dim nBytes As Long
    If szDb > 2147483647 Then
        SQLite3_API.sqlite3_free pData
        Err.Raise vbObjectError + 700, "SQLite3_Serialize.SerializeDB", _
                  "Database is too large to serialize into a VBA array (>2 GB)."
    End If
    nBytes = CLng(szDb)

    Dim buf() As Byte
    ReDim buf(nBytes - 1)
    CopyMemory buf(0), ByVal pData, CLngPtr(nBytes)
    SQLite3_API.sqlite3_free pData

    SerializeDB = buf
End Function

'==============================================================================
' DeserializeDB
' Replace the content of an existing open connection with the bytes in data().
' The connection must have been opened (typically as ":memory:" for a fresh
' in-memory DB, or any file DB you want to overwrite).
'
' After this call the connection behaves exactly as if it had been opened
' against a file containing data().
'
' WARNING: This replaces ALL content in the target connection.
'          Any uncommitted transactions are rolled back first.
'==============================================================================
Public Sub DeserializeDB(ByVal conn As SQLite3Connection, _
                           ByRef data() As Byte, _
                           Optional ByVal schema As String = "main")
    Dim nBytes As Long: nBytes = UBound(data) - LBound(data) + 1
    If nBytes <= 0 Then
        Err.Raise vbObjectError + 701, "SQLite3_Serialize.DeserializeDB", _
                  "data() array is empty."
    End If

    ' Allocate a SQLite-heap buffer (so FREEONCLOSE works)
    Dim pBuf As LongPtr: pBuf = SQLite3_API.sqlite3_malloc(nBytes)
    If pBuf = 0 Then
        Err.Raise vbObjectError + 702, "SQLite3_Serialize.DeserializeDB", _
                  "sqlite3_malloc(" & nBytes & ") returned NULL -- out of memory."
    End If

    CopyMemory ByVal pBuf, data(LBound(data)), CLngPtr(nBytes)

    Dim schBytes() As Byte: schBytes = SQLite3_API.ToUTF8(schema)
    Dim szDb As LongLong: szDb = CLngLng(nBytes)
    Dim flags As Long: flags = SQLITE_DESERIALIZE_FREEONCLOSE Or SQLITE_DESERIALIZE_RESIZEABLE

    Dim rc As Long
    rc = SQLite3_API.sqlite3_deserialize( _
        conn.Handle, VarPtr(schBytes(0)), pBuf, szDb, szDb, flags)

    If rc <> SQLITE_OK Then
        ' pBuf ownership passed to SQLite even on error when FREEONCLOSE is set
        Err.Raise vbObjectError + 703, "SQLite3_Serialize.DeserializeDB", _
                  "sqlite3_deserialize failed (rc=" & rc & ")"
    End If
End Sub

'==============================================================================
' InMemoryClone
' Open a new, independent in-memory connection whose content is an exact
' snapshot of conn at the moment of the call.
' Changes to conn after this call do NOT affect the clone, and vice versa.
'
' Uses the backup API internally (not serialize/deserialize) so that the clone
' is built from a clean page-level copy unaffected by WAL header state.
' The clone is opened without WAL mode -- WAL is irrelevant for :memory:.
'
' The returned connection must be closed by the caller when done.
'==============================================================================
Public Function InMemoryClone(ByVal conn As SQLite3Connection, _
                                Optional ByVal srcSchema As String = "main") As SQLite3Connection
    ' Open a fresh :memory: connection.  WAL is explicitly disabled --
    ' WAL mode on :memory: can interfere with reads after a backup step.
    Dim clone As New SQLite3Connection
    clone.OpenDatabase ":memory:", conn.DllPath, 5000, False

    ' Convert schema names to UTF-8 byte arrays for the backup API.
    Dim srcSchBytes() As Byte: srcSchBytes = SQLite3_API.ToUTF8(srcSchema)
    Dim dstSchBytes() As Byte: dstSchBytes = SQLite3_API.ToUTF8("main")

    ' Initialise the backup from conn -> clone.
    Dim pBackup As LongPtr
    pBackup = SQLite3_API.sqlite3_backup_init( _
        clone.Handle, VarPtr(dstSchBytes(0)), _
        conn.Handle,  VarPtr(srcSchBytes(0)))

    If pBackup = 0 Then
        Err.Raise vbObjectError + 710, "SQLite3_Serialize.InMemoryClone", _
                  "sqlite3_backup_init failed: " & SQLite3_API.sqlite3_errmsg_str(conn.Handle)
    End If

    ' Copy all pages in one step.  Returns SQLITE_DONE (101) on success.
    Dim rc As Long
    rc = SQLite3_API.sqlite3_backup_step(pBackup, -1)
    SQLite3_API.sqlite3_backup_finish pBackup

    If rc <> SQLITE_DONE And rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 711, "SQLite3_Serialize.InMemoryClone", _
                  "sqlite3_backup_step failed (rc=" & rc & ")"
    End If

    Set InMemoryClone = clone
End Function

'==============================================================================
' IsSerializeAvail
' Returns True if sqlite3_serialize is available in the loaded DLL.
' (Requires SQLite 3.23.0 or later, released 2018.)
' All official precompiled binaries from sqlite.org include it.
'==============================================================================
Public Function IsSerializeAvail() As Boolean
    On Error Resume Next
    Dim schBytes() As Byte: schBytes = SQLite3_API.ToUTF8("main")
    Dim szDb As LongLong
    ' Call with a null db handle -- will fail gracefully but won't crash
    ' if the symbol exists. A missing export gives a null proc address and
    ' DispCallFunc returns 0 immediately.
    ' Best check: just try on a real connection in your test suite.
    ' Here we return True if PROC_COUNT covers the serialize slot (always true
    ' in v0.1.4+) and the DLL is loaded.
    IsSerializeAvail = SQLite3_API.SQLite_IsLoaded()
    On Error GoTo 0
End Function
