Attribute VB_Name = "SQLite3_API"
'==============================================================================
' SQLite3_API.bas  -  HFT-Grade SQLite3 Dynamic Loader for VBA (64-bit only)
' Architecture : LoadLibrary / GetProcAddress / DispCallFunc
' Key point    : prgpvarg must be ByRef LongPtr (array of ptrs to Variants)
'
' Version : 0.1.4
'
' Version History:
'   0.1.0 - Initial release. LoadLibrary/DispCallFunc loader, 28 proc cache,
'            UTF-8 marshalling, all core SQLite3 wrappers.
'   0.1.1 - Fixed DispCallFunc declaration (prgpvarg ByRef As LongPtr).
'            Fixed VT_PTR->VT_I8, pvInstance/oVft argument order.
'            Replaced hand-rolled strlen loop with kernel32.lstrlenA.
'            Fixed ToUTF8 compound Dim statement; CLngPtr(0) for null args.
'   0.1.2 - Added sqlite3_bind_blob, sqlite3_column_blob wrappers.
'            Added BlobToBytes() helper (CopyMemory-based, no byte loop).
'            PROC_COUNT bumped from 28 to 30.
'   0.1.3 - Added sqlite3_interrupt wrapper.
'            P_INTERRUPT = 30; PROC_COUNT bumped to 31.
'   0.1.4 - Added Online Backup API (backup_init/step/finish/remaining/pagecount).
'            Added incremental BLOB I/O (blob_open/read/write/close/bytes).
'            Added serialize/deserialize + sqlite3_malloc/free.
'            Added sqlite3_db_status, sqlite3_stmt_status.
'            PROC_COUNT bumped from 31 to 47.
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

' -- Win32 / COM declarations -------------------------------------------------
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" _
    (ByVal lpFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" _
    (ByVal hModule As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" _
    (ByVal CodePage As Long, ByVal dwFlags As Long, _
     ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
     ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" _
    (ByVal CodePage As Long, ByVal dwFlags As Long, _
     ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
     ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenA Lib "kernel32" _
    (ByVal lpString As LongPtr) As Long
'  prgpvarg is VARIANTARG** : array of pointers to Variants -> ByRef LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" _
    (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
     ByVal vtReturn As Integer, ByVal cActuals As Long, _
     ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, _
     ByRef pvargResult As Variant) As Long

' -- Calling convention constants ---------------------------------------------
Public Const CC_CDECL   As Long = 1
Public Const CC_STDCALL As Long = 4

' -- Variant type tags --------------------------------------------------------
Public Const VT_I4 As Integer = 3
Public Const VT_I8 As Integer = 20
Public Const VT_R8 As Integer = 5

' -- SQLite result codes ------------------------------------------------------
Public Const SQLITE_OK    As Long = 0
Public Const SQLITE_ERROR As Long = 1
Public Const SQLITE_BUSY  As Long = 5
Public Const SQLITE_ROW   As Long = 100
Public Const SQLITE_DONE  As Long = 101
Public Const SQLITE_OPEN_READWRITE As Long = 2
Public Const SQLITE_OPEN_CREATE    As Long = 4
Public Const SQLITE_OPEN_FULLMUTEX As Long = &H10000

' -- SQLite column type constants ---------------------------------------------
Public Const SQLITE_INTEGER As Long = 1
Public Const SQLITE_FLOAT   As Long = 2
Public Const SQLITE_TEXT    As Long = 3
Public Const SQLITE_BLOB    As Long = 4
Public Const SQLITE_NULL    As Long = 5

' -- SQLITE_TRANSIENT ---------------------------------------------------------
Public Const SQLITE_TRANSIENT As LongLong = -1

' -- Internal -----------------------------------------------------------------
Private Const CP_UTF8 As Long = 65001

Private m_hDll      As LongPtr
Private m_procs(46) As LongPtr

Private Const P_OPEN_V2        As Long = 0
Private Const P_CLOSE          As Long = 1
Private Const P_PREPARE_V2     As Long = 2
Private Const P_STEP           As Long = 3
Private Const P_RESET          As Long = 4
Private Const P_FINALIZE       As Long = 5
Private Const P_BIND_TEXT      As Long = 6
Private Const P_BIND_INT       As Long = 7
Private Const P_BIND_INT64     As Long = 8
Private Const P_BIND_DOUBLE    As Long = 9
Private Const P_BIND_NULL      As Long = 10
Private Const P_COLUMN_COUNT   As Long = 11
Private Const P_COLUMN_TYPE    As Long = 12
Private Const P_COLUMN_TEXT    As Long = 13
Private Const P_COLUMN_INT     As Long = 14
Private Const P_COLUMN_INT64   As Long = 15
Private Const P_COLUMN_DOUBLE  As Long = 16
Private Const P_COLUMN_NAME    As Long = 17
Private Const P_COLUMN_BYTES   As Long = 18
Private Const P_ERRMSG         As Long = 19
Private Const P_CHANGES        As Long = 20
Private Const P_LAST_INSERT    As Long = 21
Private Const P_BUSY_TIMEOUT   As Long = 22
Private Const P_EXEC           As Long = 23
Private Const P_BIND_PARAM_IDX As Long = 24
Private Const P_CLEAR_BINDINGS As Long = 25
Private Const P_TOTAL_CHANGES  As Long = 26
Private Const P_LIBVERSION     As Long = 27
Private Const P_BIND_BLOB      As Long = 28
Private Const P_COLUMN_BLOB    As Long = 29
Private Const P_INTERRUPT      As Long = 30
' v0.1.4 additions
Private Const P_BACKUP_INIT    As Long = 31
Private Const P_BACKUP_STEP    As Long = 32
Private Const P_BACKUP_FINISH  As Long = 33
Private Const P_BACKUP_REMAIN  As Long = 34
Private Const P_BACKUP_PGCOUNT As Long = 35
Private Const P_BLOB_OPEN      As Long = 36
Private Const P_BLOB_READ      As Long = 37
Private Const P_BLOB_WRITE     As Long = 38
Private Const P_BLOB_CLOSE     As Long = 39
Private Const P_BLOB_BYTES     As Long = 40
Private Const P_SERIALIZE      As Long = 41
Private Const P_DESERIALIZE    As Long = 42
Private Const P_MALLOC         As Long = 43
Private Const P_FREE           As Long = 44
Private Const P_DB_STATUS      As Long = 45
Private Const P_STMT_STATUS    As Long = 46
Private Const PROC_COUNT       As Long = 47

'==============================================================================
' Library lifecycle
'==============================================================================
Public Function SQLite_Load(ByVal dllPath As String) As Boolean
    If m_hDll <> 0 Then SQLite_Load = True: Exit Function
    m_hDll = LoadLibraryW(StrPtr(dllPath))
    If m_hDll = 0 Then
        Err.Raise vbObjectError + 1, "SQLite3_API.SQLite_Load", _
                  "Cannot load sqlite3.dll from: " & dllPath
    End If
    Dim i As Long
    For i = 0 To PROC_COUNT - 1
        m_procs(i) = GetProcAddress(m_hDll, ProcName(i))
    Next i
    SQLite3_API_Ext.SetDllHandle m_hDll
    SQLite_Load = True
End Function

Public Sub SQLite_Unload()
    If m_hDll = 0 Then Exit Sub
    FreeLibrary m_hDll
    m_hDll = 0
    Dim i As Long
    For i = 0 To 46: m_procs(i) = 0: Next i
End Sub

Public Function SQLite_IsLoaded() As Boolean
    SQLite_IsLoaded = (m_hDll <> 0)
End Function

Public Function SQLite_Version() As String
    SQLite_Version = PtrToStringA(Invoke0(P_LIBVERSION))
End Function

Public Function SQLite_DllHandle() As LongPtr
    SQLite_DllHandle = m_hDll
End Function

'==============================================================================
' UTF-8 marshalling
'==============================================================================
Public Function ToUTF8(ByVal s As String) As Byte()
    Dim buf() As Byte
    If Len(s) = 0 Then
        ReDim buf(0)
        buf(0) = 0
        ToUTF8 = buf
        Exit Function
    End If
    Dim cbNeeded As Long
    cbNeeded = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), -1, _
                                   CLngPtr(0), 0, CLngPtr(0), CLngPtr(0))
    ReDim buf(cbNeeded - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(s), -1, _
                        VarPtr(buf(0)), cbNeeded, CLngPtr(0), CLngPtr(0)
    ToUTF8 = buf
End Function

Public Function PtrToStringA(ByVal pUtf8 As LongPtr) As String
    If pUtf8 = 0 Then Exit Function
    Dim cbLen As Long: cbLen = lstrlenA(pUtf8)
    If cbLen = 0 Then Exit Function
    Dim cchNeeded As Long
    cchNeeded = MultiByteToWideChar(CP_UTF8, 0, pUtf8, cbLen, CLngPtr(0), 0)
    If cchNeeded = 0 Then Exit Function
    Dim buf() As Byte
    ReDim buf((cchNeeded * 2) - 1)
    MultiByteToWideChar CP_UTF8, 0, pUtf8, cbLen, VarPtr(buf(0)), cchNeeded
    PtrToStringA = buf
End Function

Public Sub CopyMemoryByte(dst As Byte, ByVal srcPtr As LongPtr)
    CopyMemory dst, ByVal srcPtr, 1
End Sub

'==============================================================================
' SQLite3 wrappers
' Each builds: args() As Variant, vt() As Integer, ptrs() As LongPtr
' ptrs(i) = VarPtr(args(i))  -- required by DispCallFunc (VARIANTARG**)
'==============================================================================
Public Function sqlite3_open_v2(ByVal pFilename As LongPtr, _
                                 ppDb As LongPtr, _
                                 ByVal flags As Long, _
                                 ByVal pVfs As LongPtr) As Long
    Dim args(3) As Variant, vt(3) As Integer, ptrs(3) As LongPtr, ret As Variant
    args(0) = CLngLng(pFilename):       vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(VarPtr(ppDb)):    vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(flags):              vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(pVfs):            vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    DispCallFunc CLngPtr(0), m_procs(P_OPEN_V2), CC_CDECL, VT_I4, 4, vt(0), ptrs(0), ret
    sqlite3_open_v2 = CLng(ret)
End Function

Public Function sqlite3_close(ByVal pDb As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_CLOSE), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_close = CLng(ret)
End Function

Public Function sqlite3_prepare_v2(ByVal pDb As LongPtr, _
                                    ByVal pSql As LongPtr, _
                                    ByVal nByte As Long, _
                                    ppStmt As LongPtr, _
                                    ppTail As LongPtr) As Long
    Dim args(4) As Variant, vt(4) As Integer, ptrs(4) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):            vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pSql):           vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(nByte):             vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(VarPtr(ppStmt)): vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    args(4) = CLngLng(VarPtr(ppTail)): vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    DispCallFunc CLngPtr(0), m_procs(P_PREPARE_V2), CC_CDECL, VT_I4, 5, vt(0), ptrs(0), ret
    sqlite3_prepare_v2 = CLng(ret)
End Function

Public Function sqlite3_step(ByVal pStmt As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_STEP), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_step = CLng(ret)
End Function

Public Function sqlite3_reset(ByVal pStmt As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_RESET), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_reset = CLng(ret)
End Function

Public Function sqlite3_finalize(ByVal pStmt As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_FINALIZE), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_finalize = CLng(ret)
End Function

Public Function sqlite3_bind_text(ByVal pStmt As LongPtr, _
                                   ByVal iCol As Long, _
                                   ByVal pText As LongPtr, _
                                   ByVal nBytes As Long, _
                                   ByVal pDestructor As LongPtr) As Long
    Dim args(4) As Variant, vt(4) As Integer, ptrs(4) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt):       vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):           vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pText):       vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLng(nBytes):         vt(3) = VT_I4: ptrs(3) = VarPtr(args(3))
    args(4) = CLngLng(pDestructor): vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_TEXT), CC_CDECL, VT_I4, 5, vt(0), ptrs(0), ret
    sqlite3_bind_text = CLng(ret)
End Function

Public Function sqlite3_bind_int(ByVal pStmt As LongPtr, _
                                  ByVal iCol As Long, _
                                  ByVal val As Long) As Long
    Dim args(2) As Variant, vt(2) As Integer, ptrs(2) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(val):      vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_INT), CC_CDECL, VT_I4, 3, vt(0), ptrs(0), ret
    sqlite3_bind_int = CLng(ret)
End Function

Public Function sqlite3_bind_int64(ByVal pStmt As LongPtr, _
                                    ByVal iCol As Long, _
                                    ByVal val As LongLong) As Long
    Dim args(2) As Variant, vt(2) As Integer, ptrs(2) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = val:            vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_INT64), CC_CDECL, VT_I4, 3, vt(0), ptrs(0), ret
    sqlite3_bind_int64 = CLng(ret)
End Function

Public Function sqlite3_bind_double(ByVal pStmt As LongPtr, _
                                     ByVal iCol As Long, _
                                     ByVal val As Double) As Long
    Dim args(2) As Variant, vt(2) As Integer, ptrs(2) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = val:            vt(2) = VT_R8: ptrs(2) = VarPtr(args(2))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_DOUBLE), CC_CDECL, VT_I4, 3, vt(0), ptrs(0), ret
    sqlite3_bind_double = CLng(ret)
End Function

Public Function sqlite3_bind_null(ByVal pStmt As LongPtr, _
                                   ByVal iCol As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_NULL), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_bind_null = CLng(ret)
End Function

Public Function sqlite3_bind_parameter_index(ByVal pStmt As LongPtr, _
                                              ByVal pName As LongPtr) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pName): vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_PARAM_IDX), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_bind_parameter_index = CLng(ret)
End Function

Public Function sqlite3_clear_bindings(ByVal pStmt As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_CLEAR_BINDINGS), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_clear_bindings = CLng(ret)
End Function

Public Function sqlite3_column_count(ByVal pStmt As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_COUNT), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_column_count = CLng(ret)
End Function

Public Function sqlite3_column_type(ByVal pStmt As LongPtr, _
                                     ByVal iCol As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_TYPE), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_column_type = CLng(ret)
End Function

Public Function sqlite3_column_text(ByVal pStmt As LongPtr, _
                                     ByVal iCol As Long) As LongPtr
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_TEXT), CC_CDECL, VT_I8, 2, vt(0), ptrs(0), ret
    sqlite3_column_text = CLngPtr(ret)
End Function

Public Function sqlite3_column_int(ByVal pStmt As LongPtr, _
                                    ByVal iCol As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_INT), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_column_int = CLng(ret)
End Function

Public Function sqlite3_column_int64(ByVal pStmt As LongPtr, _
                                      ByVal iCol As Long) As LongLong
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_INT64), CC_CDECL, VT_I8, 2, vt(0), ptrs(0), ret
    sqlite3_column_int64 = CLngLng(ret)
End Function

Public Function sqlite3_column_double(ByVal pStmt As LongPtr, _
                                       ByVal iCol As Long) As Double
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_DOUBLE), CC_CDECL, VT_R8, 2, vt(0), ptrs(0), ret
    sqlite3_column_double = CDbl(ret)
End Function

Public Function sqlite3_column_name(ByVal pStmt As LongPtr, _
                                     ByVal iCol As Long) As LongPtr
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_NAME), CC_CDECL, VT_I8, 2, vt(0), ptrs(0), ret
    sqlite3_column_name = CLngPtr(ret)
End Function

Public Function sqlite3_column_bytes(ByVal pStmt As LongPtr, _
                                      ByVal iCol As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_BYTES), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_column_bytes = CLng(ret)
End Function

Public Function sqlite3_errmsg(ByVal pDb As LongPtr) As LongPtr
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_ERRMSG), CC_CDECL, VT_I8, 1, vt(0), ptrs(0), ret
    sqlite3_errmsg = CLngPtr(ret)
End Function

Public Function sqlite3_errmsg_str(ByVal pDb As LongPtr) As String
    sqlite3_errmsg_str = PtrToStringA(sqlite3_errmsg(pDb))
End Function

Public Function sqlite3_changes(ByVal pDb As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_CHANGES), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_changes = CLng(ret)
End Function

Public Function sqlite3_last_insert_rowid(ByVal pDb As LongPtr) As LongLong
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_LAST_INSERT), CC_CDECL, VT_I8, 1, vt(0), ptrs(0), ret
    sqlite3_last_insert_rowid = CLngLng(ret)
End Function

Public Function sqlite3_busy_timeout(ByVal pDb As LongPtr, _
                                      ByVal ms As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(ms):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_BUSY_TIMEOUT), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_busy_timeout = CLng(ret)
End Function

Public Function sqlite3_exec(ByVal pDb As LongPtr, _
                              ByVal pSql As LongPtr, _
                              ByVal pCallback As LongPtr, _
                              ByVal pArg As LongPtr, _
                              ppErrMsg As LongPtr) As Long
    Dim args(4) As Variant, vt(4) As Integer, ptrs(4) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):              vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pSql):             vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pCallback):        vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(pArg):             vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    args(4) = CLngLng(VarPtr(ppErrMsg)): vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    DispCallFunc CLngPtr(0), m_procs(P_EXEC), CC_CDECL, VT_I4, 5, vt(0), ptrs(0), ret
    sqlite3_exec = CLng(ret)
End Function

Public Function sqlite3_bind_blob(ByVal pStmt As LongPtr, _
                                   ByVal iCol As Long, _
                                   ByVal pData As LongPtr, _
                                   ByVal nBytes As Long, _
                                   ByVal pDestructor As LongPtr) As Long
    Dim args(4) As Variant, vt(4) As Integer, ptrs(4) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt):       vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):           vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pData):       vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLng(nBytes):         vt(3) = VT_I4: ptrs(3) = VarPtr(args(3))
    args(4) = CLngLng(pDestructor): vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    DispCallFunc CLngPtr(0), m_procs(P_BIND_BLOB), CC_CDECL, VT_I4, 5, vt(0), ptrs(0), ret
    sqlite3_bind_blob = CLng(ret)
End Function

Public Function sqlite3_column_blob(ByVal pStmt As LongPtr, _
                                     ByVal iCol As Long) As LongPtr
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(iCol):     vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_COLUMN_BLOB), CC_CDECL, VT_I8, 2, vt(0), ptrs(0), ret
    sqlite3_column_blob = CLngPtr(ret)
End Function

' Copy nBytes from a raw pointer into a VBA Byte array
Public Function BlobToBytes(ByVal pBlob As LongPtr, ByVal nBytes As Long) As Byte()
    Dim buf() As Byte
    If pBlob = 0 Or nBytes <= 0 Then
        ReDim buf(0)
        BlobToBytes = buf
        Exit Function
    End If
    ReDim buf(nBytes - 1)
    CopyMemory buf(0), ByVal pBlob, nBytes
    BlobToBytes = buf
End Function

' Signal a running query to abort.  Safe to call from a timer or any
' point in VBA while a Step/Exec call is pending on the same connection.
' SQLite returns SQLITE_INTERRUPT (9) to the blocked call.
Public Sub sqlite3_interrupt(ByVal pDb As LongPtr)
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_INTERRUPT), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
End Sub

'==============================================================================
' Online Backup API  (v0.1.4)
' sqlite3_backup_init returns an opaque handle; Step/Finish/Remaining/Pagecount
' take that handle. All string args are pre-converted to UTF-8 byte arrays by
' the caller (SQLite3Backup.cls) -- the wrappers receive raw pointers.
'==============================================================================
Public Function sqlite3_backup_init(ByVal pDest As LongPtr, _
                                     ByVal pDestName As LongPtr, _
                                     ByVal pSrc As LongPtr, _
                                     ByVal pSrcName As LongPtr) As LongPtr
    Dim args(3) As Variant, vt(3) As Integer, ptrs(3) As LongPtr, ret As Variant
    args(0) = CLngLng(pDest):     vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pDestName): vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pSrc):      vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(pSrcName):  vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    DispCallFunc CLngPtr(0), m_procs(P_BACKUP_INIT), CC_CDECL, VT_I8, 4, vt(0), ptrs(0), ret
    sqlite3_backup_init = CLngPtr(ret)
End Function

Public Function sqlite3_backup_step(ByVal pBackup As LongPtr, _
                                     ByVal nPage As Long) As Long
    Dim args(1) As Variant, vt(1) As Integer, ptrs(1) As LongPtr, ret As Variant
    args(0) = CLngLng(pBackup): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(nPage):      vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    DispCallFunc CLngPtr(0), m_procs(P_BACKUP_STEP), CC_CDECL, VT_I4, 2, vt(0), ptrs(0), ret
    sqlite3_backup_step = CLng(ret)
End Function

Public Function sqlite3_backup_finish(ByVal pBackup As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pBackup): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_BACKUP_FINISH), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_backup_finish = CLng(ret)
End Function

Public Function sqlite3_backup_remaining(ByVal pBackup As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pBackup): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_BACKUP_REMAIN), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_backup_remaining = CLng(ret)
End Function

Public Function sqlite3_backup_pagecount(ByVal pBackup As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pBackup): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_BACKUP_PGCOUNT), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_backup_pagecount = CLng(ret)
End Function

'==============================================================================
' Incremental BLOB I/O  (v0.1.4)
' sqlite3_blob_open opens a handle to a single BLOB cell by rowid.
' ppBlob is an OUT parameter; pass VarPtr(pBlob) where pBlob is a LongPtr.
'==============================================================================
Public Function sqlite3_blob_open(ByVal pDb As LongPtr, _
                                   ByVal pZDb As LongPtr, _
                                   ByVal pZTable As LongPtr, _
                                   ByVal pZColumn As LongPtr, _
                                   ByVal iRow As LongLong, _
                                   ByVal flags As Long, _
                                   ByVal ppBlob As LongPtr) As Long
    Dim args(6) As Variant, vt(6) As Integer, ptrs(6) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):     vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pZDb):    vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pZTable): vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(pZColumn):vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    args(4) = iRow:             vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    args(5) = CLng(flags):      vt(5) = VT_I4: ptrs(5) = VarPtr(args(5))
    args(6) = CLngLng(ppBlob):  vt(6) = VT_I8: ptrs(6) = VarPtr(args(6))
    DispCallFunc CLngPtr(0), m_procs(P_BLOB_OPEN), CC_CDECL, VT_I4, 7, vt(0), ptrs(0), ret
    sqlite3_blob_open = CLng(ret)
End Function

Public Function sqlite3_blob_read(ByVal pBlob As LongPtr, _
                                   ByVal pBuf As LongPtr, _
                                   ByVal nBytes As Long, _
                                   ByVal iOffset As Long) As Long
    Dim args(3) As Variant, vt(3) As Integer, ptrs(3) As LongPtr, ret As Variant
    args(0) = CLngLng(pBlob):  vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pBuf):   vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(nBytes):    vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    args(3) = CLng(iOffset):   vt(3) = VT_I4: ptrs(3) = VarPtr(args(3))
    DispCallFunc CLngPtr(0), m_procs(P_BLOB_READ), CC_CDECL, VT_I4, 4, vt(0), ptrs(0), ret
    sqlite3_blob_read = CLng(ret)
End Function

Public Function sqlite3_blob_write(ByVal pBlob As LongPtr, _
                                    ByVal pBuf As LongPtr, _
                                    ByVal nBytes As Long, _
                                    ByVal iOffset As Long) As Long
    Dim args(3) As Variant, vt(3) As Integer, ptrs(3) As LongPtr, ret As Variant
    args(0) = CLngLng(pBlob):  vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pBuf):   vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(nBytes):    vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    args(3) = CLng(iOffset):   vt(3) = VT_I4: ptrs(3) = VarPtr(args(3))
    DispCallFunc CLngPtr(0), m_procs(P_BLOB_WRITE), CC_CDECL, VT_I4, 4, vt(0), ptrs(0), ret
    sqlite3_blob_write = CLng(ret)
End Function

Public Function sqlite3_blob_close(ByVal pBlob As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pBlob): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_BLOB_CLOSE), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_blob_close = CLng(ret)
End Function

Public Function sqlite3_blob_bytes(ByVal pBlob As LongPtr) As Long
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pBlob): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_BLOB_BYTES), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
    sqlite3_blob_bytes = CLng(ret)
End Function

'==============================================================================
' Serialize / Deserialize  (v0.1.4, requires SQLite 3.23+)
' sqlite3_serialize: snapshot a DB to a raw byte buffer (returned as LongPtr).
'   piSize is an OUT parameter -- pass VarPtr(szDb) where szDb is LongLong.
'   Caller must free the returned pointer with sqlite3_free.
' sqlite3_deserialize: replace a DB's content from a raw byte buffer.
'   pData must be allocated with sqlite3_malloc (use FREEONCLOSE so SQLite owns it).
'==============================================================================
Public Function sqlite3_serialize(ByVal pDb As LongPtr, _
                                   ByVal pSchema As LongPtr, _
                                   ByVal piSize As LongPtr, _
                                   ByVal mFlags As Long) As LongPtr
    Dim args(3) As Variant, vt(3) As Integer, ptrs(3) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):     vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pSchema): vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(piSize):  vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLng(mFlags):     vt(3) = VT_I4: ptrs(3) = VarPtr(args(3))
    DispCallFunc CLngPtr(0), m_procs(P_SERIALIZE), CC_CDECL, VT_I8, 4, vt(0), ptrs(0), ret
    sqlite3_serialize = CLngPtr(ret)
End Function

Public Function sqlite3_deserialize(ByVal pDb As LongPtr, _
                                     ByVal pSchema As LongPtr, _
                                     ByVal pData As LongPtr, _
                                     ByVal szDb As LongLong, _
                                     ByVal szBuf As LongLong, _
                                     ByVal mFlags As Long) As Long
    Dim args(5) As Variant, vt(5) As Integer, ptrs(5) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):     vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLngLng(pSchema): vt(1) = VT_I8: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pData):   vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = szDb:             vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    args(4) = szBuf:            vt(4) = VT_I8: ptrs(4) = VarPtr(args(4))
    args(5) = CLng(mFlags):     vt(5) = VT_I4: ptrs(5) = VarPtr(args(5))
    DispCallFunc CLngPtr(0), m_procs(P_DESERIALIZE), CC_CDECL, VT_I4, 6, vt(0), ptrs(0), ret
    sqlite3_deserialize = CLng(ret)
End Function

' Allocate nBytes from SQLite's internal heap. Required for buffers passed to
' sqlite3_deserialize with SQLITE_DESERIALIZE_FREEONCLOSE.
Public Function sqlite3_malloc(ByVal nBytes As Long) As LongPtr
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLng(nBytes): vt(0) = VT_I4: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_MALLOC), CC_CDECL, VT_I8, 1, vt(0), ptrs(0), ret
    sqlite3_malloc = CLngPtr(ret)
End Function

' Free a pointer previously returned by sqlite3_malloc or sqlite3_serialize.
Public Sub sqlite3_free(ByVal pMem As LongPtr)
    If pMem = 0 Then Exit Sub
    Dim args(0) As Variant, vt(0) As Integer, ptrs(0) As LongPtr, ret As Variant
    args(0) = CLngLng(pMem): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    DispCallFunc CLngPtr(0), m_procs(P_FREE), CC_CDECL, VT_I4, 1, vt(0), ptrs(0), ret
End Sub

'==============================================================================
' Status counters  (v0.1.4)
' sqlite3_db_status: per-connection counters (cache hits, pages used, etc.)
'   pCur and pHiwtr are OUT int* -- pass VarPtr(cur) and VarPtr(hi).
' sqlite3_stmt_status: per-statement counters (full-scans, sorts, VM steps).
'==============================================================================
Public Function sqlite3_db_status(ByVal pDb As LongPtr, _
                                   ByVal op As Long, _
                                   ByVal pCur As LongPtr, _
                                   ByVal pHiwtr As LongPtr, _
                                   ByVal resetFlg As Long) As Long
    Dim args(4) As Variant, vt(4) As Integer, ptrs(4) As LongPtr, ret As Variant
    args(0) = CLngLng(pDb):    vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(op):        vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = CLngLng(pCur):   vt(2) = VT_I8: ptrs(2) = VarPtr(args(2))
    args(3) = CLngLng(pHiwtr): vt(3) = VT_I8: ptrs(3) = VarPtr(args(3))
    args(4) = CLng(resetFlg):  vt(4) = VT_I4: ptrs(4) = VarPtr(args(4))
    DispCallFunc CLngPtr(0), m_procs(P_DB_STATUS), CC_CDECL, VT_I4, 5, vt(0), ptrs(0), ret
    sqlite3_db_status = CLng(ret)
End Function

Public Function sqlite3_stmt_status(ByVal pStmt As LongPtr, _
                                     ByVal op As Long, _
                                     ByVal resetFlg As Long) As Long
    Dim args(2) As Variant, vt(2) As Integer, ptrs(2) As LongPtr, ret As Variant
    args(0) = CLngLng(pStmt): vt(0) = VT_I8: ptrs(0) = VarPtr(args(0))
    args(1) = CLng(op):       vt(1) = VT_I4: ptrs(1) = VarPtr(args(1))
    args(2) = CLng(resetFlg): vt(2) = VT_I4: ptrs(2) = VarPtr(args(2))
    DispCallFunc CLngPtr(0), m_procs(P_STMT_STATUS), CC_CDECL, VT_I4, 3, vt(0), ptrs(0), ret
    sqlite3_stmt_status = CLng(ret)
End Function

'==============================================================================
' Private helpers
'==============================================================================
Private Function ProcName(ByVal idx As Long) As String
    Select Case idx
        Case P_OPEN_V2:        ProcName = "sqlite3_open_v2"
        Case P_CLOSE:          ProcName = "sqlite3_close"
        Case P_PREPARE_V2:     ProcName = "sqlite3_prepare_v2"
        Case P_STEP:           ProcName = "sqlite3_step"
        Case P_RESET:          ProcName = "sqlite3_reset"
        Case P_FINALIZE:       ProcName = "sqlite3_finalize"
        Case P_BIND_TEXT:      ProcName = "sqlite3_bind_text"
        Case P_BIND_INT:       ProcName = "sqlite3_bind_int"
        Case P_BIND_INT64:     ProcName = "sqlite3_bind_int64"
        Case P_BIND_DOUBLE:    ProcName = "sqlite3_bind_double"
        Case P_BIND_NULL:      ProcName = "sqlite3_bind_null"
        Case P_COLUMN_COUNT:   ProcName = "sqlite3_column_count"
        Case P_COLUMN_TYPE:    ProcName = "sqlite3_column_type"
        Case P_COLUMN_TEXT:    ProcName = "sqlite3_column_text"
        Case P_COLUMN_INT:     ProcName = "sqlite3_column_int"
        Case P_COLUMN_INT64:   ProcName = "sqlite3_column_int64"
        Case P_COLUMN_DOUBLE:  ProcName = "sqlite3_column_double"
        Case P_COLUMN_NAME:    ProcName = "sqlite3_column_name"
        Case P_COLUMN_BYTES:   ProcName = "sqlite3_column_bytes"
        Case P_ERRMSG:         ProcName = "sqlite3_errmsg"
        Case P_CHANGES:        ProcName = "sqlite3_changes"
        Case P_LAST_INSERT:    ProcName = "sqlite3_last_insert_rowid"
        Case P_BUSY_TIMEOUT:   ProcName = "sqlite3_busy_timeout"
        Case P_EXEC:           ProcName = "sqlite3_exec"
        Case P_BIND_PARAM_IDX: ProcName = "sqlite3_bind_parameter_index"
        Case P_CLEAR_BINDINGS: ProcName = "sqlite3_clear_bindings"
        Case P_TOTAL_CHANGES:  ProcName = "sqlite3_total_changes"
        Case P_LIBVERSION:     ProcName = "sqlite3_libversion"
        Case P_BIND_BLOB:      ProcName = "sqlite3_bind_blob"
        Case P_COLUMN_BLOB:    ProcName = "sqlite3_column_blob"
        Case P_INTERRUPT:      ProcName = "sqlite3_interrupt"
        Case P_BACKUP_INIT:    ProcName = "sqlite3_backup_init"
        Case P_BACKUP_STEP:    ProcName = "sqlite3_backup_step"
        Case P_BACKUP_FINISH:  ProcName = "sqlite3_backup_finish"
        Case P_BACKUP_REMAIN:  ProcName = "sqlite3_backup_remaining"
        Case P_BACKUP_PGCOUNT: ProcName = "sqlite3_backup_pagecount"
        Case P_BLOB_OPEN:      ProcName = "sqlite3_blob_open"
        Case P_BLOB_READ:      ProcName = "sqlite3_blob_read"
        Case P_BLOB_WRITE:     ProcName = "sqlite3_blob_write"
        Case P_BLOB_CLOSE:     ProcName = "sqlite3_blob_close"
        Case P_BLOB_BYTES:     ProcName = "sqlite3_blob_bytes"
        Case P_SERIALIZE:      ProcName = "sqlite3_serialize"
        Case P_DESERIALIZE:    ProcName = "sqlite3_deserialize"
        Case P_MALLOC:         ProcName = "sqlite3_malloc"
        Case P_FREE:           ProcName = "sqlite3_free"
        Case P_DB_STATUS:      ProcName = "sqlite3_db_status"
        Case P_STMT_STATUS:    ProcName = "sqlite3_stmt_status"
        Case Else:             ProcName = ""
    End Select
End Function

Private Function Invoke0(ByVal idx As Long) As LongPtr
    Dim vt(0) As Integer, ptrs(0) As LongPtr, args(0) As Variant, ret As Variant
    DispCallFunc CLngPtr(0), m_procs(idx), CC_CDECL, VT_I8, 0, vt(0), ptrs(0), ret
    Invoke0 = CLngPtr(ret)
End Function
