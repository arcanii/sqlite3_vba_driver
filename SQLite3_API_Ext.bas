Attribute VB_Name = "SQLite3_API_Ext"
'==============================================================================
' SQLite3_API_Ext.bas  -  Auxiliary dispatch bridge (64-bit only)
' Holds a secondary handle copy so helper modules can call less-common
' SQLite3 functions without coupling to the private proc array in SQLite3_API.
'
' Version : 0.1.5
'
' Version History:
'   0.1.0 - Initial release. Secondary DLL handle copy, SetDllHandle,
'            DispatchProc bridge for auxiliary calls.
'   0.1.1 - Added explicit DispCallFunc Declare (prgpvarg ByRef As LongPtr).
'            Removed illegal Declare-inside-Sub and duplicate load logic.
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

' -- Win32 declaration (must be at top) --------------------------------------
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" _
    (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As Long, _
     ByVal vtReturn As Integer, ByVal cActuals As Long, _
     ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, _
     ByRef pvargResult As Variant) As Long

' -- Module state -------------------------------------------------------------
Private m_hDll          As LongPtr
Private m_pBindParamIdx As LongPtr

'==============================================================================
' Called by SQLite3_API.SQLite_Load immediately after DLL is loaded
'==============================================================================
Public Sub SetDllHandle(ByVal hDll As LongPtr)
    m_hDll          = hDll
    m_pBindParamIdx = GetProcAddress(m_hDll, "sqlite3_bind_parameter_index")
End Sub

'==============================================================================
' General-purpose dispatcher for auxiliary procs not wrapped in SQLite3_API.
' argCount, args(), vt() and retVT follow the same conventions as
' the wrappers in SQLite3_API.bas.
'==============================================================================
Public Function DispatchProc(ByVal pProc As LongPtr, _
                              ByRef args() As Variant, _
                              ByRef vt() As Integer, _
                              ByVal argCount As Integer, _
                              ByVal retVT As Integer) As Long
    If pProc = 0 Then Exit Function
    Dim ret As Variant
    DispCallFunc pProc, 0&, CC_CDECL, retVT, argCount, vt(0), args(0), ret
    DispatchProc = CLng(ret)
End Function

' Expose the bind_parameter_index proc handle for callers that need it
Public Function BindParamIdxProc() As LongPtr
    BindParamIdxProc = m_pBindParamIdx
End Function
