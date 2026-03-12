VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Backup.cls  -  Online Backup API wrapper (64-bit only)
'
' Wraps sqlite3_backup_init / _step / _finish / _remaining / _pagecount.
' The source database stays fully readable and writable during the backup.
' The destination is opened and closed by this class.
'
' Typical usage - one-shot backup:
'   Dim bak As New SQLite3Backup
'   bak.BackupToFile sourceConn, "C:\backups\mydb.bak.db"
'   Debug.Print bak.TotalPages & " pages copied"
'
' Incremental usage (e.g. with a progress callback):
'   bak.OpenBackup sourceConn, "C:\backups\mydb.bak.db"
'   Do While Not bak.IsComplete
'       bak.Step 100
'       Debug.Print bak.Progress * 100 & "% complete"
'   Loop
'   bak.CloseBackup
'
' Version : 0.1.5
'
' Version History:
'   0.1.4 - Initial release.
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

' -- Instance state -----------------------------------------------------------
Private m_pBackup    As LongPtr   ' sqlite3_backup* handle
Private m_pDestDb    As LongPtr   ' destination sqlite3* opened by this class
Private m_isOpen     As Boolean
Private m_steppedOnce As Boolean  ' True after the first sqlite3_backup_step call
Private m_totalPages As Long
Private m_lastError  As String

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Terminate()
    If m_isOpen Then CloseBackup
End Sub

'==============================================================================
' BackupToFile
' One-shot: open dest, copy everything, close. Blocks until complete.
' pagesPerStep controls how many pages are copied per internal loop iteration.
'   -1 = copy all pages in a single step (fastest but holds source write-lock
'        for the entire duration -- fine for small DBs or exclusive connections)
'   N  = copy N pages per step (yields between steps for long-running backups)
'==============================================================================
Public Sub BackupToFile(ByVal srcConn As SQLite3Connection, _
                         ByVal destPath As String, _
                         Optional ByVal pagesPerStep As Long = -1, _
                         Optional ByVal srcSchema As String = "main", _
                         Optional ByVal destSchema As String = "main")
    OpenBackup srcConn, destPath, srcSchema, destSchema
    ' sqlite3_backup_remaining returns 0 before the first Step call (per SQLite docs),
    ' so a Do-While-Not-IsComplete loop would exit immediately without copying anything.
    ' Use Do-Loop-Until to guarantee at least one Step executes.
    Dim done As Boolean
    Do
        done = Step(pagesPerStep)
    Loop Until done
    CloseBackup
End Sub

'==============================================================================
' OpenBackup
' Initialises the backup session without stepping. Call Step / CloseBackup
' manually for incremental control.
'==============================================================================
Public Sub OpenBackup(ByVal srcConn As SQLite3Connection, _
                       ByVal destPath As String, _
                       Optional ByVal srcSchema As String = "main", _
                       Optional ByVal destSchema As String = "main")
    If m_isOpen Then CloseBackup

    ' Open destination database
    Dim pathBytes() As Byte
    pathBytes = SQLite3_API.ToUTF8(destPath)
    Dim flags As Long
    flags = SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE Or SQLITE_OPEN_FULLMUTEX
    Dim rc As Long
    rc = sqlite3_open_v2(VarPtr(pathBytes(0)), m_pDestDb, flags, CLngPtr(0))
    If rc <> SQLITE_OK Then
        Err.Raise vbObjectError + 500, "SQLite3Backup.OpenBackup", _
                  "Cannot open destination DB: " & destPath & " (rc=" & rc & ")"
    End If

    ' Convert schema names to UTF-8
    Dim srcSchBytes() As Byte:  srcSchBytes  = SQLite3_API.ToUTF8(srcSchema)
    Dim destSchBytes() As Byte: destSchBytes = SQLite3_API.ToUTF8(destSchema)

    ' Init backup -- note arg order: dest first, then source
    m_pBackup = sqlite3_backup_init( _
        m_pDestDb, VarPtr(destSchBytes(0)), _
        srcConn.Handle, VarPtr(srcSchBytes(0)))

    If m_pBackup = 0 Then
        m_lastError = SQLite3_API.sqlite3_errmsg_str(m_pDestDb)
        sqlite3_close m_pDestDb
        m_pDestDb = 0
        Err.Raise vbObjectError + 501, "SQLite3Backup.OpenBackup", _
                  "sqlite3_backup_init failed: " & m_lastError
    End If

    m_isOpen = True
    m_steppedOnce = False
    m_totalPages = sqlite3_backup_pagecount(m_pBackup)
End Sub

'==============================================================================
' Step
' Copy up to nPages pages from source to destination.
'   nPages = -1 : copy all remaining pages in one call
'   nPages > 0  : copy at most nPages pages; return after (allows progress updates)
' Returns True when backup is complete (no more pages to copy).
'==============================================================================
Public Function Step(Optional ByVal nPages As Long = 100) As Boolean
    If Not m_isOpen Then
        Err.Raise vbObjectError + 502, "SQLite3Backup.Step", "Backup is not open."
    End If
    Dim rc As Long
    rc = sqlite3_backup_step(m_pBackup, nPages)
    m_steppedOnce = True
    ' Refresh total page count (may change if source DB grows during backup)
    m_totalPages = sqlite3_backup_pagecount(m_pBackup)
    Select Case rc
        Case SQLITE_OK:   Step = False  ' more pages to go
        Case SQLITE_DONE: Step = True   ' complete
        Case Else
            m_lastError = "sqlite3_backup_step rc=" & rc
            Err.Raise vbObjectError + 503, "SQLite3Backup.Step", m_lastError
    End Select
End Function

'==============================================================================
' CloseBackup
' Finalise the backup and close the destination DB.
' Must be called even if Step raised an error.
'==============================================================================
Public Sub CloseBackup()
    If Not m_isOpen Then Exit Sub
    m_isOpen = False
    If m_pBackup <> 0 Then
        sqlite3_backup_finish m_pBackup
        m_pBackup = 0
    End If
    If m_pDestDb <> 0 Then
        sqlite3_close m_pDestDb
        m_pDestDb = 0
    End If
End Sub

'==============================================================================
' Properties
'==============================================================================
' True when all pages have been copied.
' NOTE: sqlite3_backup_remaining returns 0 before the first Step call (SQLite docs),
' so IsComplete correctly returns False until at least one Step has executed.
Public Property Get IsComplete() As Boolean
    If Not m_isOpen Then IsComplete = True: Exit Property
    If Not m_steppedOnce Then IsComplete = False: Exit Property
    IsComplete = (sqlite3_backup_remaining(m_pBackup) = 0)
End Property

' Number of pages still to be copied.
Public Property Get PagesRemaining() As Long
    If Not m_isOpen Then Exit Property
    PagesRemaining = sqlite3_backup_remaining(m_pBackup)
End Property

' Total pages in the source database (updated each Step).
Public Property Get TotalPages() As Long
    TotalPages = m_totalPages
End Property

' Progress as a fraction 0.0 - 1.0.  Returns 1.0 when complete or not open.
Public Property Get Progress() As Double
    If Not m_isOpen Or m_totalPages = 0 Then Progress = 1#: Exit Property
    Dim remaining As Long: remaining = sqlite3_backup_remaining(m_pBackup)
    Progress = 1# - (CDbl(remaining) / CDbl(m_totalPages))
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m_isOpen
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property
