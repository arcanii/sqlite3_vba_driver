VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLite3Pool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' SQLite3Pool.cls  -  Connection Pool (64-bit only)
'
' VBA is single-threaded within a COM apartment. Pool provides:
'   - Re-entrant safety (checked-out connections cannot be double-leased)
'   - Automatic connection creation when pool is exhausted (up to max)
'   - Idle connection reaping after configurable timeout
'   - Auto-rollback on release for safety
'
' Usage:
'   Dim pool As New SQLite3Pool
'   pool.Initialize "C:\data\quant.db", "C:\bin\sqlite3.dll", 4
'
'   Dim conn As SQLite3Connection
'   Set conn = pool.Acquire()
'   ' ... use conn ...
'   pool.ReleaseConnection conn
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

Private Type PoolSlot
    conn     As SQLite3Connection
    inUse    As Boolean
    lastUsed As Double
End Type

Private Const MAX_POOL_SIZE As Long = 16

Private m_slots()     As PoolSlot
Private m_slotCount   As Long
Private m_dbPath      As String
Private m_dllPath     As String
Private m_maxSize     As Long
Private m_busyMs      As Long
Private m_mmapSize    As Long
Private m_enableWAL   As Boolean
Private m_idleTimeout As Double
Private m_isInit      As Boolean

'==============================================================================
' Lifecycle
'==============================================================================
Private Sub Class_Initialize()
    m_maxSize = 4
    m_busyMs = 5000
    m_mmapSize = 0
    m_enableWAL = True
    m_idleTimeout = 300
    m_slotCount = 0
    ReDim m_slots(MAX_POOL_SIZE - 1)
    m_isInit = False
End Sub

Private Sub Class_Terminate()
    ShutDown
End Sub

Public Sub Initialize(ByVal dbPath As String, _
                      Optional ByVal dllPath As String = "", _
                      Optional ByVal maxConnections As Long = 4, _
                      Optional ByVal busyTimeoutMs As Long = 5000, _
                      Optional ByVal enableWAL As Boolean = True, _
                      Optional ByVal mmapSizeBytes As Long = 0, _
                      Optional ByVal idleTimeoutSec As Double = 300)

    ' Shut down any existing connections cleanly before re-initialising
    If m_isInit Then ShutDown

    m_dbPath = dbPath
    m_dllPath = dllPath
    m_maxSize = IIf(maxConnections > MAX_POOL_SIZE, MAX_POOL_SIZE, maxConnections)
    m_busyMs = busyTimeoutMs
    m_enableWAL = enableWAL
    m_mmapSize = mmapSizeBytes
    m_idleTimeout = idleTimeoutSec

    ' Explicitly zero every slot so no stale object refs remain
    ReDim m_slots(MAX_POOL_SIZE - 1)
    Dim i As Long
    For i = 0 To MAX_POOL_SIZE - 1
        Set m_slots(i).conn = Nothing
        m_slots(i).inUse = False
        m_slots(i).lastUsed = 0
    Next i
    m_slotCount = 0
    m_isInit = True

    ' Pre-warm one connection
    Dim warm As SQLite3Connection
    Set warm = CreateConnection()
    Set m_slots(0).conn = warm
    m_slots(0).inUse = False
    m_slots(0).lastUsed = Timer()
    m_slotCount = 1
End Sub

'==============================================================================
' Acquire - get a connection from the pool (or create a new one)
'==============================================================================
Public Function Acquire() As SQLite3Connection
    AssertInit
    ReapIdleConnections

    Dim i As Long
    For i = 0 To m_slotCount - 1
        If Not m_slots(i).inUse Then
            If m_slots(i).conn.IsOpen Then
                m_slots(i).inUse = True
                m_slots(i).lastUsed = Timer()
                Set Acquire = m_slots(i).conn
                Exit Function
            End If
        End If
    Next i

    If m_slotCount < m_maxSize Then
        Dim newConn As SQLite3Connection
        Set newConn = CreateConnection()
        Set m_slots(m_slotCount).conn = newConn
        m_slots(m_slotCount).inUse = True
        m_slots(m_slotCount).lastUsed = Timer()
        m_slotCount = m_slotCount + 1
        Set Acquire = newConn
    Else
        Err.Raise vbObjectError + 700, "SQLite3Pool.Acquire", _
                  "Connection pool exhausted (max=" & m_maxSize & ")."
    End If
End Function

'==============================================================================
' ReleaseConnection - return a connection to the pool
'==============================================================================
Public Sub ReleaseConnection(ByVal conn As SQLite3Connection)
    Dim i As Long
    For i = 0 To m_slotCount - 1
        If m_slots(i).conn Is conn Then
            If conn.InTransaction Then conn.RollbackTransaction
            m_slots(i).inUse = False
            m_slots(i).lastUsed = Timer()
            Exit Sub
        End If
    Next i
    conn.CloseConnection
End Sub

'==============================================================================
' Statistics
'==============================================================================
Public Property Get PoolSize() As Long
    PoolSize = m_slotCount
End Property

Public Property Get ActiveConnections() As Long
    Dim i As Long, cnt As Long
    For i = 0 To m_slotCount - 1
        If m_slots(i).inUse Then cnt = cnt + 1
    Next i
    ActiveConnections = cnt
End Property

Public Property Get IdleConnections() As Long
    IdleConnections = m_slotCount - ActiveConnections
End Property

'==============================================================================
' ShutDown - close all pool connections
'==============================================================================
Public Sub ShutDown()
    Dim i As Long
    For i = 0 To m_slotCount - 1
        If Not m_slots(i).conn Is Nothing Then
            If m_slots(i).conn.IsOpen Then m_slots(i).conn.CloseConnection
            Set m_slots(i).conn = Nothing
        End If
    Next i
    m_slotCount = 0
    m_isInit = False
End Sub

'==============================================================================
' Private helpers
'==============================================================================
Private Function CreateConnection() As SQLite3Connection
    Dim conn As New SQLite3Connection
    conn.OpenDatabase m_dbPath, m_dllPath, m_busyMs, m_enableWAL, m_mmapSize
    Set CreateConnection = conn
End Function

Private Sub ReapIdleConnections()
    Dim i As Long, keepCount As Long
    For i = 0 To m_slotCount - 1
        If Not m_slots(i).inUse Then keepCount = keepCount + 1
    Next i
    Dim t As Double: t = Timer()
    For i = 0 To m_slotCount - 1
        If Not m_slots(i).inUse Then
            If (t - m_slots(i).lastUsed) > m_idleTimeout And keepCount > 1 Then
                m_slots(i).conn.CloseConnection
                Set m_slots(i).conn = Nothing
                CompactSlots i
                keepCount = keepCount - 1
                i = i - 1
            End If
        End If
    Next i
End Sub

Private Sub CompactSlots(ByVal removedIdx As Long)
    Dim i As Long
    For i = removedIdx To m_slotCount - 2
        Set m_slots(i).conn = m_slots(i + 1).conn
        m_slots(i).inUse = m_slots(i + 1).inUse
        m_slots(i).lastUsed = m_slots(i + 1).lastUsed
    Next i
    m_slotCount = m_slotCount - 1
    Set m_slots(m_slotCount).conn = Nothing
    m_slots(m_slotCount).inUse = False
    m_slots(m_slotCount).lastUsed = 0
End Sub

Private Sub AssertInit()
    If Not m_isInit Then
        Err.Raise vbObjectError + 710, "SQLite3Pool", _
                  "Pool not initialized. Call Initialize() first."
    End If
End Sub
