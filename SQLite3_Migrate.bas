Attribute VB_Name = "SQLite3_Migrate"
'==============================================================================
' SQLite3_Migrate.bas  -  Schema versioning and migration helpers (64-bit only)
'
' Stores the application schema version in SQLite's built-in PRAGMA user_version
' (a 32-bit integer, default 0). No extra table is required.
'
' Functions:
'   GetSchemaVersion(conn)               - read current PRAGMA user_version
'   SetSchemaVersion(conn, version)      - write PRAGMA user_version
'   ApplyMigration(conn, toVersion, sql) - apply sql if current version < toVersion,
'                                          then advance version to toVersion
'   MigrateAll(conn, migrations)         - apply an ordered array of migration steps
'
' Typical usage -- initial setup from version 0:
'
'   Dim steps(2) As MigrationStep
'   steps(0) = MakeStep(1, "CREATE TABLE accounts (id INTEGER PRIMARY KEY, name TEXT);")
'   steps(1) = MakeStep(2, "CREATE TABLE trades (id INTEGER PRIMARY KEY, acct_id INTEGER, qty REAL);")
'   steps(2) = MakeStep(3, "CREATE INDEX idx_trades_acct ON trades (acct_id);")
'   MigrateAll conn, steps
'
' Only steps whose toVersion is above the current schema version are applied,
' so it is safe to call MigrateAll on every workbook open.
'
' Version : 0.1.6
'
' Version History:
'   0.1.6 - Initial release.
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

' A single migration step: the version it brings the schema TO, plus the SQL
' that achieves it. Build with MakeStep() for convenient array literals.
Public Type MigrationStep
    toVersion As Long
    sql       As String
End Type

'==============================================================================
' MakeStep  -  convenience constructor for MigrationStep
'==============================================================================
Public Function MakeStep(ByVal toVersion As Long, ByVal sql As String) As MigrationStep
    MakeStep.toVersion = toVersion
    MakeStep.sql       = sql
End Function

'==============================================================================
' GetSchemaVersion  -  read PRAGMA user_version
'==============================================================================
Public Function GetSchemaVersion(ByVal conn As SQLite3Connection) As Long
    Dim v As Variant
    v = QueryScalar(conn, "PRAGMA user_version;")
    If IsNull(v) Or IsEmpty(v) Then
        GetSchemaVersion = 0
    Else
        GetSchemaVersion = CLng(v)
    End If
End Function

'==============================================================================
' SetSchemaVersion  -  write PRAGMA user_version
'
' PRAGMA user_version cannot be set via a bound parameter -- the value must be
' written inline. version is a Long (not user-supplied string) so injection is
' not a concern.
'==============================================================================
Public Sub SetSchemaVersion(ByVal conn As SQLite3Connection, ByVal version As Long)
    conn.ExecSQL "PRAGMA user_version = " & CStr(version) & ";"
End Sub

'==============================================================================
' ApplyMigration
' Apply sql inside a transaction if the current schema version is below
' toVersion, then advance user_version to toVersion.
'
' Parameters:
'   conn      - open SQLite3Connection
'   toVersion - the schema version this migration brings the DB to
'   sql       - DDL / DML to execute (may contain multiple statements)
'
' Returns True if the migration was applied, False if it was skipped (already
' at or above toVersion).
'==============================================================================
Public Function ApplyMigration(ByVal conn As SQLite3Connection, _
                                ByVal toVersion As Long, _
                                ByVal sql As String) As Boolean
    If GetSchemaVersion(conn) >= toVersion Then
        ApplyMigration = False
        Exit Function
    End If

    conn.BeginTransaction
    On Error GoTo fail
    conn.ExecSQL sql
    SetSchemaVersion conn, toVersion
    conn.CommitTransaction
    ApplyMigration = True
    Exit Function

fail:
    conn.RollbackTransaction
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'==============================================================================
' MigrateAll
' Apply every step in the migrations array whose toVersion exceeds the current
' schema version, in array order. Each step runs in its own transaction.
'
' Raises an error (and rolls back the failing step) if any step fails; steps
' already applied before the failure are committed and remain in effect.
'
' Parameters:
'   conn       - open SQLite3Connection
'   migrations - array of MigrationStep (build with MakeStep)
'
' Returns the number of steps actually applied.
'==============================================================================
Public Function MigrateAll(ByVal conn As SQLite3Connection, _
                             ByRef migrations() As MigrationStep) As Long
    Dim applied As Long: applied = 0
    Dim i As Long
    For i = LBound(migrations) To UBound(migrations)
        If ApplyMigration(conn, migrations(i).toVersion, migrations(i).sql) Then
            applied = applied + 1
        End If
    Next i
    MigrateAll = applied
End Function
