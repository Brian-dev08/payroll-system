Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms

Module Connection

    ' === LocalDB file path (inside /Database folder) ===
    Public dbPath As String = Path.Combine(Application.StartupPath, "Database\payroll.mdf")

    ' === Connection string using DataDirectory placeholder ===
    ' |DataDirectory| ensures compatibility when published or moved to another PC.
    Public connString As String = "Data Source=(LocalDB)\MSSQLLocalDB;" &
                                  "AttachDbFilename=|DataDirectory|\Database\payroll.mdf;" &
                                  "Integrated Security=True;" &
                                  "Connect Timeout=30;"

    ' === Shared SQL connection ===
    Public Connect As New SqlConnection(connString)

    ' === Shared objects ===
    Public Parameters As New List(Of SqlParameter)
    Public Datacount As Integer

    ' === Open connection safely ===
    Public Sub OpenConnection()
        Try
            If Connect.State = ConnectionState.Closed Then
                Connect.Open()
            End If
        Catch ex As Exception
            MessageBox.Show(
                "❌ Connection failed: " & ex.Message & vbCrLf &
                "📂 Database path: " & dbPath,
                "Database Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            )
        End Try
    End Sub

    ' === Close connection safely ===
    Public Sub CloseConnection()
        Try
            If Connect.State = ConnectionState.Open Then
                Connect.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("⚠️ Error closing connection: " & ex.Message)
        End Try
    End Sub

    ' === Helper to add SQL parameters ===
    Public Sub AddParam(ByVal key As String, ByVal value As Object)
        Parameters.Add(New SqlParameter(key, value))
    End Sub

    ' === Execute INSERT query ===
    Public Function Insert(ByVal insertQuery As String) As Boolean
        Try
            OpenConnection()
            Using command As New SqlCommand(insertQuery, Connect)
                If Parameters.Count > 0 Then
                    command.Parameters.AddRange(Parameters.ToArray())
                    Parameters.Clear()
                End If
                Datacount = command.ExecuteNonQuery()
                Return Datacount > 0
            End Using
        Catch ex As Exception
            MsgBox("Insert failed: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        Finally
            CloseConnection()
        End Try
    End Function

    ' === Execute UPDATE query ===
    Public Function Update(ByVal updateQuery As String) As Boolean
        Try
            OpenConnection()
            Using command As New SqlCommand(updateQuery, Connect)
                If Parameters.Count > 0 Then
                    command.Parameters.AddRange(Parameters.ToArray())
                    Parameters.Clear()
                End If
                Datacount = command.ExecuteNonQuery()
                Return Datacount > 0
            End Using
        Catch ex As Exception
            MsgBox("Update failed: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        Finally
            CloseConnection()
        End Try
    End Function

    ' === Execute DELETE query ===
    Public Function Delete(ByVal deleteQuery As String) As Boolean
        Try
            OpenConnection()
            Using command As New SqlCommand(deleteQuery, Connect)
                If Parameters.Count > 0 Then
                    command.Parameters.AddRange(Parameters.ToArray())
                    Parameters.Clear()
                End If
                Datacount = command.ExecuteNonQuery()
                Return Datacount > 0
            End Using
        Catch ex As Exception
            MsgBox("Delete failed: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        Finally
            CloseConnection()
        End Try
    End Function

End Module
