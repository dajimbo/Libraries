Imports System.Data.SqlClient
Public Class ParaLibrary
    Public Shared Function getConnection(server As String, database As String, uid As String, pwd As String) As SqlConnection
        Dim connectionString = "Server = " + server + ";" + "Database = " + database + ";" _
                                + "UID = " + uid + ";" + "PWD = " + pwd
        Dim Connection As New SqlConnection()
        Try
            Connection.ConnectionString = connectionString
            Connection.Open()
        Catch ex As Exception
            logErrorsToDatabase(ex)
        End Try
        Return Connection
    End Function

    Public Shared Sub executeNonQuery(query As String, connection As SqlConnection)
            Try
                Dim command As New SqlCommand(query, connection)
                command.ExecuteNonQuery()
            Catch ex As Exception
                logErrorsToDatabase(ex)
            Finally
                connection.Close()
            End Try
        End Sub
        Public Shared Function executeQuery(query As String, connection As SqlConnection) As DataTable
            Dim data As New DataTable()
            Try
                Using command As New SqlDataAdapter(query, connection)
                    command.Fill(data)
                End Using
            Catch ex As Exception
                logErrorsToDatabase(ex)
            Finally
                connection.Close()
            End Try
            Return data
        End Function


    Public Shared Sub logErrorsToDatabase(ex As Exception)
        Dim connectionString = "Server = DESKTOP-H5FMBDG\WEBDB; Database =  ParaAuditTrail; UID = sa; pwd = localadmin"
        Dim errorSqlConnection As New SqlConnection(connectionString)
        Dim command As New SqlCommand("INSERT INTO  ErrorLog VALUES ('" + ex.Source.Replace("'", "") + "','" + ex.Message.Replace("'", "") + "','" _
                                            + ex.TargetSite.Name.Replace("'", "") + "')", errorSqlConnection)
        Try
            errorSqlConnection.Open()
            command.ExecuteNonQuery()
        Catch ex
            Console.WriteLine(ex.Message)
        Finally
            errorSqlConnection.Close()
        End Try
    End Sub
End Class


