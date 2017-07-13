Imports System.Data.SqlClient

Module ModuleConnection
    Public connection As SqlConnection 'connection
    'functon for connection 
    Sub OpenConnection()
        Try
            connection = New SqlConnection("Data Source=N550JK\SQLEXPRESS01; Initial Catalog=Users; Integrated Security=True;")
            If connection.State = ConnectionState.Closed Then
                connection.Open()
            End If
        Catch ex As Exception
            MsgBox("Failed to Connect, Error " & ex.ToString)
        End Try
    End Sub
End Module
