Imports System.Data.SqlClient

Public Class register

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim TypeUser As String
        Dim Level As Integer

        TypeUser = ComboBox1.Text

        If TypeUser = "Administrator" Then
            Level = 1
        ElseIf TypeUser = "User" Then
            Level = 2
        ElseIf TypeUser = "Guest" Then
            Level = 3
        End If

        Dim sql As String = "insert into users(email, password, level)values('" & TextBox1.Text & "',CONVERT(NVARCHAR(32), HashBytes('MD5', '" & TextBox2.Text & "'),2)," & Level & ")"

        SignUp(sql)
        MsgBox("Sign Up Successful!")


    End Sub

    Private Sub SignUp(ByVal sql As String)

        Dim cmd As New SqlCommand

        OpenConnection()
        Try
            cmd.Connection = connection
            cmd.CommandText = CommandType.Text
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connection.Close()
        Catch ex As Exception
            MsgBox("Failed " & ex.ToString)
        End Try


    End Sub

    Private Sub Register_load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Text = "Administrator"
    End Sub
End Class