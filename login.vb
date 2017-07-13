Imports System.Data.SqlClient

Public Class login

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        OpenConnection()

        Dim dr As SqlDataReader
        Dim cmd As SqlCommand
        Dim sql As String

        cmd = New SqlCommand
        cmd.CommandType = CommandType.Text
        cmd.Connection = connection
        sql = "select * from users where email= '" & TextBox1.Text & "' and password=CONVERT(NVARCHAR(32), HashBytes('MD5','" & TextBox2.Text & "'),2)"
        cmd.CommandText = sql
        dr = cmd.ExecuteReader()

        If dr.HasRows Then
            dr.Read()
            If dr.Item("level") = 1 Then
                MsgBox("Welcome Admin")
            ElseIf dr.Item("level") = 2 Then
                MsgBox("Welcome User")
            ElseIf dr.Item("level") = 3 Then
                MsgBox("Welcome Guest")
            End If

        Else

            MsgBox("Access denied!")

        End If

        connection.Close()
        cmd.Dispose()

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        register.Show()
    End Sub
End Class
