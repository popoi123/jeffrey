Imports System.Data.SqlClient
Public Class LoginForm1
    Dim con As SqlConnection
    Dim dr As SqlDataReader
    Dim cmd As SqlCommand
    Dim connection As String

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If txtbx_username.Text = "" And txtbx_password.Text = "" Then
            MessageBox.Show("Enter Username or Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If txtbx_username.Text = "username" Then
            MessageBox.Show("Enter Username", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtbx_username.Focus()
            Exit Sub
        End If

        If txtbx_password.Text = "password" Then
            MessageBox.Show("Enter Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub

        End If

        connection = "data source=ADMIN-PC\SQLEXPRESS; initial catalog= db_medical ; integrated security = SSPI"

        con = New SqlConnection(connection)

        Try
            con.Open()
        Catch ex As Exception
            MessageBox.Show("connection error")
            Exit Sub
        Finally

            con.Close()
        End Try
        loginuser()

    End Sub

    Private Sub loginuser()

        cmd = New SqlCommand("select Username, Password from tbl_login where Password='" & txtbx_password.Text & "' and Username='" & txtbx_username.Text & "'", con)
        cmd.CommandType = CommandType.Text
        con.Open()
        dr = cmd.ExecuteReader
        dr.Read()
        If dr.HasRows Then
            Form1.Show()
            Me.Hide()
            txtbx_password.Clear()
            txtbx_username.Clear()
        Else
            MessageBox.Show("Incorrect Username or Password", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        con.Close()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
