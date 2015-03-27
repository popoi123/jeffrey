Imports System.Data.SqlClient
Public Class Form1
    Dim con As SqlConnection = New SqlConnection("Data Source=ADMIN-PC\SQLEXPRESS; Database=db_project; Trusted_Connection =yes;")
    Dim dr As SqlDataReader
    Dim cmd As SqlCommand
    Dim da As SqlDataAdapter
    Dim ds As DataSet
    Dim connection As String
    Friend ID As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        listviewthree()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Panel2.Hide()
        Panel4.Hide()
        Do While Panel3.Height < 390
            Panel3.Height = Panel3.Height + 10
        Loop
        Panel5.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Do While Panel3.Height > 17
            Panel3.Height = Panel3.Height - 10
        Loop
        Panel2.Show()
        Panel4.Show()
        Panel5.Hide()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If ID = Nothing Then
            MsgBox("Please choose a record to edit.", MsgBoxStyle.Exclamation)
        Else
            Panel2.Hide()
            Do While Panel4.Height < 390
                Panel4.Height = Panel4.Height + 10
            Loop
            Panel6.Show()

            Dim cmd As SqlCommand
            Dim query As String = "SELECT * FROM tbl_info WHERE stud_id='" + ListView1.SelectedItems(0).Text + "'"
            cmd = New SqlCommand(query, con)
            Try
                con.Open()
                Dim myreader As SqlDataReader = cmd.ExecuteReader()
                If myreader.Read() Then
                    Label3.Text = myreader.GetValue(0)
                    TextBox13.Text = myreader.GetValue(1)
                    TextBox12.Text = myreader.GetValue(2)
                    TextBox3.Text = myreader.GetValue(3)
                    TextBox8.Text = myreader.GetValue(5)
                    ComboBox2.Text = myreader.GetValue(6)
                    DateTimePicker4.Text = myreader.GetValue(7)
                    TextBox11.Text = myreader.GetValue(4)
                    DateTimePicker3.Text = myreader.GetValue(8)
                End If
                myreader.Close()
            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
            ID = Nothing
            listviewthree()

        End If
        con.Close()

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Do While Panel4.Height > 17
            Panel4.Height = Panel4.Height - 10
        Loop
        Panel2.Show()
        Panel6.Hide()
    End Sub
    Private Sub listviewthree()
        Dim str As String = "Select * From tbl_info"
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter
        Dim TABLE As New DataTable
        Dim i As Integer

        With cmd
            .CommandText = str
            .Connection = con
        End With
        With da
            .SelectCommand = cmd
            .Fill(TABLE)
        End With

        ListView1.Items.Clear()

        For i = 0 To TABLE.Rows.Count - 1
            With ListView1
                .Items.Add(TABLE.Rows(i)("stud_id"))

                With .Items(.Items.Count - 1).SubItems
                    'Respondent Profile
                    .Add(TABLE.Rows(i)("fisrtname"))
                    .Add(TABLE.Rows(i)("middlename"))
                    .Add(TABLE.Rows(i)("lastname"))
                    .Add(TABLE.Rows(i)("address"))
                    .Add(TABLE.Rows(i)("number"))
                    .Add(TABLE.Rows(i)("gender"))
                    .Add(TABLE.Rows(i)("birthday"))
                    .Add(TABLE.Rows(i)("date"))
                    .Add(TABLE.Rows(i)("age"))
                End With
            End With
        Next
        con.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        cmd = New SqlCommand("Insert Into tbl_info (fisrtname, middlename, lastname, age, number, gender, address, date, birthday) Values ('" + TextBox4.Text + "', '" + TextBox5.Text + "', '" + TextBox10.Text + "', '" + TextBox6.Text + "', '" + TextBox9.Text + "', '" + ComboBox1.Text + "', '" + TextBox7.Text + "', '" + DateTimePicker2.Text + "', '" + DateTimePicker1.Text + "')", con)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet()
        da.Fill(ds, "db_info")
        MsgBox("Save")
        listviewthree()

        Do While Panel3.Height > 17
            Panel3.Height = Panel3.Height - 10
        Loop
        Panel2.Show()
        Panel4.Show()
        Panel5.Hide()
    End Sub
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        With DateTimePicker1.Value
            Dim celebrate As DateTime = New DateTime(Now.Year, .Month, .Day)
            Dim age As Integer = Now.Year - .Year
            If celebrate > Now Then age -= 1
            TextBox6.Text = CStr(age)
        End With
    End Sub

    Private Sub ListView1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseClick
        ID = ListView1.SelectedItems(0).Text
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If ID = Nothing Then
            MsgBox("Please choose a record to delete.", MsgBoxStyle.Exclamation)
        Else
            Dim result As Integer = MessageBox.Show("Do you want to delete this item with ID#" + ListView1.SelectedItems(0).Text + "?", "caption", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then

                Try

                    Dim str As String = "DELETE from tbl_info where stud_id = '" + ListView1.SelectedItems(0).Text + "'"
                    Dim da As New SqlDataAdapter(str, con)
                    Dim ds As New DataSet
                    da.Fill(ds, "db_project")
                Catch ex As Exception
                    MsgBox(ex.Message)

                End Try
                MsgBox("Information Deleted!")
            End If
            ID = Nothing
            listviewthree()

        End If
    End Sub

    Private Sub DateTimePicker4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker4.ValueChanged
        With DateTimePicker4.Value
            Dim celebrate As DateTime = New DateTime(Now.Year, .Month, .Day)
            Dim age As Integer = Now.Year - .Year
            If celebrate > Now Then age -= 1
            TextBox2.Text = CStr(age)
        End With
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        cmd = New SqlCommand("Update tbl_info Set fisrtname = '" + TextBox13.Text + "', middlename = '" + TextBox12.Text + "', lastname = '" + TextBox3.Text + "', number = '" + TextBox8.Text + "', Gender = '" + ComboBox2.Text + "', birthday = '" + DateTimePicker4.Text + "',age = '" + TextBox2.Text + "',address = '" + TextBox11.Text + "', date = '" + DateTimePicker3.Text + "' where stud_id = '" + Label3.Text + "'", con)
        da = New SqlDataAdapter(cmd)
        ds = New DataSet()
        da.Fill(ds, "db_info")
        MsgBox("Save")
        listviewthree()

        Do While Panel4.Height > 17
            Panel4.Height = Panel4.Height - 10
        Loop
        Panel2.Show()
        Panel6.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim str As String = "Select * From tbl_info where fisrtname = '" + TextBox1.Text + "' Or middlename = '" + TextBox1.Text + "'Or lastname = '" + TextBox1.Text + "'"
        Dim cmd As New SqlCommand
        Dim da As New SqlDataAdapter
        Dim TABLE As New DataTable
        Dim i As Integer

        With cmd
            .CommandText = str
            .Connection = con
        End With

        With da
            .SelectCommand = cmd
            .Fill(TABLE)
        End With

        ListView1.Items.Clear()

        For i = 0 To TABLE.Rows.Count - 1
            With ListView1
                .Items.Add(TABLE.Rows(i)("stud_id"))

                With .Items(.Items.Count - 1).SubItems
                    'Respondent Profile
                    .Add(TABLE.Rows(i)("fisrtname"))
                    .Add(TABLE.Rows(i)("middlename"))
                    .Add(TABLE.Rows(i)("lastname"))
                    .Add(TABLE.Rows(i)("address"))
                    .Add(TABLE.Rows(i)("number"))
                    .Add(TABLE.Rows(i)("gender"))
                    .Add(TABLE.Rows(i)("birthday"))
                    .Add(TABLE.Rows(i)("date"))
                    .Add(TABLE.Rows(i)("age"))

                End With
            End With
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        listviewthree()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        LoginForm1.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        Dim selStart As Integer = TextBox9.SelectionStart
        Dim selMoveLeft As Integer = 0
        Dim newStr As String = "" 'Build a new string by copying each valid character from the existing string. The new string starts as blank and valid characters are added 1 at a time.

        For i As Integer = 0 To TextBox9.Text.Length - 1

            If "0123456789".IndexOf(TextBox9.Text(i)) <> -1 Then 'Characters that are in the allowed set will be added to the new string.
                newStr = newStr & TextBox9.Text(i)

            ElseIf i < selStart Then 'Characters that are not valid are removed - if these characters are before the cursor, we need to move the cursor left to account for their removal.
                selMoveLeft = selMoveLeft + 1

            End If
        Next

        TextBox9.Text = newStr 'Place the new text into the textbox.
        TextBox9.SelectionStart = selStart - selMoveLeft 'Move the cursor to the appropriate location.

    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Dim selStart As Integer = TextBox8.SelectionStart
        Dim selMoveLeft As Integer = 0
        Dim newStr As String = "" 'Build a new string by copying each valid character from the existing string. The new string starts as blank and valid characters are added 1 at a time.

        For i As Integer = 0 To TextBox8.Text.Length - 1

            If "0123456789".IndexOf(TextBox8.Text(i)) <> -1 Then 'Characters that are in the allowed set will be added to the new string.
                newStr = newStr & TextBox8.Text(i)

            ElseIf i < selStart Then 'Characters that are not valid are removed - if these characters are before the cursor, we need to move the cursor left to account for their removal.
                selMoveLeft = selMoveLeft + 1

            End If
        Next

        TextBox8.Text = newStr 'Place the new text into the textbox.
        TextBox8.SelectionStart = selStart - selMoveLeft 'Move the cursor to the appropriate location.

    End Sub

    Private Sub TextBox13_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox13.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub TextBox12_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox12.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub TextBox4_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub

    Private Sub TextBox10_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox10.KeyPress
        If (Microsoft.VisualBasic.Asc(e.KeyChar) < 65) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 90) _
            And (Microsoft.VisualBasic.Asc(e.KeyChar) < 97) _
            Or (Microsoft.VisualBasic.Asc(e.KeyChar) > 122) Then
            'Allowed space
            If (Microsoft.VisualBasic.Asc(e.KeyChar) <> 32) Then
                e.Handled = True
            End If
        End If
        ' Allowed backspace
        If (Microsoft.VisualBasic.Asc(e.KeyChar) = 8) Then
            e.Handled = False
        End If
    End Sub
End Class
