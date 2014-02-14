Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.IO

Public Class UpdateForm
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim mymatch1 As Match = System.Text.RegularExpressions.Regex.Match(TextBox3.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch2 As Match = System.Text.RegularExpressions.Regex.Match(TextBox4.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch3 As Match = System.Text.RegularExpressions.Regex.Match(TextBox5.Text, "^\d+$")

        Dim mymatch8 As Match = System.Text.RegularExpressions.Regex.Match(TextBox17.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch9 As Match = System.Text.RegularExpressions.Regex.Match(TextBox18.Text, "^\w{1,50}|\w+\s{1}\w$")

        If Not mymatch1.Success Then
            MessageBox.Show("Please Enter Party-Name OR Select Exisiting Entry. ", "Error")
        ElseIf Not mymatch2.Success Then
            MessageBox.Show("Please Enter Party Address OR Select Exisiting Entry.", "Error")
        ElseIf Not mymatch3.Success Then
            MessageBox.Show("Please Enter Contact Number OR Select Exisiting Entry.", "Error")

        ElseIf Not mymatch8.Success Then
            MessageBox.Show("Please Enter Vendor Code OR Select Exisiting Entry.", "Error")
        ElseIf Not mymatch9.Success Then
            MessageBox.Show("Please Enter Client Number OR Select Exisiting Entry.", "Error")
        Else

            Dim con As OleDbConnection
            Dim in_srno As Integer = 0
            Dim srno As Integer = 0

            Dim p_name, p_code, p_addr, ctct, v_code, client_name, email, vat_num As String
            Dim strsql As String

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            Try
                con.Open()
                p_code = MainForm.TextBox2.Text
                p_name = Me.TextBox3.Text
                p_addr = Me.TextBox4.Text
                ctct = Me.TextBox5.Text
                v_code = Me.TextBox17.Text
                client_name = Me.TextBox18.Text
                email = Me.TextBox25.Text
                vat_num = Me.TextBox26.Text

                Dim str1 As String = CStr(srno)
                Dim str2 As String = "PAR"

                strsql = "Update atable set v_code='" & v_code & "',p_name='" & p_name & "',p_addr='" & p_addr & "',client_name='" & client_name & "',ctct='" & ctct & "',email='" & email & "',vat_num='" & vat_num & "' where p_code='" & p_code & "'"

                Dim x As Integer
                Dim sql As New OleDbCommand(strsql, con)
                x = sql.ExecuteNonQuery()

            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error")
                con.Close()
            End Try
            MsgBox("Record updated successfully.", MsgBoxStyle.Information, "Success")
            con.Close()
            MainForm.ListBox1_SelectedIndexChanged(sender, e)
            Me.Close()
        End If
    End Sub

    Private Sub UpdateForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim con As OleDbConnection
            Dim cmd As OleDbCommand
            Dim rs As OleDbDataReader

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            con.Open()
            cmd = New OleDbCommand("select * from atable where p_code = '" & MainForm.TextBox2.Text & "'", con)
            rs = cmd.ExecuteReader()

            While rs.Read()
                TextBox2.Text = rs(1)
                TextBox17.Text = rs(2)
                TextBox3.Text = rs(3)
                TextBox4.Text = rs(4)
                TextBox18.Text = rs(5)
                TextBox5.Text = rs(6)
                TextBox25.Text = rs(7)
                TextBox26.Text = rs(8)

            End While
            con.Close()
        Catch ex As Exception
            MsgBox("Error in Populating UpdateForm.")
        End Try
    End Sub
End Class