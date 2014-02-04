Option Explicit On
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.IO

Public Class Form3

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim cnt As Integer = 0
        Label30.Text = cnt
        While cnt = CInt(Form1.ComboBox3.SelectedItem)
            Label30.Text = cnt

            Dim flag As Integer = 0
            Dim st As Integer = 0

            Dim mymatch1 As Match = System.Text.RegularExpressions.Regex.Match(Form1.TextBox3.Text, "^\w{1,50}|\w+\s{1}\w$")
            Dim mymatch2 As Match = System.Text.RegularExpressions.Regex.Match(Form1.TextBox4.Text, "^\w{1,50}|\w+\s{1}\w$")
            Dim mymatch3 As Match = System.Text.RegularExpressions.Regex.Match(Form1.TextBox5.Text, "^\d+$")

            Dim mymatch4 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\d+$")
            Dim mymatch5 As Match = System.Text.RegularExpressions.Regex.Match(TextBox10.Text, "^\d+$")
            Dim mymatch6 As Match = System.Text.RegularExpressions.Regex.Match(TextBox11.Text, "^\d+$")
            Dim mymatch7 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\w{1,50}|\w+\s{1}\w$")



            If Not mymatch1.Success Then
                MessageBox.Show("Please Enter Party-Name OR Select Exisiting Entry. ", "Error")
            ElseIf Not mymatch2.Success Then
                MessageBox.Show("Please Enter Party Address OR Select Exisiting Entry.", "Error")
            ElseIf Not mymatch3.Success Then
                MessageBox.Show("Please Enter Contact Number OR Select Exisiting Entry.", "Error")

            Else
                If Not mymatch7.Success Then
                    MessageBox.Show("Please Enter the valid Description. ", "Error")
                ElseIf Not mymatch4.Success Then
                    MessageBox.Show("Please Enter the valid Quantity. ", "Error")
                ElseIf Not mymatch5.Success Then
                    MessageBox.Show("Please Enter the valid Price. ", "Error")
                ElseIf Not mymatch6.Success Then
                    MessageBox.Show("Please Enter the valid VAT. ", "Error")
                ElseIf TextBox12.Text = "" Then
                    MessageBox.Show("Please Press the Total Amount Button. ", "Error")
                Else
                    st = 1
                End If


                If st = 1 Then
                    Try
                        Dim con As OleDbConnection
                        Dim cmd As OleDbCommand
                        Dim rs As OleDbDataReader
                        Dim in_srno As Integer = 0
                        Dim invoice As Integer = 0
                        Dim srno As Integer = 0


                        Dim ref_no, credit, p_name, p_code, p_addr, desc, ctct, status, amt_type, c_date, amt_date As String
                        Dim qty, rate, vat As Integer
                        Dim tamt, amt As Double

                        Dim strsql, strsql2 As String

                        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
                        con.Open()


                        p_code = "NA"
                        If Form1.RadioButton2.Checked = True Then
                            p_code = Form1.TextBox2.Text
                        End If
                        p_name = Form1.TextBox3.Text
                        p_addr = Form1.TextBox4.Text
                        ctct = Form1.TextBox5.Text


                        If Form1.RadioButton2.Checked = False Then

                            cmd = New OleDbCommand("select max(srno) from atable ", con)
                            rs = cmd.ExecuteReader()

                            While rs.Read()
                                srno = rs(0)
                            End While
                            rs.Close()

                            srno = srno + 1
                            Dim str1 As String = CStr(srno)
                            Dim str2 As String = "PAR"
                            p_code = str2 + str1


                            strsql = "insert into atable values(" & srno & ",'" & p_code & "','" & p_name & "','" & p_addr & "','" & ctct & "')"

                            Dim x As Integer
                            Dim sql As New OleDbCommand(strsql, con)
                            x = sql.ExecuteNonQuery()

                        End If




                        cmd = New OleDbCommand("select max(in_srno) from acctable ", con)
                        rs = cmd.ExecuteReader()

                        While rs.Read()
                            in_srno = rs(0)
                        End While
                        rs.Close()

                        in_srno = in_srno + 1

                        desc = TextBox9.Text
                        qty = CInt(TextBox9.Text)
                        rate = CInt(TextBox10.Text)
                        vat = CDbl(TextBox11.Text)
                        tamt = CDbl(TextBox12.Text)
                        credit = DateTimePicker1.Text.ToString

                        c_date = Form1.TextBox6.Text
                        status = "Balance"
                        amt_type = "NA"
                        amt = 0
                        ref_no = "NA"
                        amt_date = "NA"

                        If Form1.RadioButton4.Checked = True Then

                            status = "Paid"
                            amt_type = Form1.ComboBox1.SelectedItem.ToString

                            If Form1.TextBox7.Enabled = False Then
                                ref_no = "NA"
                            Else
                                ref_no = Form1.TextBox7.Text
                            End If


                            amt = CDbl(Form1.TextBox8.Text)
                            amt_date = Form1.DateTimePicker2.Text.ToString



                            'invoice generation
                            cmd = New OleDbCommand("select max(invoice) from acctable ", con)
                            rs = cmd.ExecuteReader()

                            While rs.Read()
                                invoice = rs(0)
                            End While
                            rs.Close()
                            invoice = invoice + 1


                        End If

                        strsql2 = "insert into acctable values(" & in_srno & "," & invoice & ",'" & p_code & "','" & desc & "'," & qty & "," & rate & "," & vat & "," & tamt & ",'" & status & "','" & c_date & "','" & amt_type & "','" & ref_no & "'," & amt & ",'" & amt_date & "','" & credit & "')"

                        Dim x2 As Integer
                        Dim sql2 As New OleDbCommand(strsql2, con)
                        x2 = sql2.ExecuteNonQuery()



                        con.Close()

                    Catch ex As Exception
                        flag = 1
                    End Try


                    If flag <> 1 Then
                        MsgBox("Record Inserted Succesfully.", MsgBoxStyle.Information, " Success")
                        'Button2_Click(sender, e)
                    Else
                        MsgBox("Error in Inserting Record.", MsgBoxStyle.Critical, " Failure")
                    End If
                End If
            End If
            cnt = cnt + 1
        End While

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.TextBox9.Clear()
        Me.TextBox10.Clear()
        Me.TextBox11.Clear()
        Me.TextBox12.Clear()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            Dim qty As Integer = CInt(TextBox9.Text)
            Dim rate As Integer = CInt(TextBox10.Text)
            Dim vat As Double = CDbl(TextBox11.Text)
            Dim tamt As Double = rate * qty
            Dim tvat As Double = tamt * (vat / 100)
            TextBox12.Text = tamt + tvat
        Catch ex As Exception
            MsgBox("Error! Please Check All Fields of Object Description", MsgBoxStyle.Critical, " Failure")
        End Try
    End Sub
End Class