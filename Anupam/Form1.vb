Option Explicit On
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.IO
Public Class MainForm
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Button2_Click(sender, e)
        GroupBox9.Enabled = False
        If RadioButton2.Checked = True Then
            GroupBox1.Enabled = True
            ListBox1.Enabled = True
            ListBox1.Items.Clear()
            GroupBox2.Enabled = False
            RadioButton3.Checked = True

            Dim con As OleDbConnection
            Dim cmd As OleDbCommand
            Dim dr As OleDbDataReader

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            cmd = New OleDbCommand("select * from atable", con)

            con.Open()
            dr = cmd.ExecuteReader()
            While (dr.Read) = True
                ListBox1.Items.Add(dr("p_code"))
            End While
        Else
            GroupBox1.Enabled = False
            ListBox1.Enabled = False
            GroupBox2.Enabled = True
            RadioButton3.Checked = False
        End If
        ComboBox3.Items.Clear()
        ComboBox3.Items.Insert(0, 1)
        ComboBox3.Items.Insert(1, 2)
        ComboBox3.Items.Insert(2, 3)
        ComboBox3.Items.Insert(3, 4)
        ComboBox3.Items.Insert(4, 5)
        ComboBox3.Items.Insert(5, 6)
        ComboBox3.Items.Insert(6, 7)
        ComboBox3.Items.Insert(7, 8)
        ComboBox3.Items.Insert(8, 9)
        ComboBox3.SelectedIndex = 0
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            'Me.Visible = False
            'TODO: This line of code loads data into the 'AdataDataSet2.exp' table. You can move, or remove it, as needed.
            Me.ExpTableAdapter.Fill(Me.AdataDataSet2.exp)
            'TODO: This line of code loads data into the 'AdataDataSet1.acctable' table. You can move, or remove it, as needed.
            Me.AcctableTableAdapter1.Fill(Me.AdataDataSet1.acctable)
            'TODO: This line of code loads data into the 'AdataDataSet.atable' table. You can move, or remove it, as needed.
            Me.AtableTableAdapter.Fill(Me.AdataDataSet.atable)
            'TODO: This line of code loads data into the 'AdataDataSet.acctable' table. You can move, or remove it, as needed.
            Me.AcctableTableAdapter.Fill(Me.AdataDataSet.acctable)

            ComboBox1.Items.Insert(0, "Select")
            ComboBox1.SelectedIndex = 0
            ComboBox1.Items.Add("Cash")
            ComboBox1.Items.Add("Cheque")
            ComboBox1.Items.Add("Other")

            TabControl1.Enabled = True
            TextBox6.Text = Now.Date

            Dim tdate As Date
            tdate = CDate(Now.Date)

            Dim tday As Double = 0
            Dim con As OleDbConnection
            Dim cmd2 As OleDbCommand
            Dim rs2 As OleDbDataReader

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            con.Open()

            cmd2 = New OleDbCommand("select e_amt from exp where e_date = FORMAT('" & tdate & "','DD-MM-YYYY')", con)
            rs2 = cmd2.ExecuteReader()

            While rs2.Read()
                tday = tday + CDbl(rs2(0))
            End While
            con.Close()

            Label27.Text = tday
            GroupBox9.Enabled = False
            Label30.Text = 0

            CheckBox1_CheckedChanged(sender, e)
            CheckBox2_CheckedChanged(sender, e)
            CheckBox3_CheckedChanged(sender, e)
            CheckBox4_CheckedChanged(sender, e)
            CheckBox5_CheckedChanged(sender, e)
        Catch ex As Exception
            MsgBox("Please Undo any changes done with Database. Restart the Application." + vbNewLine + +ex.ToString, MsgBoxStyle.Critical, "Critical Error")
        End Try
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        GroupBox9.Enabled = False
        If RadioButton4.Checked = True Then

            GroupBox3.Enabled = True
            GroupBox1.Enabled = False
            Button1.Enabled = False
            GroupBox9.Enabled = False

            If RadioButton1.Checked = True Then
                MsgBox("Please Select the Exisiting Record.", MsgBoxStyle.Exclamation, " Informative")
                RadioButton3.Checked = True
                Button1.Enabled = True

                Me.Button2_Click(sender, e)
            End If

        Else
            Button1.Enabled = True
            If RadioButton2.Checked = True Then
                GroupBox1.Enabled = True
            End If
            GroupBox3.Enabled = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim cnt As Integer = CInt(Label30.Text)
        Dim flag As Integer = 0
        Dim st As Integer = 0

        Dim mymatch4 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\d+$")
        Dim mymatch5 As Match = System.Text.RegularExpressions.Regex.Match(TextBox10.Text, "^\d+\.\d+$")

        Dim mymatch7 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\w{1,50}|\w+\s{1}\w$")

        If Not mymatch7.Success Then
            MessageBox.Show("Please Enter the valid Description. ", "Error")
        ElseIf Not mymatch4.Success Then
            MessageBox.Show("Please Enter the valid Quantity. ", "Error")
        ElseIf Not mymatch5.Success Then
            MessageBox.Show("Please Enter the valid Price. Check precision. Example 100.0 ", "Error")
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
                Dim invoice As Integer = CInt(TextBox20.Text)
                Dim srno As Integer = 0
                'Dim obj_no As Integer = CInt(Label30.Text)


                Dim ref_no, credit, p_name, p_code, p_addr, desc, ctct, status, amt_type, c_date, amt_date, v_code, client_name As String
                Dim qty As Integer
                Dim vat, rate As Decimal
                Dim tamt, amt As Double

                Dim strsql2 As String

                con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
                con.Open()


                p_code = "NA"
                If RadioButton2.Checked = True Then
                    p_code = TextBox2.Text
                End If
                p_name = TextBox3.Text
                p_addr = TextBox4.Text
                ctct = TextBox5.Text
                v_code = TextBox17.Text
                client_name = TextBox18.Text

                'insertion in account table.

                cmd = New OleDbCommand("select max(in_srno) from acctable ", con)
                rs = cmd.ExecuteReader()

                While rs.Read()
                    in_srno = rs(0)
                End While
                rs.Close()

                in_srno = in_srno + 1

                desc = TextBox13.Text
                qty = CInt(TextBox9.Text)
                rate = Decimal.Parse(TextBox10.Text)
                vat = Decimal.Parse(TextBox11.Text)
                tamt = CDbl(TextBox12.Text)
                credit = DateTimePicker1.Text.ToString

                c_date = TextBox6.Text
                status = "Balance"
                amt_type = "NA"
                amt = 0
                ref_no = "NA"
                amt_date = "NA"

                If RadioButton4.Checked = True Then
                    status = "Paid"
                    amt_type = ComboBox1.SelectedItem.ToString

                    If TextBox7.Enabled = False Then
                        ref_no = "NA"
                    Else
                        ref_no = TextBox7.Text
                    End If

                    amt = CDbl(TextBox8.Text)
                    amt_date = DateTimePicker2.Text.ToString
                End If

                'invoice generation

                strsql2 = "insert into acctable values(" & in_srno & "," & invoice & ",'" & p_code & "','" & desc & "'," & qty & "," & rate & "," & vat & "," & tamt & ",'" & status & "','" & c_date & "','" & amt_type & "','" & ref_no & "'," & amt & ",'" & amt_date & "','" & credit & "')"

                Dim x2 As Integer
                Dim sql2 As New OleDbCommand(strsql2, con)
                x2 = sql2.ExecuteNonQuery()
                con.Close()

            Catch ex As Exception
                flag = 1
                MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Error")

            End Try
        End If

        If cnt >= CInt(ComboBox3.SelectedItem) And st = 1 Then
            MsgBox("Record Inserted and Current Order Completed Successfully.", MsgBoxStyle.Information, " Success")
            Button9.Enabled = True
            GroupBox9.Enabled = False
            Me.Button2_Click(sender, e)

            flag = 2
        End If

        If flag <> 1 And flag <> 2 And st = 1 Then
            MsgBox("Record Inserted Successfully. Please Complete the Order.", MsgBoxStyle.Exclamation, " Success")

            Me.TextBox7.Clear()
            Me.TextBox8.Clear()
            Me.TextBox9.Clear()
            Me.TextBox10.Clear()
            'Me.TextBox11.Clear()
            Me.TextBox12.Clear()
            Me.TextBox13.Clear()
            cnt = cnt + 1
            Label30.Text = cnt

        ElseIf flag <> 0 And flag <> 2 Then
            MsgBox("Error in Inserting Record.", MsgBoxStyle.Critical, " Failure")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "Auto-Generated"
        Me.TextBox2.Text = "Auto-Generated"
        Me.TextBox3.Clear()
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        'Me.TextBox6.Clear() Date field
        Me.TextBox7.Clear()
        Me.TextBox8.Clear()
        Me.TextBox9.Clear()
        Me.TextBox10.Clear()
        Me.TextBox11.Clear()
        Me.TextBox12.Clear()
        Me.TextBox13.Clear()

        Me.TextBox17.Clear()
        Me.TextBox18.Clear()

        Me.ComboBox2.Items.Clear()
        'Me.ComboBox4.Items.Clear()
        Label30.Text = 0
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem = "Cash" Then
            TextBox7.Enabled = False
        Else
            TextBox7.Enabled = True
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            Dim qty As Integer = CInt(TextBox9.Text)
            Dim rate As Double = CDbl(TextBox10.Text)
            Dim vat As Double = CDbl(TextBox11.Text)
            Dim tamt As Double = rate * qty
            Dim tvat As Double = tamt * (vat / 100)
            tamt = Format(tamt, "#,##0.00")
            tvat = Format(tvat, "#,##0.00")
            TextBox12.Text = tamt + tvat
        Catch ex As Exception
            MsgBox("Error! Please Check All Fields of Object Description", MsgBoxStyle.Critical, " Failure")
        End Try
    End Sub

    Private Sub rem_dup(ByVal combo As ComboBox)
        combo.Sorted = True
        combo.Refresh()

        Dim index As Integer
        Dim itemcount As Integer = combo.Items.Count

        If itemcount > 1 Then
            Dim lastitem As String = combo.Items(itemcount - 1)

            For index = itemcount - 2 To 0 Step -1
                If combo.Items(index) = lastitem Then
                    combo.Items.RemoveAt(index)
                Else
                    lastitem = combo.Items(index)
                End If
            Next
        End If
    End Sub

    Public Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Button2_Click(sender, e)
        Dim p_code As String
        p_code = CStr(ListBox1.SelectedItem)

        Try

            Dim con As OleDbConnection
            Dim cmd, cmd1 As OleDbCommand
            Dim rs, rs1 As OleDbDataReader

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            con.Open()
            cmd = New OleDbCommand("select * from atable where p_code = '" & p_code & "'", con)
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

            If RadioButton4.Checked = True Then

                ComboBox2.Items.Clear()

                cmd1 = New OleDbCommand("select * from acctable where p_code = '" & p_code & "' AND status='Balance'", con)
                rs1 = cmd1.ExecuteReader()

                While rs1.Read()
                    ComboBox2.Items.Add(rs1(1))

                    TextBox1.Text = rs1(1)
                    'Label30.Text = rs1(2)
                    TextBox9.Text = rs1(4)
                    TextBox10.Text = CDbl(rs1(5))
                    TextBox11.Text = CDbl(rs1(6))
                    TextBox12.Text = rs1(7)
                    TextBox13.Text = rs1(3)
                    DateTimePicker1.Text = rs1(14)

                End While
                rem_dup(ComboBox2)
            End If
            con.Close()
        Catch ex As Exception
            MsgBox("Error in Populating ListBox.")
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Button2_Click(sender, e)
        GroupBox9.Enabled = False

        If RadioButton1.Checked = True Then
            GroupBox1.Enabled = False
            GroupBox2.Enabled = True
            GroupBox3.Enabled = False
            ListBox1.Enabled = False
            RadioButton3.Checked = True

            'ComboBox3.Items.Clear()
            'ComboBox3.Items.Insert(0, 1)
            'ComboBox3.SelectedIndex = 0

        Else
            GroupBox1.Enabled = True
            GroupBox2.Enabled = False
            GroupBox3.Enabled = True
            ListBox1.Enabled = True
            RadioButton3.Checked = False

        End If
        'Updated 16-Dec-13. New and Existing order can have multiple objects.
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' ComboBox4.Items.Clear()
        Dim p_code As String
        p_code = CStr(ListBox1.SelectedItem)
        Dim invoice As Integer
        invoice = CInt(ComboBox2.SelectedItem)

        Dim con As OleDbConnection
        Dim cmd1 As OleDbCommand
        Dim rs1 As OleDbDataReader

        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
        con.Open()

        If RadioButton4.Checked = True Then

            cmd1 = New OleDbCommand("select * from acctable where p_code = '" & p_code & "' AND status='Balance' AND invoice = " & invoice & " ", con)
            rs1 = cmd1.ExecuteReader()

            While rs1.Read()
                ' ComboBox4.Items.Add(rs1(2))
                TextBox1.Text = rs1(1)
                ' Label30.Text = rs1(2) 'Object Number Display
                TextBox9.Text = rs1(4)
                TextBox10.Text = CDbl(rs1(5))
                TextBox11.Text = CDbl(rs1(6))
                TextBox12.Text = CDbl(rs1(7))
                TextBox13.Text = rs1(3)
                DateTimePicker1.Text = rs1(14)
            End While
            'rem_dup(ComboBox4)
        End If
        con.Close()
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        Button2_Click(sender, e)
        If RadioButton3.Checked = True Then
            GroupBox3.Enabled = False
            If RadioButton2.Checked = True Then
                GroupBox1.Enabled = True
            End If
        Else
            GroupBox3.Enabled = True
            GroupBox1.Enabled = False
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim st As Integer = 0

        Dim mymatch1 As Match = System.Text.RegularExpressions.Regex.Match(TextBox3.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch2 As Match = System.Text.RegularExpressions.Regex.Match(TextBox4.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch3 As Match = System.Text.RegularExpressions.Regex.Match(TextBox5.Text, "^\d+$")

        Dim mymatch4 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\d+$")
        Dim mymatch5 As Match = System.Text.RegularExpressions.Regex.Match(TextBox10.Text, "^[1-9]\d*(\.\d+)?$")
        Dim mymatch6 As Match = System.Text.RegularExpressions.Regex.Match(TextBox11.Text, "^[1-9]\d*(\.\d+)?$")
        Dim mymatch7 As Match = System.Text.RegularExpressions.Regex.Match(TextBox9.Text, "^\w{1,50}|\w+\s{1}\w$")

        Dim mymatch8 As Match = System.Text.RegularExpressions.Regex.Match(TextBox7.Text, "^\d+$")
        Dim mymatch9 As Match = System.Text.RegularExpressions.Regex.Match(TextBox8.Text, "^\d+\.\d+$")

        Dim mymatch10 As Match = System.Text.RegularExpressions.Regex.Match(TextBox17.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch11 As Match = System.Text.RegularExpressions.Regex.Match(TextBox18.Text, "^\w{1,50}|\w+\s{1}\w$")


        If RadioButton4.Checked = True Then
            If Not mymatch1.Success Then
                MessageBox.Show("Please Enter Party-Name OR Select Exisiting Entry. ", "Error")
            ElseIf Not mymatch2.Success Then
                MessageBox.Show("Please Enter Party Address OR Select Exisiting Entry.", "Error")
            ElseIf Not mymatch3.Success Then
                MessageBox.Show("Please Enter Contact Number OR Select Exisiting Entry.", "Error")

            ElseIf Not mymatch10.Success Then
                MessageBox.Show("Please Enter Vendor-Code OR Select Exisiting Entry.", "Error")
            ElseIf Not mymatch11.Success Then
                MessageBox.Show("Please Enter Client Name OR Select Exisiting Entry.", "Error")

            ElseIf Not mymatch7.Success Then
                MessageBox.Show("Please Enter the valid Quantity. ", "Error")
            ElseIf Not mymatch4.Success Then
                MessageBox.Show("Please Enter the valid Quantity. ", "Error")
            ElseIf Not mymatch5.Success Then
                MessageBox.Show("Please Enter the valid Price. Check for Precision" + vbNewLine + "Example: 100.00", "Error")
            ElseIf Not mymatch6.Success Then
                MessageBox.Show("Please Enter the valid VAT. ", "Error")
            ElseIf TextBox12.Text = "" Then
                MessageBox.Show("Please Press the Total Amount Button. ", "Error")

            ElseIf Not IsNumeric(ComboBox2.SelectedItem) Then
                MessageBox.Show("Please Select the Invoice-Number", "Error")
            ElseIf ComboBox1.SelectedItem.ToString = "Select" Then
                MessageBox.Show("Please Select the Amount-Type", "Error")
            ElseIf Not mymatch8.Success And TextBox7.Enabled = True Then
                MessageBox.Show("Please Enter Valid Reference/Cheque No. " & vbCrLf & "Enter 0000 if Amount type is CASH", "Error")
            ElseIf Not mymatch9.Success Then
                MessageBox.Show("Please Enter Valid Amount" + vbNewLine + "Example: 100.00", "Error")
            Else
                st = 1
            End If
        End If

        If st = 1 Then
            Dim flag As Integer = 0
            Try
                Dim p_code As String
                p_code = CStr(ListBox1.SelectedItem)
                Dim invoice As Integer
                invoice = CInt(ComboBox2.SelectedItem)
                '  Dim obj_no As Integer
                '  obj_no = CInt(ComboBox3.SelectedItem)

                'DECL.

                Dim ref_no, desc, status, amt_type, c_date, amt_date, credit, bank As String
                Dim qty As Integer
                Dim tamt, amt, rate, vat As Double
                Dim strsql2 As String

                Dim con As OleDbConnection

                con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
                con.Open()

                'Amt fill
                desc = TextBox13.Text
                qty = CInt(TextBox9.Text)
                rate = CDbl(TextBox10.Text)
                vat = CDbl(TextBox11.Text)
                tamt = CDbl(TextBox12.Text)
                credit = DateTimePicker1.Text.ToString

                c_date = TextBox6.Text
                status = "Balance"
                amt_type = "NA"
                amt = 0
                ref_no = "NA"
                amt_date = "NA"
                bank = "NA"

                If RadioButton4.Checked = True Then

                    status = "Paid"
                    amt_type = ComboBox1.SelectedItem.ToString

                    If TextBox7.Enabled = False Then
                        ref_no = "NA"
                    Else
                        ref_no = TextBox7.Text
                        bank = TextBox27.Text
                    End If


                    amt = CDbl(TextBox8.Text)
                    amt_date = DateTimePicker2.Text.ToString

                End If
                'AND obj_no=" & obj_no & " No need of this. update for all objects one time.
                strsql2 = "Update acctable set status='" & status & "',c_date='" & c_date & "',amt_type='" & amt_type & "',ref_no='" & ref_no & "',amt=" & amt & ",amt_date='" & amt_date & "',bank='" & bank & "' where invoice= " & invoice & " "
                Dim x2 As Integer
                Dim sql2 As New OleDbCommand(strsql2, con)
                x2 = sql2.ExecuteNonQuery()

                con.Close()
            Catch ex As Exception
                flag = 1
            End Try

            If flag <> 1 Then
                MsgBox("Record Updated Succesfully.", MsgBoxStyle.Information, " Success")
                'GroupBox3.Enabled = False
                RadioButton3.Checked = True
            Else
                MsgBox("Error in Updating Record. Please Check All Fields", MsgBoxStyle.Critical, " Failure")
                Me.Button2_Click(sender, e)
                'GroupBox3.Enabled = False
                RadioButton3.Checked = True
            End If
            Button2_Click(sender, e)
        End If

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim mymatch1 As Match = System.Text.RegularExpressions.Regex.Match(TextBox14.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch2 As Match = System.Text.RegularExpressions.Regex.Match(TextBox15.Text, "^\d+$")
        Dim flag As Integer = 0
        Try

            If Not mymatch1.Success Then
                MessageBox.Show("Please Enter Alphanumeric Chracters Only in Description. ", "Error")
                flag = 1
            ElseIf Not mymatch2.Success Then
                MessageBox.Show("Please Enter Expense-Amount.", "Error")
                flag = 1

            Else

                Dim con As OleDbConnection
                Dim cmd, cmd2 As OleDbCommand
                Dim rs, rs2 As OleDbDataReader
                Dim srno As Integer = 0
                Dim tday As Integer = 0

                Dim e_desc As String
                Dim e_date As Date
                Dim e_amt As Double

                e_desc = TextBox14.Text
                e_date = CDate(DateTimePicker3.Text)
                e_amt = CDbl(TextBox15.Text)

                Dim strsql As String

                con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
                con.Open()

                cmd = New OleDbCommand("select max(srno) from exp ", con)
                rs = cmd.ExecuteReader()

                While rs.Read()
                    srno = rs(0)
                End While
                rs.Close()

                srno = srno + 1

                strsql = "insert into exp values(" & srno & ",'" & e_desc & "','" & e_date & "'," & e_amt & ")"

                Dim x As Integer
                Dim sql As New OleDbCommand(strsql, con)
                x = sql.ExecuteNonQuery()

                Dim tdate As Date
                tdate = CDate(Now.Date)

                cmd2 = New OleDbCommand("select e_amt from exp where e_date = FORMAT('" & tdate & "','DD-MM-YYYY')", con)
                rs2 = cmd2.ExecuteReader()

                While rs2.Read()
                    tday = tday + CDbl(rs2(0))
                End While
                con.Close()

                Label27.Text = tday
                'Refresh Datagridview
                ExpTableAdapter.Fill(AdataDataSet2.exp)
            End If
        Catch ex As Exception
            flag = 1
        End Try

        If flag <> 1 Then
            MsgBox("Record Inserted Succesfully.", MsgBoxStyle.Information, " Success")
            Me.Button3_Click(sender, e)
        Else
            MsgBox("Error in Inserting Record.", MsgBoxStyle.Critical, " Failure")
            Me.Button3_Click(sender, e)
        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim rs As OleDbDataReader
        Dim srno As Integer = 0

        Dim f_date, t_date As Date
        Dim t_amt As Double = 0

        f_date = CDate(DateTimePicker4.Text)
        t_date = CDate(DateTimePicker5.Text)

        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
        con.Open()

        cmd = New OleDbCommand("select e_amt from exp where e_date >= FORMAT('" & f_date & "','DD-MM-YYYY') AND e_date <= FORMAT('" & t_date & "','DD-MM-YYYY')", con)
        rs = cmd.ExecuteReader()

        While rs.Read()
            t_amt = t_amt + CDbl(rs(0))
        End While

        rs.Close()
        con.Close()
        TextBox16.Text = t_amt
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox14.Clear()
        TextBox15.Clear()
        TextBox15.Clear()
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try

            AtableTableAdapter.Fill(AdataDataSet.atable)
            AcctableTableAdapter.Fill(AdataDataSet.acctable)
            AcctableTableAdapter1.Fill(AdataDataSet1.acctable)
        Catch ex As Exception
            MsgBox("Error in Refreshing", MsgBoxStyle.Critical, "Failure")
        End Try

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim mymatch1 As Match = System.Text.RegularExpressions.Regex.Match(TextBox3.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch2 As Match = System.Text.RegularExpressions.Regex.Match(TextBox4.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch3 As Match = System.Text.RegularExpressions.Regex.Match(TextBox5.Text, "^\d+$")


        Dim mymatch4 As Match = System.Text.RegularExpressions.Regex.Match(TextBox17.Text, "^\w{1,50}|\w+\s{1}\w$")
        Dim mymatch5 As Match = System.Text.RegularExpressions.Regex.Match(TextBox18.Text, "^\w{1,50}|\w+\s{1}\w$")

        Dim mymatch6 As Match = System.Text.RegularExpressions.Regex.Match(TextBox11.Text, "^\d+\.\d+$")

        If Not mymatch1.Success Then
            MessageBox.Show("Please Enter Party-Name OR Select Exisiting Entry. ", "Error")
        ElseIf Not mymatch2.Success Then
            MessageBox.Show("Please Enter Party Address OR Select Exisiting Entry.", "Error")
        ElseIf Not mymatch3.Success Then
            MessageBox.Show("Please Enter Contact Number OR Select Exisiting Entry.", "Error")

        ElseIf Not mymatch4.Success Then
            MessageBox.Show("Please Enter Vendor-Code OR Select Exisiting Entry.", "Error")
        ElseIf Not mymatch5.Success Then
            MessageBox.Show("Please Enter Client Name OR Select Exisiting Entry.", "Error")

        ElseIf Not IsNumeric(ComboBox3.SelectedItem) Then
            MsgBox("Please Select Number of Objects.", MsgBoxStyle.Critical, " Failure")
            GroupBox9.Enabled = False
        ElseIf Not mymatch6.Success Then
            MessageBox.Show("Please Enter the valid VAT with Decimal Precision. Example: 2.0 ", "Error")
            GroupBox9.Enabled = False
        Else
            GroupBox1.Enabled = False
            Dim invoice As Integer = 0
            GroupBox9.Enabled = True

            Label30.Text = 1

            Dim con As OleDbConnection
            Dim cmd As OleDbCommand
            Dim rs As OleDbDataReader

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            con.Open()

            'multiple obj logic

            cmd = New OleDbCommand("select max(invoice) from acctable ", con)
            rs = cmd.ExecuteReader()

            While rs.Read()
                invoice = rs(0)
            End While
            rs.Close()

            invoice = invoice + 1
            TextBox20.Text = invoice
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Button2_Click(sender, e)
        GroupBox1.Enabled = True
        GroupBox9.Enabled = False
    End Sub

    Private Sub rem_dup_list(ByVal list As ListBox)
        list.Sorted = True
        list.Refresh()

        Dim index As Integer
        Dim itemcount As Integer = list.Items.Count

        If itemcount > 1 Then
            Dim lastitem As String = list.Items(itemcount - 1)

            For index = itemcount - 2 To 0 Step -1
                If list.Items(index) = lastitem Then
                    list.Items.RemoveAt(index)
                Else
                    lastitem = list.Items(index)
                End If
            Next
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dr As OleDbDataReader
        Try
            AcctableTableAdapter1.Fill(AdataDataSet1.acctable)
            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            cmd = New OleDbCommand("select * from atable", con)

            con.Open()
            dr = cmd.ExecuteReader()
            While (dr.Read) = True
                ComboBox5.Items.Add(dr("p_code"))
            End While
            rem_dup_list(ListBox2)
            rem_dup(ComboBox5)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Failure")
        End Try

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim invoice As Integer
        Dim cnt As Integer = 1
        Dim gamt As Double = 0
        Dim tvamt As Double = 0
        Dim tamt As Double = 0
        Dim amt As Double = 0
        Dim p_code As String

        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dr As OleDbDataReader

        'Print Document
        Dim wd As Word.Application
        Dim wdDoc As Word.Document


        If ComboBox5.SelectedItem = "" Then
            MessageBox.Show("Please Select Party-Code. ", "Error")
        ElseIf Not IsNumeric(ListBox2.SelectedItem) Then
            MessageBox.Show("Please Select Invoice-Number. ", "Error")

        Else
            Try
                invoice = CInt(ListBox2.SelectedItem)
                p_code = ComboBox5.SelectedItem.ToString
                Dim vat As Double

                wd = New Word.Application
                wd.Visible = True
                wdDoc = wd.Documents.Add("" & Directory.GetCurrentDirectory & "\Sample.doc") 'Add document to Word
                With wdDoc

                    con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
                    cmd = New OleDbCommand("select * from atable where p_code='" & p_code & "'", con)

                    con.Open()
                    dr = cmd.ExecuteReader()

                    While (dr.Read) = True
                        'Setting text from VB to Word
                        .FormFields("text1").Range.Text = dr(3)
                        .FormFields("text2").Range.Text = dr(4)
                        .FormFields("text3").Range.Text = invoice
                        .FormFields("text4").Range.Text = dr(2)

                        If CheckBox4.Checked = False Then
                            .FormFields("text11").Range.Text = ""
                        Else
                            .FormFields("text11").Range.Text = TextBox22.Text
                        End If

                        If CheckBox5.Checked = False Then
                            .FormFields("text12").Range.Text = ""
                        Else
                            .FormFields("text12").Range.Text = TextBox23.Text
                        End If

                        If CheckBox1.Checked = False Then
                            .FormFields("text5").Range.Text = ""
                        Else
                            .FormFields("text5").Range.Text = DateTimePicker7.Text
                        End If

                        If CheckBox2.Checked = False Then
                            .FormFields("text6").Range.Text = ""
                        Else
                            .FormFields("text6").Range.Text = DateTimePicker8.Text
                        End If

                        If CheckBox3.Checked = False Then
                            .FormFields("text7").Range.Text = ""
                        Else
                            .FormFields("text7").Range.Text = DateTimePicker9.Text
                        End If

                    End While

                    cmd = New OleDbCommand("select * from acctable where invoice=" & invoice & " AND p_code='" & p_code & "'", con)


                    dr = cmd.ExecuteReader()
                    'Changes made on 16-12-13
                    While (dr.Read) = True
                        .FormFields("text" + CStr(cnt) + "1").Range.Text = cnt 'dr(1)
                        .FormFields("text" + CStr(cnt) + "2").Range.Text = dr(3)
                        .FormFields("text" + CStr(cnt) + "3").Range.Text = dr(4)
                        .FormFields("text" + CStr(cnt) + "4").Range.Text = dr(5)

                        Dim qty As Integer = CInt(dr(4))
                        Dim rate As Double = CDec(dr(5))
                        vat = CDec(dr(6))
                        amt = rate * qty

                        'Amt
                        .FormFields("text" + CStr(cnt) + "5").Range.Text = amt

                        'T.AMT
                        tamt = tamt + amt

                        cnt = cnt + 1
                    End While
                    'T.AMT out
                    .FormFields("text8").Range.Text = tamt

                    'VAT
                    .FormFields("text9").Range.Text = vat

                    'G.AMT
                    tvamt = tamt * (vat / 100)
                    gamt = tamt + tvamt

                    gamt = Format(gamt, "#,##0.00")
                    .FormFields("text10").Range.Text = gamt

                    'T.VAT.AMT
                    .FormFields("tvamt").Range.Text = Format(tvamt, "#,##0.00")

                    'amtinwords
                    .FormFields("amt").Range.Text = SpellNumber(gamt)
                End With
                con.Close()

                TextBox21.Text = SpellNumber(gamt) 'AmtInWord(gamt)
                Button13_Click(sender, e)

                'wdDoc.Save()
                wdDoc.SaveAs(Directory.GetCurrentDirectory & "\InvoiceBook\" & invoice)
                Me.WindowState = FormWindowState.Minimized
            Catch ex As Exception
                MsgBox("Error in Creating Word Document." + ex.ToString, MsgBoxStyle.Critical, "Failure")
            End Try
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        TextBox22.Clear()
        TextBox23.Clear()
        ListBox2.Items.Clear()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        Dim con As OleDbConnection
        Dim cmd As OleDbCommand
        Dim dr As OleDbDataReader
        Try

            Dim p_code As String = CStr(ComboBox5.SelectedItem)
            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            cmd = New OleDbCommand("select * from acctable where p_code='" & p_code & "'", con)

            con.Open()
            dr = cmd.ExecuteReader()
            While (dr.Read) = True
                ListBox2.Items.Add(dr(1))
            End While

            rem_dup_list(ListBox2)
            rem_dup(ComboBox5)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Failure")
        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            DateTimePicker7.Enabled = True
        Else
            DateTimePicker7.Enabled = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            DateTimePicker8.Enabled = True
        Else
            DateTimePicker8.Enabled = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            DateTimePicker9.Enabled = True
        Else
            DateTimePicker9.Enabled = False
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            TextBox22.Enabled = True
        Else
            TextBox22.Enabled = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            TextBox23.Enabled = True
        Else
            TextBox23.Enabled = False
        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
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
            Dim cmd As OleDbCommand
            Dim rs As OleDbDataReader
            Dim in_srno As Integer = 0
            Dim srno As Integer = 0

            Dim p_name, p_code, p_addr, ctct, v_code, client_name, email, vat_num As String
            Dim strsql As String

            con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Directory.GetCurrentDirectory & "\Adata.mdb;")
            Try
                con.Open()

                p_code = "NA"
                If RadioButton2.Checked = True Then
                    p_code = TextBox2.Text
                End If
                p_name = TextBox3.Text
                p_addr = TextBox4.Text
                ctct = TextBox5.Text
                v_code = TextBox17.Text
                client_name = TextBox18.Text
                email = TextBox25.Text
                vat_num = TextBox26.Text


                If RadioButton2.Checked = False Then

                    cmd = New OleDbCommand("select max(srno) from atable ", con)
                    rs = cmd.ExecuteReader()

                    While rs.Read()
                        srno = rs(0)
                    End While
                    rs.Close()

                    srno = srno + 1
                    Dim str1 As String = CStr(srno)
                    Dim str2 As String = "PAR"
                    p_code = str2 + str1 + "-" + p_name


                    strsql = "insert into atable values(" & srno & ",'" & p_code & "','" & v_code & "','" & p_name & "','" & p_addr & "','" & client_name & "','" & ctct & "','" & email & "','" & vat_num & "')"

                    Dim x As Integer
                    Dim sql As New OleDbCommand(strsql, con)
                    x = sql.ExecuteNonQuery()

                End If
            Catch ex As Exception
                MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error")
                con.Close()
            End Try
            MsgBox("Record inserted successfully.", MsgBoxStyle.Information, "Success")
            con.Close()
            RadioButton1.Checked = False
            RadioButton2.Checked = True
            Button2_Click(sender, e)
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        UpdateForm.Show()
    End Sub

    Private Sub FillByToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.AcctableTableAdapter1.FillBy(Me.AdataDataSet1.acctable)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
    End Sub


End Class
