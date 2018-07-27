Imports System.Data.Odbc
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Dim sql, x As String
    Dim srl, rindx As Integer
    Dim con As OdbcConnection
    Dim cmd As OdbcCommand
    Dim dr As OdbcDataReader
    Dim ds As New DataSet1

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        Panel1.Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Panel1.Visible = True
    End Sub

    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.Click

    End Sub
    Private Sub ListBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        If ListBox1.Text = "" Then
            MsgBox("List is empty.", MsgBoxStyle.Information, "INFO")
        Else
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        End If
    End Sub

    Private Sub txtcp_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtcp.DoubleClick
        If Trim(ComboBox1.Text) = "" Then
            MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        ElseIf Trim(Val(txtcp.Text)) = "" Then
            MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        Else
            ListBox1.Items.Add(Trim(ComboBox1.Text) & "=" & Trim(Val(txtcp.Text)))
            ComboBox1.Focus()
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub txtpamt_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpamt.DoubleClick
        txtdamt.Text = Trim(Val(txttamt.Text)) - Trim(Val(txtpamt.Text))
    End Sub

    Private Sub txtpamt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtpamt.TextChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim report As New ReportDocument
        ''Dim path As String = Application.StartupPath  & "Reports\CrystalReport1.rpt"
        'Dim path As String = "E:\MKB_PC_BackUP\V B .Net\My Project\PhotoStudio\PhotoStudio\Reports\CrystalReport2.rpt"


        'If Trim(txtsrl.Text) = "" Then
        '    MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        '    txtsrl.Focus()
        '    'End If
        'ElseIf Trim(txtname.Text) = "" Then
        '    MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        '    txtname.Focus()
        '    'End If
        'ElseIf Trim(ComboBox1.Text) = "" Then
        '    MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        '    ComboBox1.Focus()
        '    'End If
        'ElseIf Trim(txtcp.Text) = "" Or IsNumeric(txtcp.Text) = False Then
        '    MsgBox("You must have to fill the asterick marked * field.", MsgBoxStyle.Information, "INFO")
        '    txtcp.Focus()
        'End If

        ''Saving Information into database
        'x = ""
        'For i = 0 To ListBox1.Items.Count - 1
        '    x = ListBox1.Items(i) & "," & x
        'Next
        'If IsNumeric(txttamt.Text) = False Then
        '    MsgBox("Please enter the valid amount.", MsgBoxStyle.Information, "Info")
        'End If
        'If IsNumeric(txtpamt.Text) = False Then
        '    MsgBox("Please enter the valid amount.", MsgBoxStyle.Information, "Info")
        'End If
        'If IsNumeric(txtcp.Text) = False Then
        '    MsgBox("Please enter the valid entry.", MsgBoxStyle.Information, "Info")
        'End If
        'sql = "INSERT INTO std VALUES(" & txtsrl.Text & ",'" & txtname.Text & "','" & txtadd.Text & "','" & MaskedTextBox1.Text & "','" & MaskedTextBox2.Text & "','" & txttamt.Text & "','" & txtpamt.Text & "','" & txtdamt.Text & "','" & x & "')"
        'cmd = New OdbcCommand(sql, con)
        'cmd.ExecuteNonQuery()
        'cmd.Dispose()

        'report.Load(path)
        'report.SetParameterValue("srlno", txtsrl.Text)
        'report.SetParameterValue("dopic", MaskedTextBox1.Text)
        'report.SetParameterValue("delivery", MaskedTextBox2.Text)
        'report.SetParameterValue("name", txtname.Text)
        'report.SetParameterValue("add", txtadd.Text)
        'report.SetParameterValue("totalphoto", x)
        'report.SetParameterValue("tamt", txttamt.Text)
        'report.SetParameterValue("pamt", txtpamt.Text)
        'report.SetParameterValue("damt", txtdamt.Text)
        'Form2.CrystalReportViewer1.ReportSource = report
        'Form2.Show()
        'txtsrl.Clear()
        ''txtsrl.Focus()
        'MaskedTextBox1.Clear()
        'MaskedTextBox2.Clear()
        'txtname.Clear()
        'txtadd.Clear()
        'txttamt.Clear()
        'txtpamt.Clear()
        'txtdamt.Clear()
        'txtcp.Clear()
        'ListBox1.Items.Clear()

    End Sub

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        con.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OdbcConnection("Driver={Microsoft Access Driver (*.mdb, *.accdb)}; dbq=" & Application.StartupPath & "\studio.accdb")
        con.Open()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        'BLANK TEXT BOX VALIDATION CHECKING
        If Trim(txt1.Text) = "" Then
            MsgBox("Please enter the valid name.", MsgBoxStyle.Information, "Info")
            'ElseIf IsInputChar(txt1.Text) = False Then
            '   MsgBox("Please enter the valid name.", MsgBoxStyle.Information, "Info")
        Else
            sql = "SELECT * FROM std WHERE ename LIKE '%" & txt1.Text & "%'"
            cmd = New OdbcCommand(sql, con)
            dr = cmd.ExecuteReader
            ds.Tables("record").Rows.Clear()
            While dr.Read
                DataGridView1.Rows.Add(New String() {Convert.ToString(dr.Item(0)), Convert.ToString(dr.Item(1)), Convert.ToString(dr.Item(4)), Convert.ToString(dr.Item(7)), Convert.ToString(dr.Item(8))})
                ds.Tables("record").Rows.Add(New String() {Convert.ToString(dr.Item(0)), Convert.ToString(dr.Item(1)), Convert.ToString(dr.Item(2)), Convert.ToString(dr.Item(3)), Convert.ToString(dr.Item(4)), Convert.ToString(dr.Item(5)), Convert.ToString(dr.Item(6)), Convert.ToString(dr.Item(7)), Convert.ToString(dr.Item(8))})
            End While
            dr.Close()
            cmd.Dispose()
            txt1.Clear()
            Panel1.Visible = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'BLANK TEXT BOX VALIDATION CHECKING
        If Trim(txt2.Text) = "" Then
            MsgBox("Please enter the valid Serial no..", MsgBoxStyle.Information, "Info")
        ElseIf IsNumeric(txt2.Text) = False Then
            MsgBox("Please enter the valid Serial no..", MsgBoxStyle.Information, "Info")
        Else
            sql = "SELECT * FROM std WHERE slno=(" & txt2.Text & ")"
            cmd = New OdbcCommand(sql, con)
            dr = cmd.ExecuteReader
            DataGridView1.Rows.Clear()
            ds.Tables("record").Rows.Clear()
            DataGridView1.Rows.Clear()
            While dr.Read
                DataGridView1.Rows.Add(New String() {Convert.ToString(dr.Item(0)), Convert.ToString(dr.Item(1)), Convert.ToString(dr.Item(4)), Convert.ToString(dr.Item(7)), Convert.ToString(dr.Item(8))})
                ds.Tables("record").Rows.Add(New String() {Convert.ToString(dr.Item(0)), Convert.ToString(dr.Item(1)), Convert.ToString(dr.Item(2)), Convert.ToString(dr.Item(3)), Convert.ToString(dr.Item(4)), Convert.ToString(dr.Item(5)), Convert.ToString(dr.Item(6)), Convert.ToString(dr.Item(7)), Convert.ToString(dr.Item(8))})
            End While
            dr.Close()
            cmd.Dispose()
            txt2.Clear()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Trim(txtsrl.Text) = "" Or Trim(MaskedTextBox1.Text) = "" Or Trim(MaskedTextBox2.Text) = "" Or Trim(ListBox1.SelectedIndex) = "" Then
            MsgBox("Field is empty", MsgBoxStyle.OkOnly, "Info")
        ElseIf MsgBox("Are you sure wants to update", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Info") = MsgBoxResult.Yes Then
            x = ""
            For i = 0 To ListBox1.Items.Count - 1
                x = ListBox1.Items(i) & "," & x
            Next
            sql = "UPDATE std set slno=" & txtsrl.Text & ",ename='" & txtname.Text & "',dopic='" & MaskedTextBox1.Text & "',delivery='" & MaskedTextBox2.Text & "', totalphoto='" & x & "' WHERE slno=" & txtsrl.Text & ""
            cmd = New OdbcCommand(sql, con)
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            DataGridView1.Rows.Clear()
        End If
        txtsrl.Clear()
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        txtname.Clear()
        txtadd.Clear()
        txttamt.Clear()
        txtpamt.Clear()
        txtdamt.Clear()
        txtcp.Clear()
        ListBox1.Items.Clear()
        txtsrl .Focus()
    End Sub

    Private Sub DataGridView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.Click
        'Panel1.Visible = True
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        MaskedTextBox1.Text = DateTimePicker1.Text
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        MaskedTextBox2.Text = DateTimePicker2.Text
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.RowCount = 0 Then
            MsgBox("No value is store.", MsgBoxStyle.Information, "Info")
            Exit Sub
        End If
        rindx = DataGridView1.CurrentRow.Index
        srl = DataGridView1.Item(0, rindx).Value
        MsgBox(srl)
        sql = "SELECT * FROM std WHERE slno=" & srl & ""
        cmd = New OdbcCommand(sql, con)
        dr = cmd.ExecuteReader

        If dr.Read Then
            txtsrl.Text = Convert.ToString(dr.Item(0))
            txtname.Text = Convert.ToString(dr.Item(1))
            txtadd.Text = Convert.ToString(dr.Item(2))
            MaskedTextBox1.Text = Convert.ToString(dr.Item(3))
            MaskedTextBox2.Text = Convert.ToString(dr.Item(4))
            txttamt.Text = Convert.ToString(dr.Item(5))
            txtpamt.Text = Convert.ToString(dr.Item(6))
            txtdamt.Text = Convert.ToString(dr.Item(7))
            ListBox1.Text = Convert.ToString(dr.Item(8))
        End If

        While dr.Read
            ds.Tables("record").Rows.Add(New String() {Convert.ToString(dr.Item(0)), Convert.ToString(dr.Item(1)), Convert.ToString(dr.Item(2)), Convert.ToString(dr.Item(3)), Convert.ToString(dr.Item(4)), Convert.ToString(dr.Item(5)), Convert.ToString(dr.Item(6)), Convert.ToString(dr.Item(7))})
        End While
        

        dr.Close()
        cmd.Dispose()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'Dim report As New ReportDocument
        ''Dim path As String = Application.StartupPath  & "Reports\CrystalReport1.rpt"
        'Dim path As String = "E:\MKB_PC_BackUP\V B .Net\My Project\PhotoStudio\PhotoStudio\Reports\CrystalReport2.rpt"

        'report.Load(path)
        'report.SetParameterValue("srlno", txtsrl.Text)
        'report.SetParameterValue("dopic", MaskedTextBox1.Text)
        'report.SetParameterValue("delivery", MaskedTextBox2.Text)
        'report.SetParameterValue("name", txtname.Text)
        'report.SetParameterValue("add", txtadd.Text)
        ''report.SetParameterValue("totalphoto", x)
        'report.SetParameterValue("tamt", txttamt.Text)
        'report.SetParameterValue("pamt", txtpamt.Text)
        'report.SetParameterValue("damt", txtdamt.Text)
        'Form2.CrystalReportViewer1.ReportSource = report
        'Form2.Show()
        'report.SetDataSource(ds.Tables("record"))
        'Form2.CrystalReportViewer1.ReportSource = report
        'Form2.Show()
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PictureBox1.Image = Image.FromFile(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        
        'Form name change automaticaly to image name selected from file
        If Trim(txtsrl.Text) <> "" Then
            Form3.Text = "Image-" & txtsrl.Text
            Form3.PictureBox1.BackgroundImage = PictureBox1.BackgroundImage
            Form3.Show()
        End If
    End Sub

    Private Sub txtcp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcp.TextChanged

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

    End Sub
End Class
