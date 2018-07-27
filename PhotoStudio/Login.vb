Public Class Login

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Trim(TextBox1.Text) = "" Then
            MsgBox("To Login Please Enter User_ID")
        ElseIf Trim(TextBox2.Text) = "" Then
            MsgBox("Please Enter Your Password")
        Else
            Form1.Show()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Trim(TextBox1.Text) = "" Then
            MsgBox("To Login Please Enter User_ID")
        ElseIf Trim(TextBox2.Text) = "" Then
            MsgBox("Please Enter Your Password")
        Else
            Form1.Show()
        End If

    End Sub
End Class