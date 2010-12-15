Public Class Generator
    Dim T As Object
    Public DT As String = "Dawg"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Viewer.TextBox1.Text = ComboBox1.Text + " " + Label1.Text + " " + TextBox1.Text + " " + Label2.Text + " " + TextBox2.Text + " " + Label3.Text + " " + TextBox3.Text + " " + Label4.Text + " " + TextBox4.Text + " " + Label5.Text + " " + TextBox5.Text + "."
        Viewer.Show()
    End Sub

    Private Sub Generator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Enter
        T = TextBox2
        AcceptButton = Button2
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.T.Focus()
        Me.T.SelectionStart = Me.T.TextLength
    End Sub

    Private Sub ComboBox1_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Enter
        T = TextBox1
        AcceptButton = Button2
    End Sub

    Private Sub TextBox2_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Enter
        T = TextBox3
        AcceptButton = Button2
    End Sub

    Private Sub TextBox3_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Enter
        T = TextBox4
        AcceptButton = Button2
    End Sub

    Private Sub TextBox4_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.Enter
        T = TextBox5
        AcceptButton = Button2
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim R As String = InputBox("Name", "Name for the bookmark", DT)
        If R <> "" Then
            Dim D As String = ComboBox1.Text + "~" + TextBox1.Text + "~" + TextBox2.Text + "~" + TextBox3.Text + "~" + TextBox4.Text + "~" + TextBox5.Text
            SetCle(Application.StartupPath + "\YoDawgs.ini", "Dawgs", R, D)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Loader.Show()
    End Sub
End Class
