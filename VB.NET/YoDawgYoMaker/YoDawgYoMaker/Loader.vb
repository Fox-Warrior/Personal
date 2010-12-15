Public Class Loader

    Private Sub Loader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SW As Boolean = False
        Dim Joined = ""
        For Each A As String In GetSectionCles(Application.StartupPath + "\YoDawgs.ini", "Dawgs")
            If SW = False Then
                Dim Split = Strings.Split(GetCle(Application.StartupPath + "\YoDawgs.ini", "Dawgs", Strings.Left(A, InStr(A, "=") - 1)), "~")
                Joined = Split(0) + " " + Generator.Label1.Text + " " + _
                Split(1) + " " + Generator.Label2.Text + " " + Split(2) + _
                Generator.Label3.Text + " " + Split(3) + " " + Generator.Label4.Text + " " + _
                Split(4) + " " + Generator.Label5.Text + " " + Split(5) + " " + Generator.Label6.Text
                SW = True
            End If
            ListView1.Items.Add(Strings.Left(A, InStr(A, "=") - 1)).SubItems.Add(Joined)
        Next
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        Dim Split = Strings.Split(GetCle(Application.StartupPath + "\YoDawgs.ini", "Dawgs", ListView1.SelectedItems.Item(0).Text), "~")
        Generator.ComboBox1.Text = Split(0)
        Generator.TextBox1.Text = Split(1)
        Generator.TextBox2.Text = Split(2)
        Generator.TextBox3.Text = Split(3)
        Generator.TextBox4.Text = Split(4)
        Generator.TextBox5.Text = Split(5)
        Generator.DT = ListView1.SelectedItems.Item(0).Text
        Me.Close()
    End Sub
End Class