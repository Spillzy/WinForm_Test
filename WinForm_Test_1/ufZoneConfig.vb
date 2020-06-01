Public Class ufZoneConfig


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Me.Text = TextBox1.Text
        ufSummary_Overview.Label1.Text = TextBox1.Text

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub ufZoneConfig_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        ufSummary_Overview.BackColor = Color.AliceBlue

    End Sub
End Class