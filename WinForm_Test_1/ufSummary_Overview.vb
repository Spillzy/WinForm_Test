Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop


Public Class ufSummary_Overview
    Private Sub tbCustomer_TextChanged(sender As Object, e As EventArgs) Handles tbCustomer.TextChanged

        Dim objText As Excel.Range
        'objText = MainClass.


        'Dim objWorksheet As Excel.Worksheet
        'Dim rngCustomer As Excel.Range

        'objWorksheet = Excel
        'rngCustomer = ActiveSheet.Cells

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ufZoneConfig.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ufZoneConfig.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ufZoneConfig.BackColor = Color.AliceBlue
    End Sub
End Class
