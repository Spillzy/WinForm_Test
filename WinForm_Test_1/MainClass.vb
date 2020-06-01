Imports Microsoft.Office.Interop

Public Class MainClass

    Shared Sub Main()

        'Start Excel and create New Workbook
        'Dim objExcel As New Excel.Application
        'Dim objWorkbook As Excel.Workbook = objExcel.Workbooks.Add
        'Dim objWorksheet As New Excel.Worksheet

        Dim objExcel As Object
        objExcel = CreateObject("Excel.Application")
        Dim objWorkbook As Excel.Workbook
        Dim objWorksheet As Excel.Worksheet
        objWorkbook = objExcel.Workbooks.Add
        objWorksheet = objExcel.Worksheets.add()

        'Check if Excel is installed
        If objExcel Is Nothing Then
            MessageBox.Show("Excel needs to be installed in order for the Configurator to run.", "Buy-off Configurator", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        'Show Excel
        objExcel.Visible = True
        objExcel.UserControl = False

        'MessageBox.Show("Staring Excel...", "Buy-off Configurator", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'Name the first sheet that was created
        objWorksheet = objExcel.ActiveSheet
        With objWorksheet
            .Name = "Overview Summary"
        End With

        'Set the Background of the worksheet to Light Grey
        Dim rngWholeSheet As Excel.Range
        rngWholeSheet = objWorksheet.Cells
        rngWholeSheet.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(231, 230, 230))  'Light Grey Background

        'Writes and Formats the FANUC Header
        Dim rngFanucHeader As Excel.Range = objWorksheet.Range("H2: M3")
        With rngFanucHeader
            .Merge()
            .Value = "Fanuc"
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlCenter
            .RowHeight = 20
            .Font.Name = "Calibri"
            .Font.Size = 36
            .Font.Color = -16776961 'Changes Font to Red
            .Font.Bold = True
        End With

        'Load the First Windows Form
        Dim objSummary_Overview As New ufSummary_Overview
        Application.Run(ufSummary_Overview)

        MessageBox.Show("Buyoff Worksheet is Generated", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)

        'Test format -To be Removed
        rngFanucHeader = objWorksheet.Range("H2:O3")
        rngFanucHeader.Merge()

        objExcel.UserControl = True

        'Close this Workbook and quit this instance of Excel
        objWorkbook.Close()
        objExcel.Quit()
        releaseObject(objWorksheet)
        releaseObject(objWorkbook)
        releaseObject(objExcel)

    End Sub

    Shared Sub releaseObject(ByVal obj As Object)
        'Handles any errors caused when attempted to close Excel.
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class

