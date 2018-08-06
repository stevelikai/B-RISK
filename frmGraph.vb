Imports System.Drawing.Printing
Imports System.Drawing.Imaging
'Imports System.Windows.Forms.DataVisualization.Charting

Public Class frmGraph
    Private PrintPageSettings As New PageSettings
    Private StringToPrint As String
    Private PrintFont As New Font("QuickType II Mono", 8.25)

    Private Sub PrintGraphic(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        'create the graphic image
        ev.Graphics.DrawImage(Image.FromFile("Image.Jpeg"), ev.Graphics.VisibleClipBounds)
        'specify that this is the last page to print
        ev.HasMorePages = False
    End Sub
    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click

        'Dim bmpGraph As Bitmap
        'AxMSChart1.EditCopy() 'put the chart in the clipboard
        'bmpGraph = System.Windows.Forms.Clipboard.GetDataObject().GetData(System.Windows.Forms.DataFormats.Bitmap)

        Try
            'specify current page settings
            PrintDocument1.DefaultPageSettings = PrintPageSettings
            'specify doc for print dialog box and show
            'StringToPrint = RichTextBox1.Text
            PrintDialog1.Document = PrintDocument1
            Dim result As DialogResult = PrintDialog1.ShowDialog()
            'if click ok, send doc to printer
            If result = DialogResult.OK Then
                AddHandler PrintDocument1.PrintPage, AddressOf Me.PrintGraphic
                PrintDocument1.Print()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub mnu_exportgraphdata_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_exportgraphdata.Click


    End Sub
End Class