Imports System.Drawing.Printing
Public Class frmViewDocs
    Private PrintPageSettings As New PageSettings
    Private StringToPrint As String
    Private PrintFont As New Font("Courier New", 8.25)
 
    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim numChars As Integer
        Dim numLines As Integer
        Dim stringForPage As String
        Dim strFormat As New StringFormat
        'based on page setup, define drawable rectangle on page
        Dim rectDraw As New RectangleF(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height)
        'define area to determine how much text can display on a page
        'make height one line shorter to ensure text doesn't clip
        Dim sizeMeasure As New SizeF(e.MarginBounds.Width, e.MarginBounds.Height - PrintFont.GetHeight(e.Graphics))

        'when drawing long strings, break between words
        strFormat.Trimming = StringTrimming.Word
        'compute how many chars and lines can fit based on sizemeasure
        e.Graphics.MeasureString(StringToPrint, PrintFont, sizeMeasure, strFormat, numChars, numLines)
        'compute string that will fit on a page
        stringForPage = StringToPrint.Substring(0, numChars)
        'print string on current page
        e.Graphics.DrawString(stringForPage, PrintFont, Brushes.Black, rectDraw, strFormat)
        'if there is more text indicate there are more pages
        If numChars < StringToPrint.Length Then
            'subtract text from string that has been printed
            StringToPrint = StringToPrint.Substring(numChars)
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            'all text has been printed so restore string
            StringToPrint = RichTextBox1.Text
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub PrintToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem1.Click
        Try

            If My.Computer.Clipboard.ContainsImage() Then
                Dim grabpicture As System.Drawing.Image
                grabpicture = My.Computer.Clipboard.GetImage()
                'PictureBox1.Image = grabpicture
            End If

            'specify current page settings
            PrintDocument1.DefaultPageSettings = PrintPageSettings
            'specify doc for print dialog box and show
            StringToPrint = RichTextBox1.Text
            PrintDialog1.Document = PrintDocument1
            PrintPreviewDialog1.BringToFront()
            Dim result As DialogResult = PrintDialog1.ShowDialog()
            'if click ok, send doc to printer
            If result = DialogResult.OK Then
                PrintDocument1.Print()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PageSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageSetupToolStripMenuItem.Click
        Try
            'load page settings and display dialog box
            PageSetupDialog1.PageSettings = PrintPageSettings
            PrintPreviewDialog1.BringToFront()
            PageSetupDialog1.ShowDialog()
        Catch ex As Exception
            'display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Try
            'specify current page settings
            PrintDocument1.DefaultPageSettings = PrintPageSettings
            'specify doc for print preview dialog box and show
            StringToPrint = RichTextBox1.Text
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.BringToFront()
            PrintPreviewDialog1.ShowDialog()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub SelectVariablesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectVariablesToolStripMenuItem.Click
        frmprintvar.Show()

    End Sub

    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        Try
            Clipboard.SetText(RichTextBox1.Text)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmViewDocs.vb ToolStripMenuItem10_Click")
        End Try
    End Sub

    Private Sub SaveLogTortfFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveLogTortfFileToolStripMenuItem.Click
        Try

            Dim tmp As String
            tmp = basefile.Replace(".xml", "_wallventflow.rtf")
            'If My.Computer.FileSystem.FileExists(tmp) Then
            RichTextBox1.SaveFile(tmp)
            'End If



        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in frmViewDocs.vb SaveLogTortfFileToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub frmViewDocs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class