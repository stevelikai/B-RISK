﻿Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Drawing.Printing
Imports BRISK.Branzfire.System.Windows.Forms.DataVisualization.Charting.Utilities

Public Class frmPlot



    Private Sub frmPlot_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = EventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = EventArgs.CloseReason
        eventArgs.Cancel = Cancel
    End Sub

    Private Sub frmPlot_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        For i = 1 To 12
            mnuLineColor.DropDownItems(i - 1).Visible = False
        Next
        For i = 1 To NumberRooms
            mnuLineColor.DropDownItems(i - 1).Visible = True
        Next
        If timeunit = 60 Then
            MinutesToolStripMenuItem.Checked = True
            SecondsToolStripMenuItem.Checked = False
        Else
            MinutesToolStripMenuItem.Checked = False
            SecondsToolStripMenuItem.Checked = True
        End If
        Me.CenterToParent()
    End Sub

    Private Sub HistogramTestToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            ' Populate "Default" chart series with random data
            Dim rand As Random
            rand = New Random()

            Chart1.Series.Clear()
            Chart1.Series.Add("Series1")

            Dim index As Integer
            For index = 1 To 200
                Chart1.Series("Series1").Points.AddY(rand.Next(1, 1000))
            Next index

            ' HistogramChartHelper is a helper class found in the samples Utilities folder. 
            Dim histogramHelper As New HistogramChartHelper()

            ' Show the percent frequency on the right Y axis.
            histogramHelper.ShowPercentOnSecondaryYAxis = False

            ' Specify number of segment intervals
            histogramHelper.SegmentIntervalNumber = 10

            ' Or you can specify the exact length of the interval
            ' histogramHelper.SegmentIntervalWidth = 15;
            ' Create histogram series    
            histogramHelper.CreateHistogram(Chart1, "Series1", "Histogram")
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in HistogramTestToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub InputsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub PageSetupToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageSetupToolStripMenuItem.Click
        ' Show Page Setup dialog
        Me.Chart1.Printing.PrintDocument = PrintDocument1
        Me.Chart1.Printing.PageSetup()
    End Sub

    Private Sub PreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PreToolStripMenuItem.Click
        ' Print preview chart
        'Me.Chart1.Printing.PrintPreview()
        Me.Chart1.BackColor = Color.White

        PrintPreviewDialog1.ShowDialog()
        Me.Chart1.BackColor = Color.AliceBlue
    End Sub
    Private Sub PrintToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem1.Click
        Try
            ' Print chart (with Printer dialog)
            Me.Chart1.Printing.PrintDocument = PrintDocument1
            Me.Chart1.BackColor = Color.White
            Me.Chart1.Printing.Print(True)
            Me.Chart1.BackColor = Color.AliceBlue
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in PrintToolStripMenuItem1_Click")
        End Try
    End Sub
    Private Sub MoreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Try
            ' Create and initialize print font 
            Dim printFont As New System.Drawing.Font("Arial", 10)
            ' Create Rectangle structure, used to set the position of the chart 
            Dim myRec As New System.Drawing.Rectangle(e.MarginBounds.Left, e.MarginBounds.Top, e.MarginBounds.Width, e.MarginBounds.Height / 3)

            ' Draw a line of text, followed by the chart, and then another line of text 
            'ev.Graphics.DrawString("Line before chart", printFont, Brushes.Black, ev.MarginBounds.Left, 10)
            Me.Chart1.Printing.PrintPaint(e.Graphics, myRec)
            'ev.Graphics.DrawString("Line after chart", printFont, Brushes.Black, ev.MarginBounds.Left, 350)
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in PrintDocument1_PrintPage")
        End Try
    End Sub

    Private Sub CopyToClipboardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToClipboardToolStripMenuItem.Click
        Try
            'get a temp name
            Dim tmp As String = My.Computer.FileSystem.GetTempFileName()

            Me.Chart1.BackColor = Color.White

            'save chart as image in temp file
            Chart1.SaveImage(tmp, ChartImageFormat.Bmp)

            Me.Chart1.BackColor = Color.AliceBlue

            'copy the chart to the clipboard
            Clipboard.SetImage(Image.FromFile(tmp))
        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in  CopyToClipboardToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Room1ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room1ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(0)
        Me.ColorDialog1.ShowDialog()
        roomcolor(0) = Me.ColorDialog1.Color
        My.Settings.color1 = roomcolor(0)
        My.Settings.Save()
        Me.Chart1.Series(0).Color = roomcolor(0)
    End Sub

    Private Sub Room2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room2ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(1)
        Me.ColorDialog1.ShowDialog()
        roomcolor(1) = Me.ColorDialog1.Color
        My.Settings.color2 = roomcolor(1)
        My.Settings.Save()
        Me.Chart1.Series(1).Color = roomcolor(1)
    End Sub

    Private Sub Room3ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room3ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(2)
        Me.ColorDialog1.ShowDialog()
        roomcolor(2) = Me.ColorDialog1.Color
        My.Settings.color3 = roomcolor(2)
        My.Settings.Save()
        Me.Chart1.Series(2).Color = roomcolor(2)
    End Sub

    Private Sub RooToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RooToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(3)
        Me.ColorDialog1.ShowDialog()
        roomcolor(3) = Me.ColorDialog1.Color
        My.Settings.color4 = roomcolor(3)
        My.Settings.Save()
        Me.Chart1.Series(3).Color = roomcolor(3)
    End Sub

    Private Sub Room5ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room5ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(4)
        Me.ColorDialog1.ShowDialog()
        roomcolor(4) = Me.ColorDialog1.Color
        My.Settings.color5 = roomcolor(4)
        My.Settings.Save()
        Me.Chart1.Series(4).Color = roomcolor(4)
    End Sub

    Private Sub Room6ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room6ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(5)
        Me.ColorDialog1.ShowDialog()
        roomcolor(5) = Me.ColorDialog1.Color
        My.Settings.color6 = roomcolor(5)
        My.Settings.Save()
        Me.Chart1.Series(5).Color = roomcolor(5)
    End Sub

    Private Sub Room7ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room7ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(6)
        Me.ColorDialog1.ShowDialog()
        roomcolor(6) = Me.ColorDialog1.Color
        My.Settings.color7 = roomcolor(6)
        My.Settings.Save()
        Me.Chart1.Series(6).Color = roomcolor(6)
    End Sub

    Private Sub Room8ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room8ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(7)
        Me.ColorDialog1.ShowDialog()
        roomcolor(7) = Me.ColorDialog1.Color
        My.Settings.color8 = roomcolor(7)
        My.Settings.Save()
        Me.Chart1.Series(7).Color = roomcolor(7)
    End Sub

    Private Sub Room9ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room9ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(8)
        Me.ColorDialog1.ShowDialog()
        roomcolor(8) = Me.ColorDialog1.Color
        My.Settings.color9 = roomcolor(8)
        My.Settings.Save()
        Me.Chart1.Series(8).Color = roomcolor(8)
    End Sub

    Private Sub Room10ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room10ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(9)
        Me.ColorDialog1.ShowDialog()
        roomcolor(9) = Me.ColorDialog1.Color
        My.Settings.color10 = roomcolor(9)
        My.Settings.Save()
        Me.Chart1.Series(9).Color = roomcolor(9)
    End Sub

    Private Sub Room11ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room11ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(10)
        Me.ColorDialog1.ShowDialog()
        roomcolor(10) = Me.ColorDialog1.Color
        My.Settings.color11 = roomcolor(10)
        My.Settings.Save()
        Me.Chart1.Series(10).Color = roomcolor(10)
    End Sub

    Private Sub Room12ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Room12ToolStripMenuItem.Click
        Me.ColorDialog1.Color = roomcolor(11)
        Me.ColorDialog1.ShowDialog()
        roomcolor(11) = Me.ColorDialog1.Color
        My.Settings.color12 = roomcolor(11)
        My.Settings.Save()
        Me.Chart1.Series(11).Color = roomcolor(11)

    End Sub

    Private Sub AutoResetToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoResetToolStripMenuItem.Click
        For i = 1 To 12
            roomcolor(i - 1) = System.Drawing.Color.Empty
        Next
        My.Settings.color1 = System.Drawing.Color.Empty
        My.Settings.color2 = System.Drawing.Color.Empty
        My.Settings.color3 = System.Drawing.Color.Empty
        My.Settings.color4 = System.Drawing.Color.Empty
        My.Settings.color5 = System.Drawing.Color.Empty
        My.Settings.color6 = System.Drawing.Color.Empty
        My.Settings.color7 = System.Drawing.Color.Empty
        My.Settings.color8 = System.Drawing.Color.Empty
        My.Settings.color9 = System.Drawing.Color.Empty
        My.Settings.color10 = System.Drawing.Color.Empty
        My.Settings.color11 = System.Drawing.Color.Empty
        My.Settings.color12 = System.Drawing.Color.Empty
        My.Settings.Save()

        For i = 1 To NumberRooms
            Me.Chart1.Series(i - 1).Color = roomcolor(i - 1)
        Next

    End Sub

    Private Sub ExportDataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExportDataToolStripMenuItem.Click

        Dim i, j, k As Integer
        Dim sname As String = ""
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        Try

            Dim getname As String = ""


            Select Case Chart1.ChartAreas("ChartArea1").AxisY.Title
                Case "ISO 834 - Surface Temperature (C)"
                    getname = RiskDataDirectory & "excel_furnaceST_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Surface Temperature"
                Case "ISO 834 - Net Heat Flux (kW/m2)"
                    getname = RiskDataDirectory & "excel_furnace_qnet_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Net Heat Flux"
                Case "ISO 834 - Normalised Heat Load (s1/2 K)"
                    getname = RiskDataDirectory & "excel_furnaceNHL_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Normalised Heat Load"
                Case "Nodal Ceiling Temperatures (C)"
                    getname = RiskDataDirectory & "excel_Cnodaltemps_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Nodal Ceiling Temperatures"
                    'Exit Sub
                Case "Nodal Upper Wall Temperatures (C)"
                    getname = RiskDataDirectory & "excel_UWnodaltemps_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Nodal Upper Wall Temperatures"
                    Exit Sub
                Case "Ceiling char depth (mm) - based on 300 C isotherm"
                    getname = RiskDataDirectory & "excel_Cchardepth_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Ceiling Char Depth"
                    Exit Sub
                Case "Upper wall char depth (mm) - based on 300 C isotherm"
                    getname = RiskDataDirectory & "excel_UWchardepth_" & Convert.ToString(frmInputs.txtBaseName.Text)
                    sname = "Upper Wall Char Depth"
                    Exit Sub
                Case Else
                    MsgBox("Not currently available, sorry.")
                    Exit Sub
            End Select

            'Start a new workbook in Excel
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add

            'Create an array
            'Dim DataArray(0 To SimTime / Timestep + 4, 0 To 3) As Object
            Dim DataArray(0 To 6 * 3600 / Timestep + 4, 0 To 3) As Object

            'Add headers to the worksheet on row 1
            oSheet = oBook.Worksheets(j + 1)
            oSheet.name = sname
            oSheet.Range("A1").Value = "time (sec)"
            oSheet.Range("B1").Value = "ceiling"
            oSheet.Range("C1").Value = "wall"
            oSheet.Range("D1").Value = "floor"

            For j = 0 To 2 'surface

                'For m = 0 To SimTime / Timestep
                For m = 0 To 6 * 3600
                    DataArray(m, 0) = m

                    Select Case Chart1.ChartAreas("ChartArea1").AxisY.Title
                        Case "ISO 834 - Surface Temperature (C)"
                            DataArray(m, j + 1) = furnaceST(j, m / Timestep) - 273
                        Case "ISO 834 - Net Heat Flux (kW/m2)"
                            DataArray(m, j + 1) = furnaceqnet(j, m / Timestep)
                        Case "ISO 834 - Normalised Heat Load (s1/2 K)"
                            DataArray(m, j + 1) = furnaceNHL(j, m / Timestep)
                        Case "ISO 834 - Normalised Heat Load (s1/2 K)"
                            DataArray(m, j + 1) = furnaceNHL(j, m / Timestep)
                        Case "ISO 834 - Normalised Heat Load (s1/2 K)"
                            DataArray(m, j + 1) = furnaceNHL(j, m / Timestep)

                    End Select

                Next

            Next
            'Transfer the array to the worksheet starting at cell A2
            'oSheet.Range("A2").Resize(SimTime / Timestep + 1, 4).Value = DataArray
            oSheet.Range("A2").Resize(6 * 3600 + 1, 4).Value = DataArray
            oBook.worksheets.add()

            'Save the Workbook and Quit Excel
            If Not String.IsNullOrEmpty(getname) Then oBook.SaveAs(getname)
            oBook.Close(SaveChanges:=False)
            oExcel.Quit()
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing
            Me.Close()
            MsgBox("File " & getname & " saved.", MsgBoxStyle.Information)

            Exit Sub

        Catch ex As Exception
            Me.Close()
        End Try

    End Sub

    Private Sub MinutesToolStripMenuItem_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MinutesToolStripMenuItem.CheckedChanged


    End Sub

    Private Sub MinutesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinutesToolStripMenuItem.Click
        MinutesToolStripMenuItem.Checked = True
        SecondsToolStripMenuItem.Checked = False
        timeunit = 60
        Me.Close()
    End Sub



    Private Sub SecondsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SecondsToolStripMenuItem.Click
        MinutesToolStripMenuItem.Checked = False
        SecondsToolStripMenuItem.Checked = True
        timeunit = 1
        Me.Close()
    End Sub
End Class