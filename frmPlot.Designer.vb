<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPlot
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Title1 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPlot))
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PageSetupToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.CopyToClipboardToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuLineColor = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room1ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room3ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RooToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room5ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room6ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room7ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room8ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room9ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room10ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room11ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Room12ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AutoResetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TimeUnitsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SecondsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MinutesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NumericUpDownTime = New System.Windows.Forms.NumericUpDown()
        Me.lblUpDownTime = New System.Windows.Forms.Label()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.NumericUpDownTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Chart1
        '
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Me.Chart1.Dock = System.Windows.Forms.DockStyle.Fill
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(0, 24)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Legend = "Legend1"
        Series1.Name = "Series1"
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Size = New System.Drawing.Size(641, 340)
        Me.Chart1.TabIndex = 0
        Me.Chart1.Text = "Chart1"
        Title1.Name = "Title1"
        Me.Chart1.Titles.Add(Title1)
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PrintToolStripMenuItem, Me.CopyToClipboardToolStripMenuItem, Me.mnuLineColor, Me.ExportDataToolStripMenuItem, Me.TimeUnitsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(641, 24)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PageSetupToolStripMenuItem, Me.PreToolStripMenuItem, Me.PrintToolStripMenuItem1})
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'PageSetupToolStripMenuItem
        '
        Me.PageSetupToolStripMenuItem.Name = "PageSetupToolStripMenuItem"
        Me.PageSetupToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.PageSetupToolStripMenuItem.Text = "Page Setup"
        '
        'PreToolStripMenuItem
        '
        Me.PreToolStripMenuItem.Name = "PreToolStripMenuItem"
        Me.PreToolStripMenuItem.Size = New System.Drawing.Size(133, 22)
        Me.PreToolStripMenuItem.Text = "Preview"
        '
        'PrintToolStripMenuItem1
        '
        Me.PrintToolStripMenuItem1.Name = "PrintToolStripMenuItem1"
        Me.PrintToolStripMenuItem1.Size = New System.Drawing.Size(133, 22)
        Me.PrintToolStripMenuItem1.Text = "Print"
        '
        'CopyToClipboardToolStripMenuItem
        '
        Me.CopyToClipboardToolStripMenuItem.Name = "CopyToClipboardToolStripMenuItem"
        Me.CopyToClipboardToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.CopyToClipboardToolStripMenuItem.Size = New System.Drawing.Size(116, 20)
        Me.CopyToClipboardToolStripMenuItem.Text = "Copy to Clipboard"
        '
        'mnuLineColor
        '
        Me.mnuLineColor.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Room1ToolStripMenuItem, Me.Room2ToolStripMenuItem, Me.Room3ToolStripMenuItem, Me.RooToolStripMenuItem, Me.Room5ToolStripMenuItem, Me.Room6ToolStripMenuItem, Me.Room7ToolStripMenuItem, Me.Room8ToolStripMenuItem, Me.Room9ToolStripMenuItem, Me.Room10ToolStripMenuItem, Me.Room11ToolStripMenuItem, Me.Room12ToolStripMenuItem, Me.AutoResetToolStripMenuItem})
        Me.mnuLineColor.Name = "mnuLineColor"
        Me.mnuLineColor.Size = New System.Drawing.Size(85, 20)
        Me.mnuLineColor.Text = "Line Colours"
        '
        'Room1ToolStripMenuItem
        '
        Me.Room1ToolStripMenuItem.Name = "Room1ToolStripMenuItem"
        Me.Room1ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room1ToolStripMenuItem.Text = "Room 1"
        '
        'Room2ToolStripMenuItem
        '
        Me.Room2ToolStripMenuItem.Name = "Room2ToolStripMenuItem"
        Me.Room2ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room2ToolStripMenuItem.Text = "Room 2"
        '
        'Room3ToolStripMenuItem
        '
        Me.Room3ToolStripMenuItem.Name = "Room3ToolStripMenuItem"
        Me.Room3ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room3ToolStripMenuItem.Text = "Room 3"
        '
        'RooToolStripMenuItem
        '
        Me.RooToolStripMenuItem.Name = "RooToolStripMenuItem"
        Me.RooToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.RooToolStripMenuItem.Text = "Room 4"
        '
        'Room5ToolStripMenuItem
        '
        Me.Room5ToolStripMenuItem.Name = "Room5ToolStripMenuItem"
        Me.Room5ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room5ToolStripMenuItem.Text = "Room 5"
        '
        'Room6ToolStripMenuItem
        '
        Me.Room6ToolStripMenuItem.Name = "Room6ToolStripMenuItem"
        Me.Room6ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room6ToolStripMenuItem.Text = "Room 6"
        '
        'Room7ToolStripMenuItem
        '
        Me.Room7ToolStripMenuItem.Name = "Room7ToolStripMenuItem"
        Me.Room7ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room7ToolStripMenuItem.Text = "Room 7"
        '
        'Room8ToolStripMenuItem
        '
        Me.Room8ToolStripMenuItem.Name = "Room8ToolStripMenuItem"
        Me.Room8ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room8ToolStripMenuItem.Text = "Room 8"
        '
        'Room9ToolStripMenuItem
        '
        Me.Room9ToolStripMenuItem.Name = "Room9ToolStripMenuItem"
        Me.Room9ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room9ToolStripMenuItem.Text = "Room 9"
        '
        'Room10ToolStripMenuItem
        '
        Me.Room10ToolStripMenuItem.Name = "Room10ToolStripMenuItem"
        Me.Room10ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room10ToolStripMenuItem.Text = "Room 10"
        '
        'Room11ToolStripMenuItem
        '
        Me.Room11ToolStripMenuItem.Name = "Room11ToolStripMenuItem"
        Me.Room11ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room11ToolStripMenuItem.Text = "Room 11"
        '
        'Room12ToolStripMenuItem
        '
        Me.Room12ToolStripMenuItem.Name = "Room12ToolStripMenuItem"
        Me.Room12ToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.Room12ToolStripMenuItem.Text = "Room 12"
        '
        'AutoResetToolStripMenuItem
        '
        Me.AutoResetToolStripMenuItem.Name = "AutoResetToolStripMenuItem"
        Me.AutoResetToolStripMenuItem.Size = New System.Drawing.Size(134, 22)
        Me.AutoResetToolStripMenuItem.Text = "Auto Select"
        '
        'ExportDataToolStripMenuItem
        '
        Me.ExportDataToolStripMenuItem.Name = "ExportDataToolStripMenuItem"
        Me.ExportDataToolStripMenuItem.Size = New System.Drawing.Size(79, 20)
        Me.ExportDataToolStripMenuItem.Text = "Export Data"
        '
        'TimeUnitsToolStripMenuItem
        '
        Me.TimeUnitsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SecondsToolStripMenuItem, Me.MinutesToolStripMenuItem})
        Me.TimeUnitsToolStripMenuItem.Name = "TimeUnitsToolStripMenuItem"
        Me.TimeUnitsToolStripMenuItem.Size = New System.Drawing.Size(75, 20)
        Me.TimeUnitsToolStripMenuItem.Text = "Time units"
        '
        'SecondsToolStripMenuItem
        '
        Me.SecondsToolStripMenuItem.Checked = True
        Me.SecondsToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.SecondsToolStripMenuItem.Name = "SecondsToolStripMenuItem"
        Me.SecondsToolStripMenuItem.Size = New System.Drawing.Size(117, 22)
        Me.SecondsToolStripMenuItem.Text = "seconds"
        '
        'MinutesToolStripMenuItem
        '
        Me.MinutesToolStripMenuItem.Name = "MinutesToolStripMenuItem"
        Me.MinutesToolStripMenuItem.Size = New System.Drawing.Size(117, 22)
        Me.MinutesToolStripMenuItem.Text = "minutes"
        '
        'NumericUpDownTime
        '
        Me.NumericUpDownTime.AutoSize = True
        Me.NumericUpDownTime.Location = New System.Drawing.Point(526, 4)
        Me.NumericUpDownTime.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDownTime.Name = "NumericUpDownTime"
        Me.NumericUpDownTime.Size = New System.Drawing.Size(76, 20)
        Me.NumericUpDownTime.TabIndex = 2
        Me.NumericUpDownTime.Value = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDownTime.Visible = False
        '
        'lblUpDownTime
        '
        Me.lblUpDownTime.AutoSize = True
        Me.lblUpDownTime.BackColor = System.Drawing.SystemColors.Desktop
        Me.lblUpDownTime.Location = New System.Drawing.Point(464, 11)
        Me.lblUpDownTime.Name = "lblUpDownTime"
        Me.lblUpDownTime.Size = New System.Drawing.Size(56, 13)
        Me.lblUpDownTime.TabIndex = 3
        Me.lblUpDownTime.Text = "Time (sec)"
        Me.lblUpDownTime.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblUpDownTime.Visible = False
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Document = Me.PrintDocument1
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'frmPlot
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(641, 364)
        Me.Controls.Add(Me.lblUpDownTime)
        Me.Controls.Add(Me.NumericUpDownTime)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmPlot"
        Me.TopMost = True
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.NumericUpDownTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Chart1 As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PageSetupToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NumericUpDownTime As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblUpDownTime As System.Windows.Forms.Label
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents CopyToClipboardToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuLineColor As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room1ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room2ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room3ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RooToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room5ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room6ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room7ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents Room8ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room9ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room10ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room11ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Room12ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AutoResetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExportDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TimeUnitsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SecondsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MinutesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
