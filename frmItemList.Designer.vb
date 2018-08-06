<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmItemList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.lstItems = New System.Windows.Forms.ListBox()
        Me.cmdAdd = New System.Windows.Forms.Button()
        Me.cmdRemove = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdIgniteFirst = New System.Windows.Forms.Button()
        Me.cmdRandomIgnite = New System.Windows.Forms.Button()
        Me.cmdRoomPopulate = New System.Windows.Forms.Button()
        Me.txtFLED = New System.Windows.Forms.TextBox()
        Me.cmdCopy = New System.Windows.Forms.Button()
        Me.cmdImport = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdAlphaT = New System.Windows.Forms.Button()
        Me.chkAlphaTFire = New System.Windows.Forms.CheckBox()
        Me.lstFireRoom2 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cmdDist_FLED = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.optFREoff = New System.Windows.Forms.RadioButton()
        Me.optFREon = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 9)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(340, 17)
        Me.ListBox1.TabIndex = 0
        '
        'lstItems
        '
        Me.lstItems.FormattingEnabled = True
        Me.lstItems.Location = New System.Drawing.Point(12, 35)
        Me.lstItems.Name = "lstItems"
        Me.lstItems.Size = New System.Drawing.Size(340, 108)
        Me.lstItems.TabIndex = 1
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(362, 12)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
        Me.cmdAdd.TabIndex = 2
        Me.cmdAdd.Text = "Add"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdRemove
        '
        Me.cmdRemove.Location = New System.Drawing.Point(362, 41)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(75, 23)
        Me.cmdRemove.TabIndex = 3
        Me.cmdRemove.Text = "Remove"
        Me.cmdRemove.UseVisualStyleBackColor = True
        '
        'cmdEdit
        '
        Me.cmdEdit.Location = New System.Drawing.Point(362, 70)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(75, 23)
        Me.cmdEdit.TabIndex = 4
        Me.cmdEdit.Text = "Edit"
        Me.cmdEdit.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(374, 395)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(75, 23)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'cmdIgniteFirst
        '
        Me.cmdIgniteFirst.Location = New System.Drawing.Point(13, 157)
        Me.cmdIgniteFirst.Name = "cmdIgniteFirst"
        Me.cmdIgniteFirst.Size = New System.Drawing.Size(91, 23)
        Me.cmdIgniteFirst.TabIndex = 9
        Me.cmdIgniteFirst.Text = "First Ignite"
        Me.ToolTip1.SetToolTip(Me.cmdIgniteFirst, "Item selected in list will be the first item to be ignite")
        Me.cmdIgniteFirst.UseVisualStyleBackColor = True
        '
        'cmdRandomIgnite
        '
        Me.cmdRandomIgnite.Location = New System.Drawing.Point(110, 157)
        Me.cmdRandomIgnite.Name = "cmdRandomIgnite"
        Me.cmdRandomIgnite.Size = New System.Drawing.Size(91, 23)
        Me.cmdRandomIgnite.TabIndex = 10
        Me.cmdRandomIgnite.Text = "Random Ignite"
        Me.ToolTip1.SetToolTip(Me.cmdRandomIgnite, "The first item ignited will be randomly selected from the items listed")
        Me.cmdRandomIgnite.UseVisualStyleBackColor = True
        '
        'cmdRoomPopulate
        '
        Me.cmdRoomPopulate.Location = New System.Drawing.Point(252, 157)
        Me.cmdRoomPopulate.Name = "cmdRoomPopulate"
        Me.cmdRoomPopulate.Size = New System.Drawing.Size(101, 23)
        Me.cmdRoomPopulate.TabIndex = 24
        Me.cmdRoomPopulate.Text = "Room Populate"
        Me.ToolTip1.SetToolTip(Me.cmdRoomPopulate, "The first item ignited will be randomly selected from the items listed")
        Me.cmdRoomPopulate.UseVisualStyleBackColor = True
        '
        'txtFLED
        '
        Me.txtFLED.AcceptsReturn = True
        Me.txtFLED.BackColor = System.Drawing.SystemColors.Info
        Me.txtFLED.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFLED.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFLED.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFLED.Location = New System.Drawing.Point(290, 22)
        Me.txtFLED.MaxLength = 0
        Me.txtFLED.Name = "txtFLED"
        Me.txtFLED.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFLED.Size = New System.Drawing.Size(49, 20)
        Me.txtFLED.TabIndex = 21
        Me.txtFLED.Text = "400"
        Me.ToolTip1.SetToolTip(Me.txtFLED, "Total FLED for the room includes burning objects")
        '
        'cmdCopy
        '
        Me.cmdCopy.Location = New System.Drawing.Point(361, 99)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.Size = New System.Drawing.Size(75, 23)
        Me.cmdCopy.TabIndex = 8
        Me.cmdCopy.Text = "Copy"
        Me.cmdCopy.UseVisualStyleBackColor = True
        '
        'cmdImport
        '
        Me.cmdImport.Location = New System.Drawing.Point(361, 128)
        Me.cmdImport.Name = "cmdImport"
        Me.cmdImport.Size = New System.Drawing.Size(75, 23)
        Me.cmdImport.TabIndex = 11
        Me.cmdImport.Text = "Import"
        Me.cmdImport.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(362, 157)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "web"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'cmdAlphaT
        '
        Me.cmdAlphaT.Location = New System.Drawing.Point(40, 43)
        Me.cmdAlphaT.Name = "cmdAlphaT"
        Me.cmdAlphaT.Size = New System.Drawing.Size(91, 23)
        Me.cmdAlphaT.TabIndex = 1
        Me.cmdAlphaT.Text = "Settings"
        Me.cmdAlphaT.UseVisualStyleBackColor = True
        '
        'chkAlphaTFire
        '
        Me.chkAlphaTFire.AutoSize = True
        Me.chkAlphaTFire.Location = New System.Drawing.Point(18, 18)
        Me.chkAlphaTFire.Name = "chkAlphaTFire"
        Me.chkAlphaTFire.Size = New System.Drawing.Size(234, 17)
        Me.chkAlphaTFire.TabIndex = 13
        Me.chkAlphaTFire.Text = "Use Power Law Design Fire (overrides DFG)"
        Me.chkAlphaTFire.UseVisualStyleBackColor = True
        '
        'lstFireRoom2
        '
        Me.lstFireRoom2.BackColor = System.Drawing.SystemColors.Info
        Me.lstFireRoom2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstFireRoom2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.lstFireRoom2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstFireRoom2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstFireRoom2.Location = New System.Drawing.Point(380, 236)
        Me.lstFireRoom2.MaxDropDownItems = 10
        Me.lstFireRoom2.Name = "lstFireRoom2"
        Me.lstFireRoom2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstFireRoom2.Size = New System.Drawing.Size(57, 22)
        Me.lstFireRoom2.TabIndex = 14
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkAlphaTFire)
        Me.GroupBox1.Controls.Add(Me.cmdAlphaT)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 187)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(347, 82)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(366, 261)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(71, 17)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Fire Room"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdDist_FLED)
        Me.GroupBox2.Controls.Add(Me.txtFLED)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Location = New System.Drawing.Point(13, 275)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(436, 59)
        Me.GroupBox2.TabIndex = 159
        Me.GroupBox2.TabStop = False
        '
        'cmdDist_FLED
        '
        Me.cmdDist_FLED.Location = New System.Drawing.Point(356, 24)
        Me.cmdDist_FLED.Name = "cmdDist_FLED"
        Me.cmdDist_FLED.Size = New System.Drawing.Size(68, 19)
        Me.cmdDist_FLED.TabIndex = 154
        Me.cmdDist_FLED.Text = "distribution"
        Me.cmdDist_FLED.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(15, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(265, 17)
        Me.Label2.TabIndex = 24
        Me.Label2.Text = "Fire Load Energy Density (per unit floor area) MJ/m²"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.optFREoff)
        Me.GroupBox3.Controls.Add(Me.optFREon)
        Me.GroupBox3.Location = New System.Drawing.Point(13, 340)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(347, 78)
        Me.GroupBox3.TabIndex = 172
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Fuel Response Effects"
        Me.GroupBox3.Visible = False
        '
        'optFREoff
        '
        Me.optFREoff.AutoSize = True
        Me.optFREoff.Checked = True
        Me.optFREoff.Location = New System.Drawing.Point(97, 34)
        Me.optFREoff.Name = "optFREoff"
        Me.optFREoff.Size = New System.Drawing.Size(39, 17)
        Me.optFREoff.TabIndex = 1
        Me.optFREoff.TabStop = True
        Me.optFREoff.Text = "Off"
        Me.optFREoff.UseVisualStyleBackColor = True
        '
        'optFREon
        '
        Me.optFREon.AutoSize = True
        Me.optFREon.Location = New System.Drawing.Point(19, 34)
        Me.optFREon.Name = "optFREon"
        Me.optFREon.Size = New System.Drawing.Size(39, 17)
        Me.optFREon.TabIndex = 0
        Me.optFREon.Text = "On"
        Me.optFREon.UseVisualStyleBackColor = True
        '
        'frmItemList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(464, 426)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdRoomPopulate)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lstFireRoom2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.cmdImport)
        Me.Controls.Add(Me.cmdRandomIgnite)
        Me.Controls.Add(Me.cmdIgniteFirst)
        Me.Controls.Add(Me.cmdCopy)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.cmdRemove)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.lstItems)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmItemList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Items"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lstItems As System.Windows.Forms.ListBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdCopy As System.Windows.Forms.Button
    Friend WithEvents cmdIgniteFirst As System.Windows.Forms.Button
    Friend WithEvents cmdRandomIgnite As System.Windows.Forms.Button
    Friend WithEvents cmdImport As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents cmdAlphaT As System.Windows.Forms.Button
    Friend WithEvents chkAlphaTFire As System.Windows.Forms.CheckBox
    Public WithEvents lstFireRoom2 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdRoomPopulate As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdDist_FLED As System.Windows.Forms.Button
    Public WithEvents txtFLED As System.Windows.Forms.TextBox
    Public WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents optFREoff As RadioButton
    Friend WithEvents optFREon As RadioButton
End Class
