<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRoomList
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
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.lstRooms = New System.Windows.Forms.ListBox
        Me.cmdAddRoom = New System.Windows.Forms.Button
        Me.cmdEditRoom = New System.Windows.Forms.Button
        Me.cmdCopyRoom = New System.Windows.Forms.Button
        Me.cmdRemoveRoom = New System.Windows.Forms.Button
        Me.cmdExit = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.ListBox1.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Location = New System.Drawing.Point(12, 12)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(462, 18)
        Me.ListBox1.TabIndex = 6
        '
        'lstRooms
        '
        Me.lstRooms.ColumnWidth = 25
        Me.lstRooms.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstRooms.FormattingEnabled = True
        Me.lstRooms.ItemHeight = 14
        Me.lstRooms.Location = New System.Drawing.Point(12, 32)
        Me.lstRooms.Name = "lstRooms"
        Me.lstRooms.Size = New System.Drawing.Size(462, 186)
        Me.lstRooms.TabIndex = 7
        '
        'cmdAddRoom
        '
        Me.cmdAddRoom.Location = New System.Drawing.Point(480, 32)
        Me.cmdAddRoom.Name = "cmdAddRoom"
        Me.cmdAddRoom.Size = New System.Drawing.Size(127, 23)
        Me.cmdAddRoom.TabIndex = 8
        Me.cmdAddRoom.Text = "Add Room"
        Me.cmdAddRoom.UseVisualStyleBackColor = True
        '
        'cmdEditRoom
        '
        Me.cmdEditRoom.Location = New System.Drawing.Point(480, 61)
        Me.cmdEditRoom.Name = "cmdEditRoom"
        Me.cmdEditRoom.Size = New System.Drawing.Size(127, 24)
        Me.cmdEditRoom.TabIndex = 9
        Me.cmdEditRoom.Text = "Edit"
        Me.cmdEditRoom.UseVisualStyleBackColor = True
        '
        'cmdCopyRoom
        '
        Me.cmdCopyRoom.Location = New System.Drawing.Point(480, 91)
        Me.cmdCopyRoom.Name = "cmdCopyRoom"
        Me.cmdCopyRoom.Size = New System.Drawing.Size(127, 23)
        Me.cmdCopyRoom.TabIndex = 10
        Me.cmdCopyRoom.Text = "Copy"
        Me.cmdCopyRoom.UseVisualStyleBackColor = True
        '
        'cmdRemoveRoom
        '
        Me.cmdRemoveRoom.Location = New System.Drawing.Point(480, 120)
        Me.cmdRemoveRoom.Name = "cmdRemoveRoom"
        Me.cmdRemoveRoom.Size = New System.Drawing.Size(127, 23)
        Me.cmdRemoveRoom.TabIndex = 11
        Me.cmdRemoveRoom.Text = "Remove"
        Me.cmdRemoveRoom.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(480, 211)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(127, 23)
        Me.cmdExit.TabIndex = 12
        Me.cmdExit.Text = "Close"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'frmRoomList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(615, 255)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdRemoveRoom)
        Me.Controls.Add(Me.cmdCopyRoom)
        Me.Controls.Add(Me.cmdEditRoom)
        Me.Controls.Add(Me.cmdAddRoom)
        Me.Controls.Add(Me.lstRooms)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MinimizeBox = False
        Me.Name = "frmRoomList"
        Me.Text = "Manage Rooms"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lstRooms As System.Windows.Forms.ListBox
    Friend WithEvents cmdAddRoom As System.Windows.Forms.Button
    Friend WithEvents cmdEditRoom As System.Windows.Forms.Button
    Friend WithEvents cmdCopyRoom As System.Windows.Forms.Button
    Friend WithEvents cmdRemoveRoom As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
End Class
