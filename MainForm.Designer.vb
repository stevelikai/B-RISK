﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMainForm
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboCountries1 = New System.Windows.Forms.ComboBox()
        Me.cmdWhere2 = New System.Windows.Forms.Button()
        Me.txtExecutingQueryStatement = New System.Windows.Forms.TextBox()
        Me.cmdWhere1 = New System.Windows.Forms.Button()
        Me.cboItems = New System.Windows.Forms.ComboBox()
        Me.cmdRunCommand = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.cboCountries1)
        Me.Panel1.Controls.Add(Me.cmdWhere2)
        Me.Panel1.Controls.Add(Me.txtExecutingQueryStatement)
        Me.Panel1.Controls.Add(Me.cmdWhere1)
        Me.Panel1.Controls.Add(Me.cboItems)
        Me.Panel1.Controls.Add(Me.cmdRunCommand)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 196)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(664, 154)
        Me.Panel1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(509, 128)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "<-- Procedure"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(509, 92)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "<-- View"
        '
        'cboCountries1
        '
        Me.cboCountries1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboCountries1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCountries1.FormattingEnabled = True
        Me.cboCountries1.Location = New System.Drawing.Point(281, 89)
        Me.cboCountries1.Name = "cboCountries1"
        Me.cboCountries1.Size = New System.Drawing.Size(142, 21)
        Me.cboCountries1.TabIndex = 3
        '
        'cmdWhere2
        '
        Me.cmdWhere2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdWhere2.Location = New System.Drawing.Point(429, 124)
        Me.cmdWhere2.Name = "cmdWhere2"
        Me.cmdWhere2.Size = New System.Drawing.Size(75, 23)
        Me.cmdWhere2.TabIndex = 6
        Me.cmdWhere2.Text = "Where 2"
        Me.cmdWhere2.UseVisualStyleBackColor = True
        '
        'txtExecutingQueryStatement
        '
        Me.txtExecutingQueryStatement.BackColor = System.Drawing.Color.LightGray
        Me.txtExecutingQueryStatement.ForeColor = System.Drawing.Color.Maroon
        Me.txtExecutingQueryStatement.Location = New System.Drawing.Point(12, 6)
        Me.txtExecutingQueryStatement.Multiline = True
        Me.txtExecutingQueryStatement.Name = "txtExecutingQueryStatement"
        Me.txtExecutingQueryStatement.ReadOnly = True
        Me.txtExecutingQueryStatement.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtExecutingQueryStatement.Size = New System.Drawing.Size(640, 77)
        Me.txtExecutingQueryStatement.TabIndex = 0
        '
        'cmdWhere1
        '
        Me.cmdWhere1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdWhere1.Location = New System.Drawing.Point(429, 88)
        Me.cmdWhere1.Name = "cmdWhere1"
        Me.cmdWhere1.Size = New System.Drawing.Size(75, 23)
        Me.cmdWhere1.TabIndex = 4
        Me.cmdWhere1.Text = "Where 1"
        Me.cmdWhere1.UseVisualStyleBackColor = True
        '
        'cboItems
        '
        Me.cboItems.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cboItems.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItems.FormattingEnabled = True
        Me.cboItems.Location = New System.Drawing.Point(12, 89)
        Me.cboItems.Name = "cboItems"
        Me.cboItems.Size = New System.Drawing.Size(142, 21)
        Me.cboItems.TabIndex = 1
        '
        'cmdRunCommand
        '
        Me.cmdRunCommand.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdRunCommand.Location = New System.Drawing.Point(164, 88)
        Me.cmdRunCommand.Name = "cmdRunCommand"
        Me.cmdRunCommand.Size = New System.Drawing.Size(75, 23)
        Me.cmdRunCommand.TabIndex = 2
        Me.cmdRunCommand.Text = "Run"
        Me.cmdRunCommand.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(664, 196)
        Me.DataGridView1.TabIndex = 0
        '
        'frmMainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(664, 350)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmMainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Execute queries stored in MS-Access database demo"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents cmdRunCommand As System.Windows.Forms.Button
    Friend WithEvents cboItems As System.Windows.Forms.ComboBox
    Friend WithEvents cmdWhere1 As System.Windows.Forms.Button
    Friend WithEvents txtExecutingQueryStatement As System.Windows.Forms.TextBox
    Friend WithEvents cboCountries1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmdWhere2 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
