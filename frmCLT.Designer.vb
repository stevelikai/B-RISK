<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCLT
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCLT))
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.cmdClose = New System.Windows.Forms.Button()
        Me.optCLTOFF = New System.Windows.Forms.RadioButton()
        Me.optCLTON = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.numeric_wallareapercent = New System.Windows.Forms.NumericUpDown()
        Me.numeric_ceilareapercent = New System.Windows.Forms.NumericUpDown()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCharTemp = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtLamellaDepth = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtCritFlux = New System.Windows.Forms.TextBox()
        Me.txtFlameFlux = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtCLTLoG = New System.Windows.Forms.TextBox()
        Me.txtCLTigtemp = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtCLTcalibration = New System.Windows.Forms.TextBox()
        Me.Panel_integralmodel = New System.Windows.Forms.Panel()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.lblDebondTemp = New System.Windows.Forms.Label()
        Me.txtDebondTemp = New System.Windows.Forms.TextBox()
        Me.RB_dynamic = New System.Windows.Forms.RadioButton()
        Me.RB_Integral = New System.Windows.Forms.RadioButton()
        Me.RB_Kinetic = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numeric_wallareapercent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numeric_ceilareapercent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_integralmodel.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'cmdClose
        '
        Me.cmdClose.Location = New System.Drawing.Point(290, 344)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(75, 23)
        Me.cmdClose.TabIndex = 0
        Me.cmdClose.Text = "Close"
        Me.cmdClose.UseVisualStyleBackColor = True
        '
        'optCLTOFF
        '
        Me.optCLTOFF.AutoSize = True
        Me.optCLTOFF.Checked = True
        Me.optCLTOFF.Location = New System.Drawing.Point(28, 21)
        Me.optCLTOFF.Name = "optCLTOFF"
        Me.optCLTOFF.Size = New System.Drawing.Size(101, 17)
        Me.optCLTOFF.TabIndex = 1
        Me.optCLTOFF.TabStop = True
        Me.optCLTOFF.Text = "CLT model is off"
        Me.optCLTOFF.UseVisualStyleBackColor = True
        '
        'optCLTON
        '
        Me.optCLTON.AutoSize = True
        Me.optCLTON.Location = New System.Drawing.Point(148, 21)
        Me.optCLTON.Name = "optCLTON"
        Me.optCLTON.Size = New System.Drawing.Size(101, 17)
        Me.optCLTON.TabIndex = 2
        Me.optCLTON.Text = "CLT model is on"
        Me.optCLTON.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 3
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(36, 60)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(293, 145)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = resources.GetString("TextBox1.Text")
        '
        'numeric_wallareapercent
        '
        Me.numeric_wallareapercent.Location = New System.Drawing.Point(210, 29)
        Me.numeric_wallareapercent.Name = "numeric_wallareapercent"
        Me.numeric_wallareapercent.Size = New System.Drawing.Size(59, 20)
        Me.numeric_wallareapercent.TabIndex = 5
        Me.numeric_wallareapercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'numeric_ceilareapercent
        '
        Me.numeric_ceilareapercent.Location = New System.Drawing.Point(210, 59)
        Me.numeric_ceilareapercent.Name = "numeric_ceilareapercent"
        Me.numeric_ceilareapercent.Size = New System.Drawing.Size(59, 20)
        Me.numeric_ceilareapercent.TabIndex = 6
        Me.numeric_ceilareapercent.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(76, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Contribuing wall area (%)  "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(64, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(137, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Contribuing ceiling area (%) "
        '
        'txtCharTemp
        '
        Me.txtCharTemp.Location = New System.Drawing.Point(210, 90)
        Me.txtCharTemp.Name = "txtCharTemp"
        Me.txtCharTemp.Size = New System.Drawing.Size(59, 20)
        Me.txtCharTemp.TabIndex = 9
        Me.txtCharTemp.Text = "300"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(97, 93)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Char temperature (C)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(91, 161)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(108, 13)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Lamella thickness (m)"
        '
        'txtLamellaDepth
        '
        Me.txtLamellaDepth.Location = New System.Drawing.Point(210, 158)
        Me.txtLamellaDepth.Name = "txtLamellaDepth"
        Me.txtLamellaDepth.Size = New System.Drawing.Size(59, 20)
        Me.txtLamellaDepth.TabIndex = 13
        Me.txtLamellaDepth.Text = "0.020"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(70, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(126, 13)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Critical heat flux (kW/m2)"
        '
        'txtCritFlux
        '
        Me.txtCritFlux.Location = New System.Drawing.Point(207, 19)
        Me.txtCritFlux.Name = "txtCritFlux"
        Me.txtCritFlux.Size = New System.Drawing.Size(60, 20)
        Me.txtCritFlux.TabIndex = 15
        Me.txtCritFlux.Text = "16"
        '
        'txtFlameFlux
        '
        Me.txtFlameFlux.Location = New System.Drawing.Point(208, 49)
        Me.txtFlameFlux.Name = "txtFlameFlux"
        Me.txtFlameFlux.Size = New System.Drawing.Size(59, 20)
        Me.txtFlameFlux.TabIndex = 16
        Me.txtFlameFlux.Text = "17"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(97, 52)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(99, 13)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Flame flux (kW/m2)"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(36, 83)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(160, 13)
        Me.Label8.TabIndex = 18
        Me.Label8.Text = "Latent heat of gasification (kJ/g)"
        '
        'txtCLTLoG
        '
        Me.txtCLTLoG.Location = New System.Drawing.Point(208, 80)
        Me.txtCLTLoG.Name = "txtCLTLoG"
        Me.txtCLTLoG.Size = New System.Drawing.Size(59, 20)
        Me.txtCLTLoG.TabIndex = 19
        Me.txtCLTLoG.Text = "1.6"
        '
        'txtCLTigtemp
        '
        Me.txtCLTigtemp.Location = New System.Drawing.Point(208, 113)
        Me.txtCLTigtemp.Name = "txtCLTigtemp"
        Me.txtCLTigtemp.Size = New System.Drawing.Size(59, 20)
        Me.txtCLTigtemp.TabIndex = 20
        Me.txtCLTigtemp.Text = "384"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(80, 116)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(116, 13)
        Me.Label9.TabIndex = 21
        Me.Label9.Text = "Ignition temperature (C)"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(110, 149)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(86, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Calibration factor"
        '
        'txtCLTcalibration
        '
        Me.txtCLTcalibration.Location = New System.Drawing.Point(207, 146)
        Me.txtCLTcalibration.Name = "txtCLTcalibration"
        Me.txtCLTcalibration.Size = New System.Drawing.Size(60, 20)
        Me.txtCLTcalibration.TabIndex = 23
        Me.txtCLTcalibration.Text = "1"
        '
        'Panel_integralmodel
        '
        Me.Panel_integralmodel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTcalibration)
        Me.Panel_integralmodel.Controls.Add(Me.Label10)
        Me.Panel_integralmodel.Controls.Add(Me.Label9)
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTigtemp)
        Me.Panel_integralmodel.Controls.Add(Me.txtCLTLoG)
        Me.Panel_integralmodel.Controls.Add(Me.Label8)
        Me.Panel_integralmodel.Controls.Add(Me.Label7)
        Me.Panel_integralmodel.Controls.Add(Me.txtFlameFlux)
        Me.Panel_integralmodel.Controls.Add(Me.txtCritFlux)
        Me.Panel_integralmodel.Controls.Add(Me.Label6)
        Me.Panel_integralmodel.Location = New System.Drawing.Point(383, 224)
        Me.Panel_integralmodel.Name = "Panel_integralmodel"
        Me.Panel_integralmodel.Size = New System.Drawing.Size(300, 182)
        Me.Panel_integralmodel.TabIndex = 24
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(250, 30)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(79, 13)
        Me.LinkLabel1.TabIndex = 25
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "See Reference"
        '
        'lblDebondTemp
        '
        Me.lblDebondTemp.AutoSize = True
        Me.lblDebondTemp.Location = New System.Drawing.Point(18, 124)
        Me.lblDebondTemp.Name = "lblDebondTemp"
        Me.lblDebondTemp.Size = New System.Drawing.Size(179, 13)
        Me.lblDebondTemp.TabIndex = 26
        Me.lblDebondTemp.Text = "Adhesive debonding temperature (C)"
        '
        'txtDebondTemp
        '
        Me.txtDebondTemp.Location = New System.Drawing.Point(210, 125)
        Me.txtDebondTemp.Name = "txtDebondTemp"
        Me.txtDebondTemp.Size = New System.Drawing.Size(58, 20)
        Me.txtDebondTemp.TabIndex = 27
        Me.txtDebondTemp.Text = "200"
        '
        'RB_dynamic
        '
        Me.RB_dynamic.AutoSize = True
        Me.RB_dynamic.Checked = True
        Me.RB_dynamic.Location = New System.Drawing.Point(19, 28)
        Me.RB_dynamic.Name = "RB_dynamic"
        Me.RB_dynamic.Size = New System.Drawing.Size(229, 17)
        Me.RB_dynamic.TabIndex = 0
        Me.RB_dynamic.TabStop = True
        Me.RB_dynamic.Text = "Global equivalence charring model (default)"
        Me.RB_dynamic.UseVisualStyleBackColor = True
        '
        'RB_Integral
        '
        Me.RB_Integral.AutoSize = True
        Me.RB_Integral.Location = New System.Drawing.Point(19, 204)
        Me.RB_Integral.Name = "RB_Integral"
        Me.RB_Integral.Size = New System.Drawing.Size(162, 17)
        Me.RB_Integral.TabIndex = 1
        Me.RB_Integral.Text = "Integral wood pyrolysis model"
        Me.RB_Integral.UseVisualStyleBackColor = True
        '
        'RB_Kinetic
        '
        Me.RB_Kinetic.AutoSize = True
        Me.RB_Kinetic.Location = New System.Drawing.Point(19, 240)
        Me.RB_Kinetic.Name = "RB_Kinetic"
        Me.RB_Kinetic.Size = New System.Drawing.Size(159, 17)
        Me.RB_Kinetic.TabIndex = 2
        Me.RB_Kinetic.Text = "Kinetic wood pyrolysis model"
        Me.RB_Kinetic.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LinkLabel1)
        Me.GroupBox1.Controls.Add(Me.RB_Kinetic)
        Me.GroupBox1.Controls.Add(Me.RB_dynamic)
        Me.GroupBox1.Controls.Add(Me.RB_Integral)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(28, 59)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(337, 276)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select wood pyrolysis model"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.txtLamellaDepth)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.txtDebondTemp)
        Me.GroupBox2.Controls.Add(Me.numeric_wallareapercent)
        Me.GroupBox2.Controls.Add(Me.lblDebondTemp)
        Me.GroupBox2.Controls.Add(Me.numeric_ceilareapercent)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.txtCharTemp)
        Me.GroupBox2.Location = New System.Drawing.Point(383, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(300, 206)
        Me.GroupBox2.TabIndex = 31
        Me.GroupBox2.TabStop = False
        '
        'frmCLT
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(695, 421)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel_integralmodel)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.optCLTON)
        Me.Controls.Add(Me.optCLTOFF)
        Me.Controls.Add(Me.cmdClose)
        Me.Name = "frmCLT"
        Me.ShowInTaskbar = False
        Me.Text = "CLT Model"
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numeric_wallareapercent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numeric_ceilareapercent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_integralmodel.ResumeLayout(False)
        Me.Panel_integralmodel.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents optCLTON As System.Windows.Forms.RadioButton
    Friend WithEvents optCLTOFF As System.Windows.Forms.RadioButton
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents numeric_ceilareapercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents numeric_wallareapercent As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCharTemp As System.Windows.Forms.TextBox
    Friend WithEvents txtLamellaDepth As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtCritFlux As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents txtFlameFlux As TextBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents txtCLTigtemp As TextBox
    Friend WithEvents txtCLTLoG As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents txtCLTcalibration As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Panel_integralmodel As Panel
    Friend WithEvents LinkLabel1 As LinkLabel
    Friend WithEvents txtDebondTemp As TextBox
    Friend WithEvents lblDebondTemp As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents RB_Kinetic As RadioButton
    Friend WithEvents RB_dynamic As RadioButton
    Friend WithEvents RB_Integral As RadioButton
    Friend WithEvents GroupBox2 As GroupBox
End Class
