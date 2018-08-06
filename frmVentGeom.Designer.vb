<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmVentGeom
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdRotateRoom2 As System.Windows.Forms.Button
	Public WithEvents cmdRotateRoom1 As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblRoom2 As System.Windows.Forms.Label
	Public WithEvents lblRoom1 As System.Windows.Forms.Label
    'Public WithEvents shpVent As Microsoft.VisualBasic.PowerPacks.RectangleShape
    'Public WithEvents _shpWall_1 As Microsoft.VisualBasic.PowerPacks.RectangleShape

    'Public WithEvents _shpWall_0 As Microsoft.VisualBasic.PowerPacks.RectangleShape
    'Public WithEvents shpWall As RectangleShapeArray
    'Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'frmVentGeom
        '
        Me.ClientSize = New System.Drawing.Size(558, 443)
        Me.Name = "frmVentGeom"
        Me.ResumeLayout(False)

    End Sub
    'Friend WithEvents RectangleShape1 As Microsoft.VisualBasic.PowerPacks.RectangleShape
    'Friend WithEvents ShapeContainer3 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
    'Friend WithEvents RectangleShape2 As Microsoft.VisualBasic.PowerPacks.RectangleShape
#End Region 
End Class