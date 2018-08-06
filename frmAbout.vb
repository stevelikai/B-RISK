Option Strict Off
Option Explicit On

Friend Class frmAbout
	Inherits System.Windows.Forms.Form

	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub
	
	Private Sub frmAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'call procedure to centre form
		Me.Text = "About " & My.Application.Info.Title
		lblVersion.Text = "Version " & Version
		lblTitle.Text = My.Application.Info.Title
        lblDescription.Text = "B-RISK is a room fire simulator with latin hypercube stratified sampling capability. Developed by BRANZ and the University of Canterbury, New Zealand."
        lblDisclaimer.Text = "While this software has been developed in good faith, the Authors do not take any  responsibility for any designs resulting from the use (proper or not) of this model or for any errors that may exist.  The software is intended to support the engineering judgment of the user, and not to be a replacement for such judgment. "
	End Sub
	
End Class