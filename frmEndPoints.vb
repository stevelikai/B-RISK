Option Strict Off
Option Explicit On
Friend Class frmEndPoints
	Inherits System.Windows.Forms.Form
	
 
	
	Private Sub frmEndPoints_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'*  =============================================
		'*      This event centres the form on the screen
		'*      when the form first loads.
		'*  =============================================
		
		'call procedure to centre form
		Centre_Form(Me)
	End Sub
	

	


    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Hide()
    End Sub
End Class