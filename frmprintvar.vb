Option Strict Off
Option Explicit On
Imports System.IO
Friend Class frmprintvar

    Inherits System.Windows.Forms.Form


    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click

        Me.Hide()

        Dim filepath As String = RiskDataDirectory & "temp_out.txt"

        Call view_output(filepath)
        Dim myfilestream As New FileStream(filepath, FileMode.Open)
        frmViewDocs.RichTextBox1.LoadFile(myfilestream, RichTextBoxStreamType.PlainText)
        myfilestream.Close()
        'Kill(UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_out.txt")
        Kill(RiskDataDirectory & "temp_out.txt")
        frmViewDocs.RichTextBox1.Select(0, 0)
        frmViewDocs.Text = "Summary of Inputs and Results"
        frmViewDocs.Show()
    End Sub

    Private Sub frmprintvar_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Centre_Form(Me)
    End Sub

End Class