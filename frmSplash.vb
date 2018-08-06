Option Strict Off
Option Explicit On

Public NotInheritable Class frmSplash
    Inherits System.Windows.Forms.Form

    Dim ApplicationPath As String

    Private Sub frmSplash_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'Me.Close()
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub frmSplash_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*  =============================================
        '*      This event centres the title form on the screen
        '*      when the form first loads, and displays the
        '*      form or a fixed period of time
        '*  18 March 1999 Colleen Wade
        '*  =============================================

        'change to the application directory
        ChDir(My.Application.Info.DirectoryPath)


        'call procedure to centre form
        Centre_Form(Me)
        Timer1.Enabled = True

        lblVersion.Text = "Version " & Version
        'App.Title = lblProductName.Text
        'MDIFrmMain.Show()

    End Sub

    Private Sub frmSplash_LostFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.LostFocus

        'MDIFrmMain.Show()
        ' Me.Close()
        ' MDIFrmMain.BringToFront()

        'frmEula.Show()
        'frmEula.BringToFront()

    End Sub


    Private Sub frmSplash_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        'MDIFrmMain.Enabled = True
        'frmEula.Show()
        'frmEula.BringToFront()

    End Sub

    Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        '*  =================================================
        '*      This event causes the title screen to display
        '*      for a fixed period of time.
        '*  =================================================

        'stop displaying after set interval


        'frmEula.Show()
        'frmEula.BringToFront()

        MDIFrmMain.Show()
        Me.Close()
        MDIFrmMain.BringToFront()

    End Sub

End Class