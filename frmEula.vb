Public Class frmEula

    Private Sub cmdAgree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAgree.Click
        MDIFrmMain.Enabled = True

        MDIFrmMain.Show()
        frmInputs.Show()
        Me.Close()

    End Sub

    Private Sub cmdDisagree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDisagree.Click
        Me.Close()
        Application.Exit()
    End Sub

    Private Sub frmEula_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      
        'get name of user application data folder
        'Call get_folders()

        'ApplicationPath = IO.Path.GetDirectoryName(Application.ExecutablePath)
        'UserAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        ' CommonFilesFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)

        Me.RichTextBox1.LoadFile(ApplicationPath & "\docs\eula.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
        'Me.RichTextBox1.LoadFile(UserAppDataFolder & gcs_folder_ext & "\docs\eula.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
        Me.Text = "B-RISK End User License Agreement"

    End Sub
End Class