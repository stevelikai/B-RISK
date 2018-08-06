'Imports System.IO
Public NotInheritable Class VersionReleaseNotes

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub VersionReleaseNotes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Dim StreamToDisplay As StreamReader
        'StreamToDisplay = New StreamReader("Resources\TextFile1.txt")
        'TextBox1.Text = StreamToDisplay.ReadToEnd
        'StreamToDisplay.Close()
        'TextBox1.Select(0, 0)

    End Sub
End Class