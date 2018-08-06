Imports System.Collections.Generic
Imports System.Net
Imports System.IO
Imports System.Xml
Public Class frmwebitemlist
    Dim oItems As List(Of oItem)
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()
    End Sub

    Private Sub frmwebitemlist_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim sep As String = vbTab
        Dim text As String
        'oItems = ItemDB.GetItemsv2()
        'Me.FillItemList()

        text = "ID" & sep & "Description"
        ListBox1.Items.Add(text)

        Me.CenterToParent()

    End Sub
End Class