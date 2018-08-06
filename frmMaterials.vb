Imports System.Collections.Generic
Imports System.Data.OleDb

Public Class frmMaterials

    Private Sub Table1BindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Table1BindingNavigatorSaveItem.Click

        'Me.Validate()

        If Table1BindingSource.Count > 0 Then
            'handle ADO.net errors
            Try

                Table1BindingSource.EndEdit()
                TableAdapterManager.UpdateAll(ThermalDataSet)
            Catch ex As ArgumentException
                MessageBox.Show(ex.Message, "Argument Exception")
                Table1BindingSource.CancelEdit()
            Catch ex As DBConcurrencyException
                MessageBox.Show("A concurrency error occurred. " & "Some rows were not updated.", "Concurrency error")
                Table1TableAdapter.Fill(ThermalDataSet.Table1)
            Catch ex As DataException
                MessageBox.Show(ex.Message, ex.GetType.ToString)
                Table1BindingSource.CancelEdit()
            Catch ex As OleDbException
                MessageBox.Show("Database error # " & ex.Message & ": " & ex.GetType.ToString)
            End Try
        Else
            Try
                TableAdapterManager.UpdateAll(ThermalDataSet)
            Catch ex As DBConcurrencyException
                MessageBox.Show("A concurrency error occurred. " & "Some rows were not updated.", "Concurrency error")
                Table1TableAdapter.Fill(ThermalDataSet.Table1)
            Catch ex As DataException
                MessageBox.Show(ex.Message, ex.GetType.ToString)
                Table1BindingSource.CancelEdit()
            Catch ex As OleDbException
                MessageBox.Show("Database error # " & ex.Message & ": " & ex.GetType.ToString)
            End Try
        End If

    End Sub


    Private Sub frmMaterials_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'capture OleDB exception
        Try
            'Dim CString = My.Settings.thermalConnectionString

            'Using cn As New OleDbConnection(connectionString:=CString)
            '    Dim cmd As OleDbCommand = New OleDbCommand
            '    cmd.Connection = cn
            '    cn.Open()
            'End Using

            Table1TableAdapter.Fill(ThermalDataSet.Table1)

        Catch ex As OleDbException
            MessageBox.Show("Database error # " & ex.ErrorCode & ": " & ex.GetType.ToString)
        Catch ex As Exception
            MessageBox.Show("Database error # " & ex.Message & ": " & ex.GetType.ToString)
        End Try

    End Sub

    Private Sub ToolStripButton_CancelItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton_CancelItem.Click
        Table1BindingSource.CancelEdit()
    End Sub


    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub

    Private Sub Table1BindingNavigator_RefreshItems(sender As Object, e As EventArgs) Handles Table1BindingNavigator.RefreshItems

    End Sub
End Class