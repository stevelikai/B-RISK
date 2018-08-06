Imports System.Math
Imports System.Collections.Generic

Public Class frmUtiskul
    Public itemid As Integer
    Public oitem As oItem
    Public oItems As List(Of oItem)
    Public oitemdistribution As oDistribution
    Public oitemdistributions As List(Of oDistribution)


    Private Function IsValidData() As Boolean

        Return Validator.IsPresent(txtDensity) AndAlso _
               Validator.IsPresent(txtDiameter) AndAlso _
               Validator.IsPresent(txtFBMLR) AndAlso _
               Validator.IsPresent(txtRamp) AndAlso _
               Validator.IsPresent(txtVolume) 

    End Function
    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        'save new item data
        If IsValidData() Then

            oItems = ItemDB.GetItemsv2()
            oitemdistributions = ItemDB.GetItemDistributions()

            For Each Me.oitem In oItems

                If Me.oitem.id = Me.Tag Then

                    oitem.pooldensity = CDbl(txtDensity.Text)
                    oitem.pooldiameter = CDbl(txtDiameter.Text)
                    oitem.poolFBMLR = CDbl(txtFBMLR.Text)
                    oitem.poolramp = CDbl(txtRamp.Text)
                    oitem.poolvolume = CDbl(txtVolume.Text)
                    oitem.poolvaptemp = CDbl(txtVtemp.Text)

                End If
            Next

            If oitem IsNot Nothing Then
                ItemDB.SaveItemsv2(oItems, oitemdistributions)
            End If

        End If
        Me.Close()
    End Sub



End Class