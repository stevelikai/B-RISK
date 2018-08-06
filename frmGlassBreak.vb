Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Public Class frmGlassBreak
    Inherits System.Windows.Forms.Form
    Dim id, idr, idc, id1, id2 As Integer
    Dim oVents As List(Of oVent)
    Dim oventdistributions As List(Of oDistribution)
    Dim oVent As New oVent


    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        '======================================================================
        ' close the glass property form
        '======================================================================

        'oVent.id = id1
        oVents = VentDB.GetVents
        oventdistributions = VentDB.GetVentDistributions
        'oVents(id1 - 1) = oVent
        oVents(id1 - 1).id = id1
        oVents(id1 - 1).glassflameflux = oVent.glassflameflux
        oVents(id1 - 1).glassdistance = CSng(txtGlassDistance.Text)
        oVents(id1 - 1).glassalpha = CSng(txtGLASSalpha.Text)
        oVents(id1 - 1).glassbreakingstress = CSng(txtGLASSbreakingstress.Text)
        oVents(id1 - 1).glassconductivity = CSng(txtGLASSconductivity.Text)
        'oVents(id1 - 1).glassemissivity = CSng(txtGLASSemissivity.Text)
        oVents(id1 - 1).glassexpansion = CSng(txtGLASSexpansion.Text)
        oVents(id1 - 1).glassfalloutime = CSng(txtGlassFalloutTime.Text)
        oVents(id1 - 1).glassshading = CSng(txtGLASSshading.Text)
        oVents(id1 - 1).glassthickness = CSng(txtGLASSthickness.Text)
        oVents(id1 - 1).glassYoungsModulus = CSng(txtGLASSYoungsModulus.Text)

        If oVents IsNot Nothing Then
            VentDB.SaveVents(oVents, oventdistributions)
            frmVentList.FillVentList()
        End If

        Me.Hide()
    End Sub

    Private Sub frmGlassBreak_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Centre_Form(Me)
    End Sub

    Private Sub optGlassNoFlameFlux_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGlassNoFlameFlux.CheckedChanged

        If eventSender.Checked Then
            If id < 1 Then Exit Sub

            If id > 0 Then
                If optGlassNoFlameFlux.Checked = True Then
                    GlassFlameFlux(idr, idc, id) = False
                    GlassFlameFlux(idc, idr, id) = False
                    oVent.glassflameflux = False
                Else
                    GlassFlameFlux(idr, idc, id) = True
                    GlassFlameFlux(idc, idr, id) = True
                    oVent.glassflameflux = True
                End If
            End If
        End If
    End Sub

    Private Sub optGlassWithFlame_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGlassWithFlame.CheckedChanged
        If eventSender.Checked Then

            If id > 0 Then
                If optGlassNoFlameFlux.Checked = True Then
                    GlassFlameFlux(idr, idc, id) = False
                    GlassFlameFlux(idc, idr, id) = False
                    oVent.glassflameflux = False
                Else
                    GlassFlameFlux(idr, idc, id) = True
                    GlassFlameFlux(idc, idr, id) = True
                    oVent.glassflameflux = True
                End If
            End If

        End If
    End Sub

    Public Sub glass_properties(ByRef oVent, ByRef OVents, ByRef oventdistributions)
        '=========================================================
        '   show a new form for describing the window glass properties
        '=========================================================
        id1 = oVent.id
        idr = oVent.fromroom
        idc = oVent.toroom

        id2 = 0
        For Each oVent2 In OVents
            If oVent2.fromroom = idr And oVent2.toroom = idc Then
                id2 = id2 + 1
                If oVent2.id = id1 Then Exit For
            End If
        Next

        id = id2 'this is the id to use in old ventarrays

        txtGLASSconductivity.Text = oVent.glassconductivity
        txtGLASSexpansion.Text = oVent.glassexpansion
        txtGLASSthickness.Text = oVent.glassthickness
        txtGlassDistance.Text = oVent.glassdistance
        txtGlassFalloutTime.Text = oVent.glassfalloutime
        txtGLASSshading.Text = oVent.glassshading
        txtGLASSbreakingstress.Text = oVent.glassbreakingstress
        txtGLASSalpha.Text = oVent.glassalpha
        txtGLASSYoungsModulus.Text = oVent.glassYoungsModulus

        If oVent.glassflameflux = False Then
            optGlassNoFlameFlux.Checked = True
        Else
            optGlassWithFlame.Checked = True
        End If

        Me.Show()

    End Sub


    Private Sub txtGLASSthickness_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSthickness.Validated

        ErrorProvider1.Clear()
        GLASSthickness(idr, idc, id) = CSng(txtGLASSthickness.Text)
        GLASSthickness(idc, idr, id) = CSng(txtGLASSthickness.Text)
        oVent.glassthickness = CSng(txtGLASSthickness.Text)

    End Sub

    Private Sub txtGLASSthickness_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSthickness.Validating
        If IsNumeric(txtGLASSthickness.Text) Then
            If (CDbl(txtGLASSthickness.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSthickness.Select(0, txtGLASSthickness.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSthickness, "Invalid Entry. Thickness must be greater than 0 mm.")
    End Sub

    Private Sub txtGLASSconductivity_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSconductivity.Validated
        ErrorProvider1.Clear()
        GLASSconductivity(idr, idc, id) = CSng(txtGLASSconductivity.Text)
        GLASSconductivity(idc, idr, id) = CSng(txtGLASSconductivity.Text)
        oVent.glassconductivity = CSng(txtGLASSconductivity.Text)
       
    End Sub

    Private Sub txtGLASSconductivity_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSconductivity.Validating
        If IsNumeric(txtGLASSconductivity.Text) Then
            If (CDbl(txtGLASSconductivity.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSconductivity.Select(0, txtGLASSconductivity.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSconductivity, "Invalid Entry. Thermal conductivity must be greater than 0 W/mK.")
    End Sub

    Private Sub txtGLASSalpha_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSalpha.Validated
       
        ErrorProvider1.Clear()
        GLASSalpha(idr, idc, id) = CSng(txtGLASSalpha.Text)
        GLASSalpha(idc, idr, id) = CSng(txtGLASSalpha.Text)
        oVent.glassalpha = CSng(txtGLASSalpha.Text)

    End Sub

    Private Sub txtGLASSalpha_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSalpha.Validating
        If IsNumeric(txtGLASSalpha.Text) Then
            If (CDbl(txtGLASSalpha.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSalpha.Select(0, txtGLASSalpha.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSalpha, "Invalid Entry. Thermal diffusivity must be greater than 0 m2/s.")
    End Sub

    Private Sub txtGLASSshading_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSshading.Validated
        ErrorProvider1.Clear()
        GLASSshading(idr, idc, id) = CSng(txtGLASSshading.Text)
        GLASSshading(idc, idr, id) = CSng(txtGLASSshading.Text)
        oVent.glassshading = CSng(txtGLASSshading.Text)
     
    End Sub

    Private Sub txtGLASSshading_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSshading.Validating
        If IsNumeric(txtGLASSshading.Text) Then
            If (CDbl(txtGLASSshading.Text) >= 2 * CDbl(txtGLASSthickness.Text)) Then

                'okay
                Exit Sub
            Else
                ' Cancel the event moving off of the control.
                e.Cancel = True

                ' Select the offending text.
                txtGLASSshading.Select(0, txtGLASSshading.Text.Length)

                ' Give the ErrorProvider the error message to display.
                ErrorProvider1.SetError(txtGLASSshading, "The depth of the shaded edge is too small for the glass fracture model. The depth must be at least twice the glass thickness.")
            End If
        End If

    End Sub

    Private Sub txtGLASSexpansion_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSexpansion.Validated
        ErrorProvider1.Clear()
        GLASSexpansion(idr, idc, id) = CSng(txtGLASSexpansion.Text)
        GLASSexpansion(idc, idr, id) = CSng(txtGLASSexpansion.Text)
        oVent.glassexpansion = CSng(txtGLASSexpansion.Text)
        
    End Sub

    Private Sub txtGLASSexpansion_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSexpansion.Validating
        If IsNumeric(txtGLASSexpansion.Text) Then
            If (CDbl(txtGLASSexpansion.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSexpansion.Select(0, txtGLASSexpansion.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSexpansion, "Invalid Entry. Thermal coefficient of expansion must be greater than 0.")
    End Sub

    Private Sub txtGLASSbreakingstress_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSbreakingstress.Validated
        ErrorProvider1.Clear()
        GLASSbreakingstress(idr, idc, id) = CSng(txtGLASSbreakingstress.Text)
        GLASSbreakingstress(idc, idr, id) = CSng(txtGLASSbreakingstress.Text)
        oVent.glassbreakingstress = CSng(txtGLASSbreakingstress.Text)
        
    End Sub

    Private Sub txtGLASSbreakingstress_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSbreakingstress.Validating
        If IsNumeric(txtGLASSbreakingstress.Text) Then
            If (CDbl(txtGLASSbreakingstress.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSbreakingstress.Select(0, txtGLASSbreakingstress.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSbreakingstress, "Invalid Entry. Breaking stress must be greater than 0 MPa.")
    End Sub

    Private Sub txtGLASSYoungsModulus_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGLASSYoungsModulus.Validated
        ErrorProvider1.Clear()
        GlassYoungsModulus(idr, idc, id) = CSng(txtGLASSYoungsModulus.Text)
        GlassYoungsModulus(idc, idr, id) = CSng(txtGLASSYoungsModulus.Text)
        oVent.glassYoungsModulus = CSng(txtGLASSYoungsModulus.Text)

    End Sub

    Private Sub txtGLASSYoungsModulus_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGLASSYoungsModulus.Validating
        If IsNumeric(txtGLASSYoungsModulus.Text) Then
            If (CDbl(txtGLASSYoungsModulus.Text) > 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGLASSYoungsModulus.Select(0, txtGLASSYoungsModulus.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGLASSYoungsModulus, "Invalid Entry. Youngs Modulus must be greater than 0 MPa.")
    End Sub

    Private Sub txtGlassFalloutTime_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGlassFalloutTime.Validated
        ErrorProvider1.Clear()
        GLASSFalloutTime(idr, idc, id) = CSng(txtGlassFalloutTime.Text)
        GLASSFalloutTime(idc, idr, id) = CSng(txtGlassFalloutTime.Text)
        oVent.glassfalloutime = CSng(txtGlassFalloutTime.Text)
          
    End Sub

    Private Sub txtGlassFalloutTime_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGlassFalloutTime.Validating
        If IsNumeric(txtGlassFalloutTime.Text) Then
            If (CDbl(txtGlassFalloutTime.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGlassFalloutTime.Select(0, txtGlassFalloutTime.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGlassFalloutTime, "Invalid Entry. Fall out time must be greater than or equal to 0 sec.")
    End Sub

    Private Sub txtGlassDistance_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGlassDistance.Validated
        ErrorProvider1.Clear()
        GLASSdistance(idr, idc, id) = CSng(txtGlassDistance.Text)
        GLASSdistance(idc, idr, id) = CSng(txtGlassDistance.Text)
        oVent.glassdistance = CSng(txtGlassDistance.Text)
       
    End Sub

    Private Sub txtGlassDistance_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtGlassDistance.Validating
        If IsNumeric(txtGlassDistance.Text) Then
            If (CDbl(txtGlassDistance.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtGlassDistance.Select(0, txtGlassDistance.Text.Length)

        ' Give the ErrorProvider the error message to display.
        ErrorProvider1.SetError(txtGlassDistance, "Invalid Entry. Distance must be greater than or equal to 0 m.")
    End Sub
End Class