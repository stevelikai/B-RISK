Option Strict Off
Option Explicit On
Imports System.Collections.Generic
Imports System.ComponentModel

Public Class frmOptions1
    Inherits System.Windows.Forms.Form

    Private Sub cboABSCoeff_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboABSCoeff.SelectedIndexChanged
        On Error Resume Next

        txtEmissionCoefficient.Enabled = False
        txtStoich.Enabled = False
        txtSootAlpha.Enabled = False
        txtSootEps.Enabled = False
        txtnC.Enabled = False
        txtnH.Enabled = False
        txtnO.Enabled = False
        txtnC.Enabled = False
        txtnN.Enabled = False
        chkHCNcalc.Enabled = True

        Select Case Me.cboABSCoeff.Text
            Case "wood"
                EmissionCoefficient = 0.8
                SootAlpha = 2.5
                SootEpsilon = 1.2
                nC = 0.95
                nH = 2.4
                nO = 1
                nN = 0
                StoichAFratio = 6.1
                'RadiantLossFraction = 0.3
            Case "ethanol"
                EmissionCoefficient = 0.8 'wood
                SootAlpha = 2.5 'wood
                SootEpsilon = 1.2 'wood
                nC = 2
                nH = 6
                nO = 1
                nN = 0
                StoichAFratio = 9
                'RadiantLossFraction = 0.25
            Case "heptane"
                EmissionCoefficient = 1.83 'heptane
                SootAlpha = 2.5 'wood
                SootEpsilon = 1.2 'wood
                nC = 7
                nH = 16
                nO = 0
                nN = 0
                StoichAFratio = 15.1
                'RadiantLossFraction = 0.25
            Case "methane"
                EmissionCoefficient = 6.45
                SootAlpha = 2.5 'wood
                SootEpsilon = 1.2 'wood
                nC = 1
                nH = 4
                nO = 0
                nN = 0
                StoichAFratio = 17.2
                'RadiantLossFraction = 0.14
            Case "plexiglas"
                EmissionCoefficient = 0.5
                SootAlpha = 2.5 'wood
                SootEpsilon = 1.2 'wood
                nC = 5 'pmma
                nH = 8 'pmma
                nO = 2 'pmma
                nN = 0 'pmma
                StoichAFratio = 0 'fix
                'RadiantLossFraction = 0.314
            Case "polystyrene"
                EmissionCoefficient = 1.2
                SootAlpha = 2.8
                SootEpsilon = 1.3
                nC = 8
                nH = 8
                nO = 0
                nN = 0
                StoichAFratio = 0 'fix
                'RadiantLossFraction = 0.56 'gm47
            Case "polyurethane foam flexible"
                EmissionCoefficient = 1.2
                SootAlpha = 2.8
                SootEpsilon = 1.3
                nC = 1
                nH = 1.74
                nO = 0.32
                nN = 0.07
                StoichAFratio = 9.5
                'RadiantLossFraction = 0.52
            Case "propane"
                EmissionCoefficient = 13.32
                SootAlpha = 2.5 'wood
                SootEpsilon = 1.2 'wood
                nC = 3
                nH = 8
                nO = 0
                nN = 0
                StoichAFratio = 0 'fix
                'RadiantLossFraction = 0.29
            Case "VM2"
                EmissionCoefficient = 1.0 'avg wood + PU foam
                SootAlpha = 2.65 'avg wood + PU foam
                SootEpsilon = 1.25 'avg wood + PU foam
                nC = 1
                nH = 2
                nO = 0.5
                nN = 0
                chkHCNcalc.Checked = False
                chkHCNcalc.Enabled = False
                StoichAFratio = 0 'fix
            Case Else
                txtEmissionCoefficient.Enabled = True
                txtSootAlpha.Enabled = True
                txtSootEps.Enabled = True
                txtnC.Enabled = True
                txtnH.Enabled = True
                txtnO.Enabled = True
                txtnC.Enabled = True
                txtnN.Enabled = True
                'txtRadiantLossFraction.Enabled = True

                If IsNumeric(Me.txtEmissionCoefficient.Text) = True Then
                    EmissionCoefficient = CSng(Me.txtEmissionCoefficient.Text)
                Else
                    EmissionCoefficient = 0.8
                End If
                If IsNumeric(Me.txtStoich.Text) = True Then
                    StoichAFratio = CSng(Me.txtStoich.Text)
                Else
                    StoichAFratio = 6.1
                End If
                If IsNumeric(Me.txtSootAlpha.Text) = True Then
                    SootAlpha = CSng(Me.txtSootAlpha.Text)
                Else
                    SootAlpha = 2.5
                End If
                If IsNumeric(Me.txtSootEps.Text) = True Then
                    SootEpsilon = CSng(Me.txtSootEps.Text)
                Else
                    SootEpsilon = 1.2
                End If
                If IsNumeric(Me.txtnC.Text) = True Then
                    nC = CSng(Me.txtnC.Text)
                Else
                    SootEpsilon = 0.95
                End If
                If IsNumeric(Me.txtnH.Text) = True Then
                    nH = CSng(Me.txtnH.Text)
                Else
                    nH = 2.4
                End If
                If IsNumeric(Me.txtnO.Text) = True Then
                    nO = CSng(Me.txtnO.Text)
                Else
                    nO = 1
                End If
                If IsNumeric(Me.txtnN.Text) = True Then
                    nN = CSng(Me.txtnN.Text)
                Else
                    nN = 0
                End If

        End Select
        Me.txtStoich.Text = CStr(StoichAFratio)
        Me.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)
        Me.txtSootAlpha.Text = CStr(SootAlpha)
        Me.txtSootEps.Text = CStr(SootEpsilon)
        Me.txtnC.Text = CStr(nC)
        Me.txtnH.Text = CStr(nH)
        Me.txtnO.Text = CStr(nO)
        Me.txtnN.Text = CStr(nN)

    End Sub


    Private Sub cmdQuintiere_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuintiere.Click
        frmQuintiere.Show()
        If UseOneCurve = True Then
            frmQuintiere.optUseOneCurve.Checked = True
        Else
            frmQuintiere.optUseAllCurves.Checked = True
        End If
        If IgnCorrelation = vbJanssens Then frmQuintiere.optJanssens.Checked = True
        If IgnCorrelation = vbFTP Then frmQuintiere.optFTP.Checked = True

        If PessimiseCombWall = True Then
            frmQuintiere.chkPessimiseWall.Checked = True
        Else
            frmQuintiere.chkPessimiseWall.Checked = False
        End If

        frmQuintiere.txtHFSlimit.Text = wallHFSlimit
        frmQuintiere.txtVFSlimit.Text = wallVFSlimit
        frmQuintiere.NumericCeilingPercent.Value = ceilingpercent
        frmQuintiere.NumericWallPercent.Value = wallpercent

        'show the wall partial coverage for wall inputs 
        If DevKey = True Then
            frmQuintiere.GroupBox1.Visible = True
        Else
            frmQuintiere.GroupBox1.Visible = False
        End If

    End Sub



    Private Sub lstRoomZone_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRoomZone.SelectedIndexChanged

        Dim id, i As Short

        'identify current room selected
        id = lstRoomZone.SelectedIndex + 1
        On Error Resume Next
        If id <= NumberRooms Then
            Me.optOneLayer.Checked = Not TwoZones(id)
            Me.optTwoLayer.Checked = TwoZones(id)
        End If

    End Sub

    Private Sub lstTimeStep_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstTimeStep.SelectedIndexChanged
        Timestep = CDbl(Me.lstTimeStep.Text)
    End Sub

    Private Sub optEnhanceOff_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optEnhanceOff.CheckedChanged
        If eventSender.Checked Then
            If optEnhanceOff.Checked = True Then
                txtLHoG.Enabled = False
                lblLHoG.Enabled = False
                txtFuelSurfaceArea.Enabled = False
                lblFuelSurfaceArea.Enabled = False

            Else
                txtLHoG.Enabled = True
                lblLHoG.Enabled = True
                txtFuelSurfaceArea.Enabled = True
                lblFuelSurfaceArea.Enabled = True

            End If

        End If
    End Sub

    Private Sub optEnhanceOn_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optEnhanceOn.CheckedChanged
        If eventSender.Checked Then
            If optEnhanceOn.Checked = True Then
                txtLHoG.Enabled = True
                lblLHoG.Enabled = True
                txtFuelSurfaceArea.Enabled = True
                lblFuelSurfaceArea.Enabled = True

            Else
                txtLHoG.Enabled = False
                lblLHoG.Enabled = False
                txtFuelSurfaceArea.Enabled = False
                lblFuelSurfaceArea.Enabled = False

            End If
        End If
    End Sub



    Private Sub optOneLayer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optOneLayer.CheckedChanged
        If eventSender.Checked Then

            Dim id As Short

            'identify current room selected
            id = lstRoomZone.SelectedIndex + 1
            If id <= NumberRooms Then
                TwoZones(id) = Not optOneLayer.Checked
                If id = fireroom Then
                    TwoZones(id) = True
                    Me.optTwoLayer.Checked = True
                    MsgBox("The fire room cannot be modelled as a single zone.")
                End If
            End If
        End If
    End Sub

    Private Sub optRCNone_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRCNone.CheckedChanged
        If eventSender.Checked Then
            If optQuintiere.Checked = True Then
                cmdQuintiere.Visible = True
            Else
                cmdQuintiere.Visible = False
            End If
        End If
    End Sub



    Private Sub optTwoLayer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optTwoLayer.CheckedChanged
        If eventSender.Checked Then

            Dim id As Short
            ReDim Preserve TwoZones(0 To NumberRooms)

            'identify current room selected
            id = lstRoomZone.SelectedIndex + 1

            TwoZones(id) = optTwoLayer.Checked

        End If
    End Sub

    Private Sub txtEmissionCoefficient_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmissionCoefficient.TextChanged
        'store the effective emission coefficient as a variable
        If IsNumeric(txtEmissionCoefficient.Text) = True Then
            EmissionCoefficient = CSng(txtEmissionCoefficient.Text)
        End If
    End Sub

    Private Sub txtEmissionCoefficient_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEmissionCoefficient.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtEmissionCoefficient.Text) = True Then
            A = CSng(txtEmissionCoefficient.Text)
            If KeyAscii = 13 Then
                If A >= 0 Then
                    EmissionCoefficient = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtEmissionCoefficient_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmissionCoefficient.Leave
        Dim A As Single
        If IsNumeric(txtEmissionCoefficient.Text) = True Then
            A = CSng(txtEmissionCoefficient.Text)
            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtEmissionCoefficient.Focus()
            End If
        End If
    End Sub

    Private Sub txtErrorControl_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtErrorControl.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Double
        If IsNumeric(txtErrorControl.Text) = True Then
            A = CDbl(txtErrorControl.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    Error_Control = A
                Else
                    txtErrorControl.Focus()
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtErrorControl_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtErrorControl.Leave

        Dim A As Double
        If IsNumeric(txtErrorControl.Text) = True Then
            A = CDbl(txtErrorControl.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtErrorControl.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtErrorControl.Focus()
        End If
    End Sub


    Private Sub txtFuelSurfaceArea_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelSurfaceArea.TextChanged
        'store the radiant loss fraction as a fraction
        If IsNumeric(txtFuelSurfaceArea.Text) = True Then
            FuelSurfaceArea = CSng(txtFuelSurfaceArea.Text)
        End If

    End Sub

    Private Sub txtFuelSurfaceArea_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFuelSurfaceArea.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtFuelSurfaceArea.Text) = True Then
            A = CSng(txtFuelSurfaceArea.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    FuelSurfaceArea = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFuelSurfaceArea_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelSurfaceArea.Leave
        Dim A As Single
        If IsNumeric(txtFuelSurfaceArea.Text) = True Then
            A = CSng(txtFuelSurfaceArea.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid Value!")
                txtFuelSurfaceArea.Focus()
            End If
        Else
            MsgBox("Invalid Value!")
            txtFuelSurfaceArea.Focus()
        End If
    End Sub



    Private Sub txtFuelThickness_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFuelThickness.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtFuelThickness.Text) = True Then
            A = CDbl(txtFuelThickness.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    Fuel_Thickness = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFuelThickness_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelThickness.Leave

        Dim A As Single
        If IsNumeric(txtFuelThickness.Text) = True Then
            A = CSng(txtFuelThickness.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid Value!")
                txtFuelThickness.Focus()
            End If
        Else
            MsgBox("Invalid Value!")
            txtFuelThickness.Focus()
        End If

    End Sub


    Private Sub txtJobNumber_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtJobNumber.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Description = txtJobNumber.Text
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtnC_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnC.TextChanged

        If IsNumeric(txtnC.Text) = True Then
            nC = CSng(txtnC.Text)
        End If

    End Sub

    Private Sub txtnC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtnC.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtnC.Text) = True Then
            A = CSng(txtnC.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    nC = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtnC_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnC.Leave

        Dim A As Single
        If IsNumeric(txtnC.Text) = True Then
            A = CSng(txtnC.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid entry")
                txtnC.Focus()
            End If
        Else
            MsgBox("Invalid entry")
            txtnC.Focus()
        End If

    End Sub

    Private Sub txtnH_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnH.TextChanged

        'store the radiant loss fraction as a fraction
        If IsNumeric(txtnH.Text) = True Then
            nH = CSng(txtnH.Text)
        End If

    End Sub

    Private Sub txtnH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtnH.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtnH.Text) = True Then
            A = CSng(txtnH.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    nH = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtnH_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnH.Leave

        Dim A As Single
        If IsNumeric(txtnH.Text) = True Then
            A = CSng(txtnH.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid entry")
                txtnH.Focus()
            End If
        Else
            MsgBox("Invalid entry")
            txtnH.Focus()
        End If

    End Sub

    Private Sub txtnN_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnN.TextChanged

        'store the radiant loss fraction as a fraction
        If IsNumeric(txtnN.Text) = True Then
            nN = CSng(txtnN.Text)
        End If

    End Sub

    Private Sub txtnN_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtnN.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtnN.Text) = True Then
            A = CSng(txtnN.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    nN = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtnN_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnN.Leave

        Dim A As Single
        If IsNumeric(txtnN.Text) = True Then
            A = CSng(txtnN.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid entry")
                txtnN.Focus()
            End If
        Else
            MsgBox("Invalid entry")
            txtnN.Focus()
        End If

    End Sub

    Private Sub txtnO_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnO.TextChanged

        'store the radiant loss fraction as a fraction
        If IsNumeric(txtnO.Text) = True Then
            nO = CSng(txtnO.Text)
        End If

    End Sub

    Private Sub txtnO_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtnO.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtnO.Text) = True Then
            A = CSng(txtnO.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    nO = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtnO_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtnO.Leave

        Dim A As Single
        If IsNumeric(txtnO.Text) = True Then
            A = CSng(txtnO.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid entry")
                txtnO.Focus()
            End If
        Else
            MsgBox("Invalid entry")
            txtnO.Focus()
        End If

    End Sub

    Private Sub txtNodes_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNodes.TextChanged
        'Dim index As Short = txtNodes.GetIndex(eventSender)

        ''store the burner as a variable
        'If IsNumeric(txtNodes(index).Text) = True Then
        '    Select Case index
        '        Case 0
        '            Ceilingnodes = CShort(txtNodes(index).Text)
        '        Case 1
        '            Wallnodes = CShort(txtNodes(index).Text)
        '        Case 2
        '            Floornodes = CShort(txtNodes(index).Text)
        '    End Select
        'End If

    End Sub

    Private Sub txtNodes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNodes.KeyPress
        'Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'Dim index As Short = txtNodes.GetIndex(eventSender)

        'Dim A As Short

        'If IsNumeric(txtNodes(index).Text) = True Then
        '    A = CShort(txtNodes(index).Text)
        '    If KeyAscii = 13 Then
        '        If A >= 10 Then
        '            Select Case index
        '                Case 0
        '                    Ceilingnodes = A
        '                Case 1
        '                    Wallnodes = CShort(txtNodes(index).Text)
        '                Case 2
        '                    Floornodes = CShort(txtNodes(index).Text)
        '            End Select
        '        End If
        '    End If
        'End If
        'eventArgs.KeyChar = Chr(KeyAscii)
        'If KeyAscii = 0 Then
        '    eventArgs.Handled = True
        'End If
    End Sub

    Private Sub txtNodes_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNodes.Leave
        'Dim index As Short = txtNodes.GetIndex(eventSender)

        'Dim A As Short

        'If IsNumeric(txtNodes(index).Text) = True Then
        '    A = CShort(txtNodes(index).Text)
        '    If A < 10 Then
        '        MsgBox("Invalid Dimension! Number of Nodes must be not less than 10.")
        '        txtNodes(index).Focus()
        '    End If
        'End If

    End Sub

    Private Sub txtpostCO_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtpostCO.TextChanged
        If IsNumeric(txtpostCO.Text) = True Then
            postCO = CSng(txtpostCO.Text)
        End If
    End Sub

    Private Sub txtpostCO_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtpostCO.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim A As Single
        If IsNumeric(txtpostCO.Text) = True Then
            A = CSng(txtpostCO.Text)
            If KeyAscii = 13 Then
                postCO = A
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtpostCO_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtpostCO.Leave
        Dim A As Single
        If IsNumeric(txtpostCO.Text) = True Then
            A = CSng(txtpostCO.Text)
            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtpostCO.Focus()
            End If
        End If

    End Sub

    Private Sub txtPostSoot_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPostSoot.TextChanged
        If IsNumeric(txtPostSoot.Text) = True Then
            postSoot = CSng(txtPostSoot.Text)
        End If
    End Sub

    Private Sub txtPostSoot_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPostSoot.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim A As Single
        If IsNumeric(txtPostSoot.Text) = True Then
            A = CSng(txtPostSoot.Text)
            If KeyAscii = 13 Then
                postSoot = A
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtPostSoot_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPostSoot.Leave
        Dim A As Single
        If IsNumeric(txtPostSoot.Text) = True Then
            A = CSng(txtPostSoot.Text)
            If A >= 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtPostSoot.Focus()
            End If
        End If

    End Sub



    Private Sub optQuintiere_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optQuintiere.CheckedChanged
        If eventSender.Checked Then

            'MDIFrmMain.mnuRun.Enabled = True
            'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.Toolbar2.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            'MDIFrmMain.Toolbar2.Items.Item(8).Visible = True 'start run
            If optQuintiere.Checked = True Then
                cmdQuintiere.Visible = True
            Else
                cmdQuintiere.Visible = False
            End If

        End If
    End Sub



    Private Sub txtBurnerWidth_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBurnerWidth.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtBurnerWidth.Text) = True Then
            A = CSng(txtBurnerWidth.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    BurnerWidth = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtBurnerWidth_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBurnerWidth.Leave

        Dim A As Single
        If IsNumeric(txtBurnerWidth.Text) = True Then
            A = CSng(txtBurnerWidth.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtBurnerWidth.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtBurnerWidth.Focus()
        End If
    End Sub

    Private Sub txtCeilingHeatFlux_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCeilingHeatFlux.TextChanged

        'store the burner as a variable
        If IsNumeric(txtCeilingHeatFlux.Text) = True Then
            CeilingHeatFlux = CSng(txtCeilingHeatFlux.Text)
        End If
    End Sub

    Private Sub txtCeilingHeatFlux_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCeilingHeatFlux.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtCeilingHeatFlux.Text) = True Then
            A = CSng(txtCeilingHeatFlux.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    CeilingHeatFlux = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtCeilingHeatFlux_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCeilingHeatFlux.Leave

        Dim A As Single
        If IsNumeric(txtCeilingHeatFlux.Text) = True Then
            A = CSng(txtCeilingHeatFlux.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtCeilingHeatFlux.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtCeilingHeatFlux.Focus()
        End If
    End Sub



    Private Sub txtFlameAreaConstant_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlameAreaConstant.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtFlameAreaConstant.Text) = True Then
            A = CSng(txtFlameAreaConstant.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    FlameAreaConstant = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlameAreaConstant_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlameAreaConstant.Leave

        Dim A As Single
        If IsNumeric(txtFlameAreaConstant.Text) = True Then
            A = CSng(txtFlameAreaConstant.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtFlameAreaConstant.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtFlameAreaConstant.Focus()
        End If
    End Sub



    Private Sub txtFlameLengthPower_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFlameLengthPower.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Double
        If IsNumeric(txtFlameLengthPower.Text) = True Then
            A = CDbl(txtFlameLengthPower.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    FlameLengthPower = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFlameLengthPower_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFlameLengthPower.Leave

        Dim A As Double
        If IsNumeric(txtFlameLengthPower.Text) = True Then
            A = CDbl(txtFlameLengthPower.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtFlameLengthPower.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtFlameLengthPower.Focus()
        End If
    End Sub



    Private Sub txtSootAlpha_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootAlpha.TextChanged
        'store the effective emission coefficient as a variable
        If IsNumeric(txtSootAlpha.Text) = True Then
            SootAlpha = CSng(txtSootAlpha.Text)
        End If

    End Sub

    Private Sub txtSootAlpha_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSootAlpha.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtSootAlpha.Text) = True Then
            A = CSng(txtSootAlpha.Text)

            If KeyAscii = 13 Then
                If A > 0 Then
                    SootAlpha = A
                End If
            End If
        End If


        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSootAlpha_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootAlpha.Leave

        Dim A As Single
        If IsNumeric(txtSootAlpha.Text) = True Then
            A = CSng(txtSootAlpha.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Value")
                txtSootAlpha.Focus()
            End If
        Else
            MsgBox("Invalid Value")
            txtSootAlpha.Focus()
        End If

    End Sub

    Private Sub txtSootEps_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootEps.TextChanged
        'store the radiant loss fraction as a fraction
        If IsNumeric(txtSootEps.Text) = True Then
            SootEpsilon = CSng(txtSootEps.Text)
        End If

    End Sub

    Private Sub txtSootEps_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSootEps.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtSootEps.Text) = True Then
            A = CSng(txtSootEps.Text)

            If KeyAscii = 13 Then
                If A > 0 Then
                    SootEpsilon = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtSootEps_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSootEps.Leave

        Dim A As Single
        If IsNumeric(txtSootEps.Text) = True Then
            A = CSng(txtSootEps.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Value")
                txtSootEps.Focus()
            End If
        Else
            MsgBox("Invalid Value")
            txtSootEps.Focus()
        End If

    End Sub


    Private Sub txtStickSpacing_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtStickSpacing.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtStickSpacing.Text) = True Then
            A = CDbl(txtStickSpacing.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    Stick_Spacing = A
                End If
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtStickSpacing_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtStickSpacing.Leave

        Dim A As Single
        If IsNumeric(txtStickSpacing.Text) = True Then
            A = CSng(txtStickSpacing.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid Value!")
                txtStickSpacing.Focus()
            End If
        Else
            MsgBox("Invalid Value!")
            txtStickSpacing.Focus()
        End If

    End Sub

    Private Sub txtWallHeatFlux_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWallHeatFlux.TextChanged
        'store the burner as a variable
        If IsNumeric(txtWallHeatFlux.Text) = True Then
            WallHeatFlux = CSng(txtWallHeatFlux.Text)
        End If
    End Sub

    Private Sub txtWallHeatFlux_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWallHeatFlux.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim A As Single
        If IsNumeric(txtWallHeatFlux.Text) = True Then
            A = CSng(txtWallHeatFlux.Text)
        End If
        If KeyAscii = 13 Then
            If A > 0 Then
                WallHeatFlux = A
            End If
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtWallHeatFlux_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWallHeatFlux.Leave

        Dim A As Single
        If IsNumeric(txtWallHeatFlux.Text) = True Then
            A = CSng(txtWallHeatFlux.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtWallHeatFlux.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtWallHeatFlux.Focus()
        End If
    End Sub



    Private Sub txtConvect_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConvect.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtConvect.Text) = True Then
            A = CSng(txtConvect.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    ConvectEndPoint = A + 273
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConvect_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConvect.Leave

        Dim A As Single
        If IsNumeric(txtConvect.Text) = True Then
            A = CSng(txtConvect.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtConvect.Focus()
            End If
        End If
    End Sub


    Private Sub txtFED_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFED.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtFED.Text) = True Then
            A = CSng(txtFED.Text)
            If KeyAscii = 13 Then
                If A > 0 And A <= 1 Then
                    FEDEndPoint = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtFED_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFED.Leave

        Dim A As Single
        If IsNumeric(txtFED.Text) = True Then
            A = CSng(txtFED.Text)

            If A > 0 And A <= 1 Then
            Else
                MsgBox("Invalid Entry!")
                txtFED.Focus()
            End If
        End If
    End Sub

    Private Sub txtMonitorHeight_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtMonitorHeight.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single

        If IsNumeric(txtMonitorHeight.Text) = True Then
            A = CSng(txtMonitorHeight.Text)

            If KeyAscii = 13 Then
                If A >= 0 Then
                    MonitorHeight = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtMonitorHeight_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtMonitorHeight.Leave

        Dim A As Single
        If IsNumeric(txtMonitorHeight.Text) = True Then
            A = CSng(txtMonitorHeight.Text)

            If A >= 0 Then
            Else
                MsgBox("Invalid Dimension!")
                txtMonitorHeight.Focus()
            End If
        End If
    End Sub

    Private Sub txtTarget_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTarget.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim A As Single
        If IsNumeric(txtTarget.Text) = True Then
            A = CSng(txtTarget.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    TargetEndPoint = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTarget_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTarget.Leave

        Dim A As Single
        If IsNumeric(txtTarget.Text) = True Then
            A = CSng(txtTarget.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtTarget.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtTarget.Focus()
        End If
    End Sub


    Private Sub txtTemp_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTemp.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtTemp.Text) = True Then
            A = CSng(txtTemp.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    TempEndPoint = A + 273
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtTemp_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTemp.Leave

        Dim A As Single
        If IsNumeric(txtTemp.Text) = True Then
            A = CSng(txtTemp.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtTemp.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtTemp.Focus()
        End If
    End Sub

    Private Sub txtVisibility_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVisibility.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        Dim A As Single
        If IsNumeric(txtVisibility.Text) = True Then
            A = CSng(txtVisibility.Text)
            If KeyAscii = 13 Then
                If A > 0 Then
                    VisibilityEndPoint = A
                End If
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtVisibility_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVisibility.Leave

        Dim A As Single
        If IsNumeric(txtVisibility.Text) = True Then
            A = CSng(txtVisibility.Text)

            If A > 0 Then
            Else
                MsgBox("Invalid Entry!")
                txtVisibility.Focus()
            End If
        Else
            MsgBox("Invalid Entry!")
            txtVisibility.Focus()
        End If
    End Sub


    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        Dim room, i As Integer
        Dim Visi_lower, Visi_upper, OpticalDensity As Single
        Dim mw_upper, mw_lower As Double
        On Error GoTo errorhandler

        If optCOman.Checked = True Then
            comode = True
        Else
            comode = False
        End If
        If optSootman.Checked = True Then
            sootmode = True
        Else
            sootmode = False
        End If

        fueltype = Me.cboABSCoeff.Text

        If chkModGER.Checked = True Then
            modGER = True
        Else
            modGER = False
        End If

        If ChkFRR.Checked = True Then
            calcFRR = True
        Else
            calcFRR = False
        End If

        If optFOtemp.Checked = True Then
            FOFluxCriteria = False
        Else
            FOFluxCriteria = True
        End If

        If optPostFlashover.Checked = True Then
            g_post = True
        Else
            g_post = False
        End If
        If optEnhanceOff.Checked = True Then
            Enhance = False
        Else
            Enhance = True
        End If

        If optQuintiere.Checked = True Then
            corner = 2
        Else
            corner = 0
        End If

        If optLightWork.Checked = True Then
            Activity = "Light"
        ElseIf optHeavyWork.Checked = True Then
            Activity = "Heavy"
        Else
            Activity = "Rest"
        End If

        If optIlluminatedSign.Checked = True Then
            illumination = True
        Else
            illumination = False
        End If

        If optFEDCO.Checked = True Then
            FEDCO = True
        Else
            FEDCO = False
        End If

        If optJET.Checked = True Then
            cjModel = cjJET
        Else
            cjModel = cjAlpert
        End If

        If Me.optRCNone.Checked = True Then
            RCNone = True
            MDIFrmMain.mnuQuintiereGraph.Visible = False
            MDIFrmMain.mnuConeHRR.Visible = False
        Else
            RCNone = False
            MDIFrmMain.mnuQuintiereGraph.Visible = True
            MDIFrmMain.mnuConeHRR.Visible = True
        End If

        If frmQuintiere.chkSpreadAdjacentRoom.CheckState = System.Windows.Forms.CheckState.Checked Then
            IgniteNextRoom = True
        Else
            IgniteNextRoom = False
        End If

        If Me.chkWallFlowDisable.CheckState = 1 Then
            nowallflow = True
        Else
            nowallflow = False
        End If

        If optGaussJor.Checked = True Then
            LEsolver = "Gauss-Jordan"
        Else
            LEsolver = "LU Decomposition"
        End If


        '....from here
        'ReDim Visibility(NumberRooms, NumberTimeSteps + 1)
        'ReDim OD_upper(NumberRooms, NumberTimeSteps + 1)
        ''ReDim SDOpticalDensity(1 To NumberRooms, 1 To NumberTimeSteps + 1)
        'ReDim OD_lower(NumberRooms, NumberTimeSteps + 1)

        'ReDim FEDSum(NumberRooms, NumberTimeSteps + 2)
        'ReDim FEDRadSum(NumberRooms, NumberTimeSteps + 1)
        'ReDim SurfaceRad(NumberRooms, NumberTimeSteps + 2)

        'If NumberTimeSteps > 0 Then
        '    If IsNumeric(COMassFraction) = True Then
        '        For room = 1 To NumberRooms
        '            'If FEDCO = True Then
        '            '    'Call FED_CO_13571(room)
        '            'Else
        '            '    Call FED_Toxicity(room)
        '            'End If
        '            'Call FED_Toxicity(room)
        '            For i = 1 To NumberTimeSteps + 1
        '                'calulate the visibility
        '                On Error Resume Next
        '                'effective molecular weight of the layers
        '                mw_upper = MolecularWeightCO * COMassFraction(room, i, 1) + MolecularWeightCO2 * CO2MassFraction(room, i, 1) + MolecularWeightH2O * H2OMassFraction(room, i, 1) + MolecularWeightHCN * HCNMassFraction(room, i, 1) + MolecularWeightO2 * O2MassFraction(room, i, 1) + MolecularWeightN2 * (1 - O2MassFraction(room, i, 1) - COMassFraction(room, i, 1) - CO2MassFraction(room, i, 1) - H2OMassFraction(room, i, 1) - HCNMassFraction(room, i, 1))
        '                mw_lower = MolecularWeightCO * COMassFraction(room, i, 2) + MolecularWeightCO2 * CO2MassFraction(room, i, 2) + MolecularWeightH2O * H2OMassFraction(room, i, 2) + MolecularWeightHCN * HCNMassFraction(room, i, 2) + MolecularWeightO2 * O2MassFraction(room, i, 2) + MolecularWeightN2 * (1 - O2MassFraction(room, i, 2) - COMassFraction(room, i, 2) - CO2MassFraction(room, i, 2) - H2OMassFraction(room, i, 2) - HCNMassFraction(room, i, 2))

        '                Visi_upper = Get_Visibility(mw_upper, RoomPressure(room, i), uppertemp(room, i), SootMassFraction(room, i, 1), FuelMassLossRate(i, room), SootMass_Rate(tim(i, 1), i), OpticalDensity)
        '                OD_upper(room, i) = OpticalDensity
        '                SPR(room, i) = 2.3 * OpticalDensity * UFlowToOutside(room, i)
        '                Visi_lower = Get_Visibility(mw_lower, RoomPressure(room, i), lowertemp(room, i), SootMassFraction(room, i, 2), FuelMassLossRate(i, room), SootMass_Rate(tim(i, 1), i), OpticalDensity)
        '                OD_lower(room, i) = OpticalDensity
        '                If layerheight(room, i) <= MonitorHeight Then
        '                    Visibility(room, i) = Visi_upper
        '                Else
        '                    Visibility(room, i) = Visi_lower
        '                End If
        '                System.Windows.Forms.Application.DoEvents()
        '            Next i
        '            'Call FED_thermal_iso13571(room)
        '        Next room

        '    End If

        '    If FEDCO = True Then
        '        Call FED_CO_iso13571_multi()
        '    Else
        '        Call FED_gases_multi()
        '    End If

        '    Call FED_thermal_iso13571_multi()
        'End If
        'to here....

        'Me.Close()
        MDIFrmMain.TopMost = True
        Me.Hide()
        MDIFrmMain.TopMost = False
        Exit Sub

errorhandler:
        Exit Sub

    End Sub

    Private Sub frmOptions1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i, room As Short
        On Error Resume Next
        Me.AddOwnedForm(frmQuintiere)
        Me.AddOwnedForm(frmCLT)


        lstTimeStep.Text = CStr(Timestep)
        Me.txtErrorControl.Text = CStr(Error_Control)
        Me.txtErrorControlVentFlow.Text = CStr(Error_Control_ventflow)

        Me.lstRoomZone.Items.Clear()
        For room = 1 To NumberRooms
            Me.lstRoomZone.Items.Add(CStr(room))
        Next room

        Me.lstRoomZone.SelectedIndex = 0
        If TwoZones(1) = True Then
            Me.optTwoLayer.Checked = True
        Else
            Me.optOneLayer.Checked = True
        End If

        If modGER = True Then
            chkModGER.Checked = True
        Else
            chkModGER.Checked = False
        End If

        If cjModel = cjAlpert Then
            Me.optAlpert.Checked = True
        Else
            Me.optJET.Checked = True
        End If

        If calcFRR = True Then
            ChkFRR.Checked = True
        Else
            ChkFRR.Checked = False
        End If

        If FEDCO = True Then
            Me.optFEDCO.Checked = True
        Else
            Me.optFEDGeneral.Checked = True
        End If

        If VM2 = True Then
            Activity = "Light"
            Me.optLightWork.Checked = True
            Me.optLightWork.Enabled = False
            Me.optAtRest.Enabled = False
            Me.optHeavyWork.Enabled = False
            Me.optFEDCO.Enabled = False
            Me.optFEDGeneral.Enabled = False
            Me.optQuintiere.Enabled = False
            Me.optRCNone.Enabled = False
            Me.Frame13.Visible = False
        Else
            Me.optAtRest.Enabled = True
            Me.optLightWork.Enabled = True
            Me.optHeavyWork.Enabled = True
            Me.optFEDCO.Enabled = True
            Me.optFEDGeneral.Enabled = True
            Me.optQuintiere.Enabled = True
            Me.optRCNone.Enabled = True
            Me.Frame13.Visible = True
        End If

        Select Case Activity
            Case "Rest"
                Me.optAtRest.Checked = True
            Case "Light"
                Me.optLightWork.Checked = True
            Case "Heavy"
                Me.optHeavyWork.Checked = True
        End Select

        If illumination = True Then
            Me.optIlluminatedSign.Checked = True
        Else
            Me.optReflectiveSign.Checked = True
        End If

        If corner = 2 Then
            Me.optQuintiere.Checked = True
            MDIFrmMain.mnuQuintiereGraph.Visible = True
            MDIFrmMain.mnuConeHRR.Visible = True
            cmdQuintiere.Visible = True
        Else
            Me.optRCNone.Checked = True
            MDIFrmMain.mnuQuintiereGraph.Visible = False
            MDIFrmMain.mnuConeHRR.Visible = False
            cmdQuintiere.Visible = False
        End If

        If LEsolver = "" Or LEsolver = "LU Decomposition" Then
            Me.optLUdecom.Checked = True
        Else
            Me.optGaussJor.Checked = True
        End If
        Me._txtNodes_0.Text = Ceilingnodes
        Me._txtNodes_1.Text = Wallnodes
        Me._txtNodes_2.Text = Floornodes

        If Enhance = True Then
            Me.optEnhanceOn.Checked = True
            txtLHoG.Enabled = True
            lblLHoG.Enabled = True
            txtFuelSurfaceArea.Enabled = True
            lblFuelSurfaceArea.Enabled = True
        Else
            Me.optEnhanceOff.Checked = True
            txtLHoG.Enabled = False
            lblLHoG.Enabled = False
            txtFuelSurfaceArea.Enabled = False
            lblFuelSurfaceArea.Enabled = False
        End If

        Me.txtStoich.Text = CStr(StoichAFratio)
        Me.txtnC.Text = CStr(nC)
        Me.txtnH.Text = CStr(nH)
        Me.txtnO.Text = CStr(nO)
        Me.txtnN.Text = CStr(nN)
        Me.txtpreCO.Text = CStr(preCO)
        Me.txtpostCO.Text = CStr(postCO)
        Me.txtpreSoot.Text = CStr(preSoot)
        Me.txtPostSoot.Text = CStr(postSoot)
        If sootmode = True Then
            Me.optSootman.Checked = True
        Else
            Me.optSootauto.Checked = True
        End If
        If comode = True Then
            Me.optCOman.Checked = True
        Else
            Me.optCOauto.Checked = True
        End If
        If FOFluxCriteria = True Then
            Me.optFOflux.Checked = True
        Else
            Me.optFOtemp.Checked = True
        End If
        If g_post = True Then
            Me.optPostFlashover.Checked = True
        Else
            Me.OptPreFlashover.Checked = True
        End If
        Me.txtHOCFuel.Text = HoC_fuel
        Me.cboABSCoeff.Text = CStr(fueltype)
        Me.txtEmissionCoefficient.Text = EmissionCoefficient
        Me.txtDescription.Text = Description
        Me.txtJobNumber.Text = JobNumber
        Me.txtInteriorTemp.Text = InteriorTemp - 273
        Me.txtExteriorTemp.Text = ExteriorTemp - 273
        Me.txtRelativeHumidity.Text = RelativeHumidity * 100
        Me.txtFED.Text = FEDEndPoint
        Me.txtTarget.Text = TargetEndPoint
        Me.txtVisibility.Text = VisibilityEndPoint
        Me.txtTemp.Text = TempEndPoint - 273
        Me.txtConvect.Text = ConvectEndPoint - 273
        Me.txtMonitorHeight.Text = MonitorHeight
        Me.txtBurnerWidth.Text = BurnerWidth
        Me.txtFlameAreaConstant.Text = FlameAreaConstant
        Me.txtFlameLengthPower.Text = FlameLengthPower
        Me.txtWallHeatFlux.Text = WallHeatFlux
        Me.txtCeilingHeatFlux.Text = CeilingHeatFlux
        Me.BringToFront()
        Me.Show()

    End Sub

    Private Sub txtDescription_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescription.TextChanged


    End Sub

    Private Sub txtDescription_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDescription.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)

        If KeyAscii = 13 Then
            Description = txtDescription.Text
            MDIFrmMain.ToolStripStatusLabel2.Text = Description
        End If

        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub



    Public Sub Update_Distributions(ByVal param As String)
        Dim odistributions As List(Of oDistribution)
        odistributions = DistributionClass.GetDistributions()

        For Each oDistribution In odistributions
            If oDistribution.varname = param Then
                Select Case param
                    Case "Interior Temperature"
                        oDistribution.varvalue = InteriorTemp
                        Exit For
                    Case "Exterior Temperature"
                        oDistribution.varvalue = ExteriorTemp
                        Exit For
                    Case "Relative Humidity"
                        oDistribution.varvalue = RelativeHumidity
                        Exit For
                    Case "Heat of Combustion PFO"
                        oDistribution.varvalue = HoC_fuel
                        Exit For
                    Case "Fire Load Energy Density"
                        oDistribution.varvalue = FLED
                        Exit For
                    Case "Soot Preflashover Yield"
                        oDistribution.varvalue = preSoot
                        Exit For
                    Case "CO Preflashover Yield"
                        oDistribution.varvalue = preCO
                        Exit For
                        'Case "Fuel Heat of Gasification"
                        '    oDistribution.varvalue = CDec(FuelHeatofGasification)
                        '    Exit For
                    Case "Sprinkler Reliability"
                        oDistribution.varvalue = SprReliability
                        Exit For
                    Case "Sprinkler Cooling Coefficient"
                        oDistribution.varvalue = SprCooling
                        Exit For
                    Case "Sprinkler Suppression Probability"
                        oDistribution.varvalue = SprSuppressionProb
                        Exit For
                    Case "Alpha T"
                        oDistribution.varvalue = AlphaT
                        Exit For
                    Case "Peak HRR"
                        oDistribution.varvalue = PeakHRR
                        Exit For
                End Select
            End If
        Next

        DistributionClass.SaveDistributions(odistributions)

    End Sub


    Private Sub cmdDist_interiortemp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_interiortemp.Click

        Dim param As String = "Interior Temperature"
        Dim units As String = "C"
        Dim instruction As String = "Interior ambient temperature in all rooms"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub cmdDist_exteriortemp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_exteriortemp.Click

        Dim param As String = "Exterior Temperature"
        Dim units As String = "C"
        Dim instruction As String = "Exterior ambient temperature in all rooms"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub cmdDist_RH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_RH.Click

        Dim param As String = "Relative Humidity"
        Dim units As String = "%"
        Dim instruction As String = "Ambient relative humidity in all rooms"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub


    Private Sub cmdDist_HOC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDist_HOC.Click

        Dim param As String = "Heat of Combustion PFO"
        Dim units As String = "kJ/g"
        Dim instruction As String = "Average heat of combustion for all room contents used in the post flashover burning stage"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim param As String = "Soot Preflashover Yield"
        Dim units As String = "g/g"
        Dim instruction As String = "Average value of fuel soot yield used for all items burning prior to flashover"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim param As String = "CO Preflashover Yield"
        Dim units As String = "g/g"
        Dim instruction As String = "Average value of fuel CO yield used for all items burning prior to flashover"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim param As String = "Fuel Heat of Gasification"
        Dim units As String = "kJ/g"
        Dim instruction As String = "Latent Fuel Heat of Gasification"

        Call frmDistParam.ShowDistributionForm(param, units, instruction)

    End Sub


    Private Sub txtInteriorTemp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInteriorTemp.Validated

        ErrorProvider1.Clear()

        InteriorTemp = CDbl(txtInteriorTemp.Text) + 273

        Update_Distributions("Interior Temperature")

    End Sub

    Private Sub txtInteriorTemp_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtInteriorTemp.Validating

        If IsNumeric(txtInteriorTemp.Text) Then
            If (CDbl(txtInteriorTemp.Text) <= 80 And CDbl(txtInteriorTemp.Text) >= -30) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtInteriorTemp.Select(0, txtInteriorTemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtInteriorTemp, "Invalid Entry. Temperature must be in the range -30C to 80C. ")


    End Sub

    Private Sub txtExteriorTemp_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtExteriorTemp.Validating

        If IsNumeric(txtExteriorTemp.Text) Then
            If (CDbl(txtExteriorTemp.Text) <= 80 And CDbl(txtExteriorTemp.Text) >= -30) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtExteriorTemp.Select(0, txtExteriorTemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtExteriorTemp, "Invalid Entry. Temperature must be in the range -30C to 80C. ")

    End Sub

    Private Sub txtExteriorTemp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExteriorTemp.Validated

        ErrorProvider1.Clear()

        ExteriorTemp = CDbl(txtExteriorTemp.Text) + 273

        Update_Distributions("Exterior Temperature")

    End Sub

    Private Sub txtRelativeHumidity_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtRelativeHumidity.Validating

        If IsNumeric(txtRelativeHumidity.Text) Then
            If (CDbl(txtRelativeHumidity.Text) <= 100 And CDbl(txtRelativeHumidity.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtRelativeHumidity.Select(0, txtRelativeHumidity.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtRelativeHumidity, "Invalid Entry. Relative Humidity must be in the range 0 to 100%. ")

    End Sub

    Private Sub txtRelativeHumidity_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRelativeHumidity.Validated

        ErrorProvider1.Clear()

        RelativeHumidity = CSng(txtRelativeHumidity.Text) / 100  '%

        Update_Distributions("Relative Humidity")

    End Sub

    Private Sub txtpreCO_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpreCO.Validated
        ErrorProvider1.Clear()

        preCO = CSng(txtpreCO.Text)

        Update_Distributions("CO Preflashover Yield")
    End Sub

    Private Sub txtpreCO_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtpreCO.Validating
        If IsNumeric(txtpreCO.Text) Then
            If (CDbl(txtpreCO.Text) <= 1 And CDbl(txtpreCO.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtpreCO.Select(0, txtpreCO.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtpreCO, "Invalid Entry. CO Yield must be in the range 0 to 1. ")

    End Sub

    Private Sub txtpreSoot_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtpreSoot.Validated
        ErrorProvider1.Clear()

        preSoot = CSng(txtpreSoot.Text)

        Update_Distributions("Soot Preflashover Yield")
    End Sub

    Private Sub txtpreSoot_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtpreSoot.Validating

        If IsNumeric(txtpreSoot.Text) Then
            If (CDbl(txtpreSoot.Text) <= 1 And CDbl(txtpreSoot.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtpreSoot.Select(0, txtpreSoot.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtpreSoot, "Invalid Entry. Soot Yield must be in the range 0 to 1. ")
    End Sub

    Private Sub txtLHoG_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLHoG.Validated
        'ErrorProvider1.Clear()

        'FuelHeatofGasification = CSng(txtLHoG.Text)

        'Update_Distributions("Fuel Heat of Gasification")
    End Sub

    Private Sub txtLHoG_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtLHoG.Validating
        If IsNumeric(txtLHoG.Text) Then
            If (CDbl(txtLHoG.Text) <= 20 And CDbl(txtLHoG.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtLHoG.Select(0, txtLHoG.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtLHoG, "Invalid Entry. Latent Heat of Gasification must be in the range 0 to 20 kJ/g. ")
    End Sub



    Private Sub txtHOCFuel_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHOCFuel.Validated
        ErrorProvider1.Clear()

        HoC_fuel = CDbl(txtHOCFuel.Text)

        Update_Distributions("Heat of Combustion PFO")
    End Sub

    Private Sub txtHOCFuel_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtHOCFuel.Validating
        If IsNumeric(txtHOCFuel.Text) Then
            If (CSng(txtHOCFuel.Text) <= 60 And CSng(txtHOCFuel.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtHOCFuel.Select(0, txtHOCFuel.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtHOCFuel, "Invalid Entry. Heat of Combustion must be in the range 0 to 60 kJ/g. ")

    End Sub


    Private Sub txtErrorControlVentFlow_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtErrorControlVentFlow.Validated
        ErrorProvider1.Clear()
        Error_Control_ventflow = CDbl(txtErrorControlVentFlow.Text)
    End Sub

    Private Sub txtErrorControlVentFlow_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtErrorControlVentFlow.Validating
        If IsNumeric(txtErrorControlVentFlow.Text) Then
            If (CDbl(txtErrorControlVentFlow.Text) <= 0.1 And CDbl(txtErrorControlVentFlow.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtErrorControlVentFlow.Select(0, txtErrorControlVentFlow.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtErrorControlVentFlow, "Invalid Entry. Value must be in the range 0 to 0.1 ")
    End Sub

    Private Sub txtJobNumber_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtJobNumber.Validated
        ErrorProvider1.Clear()
        'store the model descriptions as a text string
        JobNumber = txtJobNumber.Text
    End Sub

    Private Sub txtDescription_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.Validated
        ErrorProvider1.Clear()
        'store the model descriptions as a text string
        Description = txtDescription.Text
        MDIFrmMain.ToolStripStatusLabel2.Text = Description
    End Sub


    Private Sub txtFED_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFED.Validated
        ErrorProvider1.Clear()
        'store the FED endpoint as a variable
        If IsNumeric(txtFED.Text) = True Then
            FEDEndPoint = CSng(txtFED.Text)
        End If
    End Sub

    Private Sub txtFED_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFED.Validating
        If IsNumeric(txtFED.Text) Then
            If (CSng(txtFED.Text) <= 4 And CSng(txtFED.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFED.Select(0, txtFED.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFED, "Invalid Entry. Value must be in the range 0 to 4")
    End Sub

    Private Sub txtTarget_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTarget.Validated
        ErrorProvider1.Clear()
        'store the target incident flux endpoint as a variable
        If IsNumeric(txtTarget.Text) = True Then
            TargetEndPoint = CSng(txtTarget.Text)
        End If
    End Sub

    Private Sub txtTarget_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTarget.Validating
        If IsNumeric(txtTarget.Text) Then
            If (CSng(txtTarget.Text) <= 4 And CSng(txtTarget.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtTarget.Select(0, txtTarget.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtTarget, "Invalid Entry. Value must be in the range 0 to 4")
    End Sub

    Private Sub txtVisibility_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVisibility.Validated
        ErrorProvider1.Clear()
        'store the simulation time as a variable
        If IsNumeric(txtVisibility.Text) = True Then
            VisibilityEndPoint = CSng(txtVisibility.Text)
        End If
    End Sub

    Private Sub txtVisibility_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtVisibility.Validating
        If IsNumeric(txtVisibility.Text) Then
            If (CSng(txtVisibility.Text) <= 20 And CSng(txtVisibility.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtVisibility.Select(0, txtVisibility.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtVisibility, "Invalid Entry. Value must be in the range 0 to 20.")
    End Sub

    Private Sub txtTemp_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTemp.Validated
        ErrorProvider1.Clear()
        'store the upper temperature endpoint as a variable
        If IsNumeric(txtTemp.Text) = True Then
            TempEndPoint = CSng(txtTemp.Text) + 273 'K
        End If
    End Sub

    Private Sub txtTemp_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTemp.Validating
        If IsNumeric(txtTemp.Text) Then
            If (CSng(txtTemp.Text) <= 1300 And CSng(txtTemp.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtTemp.Select(0, txtTemp.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtTemp, "Invalid Entry. Value must be in the range 0 to 1300 C.")
    End Sub

    Private Sub txtConvect_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConvect.Validated
        ErrorProvider1.Clear()
        'store the simulation time as a variable
        If IsNumeric(txtConvect.Text) = True Then
            ConvectEndPoint = CSng(txtConvect.Text) + 273
        End If
    End Sub

    Private Sub txtConvect_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtConvect.Validating
        If IsNumeric(txtConvect.Text) Then
            If (CSng(txtConvect.Text) <= 1300 And CSng(txtConvect.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtConvect.Select(0, txtConvect.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtConvect, "Invalid Entry. Value must be in the range 0 to 1300 C.")
    End Sub

    Private Sub txtMonitorHeight_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMonitorHeight.Validated
        ErrorProvider1.Clear()
        'store monitoring height for toxicity as a variable
        If IsNumeric(txtMonitorHeight.Text) = True Then
            MonitorHeight = CSng(txtMonitorHeight.Text)
        End If
    End Sub

    Private Sub txtMonitorHeight_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMonitorHeight.Validating
        If IsNumeric(txtMonitorHeight.Text) Then
            If (CSng(txtMonitorHeight.Text) <= 100 And CSng(txtMonitorHeight.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtMonitorHeight.Select(0, txtMonitorHeight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtMonitorHeight, "Invalid Entry. Value must be in the range 0 to 100 m.")
    End Sub


    Private Sub txtBurnerWidth_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBurnerWidth.Validated
        ErrorProvider1.Clear()
        'store the burner as a variable
        BurnerWidth = CSng(txtBurnerWidth.Text)

    End Sub

    Private Sub txtBurnerWidth_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtBurnerWidth.Validating
        If IsNumeric(txtBurnerWidth.Text) Then
            If (CSng(txtBurnerWidth.Text) <= 10 And CSng(txtBurnerWidth.Text) >= 0.01) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtBurnerWidth.Select(0, txtBurnerWidth.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtBurnerWidth, "Invalid Entry. Value must be in the range 0.01 to 10 m.")
    End Sub

    Private Sub txtFlameAreaConstant_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFlameAreaConstant.Validated
        ErrorProvider1.Clear()
        'store the flame area constant as a variable
        'store the burner as a variable
        If IsNumeric(txtFlameAreaConstant.Text) = True Then
            FlameAreaConstant = CSng(txtFlameAreaConstant.Text)
        End If

    End Sub

    Private Sub txtFlameAreaConstant_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFlameAreaConstant.Validating
        If IsNumeric(txtFlameAreaConstant.Text) Then
            If (CSng(txtFlameAreaConstant.Text) <= 1 And CSng(txtFlameAreaConstant.Text) >= 0.001) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFlameAreaConstant.Select(0, txtFlameAreaConstant.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFlameAreaConstant, "Invalid Entry. Value must be in the range 0.001 to 1.")
    End Sub

    Private Sub txtFlameLengthPower_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFlameLengthPower.Validated
        ErrorProvider1.Clear()
        'store the flame length power as a variable

        FlameLengthPower = CDbl(txtFlameLengthPower.Text)

    End Sub

    Private Sub txtFlameLengthPower_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFlameLengthPower.Validating
        If IsNumeric(txtFlameLengthPower.Text) Then
            If (CDbl(txtFlameLengthPower.Text) <= 10 And CDbl(txtFlameLengthPower.Text) >= 0) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFlameLengthPower.Select(0, txtFlameLengthPower.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFlameLengthPower, "Invalid Entry. Value must be in the range 0 to 10.")
    End Sub



    Private Sub txtFuelThickness_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFuelThickness.Validated
        ErrorProvider1.Clear()

        Fuel_Thickness = CDbl(txtFuelThickness.Text)

    End Sub

    Private Sub txtFuelThickness_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFuelThickness.Validating
        If IsNumeric(txtFuelThickness.Text) Then
            If (CSng(txtFuelThickness.Text) <= 1 And CSng(txtFuelThickness.Text) >= 0.01) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtFuelThickness.Select(0, txtFuelThickness.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtFuelThickness, "Invalid Entry. Value must be in the range 0.01 to 1.")
    End Sub

    Private Sub txtStickSpacing_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStickSpacing.Validated
        ErrorProvider1.Clear()
        'store the burner as a variable

        Stick_Spacing = CDbl(txtStickSpacing.Text)

    End Sub

    Private Sub txtStickSpacing_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStickSpacing.Validating
        If IsNumeric(txtStickSpacing.Text) Then
            If (CSng(txtStickSpacing.Text) <= 1 And CSng(txtStickSpacing.Text) >= 0.01) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtStickSpacing.Select(0, txtStickSpacing.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtStickSpacing, "Invalid Entry. Value must be in the range 0.01 to 1.")
    End Sub



    Private Sub txtErrorControl_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtErrorControl.Validated
        ErrorProvider1.Clear()

        Error_Control = CDbl(txtErrorControl.Text)

    End Sub

    Private Sub txtErrorControl_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtErrorControl.Validating
        If IsNumeric(txtErrorControl.Text) Then
            If (CDbl(txtErrorControl.Text) <= 0.1 And CDbl(txtErrorControl.Text) >= 0.000001) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtErrorControl.Select(0, txtErrorControl.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtErrorControl, "Invalid Entry. Value must be in the range 0.000001 to 0.1")

    End Sub


    Private Sub _txtNodes_0_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _txtNodes_0.TextChanged

    End Sub

    Private Sub _txtNodes_0_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles _txtNodes_0.Validated
        ErrorProvider1.Clear()

        Ceilingnodes = CShort(_txtNodes_0.Text)

    End Sub

    Private Sub _txtNodes_0_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles _txtNodes_0.Validating
        If IsNumeric(_txtNodes_0.Text) Then
            If (CShort(_txtNodes_0.Text) <= 100 And CShort(_txtNodes_0.Text) >= 5) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        _txtNodes_0.Select(0, _txtNodes_0.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(_txtNodes_0, "Invalid Entry. Value must be in the range 5 to 100")

    End Sub

    Private Sub _txtNodes_1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _txtNodes_1.TextChanged

    End Sub

    Private Sub _txtNodes_1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles _txtNodes_1.Validated
        ErrorProvider1.Clear()

        Wallnodes = CShort(_txtNodes_1.Text)
    End Sub

    Private Sub _txtNodes_1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles _txtNodes_1.Validating
        If IsNumeric(_txtNodes_1.Text) Then
            If (CShort(_txtNodes_1.Text) <= 100 And CShort(_txtNodes_1.Text) >= 5) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        _txtNodes_1.Select(0, _txtNodes_1.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(_txtNodes_1, "Invalid Entry. Value must be in the range 5 to 100")

    End Sub

    Private Sub _txtNodes_2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _txtNodes_2.TextChanged

    End Sub

    Private Sub _txtNodes_2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles _txtNodes_2.Validated
        ErrorProvider1.Clear()

        Floornodes = CShort(_txtNodes_2.Text)
    End Sub

    Private Sub _txtNodes_2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles _txtNodes_2.Validating
        If IsNumeric(_txtNodes_2.Text) Then
            If (CShort(_txtNodes_2.Text) <= 100 And CShort(_txtNodes_2.Text) >= 5) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        _txtNodes_2.Select(0, _txtNodes_2.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(_txtNodes_2, "Invalid Entry. Value must be in the range 5 to 100")

    End Sub

    Private Sub optCOauto_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOauto.CheckedChanged

    End Sub

    Private Sub optCOman_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOman.CheckedChanged

    End Sub

    Private Sub txtStickSpacing_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtStickSpacing.TextChanged

    End Sub

    Private Sub txtExcessFuelFactor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtExcessFuelFactor.TextChanged

    End Sub

    Private Sub txtExcessFuelFactor_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExcessFuelFactor.Validated
        ErrorProvider1.Clear()
        'store the burner as a variable

        ExcessFuelFactor = CDbl(txtExcessFuelFactor.Text)

    End Sub

    Private Sub txtExcessFuelFactor_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtExcessFuelFactor.Validating
        If IsNumeric(txtExcessFuelFactor.Text) Then
            If (CSng(txtExcessFuelFactor.Text) <= 4.0 And CSng(txtExcessFuelFactor.Text) >= 1) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtExcessFuelFactor.Select(0, txtExcessFuelFactor.Text.Length)

        ' Give the ErrorProvider the error message to display.

        ErrorProvider1.SetError(txtExcessFuelFactor, "Invalid Entry. Value must be in the range 1.0 to 4.0.")
    End Sub


    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCribheight.Validated
        ErrorProvider1.Clear()
        'store the burner as a variable

        Cribheight = CDbl(txtCribheight.Text)
    End Sub

    Private Sub txtCribheight_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCribheight.Validating
        If IsNumeric(txtCribheight.Text) Then
            If (CSng(txtCribheight.Text) <= 2 And CSng(txtCribheight.Text) >= Fuel_Thickness) Then
                'okay
                Exit Sub
            End If
        End If

        ' Cancel the event moving off of the control.
        e.Cancel = True

        ' Select the offending text.
        txtStickSpacing.Select(0, txtCribheight.Text.Length)

        ' Give the ErrorProvider the error message to
        ' display.
        ErrorProvider1.SetError(txtCribheight, "Invalid Entry. Crib Height must be in the range 0.1 to 2.0 m")
    End Sub


    Private Sub txtHOCFuel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHOCFuel.TextChanged

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWoodOption.Click

        If useCLTmodel = True Then
            frmCLT.optCLTOFF.Checked = False
            frmCLT.optCLTON.Checked = True
        Else
            frmCLT.optCLTOFF.Checked = True
            frmCLT.optCLTON.Checked = False
        End If
        If IntegralModel = True Then
            frmCLT.RB_Integral.Checked = True
            frmCLT.panel_kinetic.Visible = False
            frmCLT.Panel_integralmodel.Visible = True
        ElseIf KineticModel = True Then
            frmCLT.RB_Kinetic.Checked = True
            frmCLT.panel_kinetic.Visible = True
            frmCLT.Panel_integralmodel.Visible = False
        Else
            frmCLT.RB_dynamic.Checked = True
            frmCLT.panel_kinetic.Visible = False
            frmCLT.Panel_integralmodel.Visible = False
        End If

        frmCLT.numeric_ceilareapercent.Value = CLTceilingpercent
        frmCLT.numeric_wallareapercent.Value = CLTwallpercent
        frmCLT.txtCharTemp.Text = chartemp
        frmCLT.txtCLTcalibration.Text = CLTcalibrationfactor
        frmCLT.txtLamellaDepth.Text = Lamella
        frmCLT.txtFlameFlux.Text = CLTflameflux
        frmCLT.txtCLTigtemp.Text = CLTigtemp
        frmCLT.txtCLTLoG.Text = CLTLoG
        frmCLT.txtCritFlux.Text = CLTQcrit
        frmCLT.txtDebondTemp.Text = DebondTemp

        frmCLT.txtA_cell.Text = Format(E_array(1), "0.00E+00")
        frmCLT.txtE_cell.Text = Format(A_array(1), "0.00E+00")
        frmCLT.txtReact_cell.Text = n_array(1)
        frmCLT.txtInit_cell.Text = mf_compinit(1)
        frmCLT.txtA_hemi.Text = Format(E_array(2), "0.00E+00")
        frmCLT.txtE_hemi.Text = Format(A_array(2), "0.00E+00")
        frmCLT.txtReact_hemi.Text = n_array(2)
        frmCLT.txtInit_hemi.Text = mf_compinit(2)
        frmCLT.txtA_lignin.Text = Format(E_array(3), "0.00E+00")
        frmCLT.txtE_lignin.Text = Format(A_array(3), "0.00E+00")
        frmCLT.txtReact_lignin.Text = n_array(3)
        frmCLT.txtInit_lignin.Text = mf_compinit(3)
        frmCLT.txtA_water.Text = Format(E_array(0), "0.00E+00")
        frmCLT.txtE_water.Text = Format(A_array(0), "0.00E+00")
        frmCLT.txtReact_water.Text = n_array(0)
        frmCLT.txtInit_water.Text = mf_compinit(0)
        frmCLT.TextBox_charyield_cell.Text = char_yield(1)
        frmCLT.TextBox_charyield_hemi.Text = char_yield(2)
        frmCLT.TextBox_charyield_lignin.Text = char_yield(3)

        Hide()

        frmCLT.Show()

    End Sub



    Private Sub txtLHoG_TextChanged(sender As Object, e As EventArgs) Handles txtLHoG.TextChanged

    End Sub
End Class