Imports System.Collections.Generic
Imports System.Math
Imports CenterSpace.NMath.Core
Imports CenterSpace.NMath.Stats

Public Class oCVent
    Implements ICloneable
    Private m_id As Integer
    Private m_lowerroom As Integer
    Private m_upperroom As Integer
    Private m_area As Double
    Private m_description As String
    Private m_integrity As Double
    Private m_maxopening As Double
    Private m_maxopeningtime As Double
    Private m_gastemp As Double
    Private m_opentime As Double
    Private m_closetime As Double
    Private m_autoopenvent As Boolean
    Private m_triggerFR As Boolean
    Private m_FRcriteria As Integer
    Private m_sdtriggerroom As Integer
    Private m_triggerHRR As Boolean
    Private m_triggerSD As Boolean
    Private m_triggerHD As Boolean
    Private m_triggerVL As Boolean
    Private m_triggerHO As Boolean
    Private m_triggerFO As Boolean
    Private m_triggerventopenduration As Double
    Private m_triggerventopendelay As Double
    Private m_HRRventopenduration As Double
    Private m_HRRventopendelay As Double
    Private m_HRRthreshold As Double
    Private m_dischargecoeff As Double
    Public Property triggerFO() As Boolean
        Get
            Return m_triggerFO
        End Get
        Set(ByVal value As Boolean)
            m_triggerFO = value
        End Set
    End Property
    Public Property triggerHO() As Boolean
        Get
            Return m_triggerHO
        End Get
        Set(ByVal value As Boolean)
            m_triggerHO = value
        End Set
    End Property
    Public Property triggerVL() As Boolean
        Get
            Return m_triggerVL
        End Get
        Set(ByVal value As Boolean)
            m_triggerVL = value
        End Set
    End Property
    Public Property sdtriggerroom() As Integer
        Get
            Return m_sdtriggerroom
        End Get
        Set(ByVal value As Integer)
            m_sdtriggerroom = value
        End Set
    End Property
    Public Property dischargecoeff() As Double
        Get
            Return m_dischargecoeff
        End Get
        Set(ByVal value As Double)
            m_dischargecoeff = value
        End Set
    End Property
    Public Property HRRthreshold() As Double
        Get
            Return m_HRRthreshold
        End Get
        Set(ByVal value As Double)
            m_HRRthreshold = value
        End Set
    End Property
    Public Property HRRventopendelay() As Double
        Get
            Return m_HRRventopendelay
        End Get
        Set(ByVal value As Double)
            m_HRRventopendelay = value
        End Set
    End Property
    Public Property HRRventopenduration() As Double
        Get
            Return m_HRRventopenduration
        End Get
        Set(ByVal value As Double)
            m_HRRventopenduration = value
        End Set
    End Property
    Public Property triggerventopendelay() As Double
        Get
            Return m_triggerventopendelay
        End Get
        Set(ByVal value As Double)
            m_triggerventopendelay = value
        End Set
    End Property
    Public Property triggerventopenduration() As Double
        Get
            Return m_triggerventopenduration
        End Get
        Set(ByVal value As Double)
            m_triggerventopenduration = value
        End Set
    End Property
    Public Property triggerHD() As Boolean
        Get
            Return m_triggerHD
        End Get
        Set(ByVal value As Boolean)
            m_triggerHD = value
        End Set
    End Property
    Public Property triggerSD() As Boolean
        Get
            Return m_triggerSD
        End Get
        Set(ByVal value As Boolean)
            m_triggerSD = value
        End Set
    End Property
    Public Property triggerHRR() As Boolean
        Get
            Return m_triggerHRR
        End Get
        Set(ByVal value As Boolean)
            m_triggerHRR = value
        End Set
    End Property
    Public Property FRcriteria() As Integer
        Get
            Return m_FRcriteria
        End Get
        Set(ByVal value As Integer)
            m_FRcriteria = value
        End Set
    End Property
    Public Property triggerFR() As Boolean
        Get
            Return m_triggerFR
        End Get
        Set(ByVal value As Boolean)
            m_triggerFR = value
        End Set
    End Property
    Public Property autoopenvent() As Boolean
        Get
            Return m_autoopenvent
        End Get
        Set(ByVal value As Boolean)
            m_autoopenvent = value
        End Set
    End Property
    Public Property opentime() As Double
        Get
            Return m_opentime
        End Get
        Set(ByVal value As Double)
            m_opentime = value
        End Set
    End Property
    Public Property closetime() As Double
        Get
            Return m_closetime
        End Get
        Set(ByVal value As Double)
            m_closetime = value
        End Set
    End Property
    Public Property gastemp() As Double
        Get
            Return m_gastemp
        End Get
        Set(ByVal value As Double)
            m_gastemp = value
        End Set
    End Property
    Public Property maxopeningtime() As Double
        Get
            Return m_maxopeningtime
        End Get
        Set(ByVal value As Double)
            m_maxopeningtime = value
        End Set
    End Property
    Public Property maxopening() As Double
        Get
            Return m_maxopening
        End Get
        Set(ByVal value As Double)
            m_maxopening = value
        End Set
    End Property
    Public Property integrity() As Double
        Get
            Return m_integrity
        End Get
        Set(ByVal value As Double)
            m_integrity = value
        End Set
    End Property
    Public Property id() As Integer
        Get
            Return m_id
        End Get
        Set(ByVal value As Integer)
            m_id = value
        End Set
    End Property
    Public Property lowerroom() As Integer
        Get
            Return m_lowerroom
        End Get
        Set(ByVal value As Integer)
            m_lowerroom = value
        End Set
    End Property
    Public Property upperroom() As Integer
        Get
            Return m_upperroom
        End Get
        Set(ByVal value As Integer)
            m_upperroom = value
        End Set
    End Property
    Public Property description() As String
        Get
            Return m_description
        End Get
        Set(ByVal value As String)
            m_description = value
        End Set
    End Property
    Public Property area() As Double
        Get
            Return m_area
        End Get
        Set(ByVal value As Double)
            m_area = value
        End Set
    End Property

    Public Sub New()
    End Sub

    Public Sub New(ByVal gastemp As Double, ByVal maxopeningtime As Single, ByVal maxopening As Single, ByVal integrity As Single, ByVal description As String, ByVal upperroom As Integer, ByVal lowerroom As Integer, ByVal area As Double, _
               ByVal opentime As Double, ByVal closetime As Double, ByVal autoopenvent As Boolean, ByVal triggerFR As Boolean, ByVal FRcriteria As Integer, ByVal dischargecoeff As Double)

        Me.lowerroom = lowerroom
        Me.upperroom = upperroom
        Me.area = area
        Me.description = description
        Me.gastemp = gastemp
        Me.maxopeningtime = maxopeningtime
        Me.maxopening = maxopening
        Me.integrity = integrity
        Me.opentime = opentime
        Me.closetime = closetime
        Me.autoopenvent = autoopenvent
        Me.triggerFR = triggerFR
        Me.FRcriteria = FRcriteria
        Me.dischargecoeff = dischargecoeff

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim v As New oCVent

        Try

            v.description = Me.description
            v.lowerroom = Me.lowerroom
            v.upperroom = Me.upperroom
            v.area = Me.area
            v.integrity = Me.integrity
            v.maxopening = Me.maxopening
            v.maxopeningtime = Me.maxopeningtime
            v.gastemp = Me.gastemp
            v.autoopenvent = Me.autoopenvent
            v.opentime = Me.opentime
            v.closetime = Me.closetime
            v.triggerFR = Me.triggerFR
            v.FRcriteria = Me.FRcriteria
            v.dischargecoeff = Me.dischargecoeff

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in cvent2.vb Clone")
        End Try

        Return v

    End Function
    Public Function GetDisplayText(ByVal j As String) As String
        Dim text As String

        If upperroom <= NumberRooms And lowerroom <= NumberRooms Then
            text = LeftJustified(j, 2) & " " & _
            RightJustified(description, 18) & _
             RightJustified(upperroom.ToString, 9) & _
             RightJustified(lowerroom.ToString, 9) & _
            RightJustified(area.ToString, 11)
        ElseIf upperroom <= NumberRooms And lowerroom > NumberRooms Then
            text = LeftJustified(j, 2) & " " & _
            RightJustified(description, 18) & _
            RightJustified(upperroom.ToString, 9) & _
            RightJustified("outside", 9) & _
            RightJustified(area.ToString, 11)
        Else
            text = LeftJustified(j, 2) & " " & _
            RightJustified(description, 18) & _
            RightJustified("outside", 9) & _
             RightJustified(lowerroom.ToString, 9) & _
           RightJustified(area.ToString, 11)
        End If

        Return text
    End Function
End Class
