Option Strict Off
Option Explicit On
Friend Class frmView
	Inherits System.Windows.Forms.Form
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		'UPGRADE_NOTE: Top was upgraded to Top_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Top_Renamed As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		'UPGRADE_NOTE: Bottom was upgraded to Bottom_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Bottom_Renamed As Integer
	End Structure
	Private Structure CharRange
		Dim cpMin As Integer ' First character of range (0 for start of doc)
		Dim cpMax As Integer ' Last character of range (-1 for end of doc)
	End Structure
	
	Private Structure FormatRange
		Dim hdc As Integer ' Actual DC to draw on
		Dim hdcTarget As Integer ' Target DC for determining text formatting
		Dim rc As RECT ' Region of the DC to draw to (in twips)
		Dim rcPage As RECT ' Region of the entire DC (page size) (in twips)
		Dim chrg As CharRange ' Range of text to draw (see above declaration)
	End Structure
	
	Private Const WM_USER As Integer = &H400
	Private Const EM_FORMATRANGE As Integer = WM_USER + 57
	Private Const EM_SETTARGETDEVICE As Integer = WM_USER + 72
	Private Const PHYSICALOFFSETX As Integer = 112
	Private Const PHYSICALOFFSETY As Integer = 113
	
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'Private Declare Function SendMessage Lib "User32"  Alias "SendMessageA"(ByVal hWnd As Integer, ByVal msg As Integer, ByVal wp As Integer, ByRef lp As Any) As Integer
	Private Declare Function CreateDC Lib "gdi32"  Alias "CreateDCA"(ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Integer, ByVal lpInitData As Integer) As Integer
	
	Private Sub cmdFonts_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFonts.Click
		'*  ==========================================================
		'*  This subprocedure shows a custom dialog box for
		'*  opening a file.
		'*  ==========================================================
		
		Dim getfont As System.Windows.Forms.Control
		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
		'UPGRADE_ISSUE: VB.Control getfont was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'getfont = frmBlank.CMDialog1
		
		'CancelError is True
		On Error GoTo errhandler
		
		'Display the Open dialog box
		'UPGRADE_WARNING: Couldn't resolve default property of object getfont.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: Constant cdlCFFixedPitchOnly was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Constant cdlCFPrinterFonts was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'getfont.flags = MSComDlg.FontsConstants.cdlCFPrinterFonts + MSComDlg.FontsConstants.cdlCFFixedPitchOnly
		'UPGRADE_WARNING: Couldn't resolve default property of object getfont.ShowFont. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'getfont.ShowFont()
		'UPGRADE_WARNING: Couldn't resolve default property of object getfont.FontName. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrinterFontName = getfont.FontName
		'UPGRADE_WARNING: Couldn't resolve default property of object getfont.FontSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'PrinterFontSize = getfont.FontSize
        'If PrinterFontName = "" Then PrinterFontName = DefaultPrinterFontName
        'If PrinterFontSize = "" Then PrinterFontSize = DefaultPrinterFontSize
		' Set the font of the RTF to a TrueType font for best results
        RichTextBox1.Font = VB6.FontChangeName(RichTextBox1.Font, PrinterFontName)
		RichTextBox1.Font = VB6.FontChangeSize(RichTextBox1.Font, CDec(PrinterFontSize))
		
        'System.Windows.Forms.Cursor.Current = Default_Renamed
		Exit Sub
		
errhandler: 
		'User pressed Cancel button
        'System.Windows.Forms.Cursor.Current = Default_Renamed
		Exit Sub
		
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		'frmView.Hide
		Me.Close()
        'frmBlank.Show()
	End Sub
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		'frmView.Hide
		frmprintvar.Show()
	End Sub
	
	Private Sub Command4_Click()
	End Sub
	
	Private Sub frmView_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim LineWidth As Integer
		Dim LeftMargin, TopMargin As Integer
		Dim RightMargin, BottomMargin As Integer
		
		On Error GoTo errorhandler
		
		' Initialize Form and Command button
		Me.Text = "Summary of Simulation Results"
		Command1.SetBounds(VB6.TwipsToPixelsX(10), VB6.TwipsToPixelsY(10), VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(380))
		Command1.Text = "&Print"
		cmdFonts.SetBounds(VB6.TwipsToPixelsX(2210), VB6.TwipsToPixelsY(10), VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(380))
		Command2.Text = "&Close"
		Command3.SetBounds(VB6.TwipsToPixelsX(1110), VB6.TwipsToPixelsY(10), VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(380))
		Command3.Text = "&Variables"
		Command2.SetBounds(VB6.TwipsToPixelsX(3310), VB6.TwipsToPixelsY(10), VB6.TwipsToPixelsX(1000), VB6.TwipsToPixelsY(380))
		cmdFonts.Text = "&Select Font"
		
		'Set the font of the RTF to a TrueType font for best results
		If PrinterFontName <> "" Then
			RichTextBox1.Font = VB6.FontChangeName(RichTextBox1.Font, PrinterFontName)
		Else
			RichTextBox1.Font = VB6.FontChangeName(RichTextBox1.Font, DefaultPrinterFontName)
		End If
		If PrinterFontSize <> "" Then
			RichTextBox1.Font = VB6.FontChangeSize(RichTextBox1.Font, CDec(PrinterFontSize))
		Else
			RichTextBox1.Font = VB6.FontChangeSize(RichTextBox1.Font, CDec(DefaultPrinterFontSize))
		End If
		
		'    'disabled 13/10/08
		'    ' Tell the RTF to base it's display off of the printer
		'    If PagePrinter Is Nothing Then
		'        Set PagePrinter = New PPrinter
		'    End If
		'    PagePrinter.Get_Margins Me.hWnd, 0, LeftMargin, TopMargin, RightMargin, BottomMargin
		'    LineWidth = WYSIWYG_RTF(RichTextBox1, LeftMargin * 1440 / 2540, RightMargin * 1440 / 2540) '1440 Twips=1 Inch
		'    Set PagePrinter = Nothing
		'    ' Set the form width to match the line width
		'    If LineWidth > 0 Then Me.Width = LineWidth + 200
		
		'call procedure to centre form
		Centre_Form(Me)
		Exit Sub
		
errorhandler: 
		MsgBox(ErrorToString(Err.Number))
		
	End Sub
	
	Private Sub frmView_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Me.Hide()
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event frmView.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmView_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		On Error Resume Next
		' Position the RTF on form
		RichTextBox1.SetBounds(VB6.TwipsToPixelsX(100), VB6.TwipsToPixelsY(500), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 200), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 600))
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		''disabled 13/10/08
		
		'    'Print the contents of the RichTextBox with a one inch margin
		'    Dim LeftMargin As Long, TopMargin As Long
		'    Dim RightMargin As Long, BottomMargin As Long
		'    If PagePrinter Is Nothing Then
		'        Set PagePrinter = New PPrinter
		'    End If
		'    PagePrinter.PrintRTF RichTextBox1, Description
		'    Set PagePrinter = Nothing
	End Sub
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'
	' WYSIWYG_RTF - Sets an RTF control to display itself the same as it
	'               would print on the default printer
	'
	' RTF - A RichTextBox control to set for WYSIWYG display.
	'
	' LeftMarginWidth - Width of desired left margin in twips
	'
	' RightMarginWidth - Width of desired right margin in twips
	'
	' Returns - The length of a line on the printer in twips
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Function WYSIWYG_RTF(ByRef RTF As System.Windows.Forms.RichTextBox, ByRef LeftMarginWidth As Integer, ByRef RightMarginWidth As Integer) As Integer
        '		Dim Printer As New Printer
        '		Dim LeftMargin, LeftOffset, RightMargin As Integer
        '		Dim LineWidth As Integer
        '		Dim PrinterhDC As Integer
        '		Dim R As Integer

        '		On Error GoTo errorhandler
        '		' Start a print job to initialize printer object
        '		'Printer.Print Space(1)
        '		Printer.ScaleMode = ScaleModeConstants.vbTwips

        '		' Get the offset to the printable area on the page in twips
        '		'UPGRADE_ISSUE: Printer property Printer.hdc was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '		LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), ScaleModeConstants.vbPixels, ScaleModeConstants.vbTwips)
        '		' Calculate the Left, and Right margins
        '		LeftMargin = LeftMarginWidth - LeftOffset
        '		RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset

        '		' Calculate the line width
        '		LineWidth = RightMargin - LeftMargin

        '		' Create an hDC on the Printer pointed to by the Printer object
        '		' This DC needs to remain for the RTF to keep up the WYSIWYG display
        '		'UPGRADE_ISSUE: Printer property Printer.DriverName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        '		PrinterhDC = CreateDC(Printer.DriverName, Printer.DeviceName, 0, 0)

        '		' Tell the RTF to base it's display off of the printer
        '		'    at the desired line width
        '		R = SendMessage(RTF.Handle.ToInt32, EM_SETTARGETDEVICE, PrinterhDC, LineWidth)

        '		' Abort the temporary print job used to get printer info
        '		'Printer.KillDoc

        '		WYSIWYG_RTF = LineWidth
        '		Exit Function

        'errorhandler: 
        '		Printer.KillDoc()
	End Function

End Class