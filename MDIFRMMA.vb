Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Drawing.Printing
Imports VB = Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Math
Imports CenterSpace.NMath.Core
'Imports iTextSharp.text.pdf
'Imports iTextSharp.text

Friend Class MDIFrmMain
    Inherits System.Windows.Forms.Form
    Private PrintPageSettings As New PageSettings
    Private StringToPrint As String
    'Private PrintFont As New Font("Courier New", 8.25)
    'Dim ApplicationPath As String = IO.Path.GetDirectoryName(Application.ExecutablePath)
    Dim oDistributions As List(Of oDistribution)

    <System.Runtime.InteropServices.DllImportAttribute("MathFuncsdll.dll")>
    Public Shared Function Add(ByVal a As Double, ByVal b As Double) As Double
    End Function

    Private Structure BrowseInfo
        'Handle to the owner window
        Dim hwndOwner As Integer
        'Address of an ITEMIDLIST structure specifying
        'the location of the root folder from which to
        'browse. Only the specified folder and its
        'subfolders appear in the dialog box. This member
        'can be NULL; in that case, the namespace root (the
        'desktop folder) is used.
        Dim pIDLRoot As Integer
        'Address of a buffer to receive the display name
        'of the folder selected by the user. The size of
        'this buffer is assumed to be MAX_PATH bytes.
        Dim pszDisplayName As Integer
        'Address of a null-terminated string that is displayed
        'above the tree view control in the dialog box. This
        'string can be used to specify instructions to the user.
        Dim lpszTitle As Integer
        'Flags specifying the options for the dialog box.  This
        'can include zero or a combination of the below values:
        Dim ulFlags As Integer
        'Address of an application-defined function that the
        'dialog box calls when an event occurs. This member
        'can be NULL.
        Dim lpfnCallback As Integer
        'Application-defined value that the dialog box passes
        'to the callback function, if one is specified.
        Dim lParam As Integer
        'Variable to receive the image associated with the
        'selected folder. The image is specified as an index
        'to the system image list.
        Dim iImage As Integer
    End Structure

    '
    ' NOTE:
    '    Many of these flags only work with
    '    certain versions of Shell32.dll.
    '

    'Only return computers. If the user selects anything
    'other than a computer, the OK button is grayed.
    Private Const BIF_BROWSEFORCOMPUTER As Integer = &H1000

    'Only return printers. If the user selects anything
    'other than a printer, the OK button is grayed.
    Private Const BIF_BROWSEFORPRINTER As Integer = &H2000

    'The browse dialog will display files as well as folders.
    Private Const BIF_BROWSEINCLUDEFILES As Integer = &H4000

    'Do not include network folders below the domain
    'level in the tree view control.
    Private Const BIF_DONTGOBELOWDOMAIN As Integer = &H2

    'Include an edit control in the dialog box.
    Private Const BIF_EDITBOX As Integer = &H10

    'Use the new user-interface providing the user with a larger
    'resizable dialog box which includes drag and drop, reordering,
    'context menus, new folders, delete, and other context menu
    'commands.
    Private Const BIF_NEWDIALOGSTYLE As Integer = &H40

    'Only return file system ancestors. If the user
    'selects anything other than a file system ancestor,
    'the OK button is grayed.
    Private Const BIF_RETURNFSANCESTORS As Integer = &H8

    'Only return file system directories. If the user selects
    'folders that are not part of the file system, the OK
    'button is grayed.
    Private Const BIF_RETURNONLYFSDIRS As Integer = &H1

    'Include a status area in the dialog box. The callback
    'function can set the status text by sending messages
    'to the dialog box.
    Private Const BIF_STATUSTEXT As Integer = &H4

    'Equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE..
    Private Const BIF_USENEWUI As Boolean = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)

    Private Const MAX_PATH As Short = 260

    ' Frees a block of task memory previously allocated
    ' through a call to the CoTaskMemAlloc or CoTaskMemRealloc
    ' function.
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Integer)

    ' Appends one string to another.
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Integer

    ' Displays a dialog box enabling the user to select a
    ' shell folder. The calling application is responsible
    ' for freeing the returned item identifier list by using
    ' the shell's task allocator.
    'UPGRADE_WARNING: Structure BrowseInfo may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef lpbi As BrowseInfo) As Integer

    ' Converts an item identifier list to a file system path.
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Integer, ByVal lpBuffer As String) As Integer

    '
    ' Desktop—virtual folder
    Const CSIDL_DESKTOP As Integer = &H0
    '
    ' User's program groups
    Const CSIDL_PROGRAMS As Integer = &H2
    '
    ' Control Panel.
    Const CSIDL_CONTROLS As Integer = &H3
    '
    ' Folder containing installed printers.
    Const CSIDL_PRINTERS As Integer = &H4
    '
    ' Folder that serves as a common repository for documents.
    Const CSIDL_PERSONAL As Integer = &H5
    '
    ' Folder that serves as a common repository for the user's favorite items.
    Const CSIDL_FAVORITES As Integer = &H6
    '
    ' Folder that corresponds to the user's Startup program group.
    Const CSIDL_STARTUP As Integer = &H7
    '
    ' User's most recently used documents.
    Const CSIDL_RECENT As Integer = &H8
    '
    ' Folder that contains Send To menu items.
    Const CSIDL_SENDTO As Integer = &H9
    '
    ' Recycle Bin.
    Const CSIDL_BITBUCKET As Integer = &HA
    '
    ' Start menu items.
    Const CSIDL_STARTMENU As Integer = &HB
    '
    ' Folder used to physically store file objects on the desktop.
    Const CSIDL_DESKTOPDIRECTORY As Integer = &H10
    '
    ' My Computer—virtual folder
    Const CSIDL_DRIVES As Integer = &H11
    '
    ' Network Neighborhood
    Const CSIDL_NETWORK As Integer = &H12
    '
    ' Network neighborhood.
    Const CSIDL_NETHOOD As Integer = &H13
    '
    ' Virtual folder containing fonts.
    Const CSIDL_FONTS As Integer = &H14
    '
    ' Document templates.
    Const CSIDL_TEMPLATES As Integer = &H15
    '
    ' Folder that contains the programs and folders that
    ' appear on the Start menu for all users.
    Const CSIDL_COMMON_STARTMENU As Integer = &H16
    '
    ' Folder  that contains the directories for the common
    ' program groups that appear on the Start menu for all users.
    Const CSIDL_COMMON_PROGRAMS As Integer = &H17
    '
    ' Folder that contains the programs that appear in the
    ' Startup folder for all users.
    Const CSIDL_COMMON_STARTUP As Integer = &H18
    '
    ' Folder that contains files and folders that
    ' appear on the desktop for all users.
    Const CSIDL_COMMON_DESKTOPDIRECTORY As Integer = &H19
    '
    ' Folder serving as a common repository for
    ' application-specificdata.
    Const CSIDL_APPDATA As Integer = &H1A
    '
    ' Folder that serves as a common repository for printer links.
    Const CSIDL_PRINTHOOD As Integer = &H1B
    Private oExcel As Object

    Private Structure SHITEMID
        Dim cb As Integer
        Dim abID As Byte
    End Structure

    Private Structure ITEMIDLIST
        Dim mkid As SHITEMID
    End Structure

    ' Retrieves the ID of a special folder.
    Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Integer, ByVal nFolder As Integer, ByRef pidl As ITEMIDLIST) As Integer

    Private Function fGetSpecialFolder(ByRef CSIDL As Integer, ByRef IDL As ITEMIDLIST) As String
        Dim sPath As String
        '
        ' Retrieve info about system folders such as the
        ' "Recent Documents" folder.  Info is stored in
        ' the IDL structure.
        '
        fGetSpecialFolder = ""
        If SHGetSpecialFolderLocation(Me.Handle.ToInt32, CSIDL, IDL) = 0 Then
            '
            ' Get the path from the ID list, and return the folder.
            '
            sPath = Space(MAX_PATH)
            If SHGetPathFromIDList(IDL.mkid.cb, sPath) Then
                fGetSpecialFolder = VB.Left(sPath, InStr(sPath, vbNullChar) - 1) & "\"
            End If
        End If
    End Function

    Private Sub MDIFrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim ans As Integer
        ans = MessageBox.Show("Are you sure you want to exit?", "Exiting", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If ans = 6 Then
            e.Cancel = False
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub MDIFrmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*  =======================================================
        '*      This event loads the main form
        '*  =======================================================

        'NMathConfiguration.LicenseKey = "XTUU-3UM1-YXFF-89WC-DAUS-VUUF-1QS1-WPUF-19SU-WPSS-8EJE-680T-FRPP-X8NS-WQSF-E842-68RT-WPTF-683S-X1D0-YR8V-19P2-9D3Y-LRTU-A10U"
        NMathConfiguration.LicenseKey = "XTUU-3UT0-CMA5-H9CE-EAU7-VUUF-1QS1-WPUF-19SU-WPSS-8EJE-680T-FRPP-X8NS-WQSF-E842-68RT-WPTF-6833-G8NA-2RUF-UC4D-LXL2-4QX6-M9"

        'call procedure to setup helpfile variables
        'Call SetAppHelp(Me.hwnd)

        ChDir((My.Application.Info.DirectoryPath))
        'call procedure to set defaults for a new model

        'get name of user application data folder
        'Call get_folders()
        UserAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        UserPersonalDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
        ApplicationPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath)
        CommonFilesFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
        'DocsFolder = ApplicationPath & "\docs\"
        DocsFolder = Application.StartupPath & "\docs\"
        DataFolder = Application.StartupPath & "\data\"

        'Call get_folders(CSIDL_PROGRAMS, ProgramFolder)
        Try

            'registry
            'DefaultRiskDataDirectory = GetSetting("BRISK", "options", "DefaultRiskDataDirectory")
            DefaultRiskDataDirectory = My.Settings.riskdatafolder
            If My.Settings.devkeycode = DevKeyCode Then
                DevKey = True
            End If

            If DevKey = True Then
                TalkToEVACNZToolStripMenuItem.Visible = True
                TalkToEVACNZToolStripMenuItem.Checked = TalkToEVACNZ
                'frmOptions1.cmdWoodOption.Visible = True
                'uwallchardepth.Visible = True
                'ceilingchardepth.Visible = True
                'frmNewVents.cmdFireResistance.Visible = True
                'frmNewCVents.cmdFireResistance.Visible = True
                ToolStripMenuItem23.Visible = True
                ToolStripMenuItem28.Visible = True
                ToolStripMenuItem29.Visible = True
                mnuNormalisedHeatLoad.Visible = True
                BatchFilesToolStripMenuItem.Visible = True
                frmCLT.Panel_integralmodel.Visible = True
            Else
                TalkToEVACNZToolStripMenuItem.Visible = False
                'frmOptions1.cmdWoodOption.Visible = False
                'uwallchardepth.Visible = False
                'ceilingchardepth.Visible = False
                'frmNewVents.cmdFireResistance.Visible = False
                'frmNewCVents.cmdFireResistance.Visible = False
                ToolStripMenuItem23.Visible = False
                ToolStripMenuItem28.Visible = False
                ToolStripMenuItem29.Visible = False
                mnuNormalisedHeatLoad.Visible = False
                BatchFilesToolStripMenuItem.Visible = False
                frmCLT.Panel_integralmodel.Visible = False
            End If

        Catch ex As Exception
            DefaultRiskDataDirectory = ""
        End Try

        If DefaultRiskDataDirectory = "" Then
            'DefaultRiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\"
            DefaultRiskDataDirectory = UserPersonalDataFolder & gcs_folder_ext & "\" & "riskdata\"
        End If
        frmInputs.Label4.Text = "Default riskdata folder is: " & DefaultRiskDataDirectory

        New_File_Start()

        If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"

        gsDatabase = FireDatabaseName
        Me.Text = ProgramTitle & " (" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision & ")"
        'Me.Text = ProgramTitle & " (" & Version & ")"
        My.MySettings.Default("thermalconnectionstring") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MaterialsDatabaseName
        My.MySettings.Default("fireconnectionstring") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FireDatabaseName

        TableLayoutPanel1.Dock = DockStyle.None
        TableLayoutPanel1.Height = Me.Height * 70

        'Me.AddOwnedForm(frmDescribeRoom)
        Me.AddOwnedForm(frmFire)
        Me.AddOwnedForm(frmprintvar)
        Me.AddOwnedForm(frmEndPoints)
        Me.AddOwnedForm(frmSmokeDetector)
        'Me.AddOwnedForm(frmExtract)
        Me.AddOwnedForm(frmMaterials)
        Me.AddOwnedForm(frmOptions1)
        Me.AddOwnedForm(frmFireObjectDB)
        Me.AddOwnedForm(frmGraph)
        Me.AddOwnedForm(frmPlot)
        Me.AddOwnedForm(frmInputs)
        'Me.AddOwnedForm(frmView)
        Me.AddOwnedForm(frmItemList)
        Me.AddOwnedForm(frmSprinklerList)
        Me.AddOwnedForm(frmSmokeDetList)
        Me.AddOwnedForm(frmVentList)
        Me.AddOwnedForm(frmFEDpath)
        Me.AddOwnedForm(frmVentOpenOptions)
        Me.AddOwnedForm(frmNewSmokeDetector)
        Me.AddOwnedForm(frmViewDocs)
        Me.AddOwnedForm(frmRoomList)

        frmInputs.Read_BaseFile_xml(RiskDataDirectory & "basemodel_default.xml", False)

        Me.Show()
        Me.Enabled = False

        mnuEULA.PerformClick()

    End Sub


    Private Sub MDIFrmMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Call Save_Registry()

        Dim i As Short
        ' Loop through the forms collection and unload
        ' each form.
        For i = (My.Application.OpenForms.Count - 1) To 1 Step -1
            My.Application.OpenForms(i).Close()
        Next i
        FileClose()

    End Sub

    Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
        '*  ===========================================================
        '*  This event displays a dialog box with information about the program
        '*  ===========================================================

        frmAbout.ShowDialog()

    End Sub

    '    Public Sub mnuBREAK1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBREAK1.Click
    '        '		'==========================================================
    '        '		'Creates an input file for BREAK1 glass fracture analysis
    '        '		'14 December 2001 by Ross Parry
    '        '		'Based on mnuCSV and mnuGlassTempGraph by Colleen Wade
    '        '		'=========================================================

    '        On Error GoTo more

    '        Dim Txt As String
    '        Dim k, j, count As Integer
    '        Dim Title As String
    '        Dim SaveBox As New SaveFileDialog
    '        Dim myStream As Stream
    '        Dim idc, idr, idv As Short
    '        Dim Default_Renamed, Message, MyValue As String
    '        Dim F7dot2, E12dot4, F8dot4, F6dot2 As String
    '        Dim F7dot1, F7dot3, F10dot6, F12dot2 As String
    '        Dim F8dot2, I10 As String
    '        Dim dblHalfWidth As Double
    '        Dim iTimeStep, iSteps As Short

    '        Title = "Display Nodal Temperatures for Which Vent?" ' Set title.

    '        Message = "Enter the first room" ' Set prompt.
    '        Default_Renamed = "1" ' Set default.
    '        MyValue = InputBox(Message, Title, Default_Renamed)
    '        idr = CShort(MyValue)

    '        Message = "Enter the second room" ' Set prompt.
    '        Default_Renamed = "2" ' Set default.
    '        MyValue = InputBox(Message, Title, Default_Renamed)
    '        idc = CShort(MyValue)

    '        Message = "Enter the vent ID" ' Set prompt.
    '        Default_Renamed = "1" ' Set default.
    '        MyValue = InputBox(Message, Title, Default_Renamed)
    '        idv = CShort(MyValue)

    '        Message = "Enter the timestep" ' Set prompt.
    '        Default_Renamed = "10" ' Set default.
    '        MyValue = InputBox(Message, Title, Default_Renamed)
    '        iTimeStep = CShort(MyValue)

    '        If AutoBreakGlass(idr, idc, idv) = False Then
    '            MsgBox("BREAK1 files can only be created for vents with glass fracture modelled.")
    '            Exit Sub
    '        End If

    '        count = 0

    '        'dialogbox title
    '        SaveBox.Title = "Save Results to BREAK1 File"

    '        'default filename
    '        If Len(DataFile) > 4 Then
    '            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 3) & "in"
    '        Else
    '            DataFile = "new.csv"
    '            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 3) & "in"
    '        End If

    '        'Set filters
    '        SaveBox.Filter = "All Files (*.*)|*.*|Files (*.in)|*.in"

    '        'Specify default filter
    '        SaveBox.FilterIndex = 2

    '        'default filename extension
    '        SaveBox.DefaultExt = "in"

    '        If VentHeight(idr, idc, idv) > VentWidth(idr, idc, idv) Then
    '            dblHalfWidth = VentHeight(idr, idc, idv) / 2
    '        Else : dblHalfWidth = VentWidth(idr, idc, idv) / 2
    '        End If

    '        'define a format string
    '        E12dot4 = ".0000E+00"
    '        F8dot4 = "00000000.0000"
    '        F7dot2 = "@@@@@.00"
    '        F7dot1 = "@@@@@@@.0"
    '        F6dot2 = "@@@@@@.00"
    '        F12dot2 = "@@@@@@@@@@@@.00"
    '        F7dot3 = "@@@@@@.000"
    '        F10dot6 = "@@@@@@@@@@.000000"
    '        F8dot2 = "@@@@@@@@.00"
    '        I10 = "@@@@@@@@@@"

    '        iSteps = Int(NumberTimeSteps * Timestep / iTimeStep)

    '        'Display the Save as dialog box.
    '        If SaveBox.ShowDialog() = DialogResult.OK Then
    '            If SaveBox.CheckFileExists = False Then
    '                'create the file
    '                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
    '            End If
    '            myStream = SaveBox.OpenFile()

    '            If (myStream IsNot Nothing) Then

    '                DataFile = SaveBox.FileName
    '                myStream.Close()

    '                FileOpen(1, DataFile, OpenMode.Output)

    '                WriteLine(1, " PHYSICAL AND MECHANICAL PROPERTIES OF GLASS ")
    '                WriteLine(1, " 1.Thermal conductivity [W/mK]=   " & VB6.Format(GLASSconductivity(idr, idc, idv), E12dot4))
    '                WriteLine(1, " 2.Thermal diffusivity [m^2/s]=   " & VB6.Format(GLASSalpha(idr, idc, idv), E12dot4))
    '                WriteLine(1, " 3.Absorption length [m]=   .1000E-02")
    '                WriteLine(1, " 4.Breaking stress [N/m^2]=   " & VB6.Format(GLASSbreakingstress(idr, idc, idv) * 1000000.0#, E12dot4))
    '                WriteLine(1, " 5.Youngs modulus [N/m^2]=   " & VB6.Format(GlassYoungsModulus(idr, idc, idv) * 1000000.0#, E12dot4))
    '                WriteLine(1, " 6.Linear coefficient of expansion [/deg C]=   " & VB6.Format(GLASSexpansion(idr, idc, idv), E12dot4))

    '                WriteLine(1, " GEOMETRY ")
    '                WriteLine(1, " 1.Glass thickness [m]=   " & VB6.Format(GLASSthickness(idr, idc, idv) / 1000, ".0000"))
    '                WriteLine(1, " 2.Shading thickness [m]=   " & VB6.Format(GLASSshading(idr, idc, idv) / 1000, ".0000"))
    '                WriteLine(1, " 3.Half-width [m]=   " & VB6.Format(dblHalfWidth, ".0000"))

    '                WriteLine(1, " COEFFICIENTS")
    '                WriteLine(1, " 1.Heat transfer coeff, unexposed [W/m^2-K]= " & VB6.Format(10, "@@@.00"))
    '                WriteLine(1, " 2.Ambient temp, unexposed [K]=  " & VB6.Format(ExteriorTemp, "@@@.0"))
    '                WriteLine(1, " 3.Emissivity of glass =  1.00")
    '                WriteLine(1, " 4.Emissivity of ambient (unexposed) =  1.00")

    '                WriteLine(1, " FLAME RADIATION")
    '                WriteLine(1, " Number of points used for flux input:  2")
    '                WriteLine(1, "     point #     time [s]             flux [W/m^2]")
    '                WriteLine(1, "         1             .00                   .00")
    '                WriteLine(1, "         2         1000.00                   .00")

    '                WriteLine(1, " GAS TEMPERATURE")
    '                WriteLine(1, " Number of points used for temperature input: " & iSteps)
    '                WriteLine(1, "     point #     time [s]             temperature [K]")
    '                WriteLine(1, VB6.Format(1, I10) & " " & VB6.Format(0, F12dot2) & "       " & VB6.Format(System.Math.Round(GLASSOtherHistory(idr, idc, idv, 1, 1)), F12dot2))

    '                For count = 1 To (iSteps - 1)
    '                    If count * iTimeStep <= UBound(GLASSOtherHistory, 5) Then
    '                        WriteLine(1, VB6.Format(count + 1, I10) & " " & VB6.Format(count * iTimeStep, F12dot2) & "       " & VB6.Format(System.Math.Round(GLASSOtherHistory(idr, idc, idv, 1, count * iTimeStep)), F12dot2))
    '                    End If
    '                Next count

    '                WriteLine(1, " HEAT TRANSFER COEFF. ON HOT LAYER SIDE")
    '                WriteLine(1, " Number of points used for heat transfer coeff input: " & iSteps)
    '                WriteLine(1, "     point #     time [s]             h2 [W/m^2-K]")
    '                WriteLine(1, VB6.Format(1, I10) & " " & VB6.Format(0, F12dot2) & "       " & VB6.Format(fnHinterior(GLASSOtherHistory(idr, idc, idv, 1, 1)), F12dot2))

    '                For count = 1 To (iSteps - 1)
    '                    If count * iTimeStep <= UBound(GLASSOtherHistory, 5) Then
    '                        WriteLine(1, VB6.Format(count + 1, I10) & " " & VB6.Format(count * iTimeStep, F12dot2) & "       " & VB6.Format(System.Math.Round(fnHinterior(GLASSOtherHistory(idr, idc, idv, 1, count * iTimeStep))), F12dot2))
    '                    End If
    '                Next count

    '                WriteLine(1, " EMISSIVITY OF HOT LAYER")
    '                WriteLine(1, " Number of points used for emissivity input: " & iSteps)
    '                WriteLine(1, "     point #     time [s]             emissivity")
    '                WriteLine(1, VB6.Format(1, I10) & " " & VB6.Format(0, F12dot2) & "                  " & VB6.Format(GLASSOtherHistory(idr, idc, idv, 2, 1), "0.00"))
    '                For count = 1 To (iSteps - 1)
    '                    If count * iTimeStep <= UBound(GLASSOtherHistory, 5) Then
    '                        WriteLine(1, VB6.Format(count + 1, I10) & " " & VB6.Format(count * iTimeStep, F12dot2) & "                  " & VB6.Format(GLASSOtherHistory(idr, idc, idv, 2, count * iTimeStep), "0.00"))
    '                    End If
    '                Next count

    '                WriteLine(1, " NUMERICAL PARAMETERS")
    '                WriteLine(1, " 1.Maximum fractional error in soln=   .000100")
    '                WriteLine(1, " 2.Size of time step [s]=  1.000")
    '                WriteLine(1, " 3.Maximum run time [s]=  250.00")
    '                WriteLine(1, " 4.Time interval for output [s]=  10.00")
    '                FileClose(1)
    '            End If
    '        End If
    '        MsgBox("Data saved in " & DataFile, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
    '        Exit Sub

    'more:
    '        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
    '        Exit Sub

    '    End Sub

    Public Sub mnuCancelBatch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCancelBatch.Click
        'cancel out of batch file runs
        flagstop = 1
        batch = False
        mnuRun.Enabled = True
        'Me.Toolbar2.Items.Item(8).Visible = True 'start run
        'Me.Toolbar2.Items.Item(9).Visible = False 'stop run
        'Me.Toolbar2.Items.Item(19).Visible = True 'graphs
    End Sub

    Public Sub mnuCeilingHoG_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCeilingHoG.Click
        '*  ======================================================
        '*  Show a graph of the Cone HRR input curve
        '*  ======================================================

        Dim result As Object

        Dim room As Integer

        If NumberRooms > 1 Then
            result = InputBox("Which room do you want?")
            If IsNumeric(result) Then
                If CInt(result) > 0 And CInt(result) <= NumberRooms Then
                    room = CInt(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        If CeilingConeDataFile(room) = "" Then CeilingConeDataFile(room) = "null.txt"
        Call Flame_Spread_Props_graph(room, 5, ceilinghigh, CeilingConeDataFile(room), Surface_Emissivity(1, room), ThermalInertiaCeiling(room), IgTempC(room), CeilingEffectiveHeatofCombustion(room), CeilingHeatofGasification(room), AreaUnderCeilingCurve(room), CeilingFTP(room), CeilingQCrit(room), Ceilingn(room))

    End Sub

    Public Sub mnuCeilingIgn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCeilingIgn.Click
        ''*  ======================================================
        ''*  Show a graph of the Cone HRR input curve
        ''*  ======================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim result As Object
        Dim DataMultiplier As Single
        Dim GraphStyle As Short
        Dim room As Integer

        'define variables
        Title = "Ceiling Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        If NumberRooms > 1 Then
            result = InputBox("Which room do you want?")
            If IsNumeric(result) Then
                If CInt(result) > 0 And CInt(result) <= NumberRooms Then
                    room = CInt(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        'If CeilingConeDataFile(room) = "" Then CeilingConeDataFile(room) = "null.txt"
        Call Flame_Spread_Props_graph(room, 2, ceilinghigh, CeilingConeDataFile(room), Surface_Emissivity(1, room), ThermalInertiaCeiling(room), IgTempC(room), CeilingEffectiveHeatofCombustion(room), CeilingHeatofGasification(room), AreaUnderCeilingCurve(room), CeilingFTP(room), CeilingQCrit(room), Ceilingn(room))

    End Sub

    Public Sub mnuCeilingJetTemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCeilingJetTemp.Click

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        If DetectorType(fireroom) = 0 Then
            MsgBox("You need to have specified a thermal detector or sprinkler in the fire room to record ceiling jet temperatures")
            Exit Sub
        End If
        'define variables
        Title = "Ceiling Jet Temp at Link (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data(1, Title, CJetTemp, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_3D_CJET(1, Title, CJetTemp, DataShift, DataMultiplier)
    End Sub

    Public Sub mnuCeilingTempGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCeilingTempGraph.Click
        '*  =======================================================
        '*  Show a graph of the ceiling temperature versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Ceiling Temp (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data
        Graph_Data_2D(Title, CeilingTemp, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuCO2Lower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCO2Lower.Click
        '*  =======================================================
        '*  Show a graph of the lower layer CO2 concentration versus time.
        '*  =======================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single


        'define variables
        Title = "Lower CO2 (%)"
        DataShift = 0
        DataMultiplier = 100
        MaxYValue = 0

        'call procedure to plot data
        '1 = upper layer
        Graph_Data_Species(2, Title, CO2VolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuCO2Upper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCO2Upper.Click
        '*  =======================================================
        '*  Show a graph of the upper layer CO2 concentration versus time.
        '*  =======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        'define variables
        Title = "Upper CO2 (%)"
        DataShift = 0
        DataMultiplier = 100


        'call procedure to plot data
        '1 = upper layer
        Graph_Data_Species(1, Title, CO2VolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuCOLower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCOLower.Click
        '*  ======================================================
        '*  Show a graph of the lower layer CO concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        'define variables
        Title = "Lower CO (ppm)"
        DataShift = 0
        DataMultiplier = 1000000


        'call procedure to plot data
        Graph_Data_Species(2, Title, COVolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuConeHRRceiling_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConeHRRceiling.Click
        '*  ======================================================
        '*  Show a graph of the Cone HRR input curve
        '*  14/8/98 C A Wade
        '*  ======================================================

        Dim Title As String, DataShift As Single, MaxYValue As Single, result As Object
        Dim DataMultiplier As Single, GraphStyle As Integer, room As Long
        '
        '    frmBlank.Show
        '
        '    'define variables
        Title = "Ceiling Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4          '2=user-defined
        MaxYValue = 0

        If NumberRooms > 1 Then
            result = InputBox("Which room to do want?")
            If IsNumeric(result) Then
                If CLng(result) > 0 And CLng(result) <= NumberRooms Then
                    room = CLng(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        'call procedure to plot data
        If CeilingConeDataFile(room) <> "null.txt" Then Graph_Cone_Data(room, ConeYC, Title, ConeXC, DataShift, DataMultiplier, GraphStyle, MaxYValue, ConeNumber_C)

    End Sub

    Public Sub mnuConeHRRfloor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConeHRRfloor.Click
        '*  ======================================================
        '*  Show a graph of the Cone HRR input curve
        '*  14/8/98 C A Wade
        '*  ======================================================

        Dim Title As String, DataShift As Single, MaxYValue As Single, result As Object
        Dim DataMultiplier As Single, GraphStyle As Integer, room As Long
        '
        '    frmBlank.Show
        '
        'define variables
        Title = "Floor Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4          '2=user-defined
        MaxYValue = 0

        If NumberRooms > 1 Then
            result = InputBox("Which room do you want?")
            If IsNumeric(result) Then
                If CLng(result) > 0 And CLng(result) <= NumberRooms Then
                    room = CLng(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        'call procedure to plot data
        If FloorConeDataFile(room) <> "null.txt" Then Graph_Cone_Data(room, ConeYF, Title, ConeXF, DataShift, DataMultiplier, GraphStyle, MaxYValue, ConeNumber_F)

    End Sub

    Public Sub mnuConeHRRwall_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuConeHRRwall.Click

        Dim Title As String, DataShift As Single, MaxYValue As Single, result As String
        Dim DataMultiplier As Single, GraphStyle As Integer, room As Long

        'define variables
        Title = "Wall Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4          '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        If NumberRooms > 1 Then
            result = InputBox("Which room to do want?")
            If IsNumeric(result) Then
                If CLng(result) > 0 And CLng(result) <= NumberRooms Then
                    room = CLng(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If
        If WallConeDataFile(room) <> "null.txt" Then Graph_Cone_Data(room, ConeYW, Title, ConeXW, DataShift, DataMultiplier, GraphStyle, MaxYValue, ConeNumber_W)

    End Sub

    Public Sub mnuContinue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuContinue.Click

        'UPGRADE_ISSUE: Form property frmBlank.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'frmBlank.Cursor = HOURGLASS

    End Sub

    Public Sub mnuCOUpper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCOUpper.Click
        '*  ======================================================
        '*  Show a graph of the upper layer CO concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single
        'define variables
        Title = "Upper CO (ppm)"
        DataShift = 0
        DataMultiplier = 1000000

        'call procedure to plot data
        Graph_Data_Species(1, Title, COVolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuCreateConeFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCreateConeFile.Click

    End Sub

    Public Sub mnuCSV_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuCSV.Click
        '===========================================================
        '   Saves results directly in a .csv file
        '===========================================================

        On Error GoTo more

        Dim s, Txt As String
        Dim room As Integer
        Dim k, j As Integer
        Dim count As Integer = 0
        Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()

        'dialogbox title
        SaveBox.Title = "Save Results to Comma Delimited File"

        'Set filters
        SaveBox.Filter = "All Files (*.*)|*.*|csv Files (*.csv)|*.csv"

        'Specify default filter
        SaveBox.FilterIndex = 2
        SaveBox.RestoreDirectory = True

        'default filename extension
        SaveBox.DefaultExt = "csv"

        'default filename
        'If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        If DataDirectory = "" Then DataDirectory = UserPersonalDataFolder & gcs_folder_ext & "\" & "data\"
        SaveBox.InitialDirectory = DataDirectory

        'default filename
        If Len(DataFile) > 4 Then
            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 3) & "csv"
        Else
            DataFile = "new.csv"
            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 3) & "csv"
        End If

        'Display the Save as dialog box.
        If SaveBox.ShowDialog() = DialogResult.OK Then
            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
            End If
            myStream = SaveBox.OpenFile()

            If (myStream IsNot Nothing) Then

                DataFile = SaveBox.FileName
                myStream.Close()

                'put stuff here to save from below 

                FileOpen(1, DataFile, OpenMode.Output)

                'define a format string
                's = "Scientific"
                s = "General Number"

                For room = 1 To NumberRooms
                    Txt = "room " & CStr(room)
                    Txt = Txt & " Time (sec)" & ",Layer (m)" & ",Upper Layer Temp (K)" & ",HRR (kW)" & ",Mass (kg/s)" & ",Plume (kg/s)"
                    Txt = Txt & ",Vent Fire (kW)" & ",CO2 Upper(%)" & ",CO Upper (%)" & ",O2 Upper (%)" & ",CO2 Lower(%)" & ",CO Lower(%)"
                    Txt = Txt & ",O2 Lower (%)" & ",FED gases(inc)" & ",Upper Wall Temp (K)" & ",Ceiling Temp (K)" & ",Rad on Floor (kW/m2)" & ",Lower Layer Temp (K)"
                    Txt = Txt & ",Lower Wall Temp (K)" & ",Floor Temp (K)" & ",Y Pyrolysis Front (m)" & ",X Pyrolysis Front (m)" & ",Z Pyrolysis Front (m)" & ",Upward Velocity (m/s)"
                    Txt = Txt & ",Lateral Velocity (m/s)" & ",Pressure (Pa)" & ",Visibility (m)" & ",Vent Flow to Upper Layer (kg/s)" & ",Vent Flow to Lower Layer (kg/s)"
                    Txt = Txt & ",Rad on Target (kW/m2)" & ",FED thermal(inc)" & ",OD upper (1/m)" & ",OD lower (1/m)" & Chr(10)
                    WriteLine(1, Txt)
                    If NumberTimeSteps > 0 Then
                        'not needed for csv file
                        'Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 29 > 32000
                        '    ExcelInterval = ExcelInterval * 2
                        'Loop
                        For j = 1 To NumberTimeSteps
                            If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                                Txt = VB6.Format(tim(j, 1), s) & "," & VB6.Format(layerheight(room, j), s) & "," & VB6.Format(uppertemp(room, j) - 273, s)
                                Txt = Txt & "," & VB6.Format(HeatRelease(room, j, 2), s) & "," & VB6.Format(FuelMassLossRate(j, 1), s) & "," & VB6.Format(massplumeflow(j, fireroom), s)
                                Txt = Txt & "," & VB6.Format(ventfire(room, j), s) & "," & VB6.Format(CO2VolumeFraction(room, j, 1) * 100, s) & "," & VB6.Format(COVolumeFraction(room, j, 1) * 100, s)
                                Txt = Txt & "," & VB6.Format(O2VolumeFraction(room, j, 1) * 100, s) & "," & VB6.Format(CO2VolumeFraction(room, j, 2) * 100, s) & "," & VB6.Format(COVolumeFraction(room, j, 2) * 100, s)
                                Txt = Txt & "," & VB6.Format(O2VolumeFraction(room, j, 2) * 100, s) & "," & VB6.Format(FEDSum(room, j), s) & "," & VB6.Format(Upperwalltemp(room, j) - 273, s)
                                Txt = Txt & "," & VB6.Format(CeilingTemp(room, j) - 273, s) & "," & VB6.Format(Target(room, j), s) & "," & VB6.Format(lowertemp(room, j) - 273, s)
                                Txt = Txt & "," & VB6.Format(LowerWallTemp(room, j) - 273, s) & "," & VB6.Format(FloorTemp(room, j) - 273, s)
                                If room = fireroom Then
                                    Txt = Txt & "," & VB6.Format(Y_pyrolysis(room, j), s) & "," & VB6.Format(X_pyrolysis(room, j), s) & "," & VB6.Format(Z_pyrolysis(room, j), s)
                                Else
                                    Txt = Txt & "," & VB6.Format(Y_pyrolysis(room, j), s) & "," & VB6.Format(0, s) & "," & VB6.Format(0, s)
                                End If
                                Txt = Txt & "," & VB6.Format(FlameVelocity(room, 1, j), s) & "," & VB6.Format(FlameVelocity(room, 2, j), s)
                                Txt = Txt & "," & VB6.Format(RoomPressure(room, j), s) & "," & VB6.Format(Visibility(room, j), s)
                                Txt = Txt & "," & VB6.Format(FlowToUpper(room, j), s) & "," & VB6.Format(FlowToLower(room, j), s)
                                Txt = Txt & "," & VB6.Format(SurfaceRad(room, j), s) & "," & VB6.Format(FEDRadSum(room, j), s)
                                Txt = Txt & "," & VB6.Format(OD_upper(room, j), s) & "," & VB6.Format(OD_lower(room, j), s)
                                WriteLine(1, Txt)
                                Txt = ""
                                System.Windows.Forms.Application.DoEvents()
                            End If
                        Next j
                    End If
                Next room
                ' the outside
                Txt = "outside "
                Txt = Txt & " Time (sec), Vent Fire (kW)"
                WriteLine(1, Txt)
                If NumberTimeSteps > 0 Then
                    For j = 1 To NumberTimeSteps
                        If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                            Txt = VB6.Format(tim(j, 1), s)
                            Txt = Txt & "," & VB6.Format(ventfire(room, j), s)
                            WriteLine(1, Txt)
                        End If
                    Next j
                End If
                FileClose(1)
            Else
                myStream.Close()
            End If
        End If

        MsgBox("Data saved in " & SaveBox.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        Exit Sub

excelerrorhandler:
        'User pressed Cancel button
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        On Error Resume Next

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        Exit Sub

    End Sub

    Public Sub mnuDatabase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '******************************************************
        '*
        '*  Show heat release database screen

        '******************************************************

        'frmBlank.Show()
        'frmFires.Show()
    End Sub


    Public Sub mnuDataFolder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.ShowDialog()
        ProjectDirectory = FolderBrowserDialog1.SelectedPath

        If ProjectDirectory = "" Then Exit Sub

        DefaultRiskDataDirectory = ProjectDirectory & "\riskdata\"
        Call Save_Registry()
        MsgBox("New default B-RISK riskdata folder is " & DefaultRiskDataDirectory)
        frmInputs.Label4.Text = "Default riskdata folder is: " & DefaultRiskDataDirectory

    End Sub

    Public Sub mnuDetectorGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDetectorGraph.Click
        '*  =======================================================
        '*  Show a graph of the detector/sprinkler link versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        'define variables
        Title = "Link Temp (C)"
        DataShift = -273
        DataMultiplier = 1

        'call procedure to plot data
        Graph_Data_2Dsprink(Title, SprinkTemp, DataShift, DataMultiplier)
    End Sub

    Public Sub mnuDmpFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDmpFile.Click
        '*  ==========================================================
        '*  This subprocedure shows a custom dialog box for
        '*  opening a file.
        '*  ==========================================================

        Dim Openbox As System.Windows.Forms.Control

        On Error GoTo errhandler

        ChDir((My.Application.Info.DirectoryPath))

        Exit Sub

errhandler:

        FileClose(1)
        Exit Sub

    End Sub

    Public Sub mnuEULA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEULA.Click
        '=============================================
        ' view EULA
        '=============================================

        ' ApplicationPath = IO.Path.GetDirectoryName(Application.ExecutablePath)
        'UserAppDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        'CommonFilesFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
        ' DocsFolder = ApplicationPath & "\docs\"
        'frmEula.RichTextBox1.LoadFile(UserAppDataFolder & gcs_folder_ext & "\docs\eula.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
        frmEula.RichTextBox1.LoadFile(DocsFolder & "eula.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
        frmEula.Text = "B-RISK End User License Agreement"
        'Me.Enabled = False
        frmEula.Show()

    End Sub

    Public Sub mnuExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExcel.Click
        '===========================================================
        '   Saves results directly in an excel file
        '===========================================================

        Dim s As String
        Dim room As Integer
        Dim k, j, count As Integer
        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        Try

            Dim SaveBox As New SaveFileDialog()

            'dialogbox title
            SaveBox.Title = "Save Results to Excel File"

            'Set filters
            SaveBox.Filter = "All Files (*.*)|*.*|xlsx Files (*.xlsx)|*.xlsx|xls Files (*.xls)|*.xls"

            'Specify default filter
            SaveBox.FilterIndex = 2
            SaveBox.RestoreDirectory = True

            'default filename extension
            'SaveBox.DefaultExt = "xlsx"

            'default filename
            DataDirectory = ProjectDirectory
            If DataDirectory = "" Then DataDirectory = UserPersonalDataFolder & gcs_folder_ext & "\" & "data\"
            SaveBox.InitialDirectory = DataDirectory

            'default filename
            If Len(DataFile) > 0 Then
                SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 4)
            Else
                MsgBox("No data, you need to load iteration first")
                Exit Sub

            End If

            count = 0

            '========================================

            'Create an array
            Dim DataArray(0 To (NumberTimeSteps * Timestep / ExcelInterval + 1), 0 To 65) As Object

            'On Error GoTo excelerrorhandler
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

            'Start a new workbook in Excel
            oExcel = CreateObject("Excel.Application")

            oBook = oExcel.Workbooks.Add

            'Add headers to the worksheet on row 1
            oSheet = oBook.Worksheets(1)

            oSheet.name = "Room 1"

            'assign values to excel cells
            'define a format string
            's = "Scientific"
            s = "0.000E+00"

            Dim rowcount As Short

            For room = 1 To NumberRooms
                'i = 1
                'For k = 0 To 10
                DataArray(0, 0) = "Time (sec)"
                DataArray(0, 1) = "Layer (m)"
                DataArray(0, 2) = "Upper Layer Temp (C)"
                DataArray(0, 3) = "HRR (kW)"
                DataArray(0, 4) = "Mass Loss Rate (kg/s)"
                DataArray(0, 5) = "Plume (kg/s)"
                DataArray(0, 6) = "Vent Fire (kW)"
                DataArray(0, 7) = "CO2 Upper(%)"
                DataArray(0, 8) = "CO Upper (ppm)"
                DataArray(0, 9) = "O2 Upper (%)"
                DataArray(0, 10) = "CO2 Lower(%)"
                DataArray(0, 11) = "CO Lower(ppm)"
                DataArray(0, 12) = "O2 Lower (%)"
                DataArray(0, 13) = "FED gases(inc)"
                DataArray(0, 14) = "Upper Wall Temp (C)"
                DataArray(0, 15) = "Ceiling Temp (C)"
                DataArray(0, 16) = "Rad on Floor (kW/m2)"
                DataArray(0, 17) = "Lower Layer Temp (C)"
                DataArray(0, 18) = "Lower Wall Temp (C)"
                DataArray(0, 19) = "Floor Temp (C)"
                DataArray(0, 20) = "Y Pyrolysis Front (m)"
                DataArray(0, 21) = "X Pyrolysis Front (m)"
                DataArray(0, 22) = "Z Pyrolysis Front (m)"
                DataArray(0, 23) = "Upward Velocity (m/s)"
                DataArray(0, 24) = "Lateral Velocity (m/s)"
                DataArray(0, 25) = "Pressure (Pa)"
                DataArray(0, 26) = "Visibility (m)"
                DataArray(0, 27) = "Vent Flow to Upper Layer (kg/s)"
                DataArray(0, 28) = "Vent Flow to Lower Layer (kg/s)"
                DataArray(0, 29) = "Rad on Target (kW/m2)"
                DataArray(0, 30) = "FED thermal(inc)"
                DataArray(0, 31) = "OD upper (1/m)"
                DataArray(0, 32) = "OD lower (1/m)"
                DataArray(0, 33) = "k upper (1/m)"
                DataArray(0, 34) = "k lower (1/m)"
                DataArray(0, 35) = "Vent Flow to Outside (m3/s)"
                DataArray(0, 36) = "HCN Upper (ppm)"
                DataArray(0, 37) = "HCN Lower (ppm)"
                DataArray(0, 38) = "SPR (m2/kg)"
                DataArray(0, 39) = "Unexposed Upper Wall Temp (C)"
                DataArray(0, 40) = "Unexposed Lower Wall Temp (C)"
                DataArray(0, 41) = "Unexposed Ceiling Temp (C)"
                DataArray(0, 42) = "Unexposed Floor Temp (C)"

                'Next
                If room = fireroom Then
                    DataArray(0, 43) = "Ceiling Jet Temp(C)"
                    DataArray(0, 44) = "Max Ceil Jet Temp(C)"
                    DataArray(0, 45) = "GER"
                    DataArray(0, 46) = "Smoke Detect OD-out (1/m)"
                    DataArray(0, 47) = "Smoke Detect OD-in (1/m)"
                    DataArray(0, 48) = "Link Temp (C)"

                    DataArray(0, 54) = "MLR free burn (kg/s)"
                    DataArray(0, 55) = "MLR ventilation effect (kg/s)"
                    DataArray(0, 56) = "MLR thermal effect (kg/s)"
                    DataArray(0, 57) = "MLR total (kg/s)"
                    DataArray(0, 58) = "Burning rate (kg/s)"
                    DataArray(0, 59) = "Fuel area shrinkage ratio (-)"
                    DataArray(0, 60) = "Unconstrained HRR (kW)"
                End If

                DataArray(0, 49) = "Normalised Heat Load (s1/2K)"
                DataArray(0, 50) = "Net ceiling flux (kW/m2)"
                DataArray(0, 51) = "Net upper wall flux (kW/m2)"
                DataArray(0, 52) = "Net lower Wall flux (kW/m2)"
                DataArray(0, 53) = "Net floor flux (kW/m2)"

                If useCLTmodel = True Then
                    DataArray(0, 61) = "Incident ceiling radiant flux (kW/m2)"
                    DataArray(0, 62) = "Incident upper wall radiant flux (kW/m2)"
                    DataArray(0, 63) = "Incident lower Wall radiant flux (kW/m2)"
                    DataArray(0, 64) = "Incident floor radiant flux (kW/m2)"

                End If
                DataArray(0, 65) = "Total fuel mass loss (kg)"
                If NumberTimeSteps > 0 Then

                    Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 59 > 32000 'the maximum number of data points able to be plotted in excel chart
                        ExcelInterval = ExcelInterval * 2
                    Loop

                    k = 2 'row
                    For j = 1 To NumberTimeSteps + 1
                        count = count + 1
                        If System.Math.Round(Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval, 4) = 0 Then
                            Me.ToolStripStatusLabel3.Text = "Saving to Excel ... Please Wait - " & Format(count / (NumberRooms * NumberTimeSteps) * 100, "0") & "%"
                            DataArray(k - 1, 0) = Format(tim(j, 1), "general number")
                            DataArray(k - 1, 1) = Format(layerheight(room, j), s)
                            DataArray(k - 1, 2) = Format(uppertemp(room, j) - 273, s)
                            DataArray(k - 1, 3) = Format(HeatRelease(room, j, 2), s)
                            DataArray(k - 1, 4) = Format(FuelMassLossRate(j, 1), s)
                            DataArray(k - 1, 5) = Format(massplumeflow(j, room), s)
                            DataArray(k - 1, 6) = Format(ventfire(room, j), s)
                            DataArray(k - 1, 7) = Format(CO2VolumeFraction(room, j, 1) * 100, s)
                            DataArray(k - 1, 8) = Format(COVolumeFraction(room, j, 1) * 1000000, s)
                            DataArray(k - 1, 9) = VB6.Format(O2VolumeFraction(room, j, 1) * 100, s)
                            DataArray(k - 1, 10) = VB6.Format(CO2VolumeFraction(room, j, 2) * 100, s)
                            DataArray(k - 1, 11) = VB6.Format(COVolumeFraction(room, j, 2) * 1000000, s)
                            DataArray(k - 1, 12) = VB6.Format(O2VolumeFraction(room, j, 2) * 100, s)
                            DataArray(k - 1, 13) = VB6.Format(FEDSum(room, j), s)
                            DataArray(k - 1, 14) = VB6.Format(Upperwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 15) = VB6.Format(CeilingTemp(room, j) - 273, s)
                            DataArray(k - 1, 16) = VB6.Format(Target(room, j), s)
                            DataArray(k - 1, 17) = VB6.Format(lowertemp(room, j) - 273, s)
                            DataArray(k - 1, 18) = VB6.Format(LowerWallTemp(room, j) - 273, s)
                            DataArray(k - 1, 19) = VB6.Format(FloorTemp(room, j) - 273, s)
                            DataArray(k - 1, 20) = VB6.Format(Y_pyrolysis(room, j), s)
                            DataArray(k - 1, 21) = VB6.Format(X_pyrolysis(room, j), s)
                            DataArray(k - 1, 22) = VB6.Format(Z_pyrolysis(room, j), s)
                            DataArray(k - 1, 23) = VB6.Format(FlameVelocity(room, 1, j), s)
                            DataArray(k - 1, 24) = VB6.Format(FlameVelocity(room, 2, j), s)
                            DataArray(k - 1, 25) = VB6.Format(RoomPressure(room, j), s)
                            DataArray(k - 1, 26) = VB6.Format(Visibility(room, j), s)
                            DataArray(k - 1, 27) = VB6.Format(FlowToUpper(room, j), s)
                            DataArray(k - 1, 28) = VB6.Format(FlowToLower(room, j), s)
                            DataArray(k - 1, 29) = VB6.Format(SurfaceRad(room, j), s)
                            DataArray(k - 1, 30) = VB6.Format(FEDRadSum(room, j), s)
                            DataArray(k - 1, 31) = VB6.Format(OD_upper(room, j), s)
                            DataArray(k - 1, 32) = VB6.Format(OD_lower(room, j), s)
                            DataArray(k - 1, 33) = VB6.Format(2.3 * OD_upper(room, j), s)
                            DataArray(k - 1, 34) = VB6.Format(2.3 * OD_lower(room, j), s)
                            DataArray(k - 1, 35) = VB6.Format(UFlowToOutside(room, j), s)
                            DataArray(k - 1, 36) = VB6.Format(HCNVolumeFraction(room, j, 1) * 1000000, s)
                            DataArray(k - 1, 37) = VB6.Format(HCNVolumeFraction(room, j, 2) * 1000000, s)
                            DataArray(k - 1, 38) = VB6.Format(SPR(room, j), s)
                            DataArray(k - 1, 39) = VB6.Format(UnexposedUpperwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 40) = VB6.Format(UnexposedLowerwalltemp(room, j) - 273, s)
                            DataArray(k - 1, 41) = VB6.Format(UnexposedCeilingtemp(room, j) - 273, s)
                            DataArray(k - 1, 42) = VB6.Format(UnexposedFloortemp(room, j) - 273, s)

                            DataArray(k - 1, 49) = VB6.Format(NHL(1, room, j), s)
                            DataArray(k - 1, 50) = VB6.Format(QCeiling(room, j), s)
                            DataArray(k - 1, 51) = VB6.Format(QUpperWall(room, j), s)
                            DataArray(k - 1, 52) = VB6.Format(QLowerWall(room, j), s)
                            DataArray(k - 1, 53) = VB6.Format(QFloor(room, j), s)

                            If room = fireroom Then
                                DataArray(k - 1, 43) = VB6.Format(CJetTemp(j, 1, 0) - 273, s)
                                DataArray(k - 1, 44) = VB6.Format(CJetTemp(j, 2, 0) - 273, s)
                                DataArray(k - 1, 45) = VB6.Format(GlobalER(j), s)

                                If NumSmokeDetectors > 0 Then
                                    DataArray(k - 1, 46) = VB6.Format(OD_outsideSD(1, j), s) 'for SD1 only
                                    DataArray(k - 1, 47) = VB6.Format(OD_insideSD(1, j), s) 'for SD1 only
                                End If

                                DataArray(k - 1, 48) = VB6.Format(LinkTemp(room, j) - 273, s)

                                DataArray(k - 1, 54) = VB6.Format(FuelBurningRate(0, room, j), s)
                                DataArray(k - 1, 55) = VB6.Format(FuelBurningRate(1, room, j), s)
                                DataArray(k - 1, 56) = VB6.Format(FuelBurningRate(2, room, j), s)
                                DataArray(k - 1, 57) = VB6.Format(FuelBurningRate(3, room, j), s)
                                DataArray(k - 1, 58) = VB6.Format(FuelBurningRate(4, room, j), s)
                                DataArray(k - 1, 59) = VB6.Format(FuelBurningRate(5, room, j), s)
                                DataArray(k - 1, 60) = Format(HeatRelease(room, j, 1), s)

                                If useCLTmodel = True Then
                                    DataArray(k - 1, 61) = Format(QCeilingAST(room, 0, j), s)
                                    DataArray(k - 1, 62) = Format(QUpperWallAST(room, 0, j), s)
                                    DataArray(k - 1, 63) = Format(QLowerWallAST(room, 0, j), s)
                                    DataArray(k - 1, 64) = Format(QFloorAST(room, 0, j), s)
                                End If

                                DataArray(k - 1, 65) = Format(TotalFuel(j), s)

                            End If

                                Application.DoEvents()
                            k = k + 1
                        End If
                    Next j

                    rowcount = k

                End If

                'Transfer the array to the worksheet starting at cell A1
                oSheet.Range("A1").Resize(rowcount - 1, 66).Value = DataArray

                If room < NumberRooms Then
                    oBook.Worksheets.add()
                    oSheet = oBook.ActiveSheet
                    oSheet.Name = "Room " & CStr(room + 1)
                End If
            Next

            oBook.Worksheets.add()
            oSheet = oBook.ActiveSheet
            oSheet.Name = "Outside"

            DataArray(0, 0) = "Time (sec)"
            DataArray(0, 1) = "Vent Fire (kW)"

            If NumberTimeSteps > 0 Then
                k = 2 'row
                For j = 1 To NumberTimeSteps + 1
                    If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                        DataArray(k - 1, 0) = Format(tim(j, 1), s)
                        DataArray(k - 1, 1) = Format(ventfire(room, j), s)
                        k = k + 1
                    End If
                Next j
            End If

            'Transfer the array to the worksheet starting at cell A1
            oSheet.Range("A1").Resize(k - 1, 2).Value = DataArray

            Me.ToolStripStatusLabel4.Text = "Saving Excel Charts... Please Wait" = "Saving Excel Charts... Please Wait"
            'If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,B:B", "Layer Height (m)", "B2", "A2:A" & CStr(rowcount), "B2:B" & CStr(rowcount))
            'If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,C:C", "Upper Layer Temp (C)", "C2", "A2:A" & CStr(rowcount), "C2:C" & CStr(rowcount))
            'If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,R:R", "Lower Layer Temp (C)", "R2", "A2:A" & CStr(rowcount), "R2:R" & CStr(rowcount))
            'If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,D:D", "Heat Release (kW)", "D2", "A2:A" & CStr(rowcount), "D2:D" & CStr(rowcount))
            'If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,E:E", "Mass Loss Rate (kg.s-1)", "E2", "A2:A" & CStr(rowcount), "E2:E" & CStr(rowcount))
            'If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,F:F", "Plume Flow (kg.s-1)", "F2", "A2:A" & CStr(rowcount), "F2:F" & CStr(rowcount))
            'If frmprintvar.chkFED.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,N:N", "FED gases", "N2", "A2:A" & CStr(rowcount), "N2:N" & CStr(rowcount))
            'If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AA:AA", "Visibility (m)", "AA2", "A2:A" & CStr(rowcount), "AA2:AA" & CStr(rowcount))
            'If frmprintvar.chkFEDrad.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AE:AE", "FED thermal", "AE2", "A2:A" & CStr(rowcount), "AE2:AE" & CStr(rowcount))
            'If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AF:AF", "Upper Layer OD (m-1)", "AF2", "A2:A" & CStr(rowcount), "AF2:AF" & CStr(rowcount))
            'If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AG:AG", "Lower Layer OD (m-1)", "AG2", "A2:A" & CStr(rowcount), "AG2:AG" & CStr(rowcount))
            'If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AU:AU", "Smoke Detector OD-OUTSIDE (m-1)", "AU2", "A2:A" & CStr(rowcount), "AU2:AU" & CStr(rowcount))
            'If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,AV:AV", "Smoke Detector OD-INSIDE (m-1)", "AV2", "A2:A" & CStr(rowcount), "AV2:AV" & CStr(rowcount))
            'If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,L:L", "Lower Layer CO (ppm)", "L2", "A2:A" & CStr(rowcount), "L2:L" & CStr(rowcount))
            'If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,I:I", "Upper Layer CO (ppm)", "I2", "A2:A" & CStr(rowcount), "I2:I" & CStr(rowcount))
            'If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(oExcel, "Room 1", "A:A,Z:Z", "Pressure (Pa)", "Z2", "A2:A" & CStr(rowcount), "Z2:Z" & CStr(rowcount))

            'save the worksheet
            'On Error Resume Next
            Dim fname As String = FileSystem.Dir(basefile)
            fname = Mid(fname, 1, Len(fname) - 4) & "_results"

            'Save the Workbook and Quit Excel
            oBook.SaveAs(DataDirectory & fname)
            oBook.Close(SaveChanges:=False)
            oExcel.Quit()

            If Err.Number = 1004 Then
                MsgBox(SaveBox.FileName & " is already Open. Please close and then try again.", MsgBoxStyle.OkOnly)
                Err.Clear()
            Else
                MsgBox("Data saved in " & DataDirectory & fname & ".xlsx", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            End If

            oExcel.Application.Quit()
            'release the objects
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing

            Me.ToolStripStatusLabel3.Text = ""
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
            Exit Sub


        Catch ex As Exception
            oExcel.Application.Quit()
            'release the objects
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing
        End Try

    End Sub

    Public Sub Add_ExcelChart(ByRef xlapp As Object, ByRef RoomName As String, ByRef exRange As String, ByRef Title As String, ByRef startcell As String, ByRef xrange As String, ByRef yrange As String)

        'add new chart to spreadsheet
        xlapp.Charts.Add()
        xlapp.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLinesNoMarkers
        xlapp.ActiveChart.SetSourceData(Source:=xlapp.Sheets(RoomName).Range(exRange), PlotBy:=Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
        xlapp.ActiveChart.Location(Where:=Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsNewSheet)
        With xlapp.ActiveChart
            .HasTitle = True
            .Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.characters.text = "Time (seconds)"
            .Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).HasTitle = True
            .Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary).AxisTitle.characters.text = Title
        End With
        xlapp.ActiveChart.PlotArea.Select()
        With xlapp.Selection.Border
            .ColorIndex = 57
            .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
            .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
        xlapp.Selection.Interior.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
        xlapp.ActiveChart.SeriesCollection(1).Select()
        With xlapp.Selection.Border
            .ColorIndex = 57
            .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
        With xlapp.Selection
            .MarkerBackgroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
            .MarkerForegroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
            .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
        xlapp.ActiveChart.SeriesCollection(1).points(3).Select()
        xlapp.ActiveChart.PlotArea.Select()
        xlapp.ActiveChart.SeriesCollection(1).Select()
        With xlapp.Selection.Border
            .ColorIndex = 57
            .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
            .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        End With
        With xlapp.Selection
            .MarkerBackgroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
            .MarkerForegroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
            .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
            .Smooth = False
            .MarkerSize = 3
            .Shadow = False
        End With
        xlapp.ActiveChart.PlotArea.Select()
        xlapp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).Select()
        xlapp.Selection.TickLabels.NumberFormat = "0"
        xlapp.ActiveChart.Legend.Select()
        xlapp.Selection.Delete()
        xlapp.ActiveChart.ChartTitle.Select()
        xlapp.Selection.Delete()
        xlapp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).AxisTitle.Select()
        xlapp.ActiveChart.ChartArea.Select()
        xlapp.Selection.AutoScaleFont = True
        With xlapp.Selection.Font
            .Name = "Arial"
            .FontStyle = "Regular"
            .Size = 12
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
            .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            .Background = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
        End With
        xlapp.ActiveChart.PlotArea.Select()
        xlapp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select()
        xlapp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory).AxisTitle.Select()
        xlapp.Selection.Font.Bold = True
        xlapp.ActiveChart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue).AxisTitle.Select()
        xlapp.Selection.Font.Bold = True
        xlapp.ActiveChart.PlotArea.Select()

        xlapp.ActiveWindow.Zoom = 100
        xlapp.ActiveSheet.Select()
        xlapp.ActiveSheet.Name = Title
        xlapp.ActiveChart.SeriesCollection(1).Select()
        If RoomDescription(1) <> "" Then
            xlapp.ActiveChart.SeriesCollection(1).Name = RoomDescription(1)
        Else
            xlapp.ActiveChart.SeriesCollection(1).Name = RoomName
        End If
        xlapp.ActiveChart.HasLegend = True
        xlapp.ActiveChart.Legend.Select()
        xlapp.Selection.Position = Microsoft.Office.Interop.Excel.Constants.xlTop

        Dim room As Short
        Dim index As Short
        If NumberRooms > 1 Then
            For room = 2 To NumberRooms
                RoomName = "Room " & CStr(room)

                xlapp.ActiveChart.PlotArea.Select()

                xlapp.ActiveChart.SeriesCollection.NewSeries()
                If RoomDescription(room) <> "" Then
                    xlapp.ActiveChart.SeriesCollection(room).Name = RoomDescription(room)
                Else
                    xlapp.ActiveChart.SeriesCollection(room).Name = RoomName
                End If

                index = xlapp.Worksheets.count - 2 - room + 1

                'need to change linetype before seriescollection method will accept values, bug in VB?
                xlapp.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine

                xlapp.ActiveChart.SeriesCollection(room).XValues = xlapp.Worksheets(index).Range(xrange)
                xlapp.ActiveChart.SeriesCollection(room).Values = xlapp.Worksheets(index).Range(yrange) 'old - wrong

                xlapp.ActiveChart.SeriesCollection(room).Select()
                With xlapp.Selection.Border
                    .ColorIndex = 57
                    .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                    .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                End With
                With xlapp.Selection
                    .MarkerBackgroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
                    .MarkerForegroundColorIndex = Microsoft.Office.Interop.Excel.Constants.xlNone
                    .MarkerStyle = Microsoft.Office.Interop.Excel.Constants.xlNone
                    .Smooth = False
                    .MarkerSize = 3
                    .Shadow = False
                End With

                'change the linetype back again
                xlapp.ActiveChart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLinesNoMarkers

            Next room
        End If

    End Sub

    Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
        '*  ==============================================================
        '*  This event exits the program, but asks if the user wants to
        '*  save the current model in a file'
        '*  ==============================================================

        Dim msg As String
        Dim answer As Short
        'Dim strTempPath As String

        msg = "Would you like to save file?"
        answer = MsgBox(msg, MB_ICONQUESTION + MB_YESNO, ProgramTitle)
        On Error GoTo ErrHandler1

        Dim i As Short
        If answer = IDYES Then
            'CancelError is True.
            On Error GoTo ErrHandler1

            Call savemodeldata()

            '	'Set filters
            '	'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            '          'frmBlank.CMDialog1Save.Filter = "All Files (*.*)|*.*|Model Files (*.mod)|*.mod"
            '	'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            '          'frmBlank.CMDialog1Save.FileName = Dir(DataFile)
            '	'Specify default filter
            '          'frmBlank.CMDialog1Save.FilterIndex = 2
            '          'frmBlank.CMDialog1Save.Title = "Save Data File"
            '	'Display the Save as dialog box
            '          'frmBlank.CMDialog1Save.ShowDialog()
            '	'Call the save file procedure
            '          'Save_File((LCase(frmBlank.CMDialog1Save.FileName)))

        ElseIf answer = IDNO Then

            Call Save_Registry()

            'delete temp files
            'strTempPath = getTempPathName() 'get name of windows temp folder

            'Kill((strTempPath & "BF*.tmp"))


            ' Loop through the forms collection and unload
            ' each form.
            For i = (My.Application.OpenForms.Count - 1) To 1 Step -1
                'UPGRADE_ISSUE: Unload Forms() was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
                'Unload(My.Application.OpenForms(i))
                My.Application.OpenForms(i).Close()
            Next i

            FileClose()
        End If

        Me.Close()
        Exit Sub

ErrHandler1:
        'User pressed Cancel button
        If Err.Number = 53 Then Resume Next
        Exit Sub

    End Sub

    Public Sub mnuExtcjet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExtcjet.Click
        '*  ===============================================================
        '*  Show a graph of OD in the ceiling jet at the detector versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Extinction Coef. at Smoke Detector (1/m)"
        DataShift = 0
        DataMultiplier = ParticleExtinction
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        'Graph_Data_2D(1, Title, SmokeConcentration, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, SmokeConcentration, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuExtl_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExtl.Click
        '*  ===============================================================
        '*  Show a graph of OD in the lower layer versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift As Single
        Dim DataMultiplier As Double

        'define variables
        Title = "Extinction Coef. lower (1/m)"
        DataShift = 0
        DataMultiplier = 2.3

        'call procedure to plot data
        Graph_Data_2D(Title, OD_lower, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuEXTu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEXTu.Click
        '*  ===============================================================
        '*  Show a graph of OD in the lower layer versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Extinction Coef. upper (1/m)"
        DataShift = 0
        DataMultiplier = 2.3
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data

        'Graph_Data_2D(1, Title, OD_upper, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, OD_upper, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFEDgases_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFEDgases.Click
        '*  =====================================================================
        '*  show a graph of the fractional effective dose for incapacitation
        '*  by carbon monoxide based on Stewart's equation and oxygen hypoxia.
        '*  An FED of 1 means incapacitation would result if someone were exposed
        '*  continuously to the upper layer.
        '*  =====================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double

        If FEDCO = True Then
            Title = "FED(CO) at " & CStr(MonitorHeight) & " m"
        Else
            Title = "Asphyxiant Gases FED (incap) at " & CStr(MonitorHeight) & " m"
        End If

        DataShift = 0
        DataMultiplier = 1
        MaxYValue = 1

        'call procedure to plot data
        Graph_Data_FED(Title, FEDSum, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuFEDradiation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFEDradiation.Click
        '*  =====================================================================
        '*
        '*  =====================================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        'define variables
        Title = "Thermal FED at " + CStr(MonitorHeight) + " m"
        DataShift = 0
        DataMultiplier = 1

        'call procedure to plot data

        Graph_Data_FED(Title, FEDRadSum, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFileOpen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFileOpen.Click
        '*  ==========================================================
        '*  This subprocedure shows a custom dialog box for
        '*  opening a file.
        '*  ==========================================================

        OpenFileDialog1.InitialDirectory = DataDirectory
        OpenFileDialog1.Title = "Open Data XML File"
        OpenFileDialog1.Filter = "All Files (*.*)|*.*|Model Files (*.xml)|*.xml"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.FileName = ""

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DataFile = OpenFileDialog1.FileName
        End If

        Call Read_File_xml(DataFile, False)

        DataDirectory = VB.Left(DataFile, Len(DataFile) - Len(Path.GetFileName(DataFile)))

        ChDir((My.Application.Info.DirectoryPath))

    End Sub

    Private Sub mnuFireParameters_Click()
        '*  ========================================================
        '*  This procedure updates the textboxes when the form
        '*  gets the focus.
        '*  ========================================================

        'frmBlank.Show()
        frmOptions1.Show()

        'frmOptions1.txtRadiantLossFraction.Text = CStr(RadiantLossFraction)
        'frmOptions1.txtMassLossUnitArea.Text = CStr(MLUnitArea)
        frmOptions1.txtEmissionCoefficient.Text = CStr(EmissionCoefficient)

    End Sub

    Public Sub mnuFloor_X_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFloor_X.Click
        '*  ===================================================================
        '*  Show a graph of the position of the X_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Floor X_Pyrolysis (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data

        Graph_Data_2D(Title, Xf_pyrolysis, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFloor_YB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFloor_YB.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = " Floor Y Burnout (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, Yf_burnout, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFloor_YP_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFloor_YP.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Floor Coverings
        '   17/4/2002 C Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Floor - Y Pyrolysis (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, Yf_pyrolysis, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFloorIgn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFloorIgn.Click
        ''*  ======================================================
        ''*  Show a graph of the Cone HRR input curve
        ''*  14/8/98 C A Wade
        ''*  ======================================================

        Dim Title, result As String

        Dim room As Integer

        'define variables
        Title = "Floor Cone HRR (kW/m2)"

        If NumberRooms > 1 Then
            result = InputBox("Which room do you want?")
            If IsNumeric(result) Then
                If CInt(result) > 0 And CInt(result) <= NumberRooms Then
                    room = CInt(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If


        If FloorConeDataFile(room) = "" Then FloorConeDataFile(room) = "null.txt"
        Call Flame_Spread_Props_graph(room, 3, floorhigh, FloorConeDataFile(room), Surface_Emissivity(4, room), ThermalInertiaFloor(room), IgTempF(room), FloorEffectiveHeatofCombustion(room), FloorHeatofGasification(room), AreaUnderFloorCurve(room), FloorFTP(room), FloorQCrit(room), Floorn(room))

    End Sub

    Public Sub mnuFlowOut_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFlowOut.Click
        '*  ===============================================================
        '*  Show a graph of vent flow
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Volume Flow to Outside (m3/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, UFlowToOutside, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFlowtoLower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFlowtoLower.Click
        '*  ===============================================================
        '*  Show a graph of vent flow
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Net Mass Flow to Lower Layer (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, FlowToLower, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuFlowtoUpper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuFlowtoUpper.Click
        '*  ===============================================================
        '*  Show a graph of vent flow
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Net Mass Flow to Upper Layer (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, FlowToUpper, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuGER_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGER.Click

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Global Equivalence Ratio"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        Graph_GER(1, Title, GlobalER, DataShift, DataMultiplier, GraphStyle, MaxYValue)

    End Sub

    Public Sub mnuGlassExcel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGlassExcel.Click
        '===========================================================
        '   Saves glass results directly in an excel file
        '   21/11/2001 C A Wade
        '===========================================================

        On Error GoTo more

        Dim s, Txt As String
        Dim room As Integer
        Dim vent, node As Short
        Dim count, j, k, P As Integer

        Dim SaveBox As System.Windows.Forms.Control

        count = 0


        'declare object variables for microsoft excel
        On Error Resume Next
        Dim xlapp As Microsoft.Office.Interop.Excel.Application
        Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

        'assign object references to the variables
        xlapp = CreateObject("Excel.Application")

        xlBook = xlapp.Workbooks.Add
        If Err.Number <> 0 Then
            GoTo more
        End If

        On Error GoTo excelerrorhandler
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        xlSheet = xlBook.Worksheets(1)
        xlSheet.Name = "Room 1"

        'assign values to excel cells
        'define a format string
        s = "Scientific"

        Dim dummy() As Double
        Dim X As Object
        For room = 1 To NumberRooms
            Txt = "Time (sec)" & ",Layer (m)" & ",Upper Layer Temp (C)" & ",Lower Layer Temp (C)" & ",HRR (kW)"
            For P = 2 To NumberRooms + 1
                For vent = 1 To NumberVents(room, P)
                    If AutoBreakGlass(room, P, vent) = True Then
                        Txt = Txt & ",Exposure Temp (C)" & ",Emissivity"
                        For node = 1 To UBound(GLASSTempHistory, 4)
                            Txt = Txt & ",Node " & CStr(node) & " (" & CStr(room) & CStr(P) & CStr(vent) & ")"
                        Next node
                    End If
                Next vent
            Next P
            xlSheet.Cells._Default(1, 1).Value = Txt

            If NumberTimeSteps > 0 Then


                Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 37 > 32000
                    ExcelInterval = ExcelInterval * 2
                Loop

                k = 2 'row
                For j = 1 To NumberTimeSteps
                    count = count + 1
                    If System.Math.Round(Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval, 4) = 0 Then
                        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        Me.ToolStripStatusLabel4.Text = "Saving to Excel ... Please Wait - " & VB6.Format(count / (NumberRooms * NumberTimeSteps) * 100, "0") & "%"

                        Txt = VB6.Format(tim(j, 1), s) & "," & VB6.Format(layerheight(room, j), s) & "," & VB6.Format(uppertemp(room, j) - 273, s)
                        Txt = Txt & "," & VB6.Format(lowertemp(room, j) - 273, s) & "," & VB6.Format(HeatRelease(room, j, 2), s)
                        For P = 2 To NumberRooms + 1
                            For vent = 1 To NumberVents(room, P)
                                If AutoBreakGlass(room, P, vent) = True Then
                                    If tim(j, 1) > VentBreakTime(room, P, vent) Then
                                        Txt = Txt & "," & ","
                                    Else
                                        Txt = Txt & "," & VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 1, j) - 273, s) & "," & VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 2, j), s)
                                        For node = 1 To UBound(GLASSTempHistory, 4)
                                            Txt = Txt & "," & VB6.Format(GLASSTempHistory(room, NumberRooms + 1, vent, node, j) - 273, s)
                                        Next node
                                    End If
                                End If
                            Next vent
                        Next P
                        xlSheet.Cells._Default(k, 1).Value = Txt
                        System.Windows.Forms.Application.DoEvents()
                        k = k + 1
                    End If
                Next j
            End If
            On Error Resume Next

            xlSheet.Columns._Default("A:A").Select()
            With xlapp.Selection
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlGeneral
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
                .WrapText = False
                .Orientation = 0
                .ShrinkToFit = False
                .MergeCells = False

            End With

            xlapp.Selection.TextToColumns(Destination:=xlSheet.Range("A1"), dataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab_Renamed:=True, Semicolon:=False, Comma:=True, Space_Renamed:=False, Other:=False)

            If room < NumberRooms Then
                xlBook.Worksheets.Add()
                xlSheet = xlBook.ActiveSheet
                xlSheet.Name = "Room " & CStr(room + 1)
            End If
        Next room

        'save the worksheet
        On Error Resume Next
        If Err.Number = 1004 Then
            Err.Clear()
        End If
        xlBook.Close(SaveChanges:=False)
        xlapp.Application.Quit()
        'release the objects
        xlapp = Nothing
        xlBook = Nothing
        xlSheet = Nothing

        Me.ToolStripStatusLabel4.Text = " "

        Exit Sub

excelerrorhandler:
        On Error Resume Next
        xlBook.Close()
        xlapp.Application.Quit()

        'release the objects
        xlapp = Nothing
        xlBook = Nothing
        xlSheet = Nothing
        Me.ToolStripStatusLabel4.Text = " "

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        'User pressed Cancel button
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        Exit Sub

    End Sub


    Public Sub mnuGlassTempGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGlassTempGraph.Click
        '*  ===============================================================
        '*  Show a graph of Nodal Glass Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer, idc As Integer, idv As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Title = "Display Nodal Temperatures for Which Vent?"   ' Set title.

        Message = "Enter the first room"   ' Set prompt.
        dDefault = "1"   ' Set default.
        ' Display message, title, and default value.
        MyValue = InputBox(Message, Title, dDefault)
        If Not IsNumeric(MyValue) Then
            Exit Sub
        End If
        idr = CInt(MyValue)
        '
        Message = "Enter the second room"   ' Set prompt.
        dDefault = "2"   ' Set default.
        ' Display message, title, and default value.
        MyValue = InputBox(Message, Title, dDefault)
        If Not IsNumeric(MyValue) Then
            Exit Sub
        End If
        idc = CInt(MyValue)

        Message = "Enter the vent ID"   ' Set prompt.
        dDefault = "1"   ' Set default.
        ' Display message, title, and default value.
        MyValue = InputBox(Message, Title, dDefault)
        idv = CInt(MyValue)

        If AutoBreakGlass(idr, idc, idv) = False Then
            MsgBox("Nodal Temperatures can only be plotted for vents with glass fracture modelled.")
            Exit Sub
        End If

        'define variables
        Title = "Nodal Glass Temperatures (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4            '2=user-defined
        MaxYValue = 0

        Dim NumberGlassNodes As Integer
        NumberGlassNodes = UBound(GLASSTempHistory, 4)

        Description = "Room " + CStr(idr) + " to Room " + CStr(idc) + " Vent " + CStr(idv)
        maxpoints = NumberGlassNodes * CInt(VentBreakTime(idr, idc, idv) / Timestep)
        If NumberTimeSteps < 2 Then Exit Sub
        numpoints = CInt(VentBreakTime(idr, idc, idv) / Timestep)
        numsets = NumberGlassNodes

        Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
        Dim curve As Integer

        curve = 1
        For i = 0 To 2 * numsets - 1 Step 2

            chdata(0, i) = "Node " & CStr(curve)

            For j = 1 To numpoints - 1
                chdata(j, i) = tim(j, 1)
                chdata(j, i + 1) = (DataMultiplier * GLASSTempHistory(idr, idc, idv, curve, j) + DataShift) 'data to be plotted
            Next

            curve = curve + 1
        Next i

        frmGraph.Text = Title
        frmGraph.Show()

    End Sub

    Public Sub mnuGraphFloorTemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphFloorTemp.Click
        '*  =======================================================
        '*  Show a graph of floor temperature versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        'define variables
        Title = "Floor Temp (C)"
        DataShift = -273
        DataMultiplier = 1

        Graph_Data_2D(Title, FloorTemp, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuGraphLowerLayerTemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphLowerLayerTemp.Click
        '*  =======================================================
        '*  Show a graph of lower layer temperature versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double

        Title = "Lower Layer Temp (C)"
        DataShift = -273
        DataMultiplier = 1

        Graph_Data_2D(Title, lowertemp, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuGraphLowerWallTemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphLowerWallTemp.Click
        '*  =======================================================
        '*  Show a graph of lower wall temperature versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double

        'define variables
        Title = "Lower Wall Temp (C)"
        DataShift = -273
        DataMultiplier = 1
        '    GraphStyle = 4                 '2=user-defined
        '    MaxYValue = 0
        '
        '    'call procedure to plot data
        Graph_Data_2D(Title, LowerWallTemp, DataShift, DataMultiplier, timeunit)

    End Sub



    Public Sub mnuGraphsOff_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphsOff.Click

        'frmOptions1.optRunTimeGraphON.Checked = False
        Me.mnuGraphsOn_HRR.Checked = False
        'frmOptions1.optRunTimeGraphOff.Checked = True
        Me.mnuGraphsOff.Checked = True

    End Sub

    Public Sub mnuGraphsOn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphsOn_HRR.Click

        'frmOptions1.optRunTimeGraphON.Checked = True
        Me.mnuGraphsOn_HRR.Checked = True
        'frmOptions1.optRunTimeGraphOff.Checked = False
        Me.mnuGraphsOff.Checked = False

    End Sub

    Public Sub mnuGraphsVisible_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuGraphsVisible.Click
        'Dim frmGraphs As Object
        'If mnuGraphsVisible.Checked = False Then
        '	'UPGRADE_WARNING: Couldn't resolve default property of object frmGraphs.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '          frmGraphs.Visible = True
        '	mnuGraphsVisible.Checked = True
        'Else
        '	'UPGRADE_WARNING: Couldn't resolve default property of object frmGraphs.Visible. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '	frmGraphs.Visible = False
        '	mnuGraphsVisible.Checked = False
        'End If
    End Sub

    Public Sub mnuH2OLower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuH2OLower.Click
        '*  ======================================================
        '*  Show a graph of the lower layer H2O concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        Title = "Lower H2O (%)"
        DataShift = 0
        DataMultiplier = 100

        Graph_Data_Species(2, Title, H2OVolumeFraction, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuH2OUpper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuH2OUpper.Click
        '*  ======================================================
        '*  Show a graph of the upper layer H2O concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        Title = "Upper H2O (%)"
        DataShift = 0
        DataMultiplier = 100

        Graph_Data_Species(1, Title, H2OVolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub



    Public Sub mnuHCNlower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHCNlower.Click
        '*  ======================================================
        '*  Show a graph of the lower layer hcn concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        Title = "Lower HCN (ppm)"
        DataShift = 0
        DataMultiplier = 1000000

        Graph_Data_Species(2, Title, HCNVolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuHCNupper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHCNupper.Click
        '*  ======================================================
        '*  Show a graph of the upper layer hcn concentration versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        Title = "Upper HCN (ppm)"
        DataShift = 0
        DataMultiplier = 1000000

        Graph_Data_Species(1, Title, HCNVolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub
    Public Sub mnuUnconstrainedHeat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUnconstrainedHeat.Click
        '*  ======================================================
        '*  Show a graph of the fire heat release rate versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        'define variables
        Title = "Unconstrained HRR (kW)"
        DataShift = 0
        DataMultiplier = 1
        'call procedure to plot data
        Graph_Data_Species(1, Title, HeatRelease, DataShift, DataMultiplier, timeunit)

    End Sub
    Public Sub mnuHeat_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHeat.Click
        '*  ======================================================
        '*  Show a graph of the fire heat release rate versus time.
        '*  ======================================================

        Dim Title As String, DataShift As Single
        Dim DataMultiplier As Single

        'define variables
        Title = "Heat Release Rate (kW)"
        DataShift = 0
        DataMultiplier = 1
        'call procedure to plot data
        Graph_Data_Species(2, Title, HeatRelease, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuHeatRelease_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        ''*  ========================================================
        ''*  This procedure updates the text boxes when the form
        ''*  gets the focus.
        ''*  ========================================================

        'Dim i As Integer
        'Dim id As Short

        'frmFire.Show()

        ''Set Room = frmFire.lstFireRoom2

        ''On Error Resume Next
        ''frmBlank.Show()
        ''frmFireSpecification.Show

        ''update list box
        ''frmFireSpecification.lstObjectID.Clear
        'frmFire.lstObjectID.Items.Clear()
        'If NumberObjects > 0 Then
        '    id = 1
        '    For i = 1 To NumberObjects
        '        frmFire.lstObjectID.Items.Add("Fire " & CStr(i) & " of " & CStr(NumberObjects))
        '    Next i
        'Else
        '    frmFire.lstObjectID.Items.Add("None")
        'End If
        'frmFire.lstObjectID.SelectedIndex = 0

        'frmItemList.lstFireRoom2.Items.Clear()
        'For i = 1 To NumberRooms
        '    frmItemList.lstFireRoom2.Items.Add(CStr(i))
        'Next i
        'frmItemList.lstFireRoom2.SelectedIndex = fireroom - 1


        'Dim fireobject As Object
        'fireobject = frmFire

        'fireobject.txtEnergyYield.Text = VB6.Format(EnergyYield(id), "0.0")
        'fireobject.txtCO2Yield.Text = VB6.Format(CO2Yield(id), "0.000")
        'fireobject.txtSootYield.Text = VB6.Format(SootYield(id), "0.000")
        'fireobject.txtHCNYield.Text = VB6.Format(HCNuserYield(id), "0.000")
        'fireobject.txtObj_CRF.Text = VB6.Format(ObjCRF(id), "0.0")
        'fireobject.txtObj_length.Text = VB6.Format(ObjLength(id), "0.000")
        'fireobject.txtObj_width.Text = VB6.Format(ObjWidth(id), "0.000")
        'fireobject.txtObj_height.Text = VB6.Format(ObjHeight(id), "0.000")
        'fireobject.txtObj_Elevation.Text = VB6.Format(ObjElevation(id), "0.000")
        ''frmFire.txtWaterVaporYield.text = Format(WaterVaporYield(id), "0.000")
        'fireobject.lblObjectDescription.Text = ObjectDescription(id)
        'fireobject.txtFireHeight.Text = VB6.Format(FireHeight(id), "0.000")

        'frmFire.Show()
    End Sub

    Public Sub mnuHelpContents_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuHelpContents.Click
        '*  ======================================================
        '*      Call a procedure to show the help contents
        '*  ======================================================

        'ShowHelpContents

        'set cancel to true
        'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        'frmBlank.CMDialog1.CancelError = True

        On Error GoTo errhandler
        'set the help command property
        'UPGRADE_ISSUE: MSComDlg.CommonDialog property CMDialog1.HelpCommand was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'frmBlank.CMDialog1.HelpCommand = MSComDlg.HelpConstants.cdlHelpForceFile
        'specify the help file
        'UPGRADE_ISSUE: MSComDlg.CommonDialog property CMDialog1.HelpFile was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'frmBlank.CMDialog1.HelpFile = "Branzfir.hlp"
        'display the windows help engine
        'UPGRADE_WARNING: The help CommonDialog is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E8158DFB-975E-4306-8E4F-43391F853639"'
        'frmBlank.CMDialog1.ShowHelp()
        Exit Sub

errhandler:
        'user pressed cancel button
    End Sub


    Public Sub mnuISO9705_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuISO9705.Click
        '*  ==============================================================
        '*      This event resets all the defaults for a new file
        '*  ==============================================================

        Dim msg As String
        Dim answer As Short
        Dim counter, L As Integer
        msg = "Would you like to save current file?"
        answer = MsgBox(msg, 52, "Save now?")
        If answer = vbYes Then
            'CancelError is True.
            On Error GoTo ErrHandler2

            SaveBaseToolStripMenuItem.PerformClick()

        ElseIf answer = 7 Then
            FileClose()
        End If

        On Error GoTo ErrHandler2

        'Call frmInputs.Read_BaseFile_xml(DataDirectory & "basemodel_ISO9705.xml", False)
        Call frmInputs.Read_BaseFile_xml(ApplicationPath & "\data\basemodel_ISO9705.xml", False)

        counter = 1

        'RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\riskdata\basemodel_ISO9705_" & counter.ToString & "\"
        RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_ISO9705_" & counter.ToString & "\"
        DataFile = RiskDataDirectory & "basemodel_ISO9705_" & counter.ToString & ".xml"


tryagain:

        If My.Computer.FileSystem.FileExists(DataFile) Then
            counter = counter + 1
            'RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\riskdata\basemodel_ISO9705_" & counter.ToString & "\"
            RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_ISO9705_" & counter.ToString & "\"
            DataFile = RiskDataDirectory & "basemodel_ISO9705_" & counter.ToString & ".xml"

            GoTo tryagain
        Else
            On Error GoTo ErrHandler2
            ProjectDirectory = RiskDataDirectory
            frmInputs.txtBaseName.Text = "ISO9705_" & counter.ToString
            Call frmInputs.Save_BaseFile_xml(DataFile)

            Dim odistribution As New oDistribution
            Dim odistributions As New List(Of oDistribution)
            odistribution = New oDistribution("Interior Temperature", "K", "None", 293, 293, 0, 0, 0, 0, 0, 0)
            odistribution.id = 1
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Exterior Temperature", "K", "None", 293, 293, 0, 0, 0, 0, 0, 0)
            odistribution.id = 2
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Relative Humidity", "-", "None", 0.5, 0.5, 0, 0, 0, 0, 0, 0)
            odistribution.id = 3
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Fire Load Energy Density", "MJ/m2", "None", 200, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 4
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Heat of Combustion PFO", "kJ/g", "None", 20, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 5
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Soot Preflashover Yield", "g/g", "None", 0.07, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 6
            odistributions.Add(odistribution)
            odistribution = New oDistribution("CO Preflashover Yield", "g/g", "None", 0.04, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 7
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Sprinkler Reliability", "-", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 8
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Sprinkler Suppression Probability", "-", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 9
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Sprinkler Cooling Coefficient", "-", "None", 1, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 10
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Fuel Heat of Gasification", "kJ/g", "None", 3, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 11
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Alpha T", "kW/s2", "None", 0.0469, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 12
            odistributions.Add(odistribution)
            odistribution = New oDistribution("Peak HRR", "kW", "None", 2000, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = 13
            odistributions.Add(odistribution)
            DistributionClass.SaveDistributions(odistributions)

            Dim oSprinklers As New List(Of oSprinkler)
            Dim osprdistributions As New List(Of oDistribution)
            SprinklerDB.SaveSprinklers2(oSprinklers, osprdistributions)

            Dim oSmokeDetectors As New List(Of oSmokeDet)
            Dim oSDdistributions As New List(Of oDistribution)
            SmokeDetDB.SaveSmokeDets(oSmokeDetectors, oSDdistributions)

            Dim oFans As New List(Of oFan)
            Dim ofandistributions As New List(Of oDistribution)
            FanDB.SaveFans(oFans, ofandistributions)

            Dim ocvents As New List(Of oCVent)
            Dim ocventdistributions As New List(Of oDistribution)
            VentDB.SaveCVents(ocvents, ocventdistributions)

            Dim ovents As New List(Of oVent)
            Dim oventdistributions As New List(Of oDistribution)
            Dim ovent As oVent
            ovent = New oVent(0, 200, False, 10, 100, 0, 1, 1, False, False, False, 0, gcs_DischargeCoeff, False, False, False, False, 0, 0, 0, 0, 500, 1, False, 0, False, 0, 0, 0, False, 72000, 0.00000042, 47, 20, 6, 0.0000083, 1, 0.937, "door opening ", 1, 2, 1, 2, 0.8, 0, 0, 0, 2.4, 2.4, False, False)
            ovent.id = 1
            ovents.Add(ovent)
            number_vents = NumberVents(ovent.fromroom, ovent.toroom) + 1
            NumberVents(ovent.fromroom, ovent.toroom) = number_vents
            NumberVents(ovent.toroom, ovent.fromroom) = number_vents
            Resize_Vents()
            VentFace(1, 2, 1) = 1
            VentOffset(1, 2, 1) = 0.8
            VentWidth(1, 2, 1) = 0.8
            VentHeight(1, 2, 1) = 2.0
            NumberVents(1, 2) = 1
            FRintegrity(1, 2, 1) = 0
            FRMaxOpening(1, 2, 1) = 100
            FRMaxOpeningTime(1, 2, 1) = 10

            odistribution = New oDistribution("", "", "None", 2, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "height"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0.8, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "width"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "prob"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "integrity"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 100, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopening"
            oventdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 10, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = ovent.id
            odistribution.varname = "maxopeningtime"
            oventdistributions.Add(odistribution)

            VentDB.SaveVents(ovents, oventdistributions)

            Dim oitems As New List(Of oItem)
            Dim oitemdistributions As New List(Of oDistribution)
            Dim oitem As oItem
            oitem = New oItem("0,0", 98.4, 0.338, 2, 684, 60, 0.101, 1, 3460, "ISO 9705 propane gas burner", "burner", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "0,0,5,100,600,100,605,300,1200,300", 0, 0, 20, 20, 0, 0.3, 0.17, 0.17, 0, 0, 0, "generic", "gas burner", 47.3, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, "None", 0, 0, 0, 0, 1, 0, 0, 0.3, 0, 0, 3, 0)
            oitem.id = 1
            oitems.Add(oitem)
            ObjDimX(1) = 0
            ObjDimY(1) = 0
            ObjLength(1) = 0.17
            ObjWidth(1) = 0.17
            ObjectMLUA(2, 1) = 3460
            odistribution = New oDistribution("", "", "None", 47.3, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "heat of combustion"
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0.024, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "soot yield"
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 2.85, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "co2 yield"
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "Latent Heat of Gasification"
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 0.3, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "Radiant Loss Fraction"
            oitemdistributions.Add(odistribution)

            odistribution = New oDistribution("", "", "None", 3460, 0, 0, 0, 0, 0, 0, 0)
            odistribution.id = oitem.id
            odistribution.varname = "HRRUA"
            oitemdistributions.Add(odistribution)

            ItemDB.SaveItemsv2(oitems, oitemdistributions)

        End If

        Dim temp As String = Path.GetFileName(DataFile)
        'basefile = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_" & CStr(frmInputs.txtBaseName.Text) & "\" & temp
        basefile = RiskDataDirectory & temp

        frmOptions1.optRCNone.Checked = CHECKED
        frmInputs.rtb_log.Clear() 'clear previous content in log screen
        frmInputs.ToolStripStatusLabel1.Text = ""
        frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory


        Exit Sub

ErrHandler2:
        'User pressed Cancel button
        Exit Sub

    End Sub

    Public Sub mnuLayerHeight_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLayerHeight.Click
        '*  ==========================================================
        '*  Show a graph of layer height above the floor versus time
        '*  ==========================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double

        'define variables
        Title = "Layer Height (m)"
        DataShift = 0
        DataMultiplier = 1
        MaxYValue = 0


        Graph_Data_2D(Title, layerheight, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuLayerTemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuLayerTemp.Click
        '*  =======================================================
        '*  Show a graph of upper layer temperature versus time
        '*  =======================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Upper Layer Temp (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 900

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, uppertemp(), DataShift, DataMultiplier

        Graph_Data_2D(Title, uppertemp, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuLoadFireDatabase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '		'*  ==========================================================
        '		'*  This subprocedure shows a custom dialog box for
        '		'*  opening a file.
        '		'*  ==========================================================

        '		Dim Openbox As System.Windows.Forms.Control
        '		'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        '		'UPGRADE_ISSUE: VB.Control Openbox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        '        'Openbox = frmBlank.CMDialog1

        '		'CancelError is True
        '		On Error GoTo errhandler

        '		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        '		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        '        'System.Windows.Forms.Cursor.Current = HOURGLASS

        '		'set directory to open in
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.InitDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.InitDir = My.Application.Info.DirectoryPath

        '		'dialogbox title
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.DialogTitle = "Load Fire Database"

        '		'Set filters
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"

        '		'Specify default filter
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.FilterIndex = 2

        '		'Display the Open dialog box
        '		If Confirm_File(FireDatabaseName, readfile, 0) = 0 Then FireDatabaseName = UserAppDataFolder & gcs_folder_ext & DefaultFireDatabaseName
        '		'If Confirm_File(FireDatabaseName, readfile, 0) = 0 Then FireDatabaseName = App.Path + "\" + DefaultFireDatabaseName
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.filename = FireDatabaseName
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'Openbox.Action = 1

        '		Dim firename As String
        '		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        '		firename = Dir(FireDatabaseName)

        '		'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar2.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        '		Me.StatusBar2.Items.Item(4).Text = firename & "   "
        '		'Call the open file procedure
        '		'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '        'FireDatabaseName = Openbox.filename
        '		'UPGRADE_ISSUE: Data property Data1.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFires.Data1.DatabaseName = FireDatabaseName
        '		'UPGRADE_ISSUE: Data property Data1.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFires.Data1.RecordSource = "fire data"
        '		frmFires.Data1.Refresh()
        '		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        '		frmFires.DBCombo1.CtlText = "Furniture"
        '		'UPGRADE_ISSUE: Data property Data1.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		'UPGRADE_NOTE: Text was upgraded to CtlText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        '		frmFires.Data1.RecordSource = "SELECT * FROM [Fire Data] " & "WHERE [Object Type] = '" & frmFires.DBCombo1.CtlText & "' order by [Description];"
        '		frmFires.Data1.Refresh()
        '		'UPGRADE_ISSUE: Data property Data2.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFires.Data2.DatabaseName = FireDatabaseName
        '		'UPGRADE_ISSUE: Data property Data2.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFires.Data2.RecordSource = "SELECT DISTINCT [Object Type] from [Fire Data] order by [Object Type]"
        '		frmFires.Data2.Refresh()
        '		'ADO here
        '		'frmFireData.datPrimaryRS.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FireDatabaseName
        '		'UPGRADE_ISSUE: Data property datPrimaryRS.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFireData.datPrimaryRS.DatabaseName = FireDatabaseName
        '		'UPGRADE_ISSUE: Data property datPrimaryRS.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFireData.datPrimaryRS.RecordSource = "fire data"
        '		'UPGRADE_ISSUE: Data property Data2.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '		frmFireData.Data2.RecordSource = "SELECT DISTINCT [Object Type] from [Fire Data] order by [Object Type]"
        '		' frmFireData.datPrimaryRS.RecordSource = "SELECT * FROM [Fire Data]"
        '		'frmFireData.datPrimaryRS.Refresh
        '		'frmFireData.Data2.Refresh

        '		frmFires.Text = "Fire Database " & "(" & FireDatabaseName & ")"

        '		ChDir((My.Application.Info.DirectoryPath))

        '		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        '		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        '		System.Windows.Forms.Cursor.Current = Default_Renamed
        '		Dim messagetxt As String
        '		messagetxt = "New Fire Database " & FireDatabaseName & " is loaded."
        '		MsgBox(messagetxt)
        '		Exit Sub

        'errhandler: 
        '		'User pressed Cancel button
        '		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        '		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        '		System.Windows.Forms.Cursor.Current = Default_Renamed
        '		'Close #1
        '		Exit Sub

    End Sub

    Public Sub mnuLoadMaterialsDatabase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '*  ==========================================================
        '*  This subprocedure shows a custom dialog box for
        '*  opening a file.
        '*  ==========================================================

        Dim Openbox As System.Windows.Forms.Control
        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        'UPGRADE_ISSUE: VB.Control Openbox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'Openbox = frmBlank.CMDialog1

        'CancelError is True
        On Error GoTo errhandler

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = HOURGLASS

        'set directory to open in
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.InitDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.InitDir = My.Application.Info.DirectoryPath

        'dialogbox title
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.DialogTitle = "Load Materials Database"

        'Set filters
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"

        'Specify default filter
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.FilterIndex = 2

        'Display the Open dialog box
        'If Confirm_File(MaterialsDatabaseName, readfile, 0) = 0 Then MaterialsDatabaseName = App.Path + "\" + DefaultMaterialsDatabaseName
        If Confirm_File(MaterialsDatabaseName, readfile, 0) = 0 Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.filename = MaterialsDatabaseName
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.Action = 1

        Dim matname As String
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'matname = Dir(MaterialsDatabaseName)
        matname = Path.GetFileName(MaterialsDatabaseName)
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar2.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.StatusBar2.Items.Item(3).Text = matname & "   "

        'Call the open file procedure
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'MaterialsDatabaseName = Openbox.filename
        'UPGRADE_ISSUE: Data property datPrimaryRS.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'frmThermal.datPrimaryRS.DatabaseName = MaterialsDatabaseName
        'UPGRADE_ISSUE: Data property datPrimaryRS.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'frmThermal.datPrimaryRS.RecordSource = "select [ID],[Material Description],[Thermal Conductivity],[Specific Heat],[Density],[Emissivity],[Comments2],[Min Temp for Spread],[Heat of Combustion],[Flame Spread Parameter],[Cone Data File],[Soot Yield],[H2O Yield],[CO2 Yield],[Calibration Factor] from [Table1] Order by [Material Description]"
        'frmThermal.datPrimaryRS.Refresh()
        'UPGRADE_ISSUE: Data property datPrimaryRS.DatabaseName was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'frmTable1.datPrimaryRS.DatabaseName = MaterialsDatabaseName
        'UPGRADE_ISSUE: Data property datPrimaryRS.RecordSource was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        'frmTable1.datPrimaryRS.RecordSource = "Table1"
        'frmTable1.datPrimaryRS.Refresh()
        ChDir((My.Application.Info.DirectoryPath))
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed

        Dim messagetxt As String
        messagetxt = "New Thermal Database " & MaterialsDatabaseName & " is loaded."
        MsgBox(messagetxt)
        Exit Sub

errhandler:
        'User pressed Cancel button
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        FileClose(1)
        Exit Sub


    End Sub


    Public Sub mnuMassLeft_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMassLeft.Click
        '*  =============================================================
        '*  Show a graph of the fuel mass remaining versus time
        '*  =============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Mass Consumed (kg)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_fuel(CShort(fireroom), Title, TotalFuel, DataShift, DataMultiplier, GraphStyle, MaxYValue, timeunit)

    End Sub

    Public Sub mnuMassLossGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMassLossGraph.Click
        '*  =============================================================
        '*  Show a graph of the fuel mass loss rate versus time
        '*  =============================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double

        'define variables
        Title = "Mass Loss Rate(kg/s)"
        DataShift = 0
        DataMultiplier = 1

        If FuelResponseEffects = True Then
            Title = "Mass Loss Rate(g/s)"
            DataShift = 0
            DataMultiplier = 1000
            Graph_Data_fuelburningrate(1, Title, FuelBurningRate, DataShift, DataMultiplier, timeunit)
        Else
            Graph_Data_2D_reverse(CInt(fireroom), Title, FuelMassLossRate, DataShift, DataMultiplier, timeunit)
        End If

    End Sub

    Public Sub mnuMaxCJettemp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMaxCJettemp.Click
        '*  ==========================================================
        '*  Show a graph ceiling jet temp at the location of the link versus time
        '*  ==========================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()
        If DetectorType(fireroom) = 0 Then
            MsgBox("You need to have specified a thermal detector or sprinkler in the fire room to record ceiling jet temperatures")
            Exit Sub
        End If

        'define variables
        Title = "Maximum Ceiling Jet Temp at Radial Position of Link (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_3D_CJET(2, Title, CJetTemp, DataShift, DataMultiplier)

    End Sub

    Public Sub mnuMechExtract_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)


    End Sub

    Public Sub mnuMultiGraphs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuMultiGraphs.Click
        '    Dim index As Integer
        '    For index = 0 To 3
        '        frmGraphs.runtimegraph1(index).Visible = True
        '        frmGraphs.Graph1.Visible = False
        '    Next index
        '    frmGraphs.Show
    End Sub

    Public Sub mnuNewFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuNewFile.Click
        '*  ==============================================================
        '*      This event resets all the defaults for a new file
        '*  ==============================================================

        Dim msg As String
        Dim answer As Short
        msg = "Would you like to save current file?"
        answer = MsgBox(msg, 52, "Save now?")
        If answer = 6 Then

            On Error GoTo ErrHandler2
            SaveBaseToolStripMenuItem.PerformClick()

        ElseIf answer = 7 Then
            FileClose()
        End If

        'RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"
        'DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        DataDirectory = DataFolder

        If My.Computer.FileSystem.FileExists(DataDirectory & "rooms.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "rooms.xml", RiskDataDirectory & "rooms.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "items.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "items.xml", RiskDataDirectory & "items.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "vents.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "vents.xml", RiskDataDirectory & "vents.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "cvents.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "cvents.xml", RiskDataDirectory & "cvents.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "fans.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "fans.xml", RiskDataDirectory & "fans.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "smokedets.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "smokedets.xml", RiskDataDirectory & "smokedets.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "sprinklers.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "sprinklers.xml", RiskDataDirectory & "sprinklers.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "distributions.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "distributions.xml", RiskDataDirectory & "distributions.xml", True)
        End If
        If My.Computer.FileSystem.FileExists(DataDirectory & "basemodel_default.xml") = True Then
            My.Computer.FileSystem.CopyFile(DataDirectory & "basemodel_default.xml", RiskDataDirectory & "basemodel_default.xml", True)
        End If

        frmInputs.rtb_log.Clear() 'clear previous content in log screen
        frmInputs.txtBaseName.Text = "default"
        frmInputs.ToolStripStatusLabel1.Text = ""
        ' basefile = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\basemodel_default.xml"
        basefile = DefaultRiskDataDirectory & "basemodel_default\basemodel_default.xml"
        ProjectDirectory = RiskDataDirectory
        Call frmInputs.Read_BaseFile_xml(basefile, False)

        Erase mc_item_hoc
        Erase mc_item_co2
        Erase mc_item_lhog
        Erase mc_item_RLF
        Erase mc_item_soot
        Erase mc_item_hrrua

        Dim oRooms As New List(Of oRoom)
        Dim oroomdistributions As New List(Of oDistribution)
        oRooms = RoomDB.GetRooms()

        NumberRooms = oRooms.Count
        Resize_Rooms()
        frmRoomList.fillroomarrays(oRooms)

        Dim oVents As New List(Of oVent)
        Dim oventdistributions As New List(Of oDistribution)
        oVents = VentDB.GetVents()

        Resize_Vents()
        Resize_CVents()

        frmInputs.rtb_log.Clear()
        frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory
        Exit Sub


ErrHandler2:
        'User pressed Cancel button
        Exit Sub
    End Sub

    Public Sub mnuO2Lower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuO2Lower.Click
        '*  ============================================================
        '*  Show a graph of the lower layer oxygen concentration versus time
        '*  ============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Lower O2 (%)"
        DataShift = 0
        DataMultiplier = 100
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, O2VolumeFraction(), DataShift, DataMultiplier, GraphStyle
        Graph_Data_Species(2, Title, O2VolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuO2Upper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuO2Upper.Click
        '*  ============================================================
        '*  Show a graph of the upper layer oxygen concentration versus time
        '*  ============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Upper O2 (%)"
        DataShift = 0
        DataMultiplier = 100
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_Species(1, Title, O2VolumeFraction, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuODcjet_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuODcjet.Click
        '*  ===============================================================
        '*  Show a graph of OD in the ceiling jet at the detector versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "OD at Smoke Detector (1/m)"
        DataShift = 0
        DataMultiplier = ParticleExtinction / 2.3
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, SmokeConcentration, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuODlower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuODlower.Click
        '*  ===============================================================
        '*  Show a graph of OD in the lower layer versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "OD lower (1/m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, OD_lower, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuODupper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuODupper.Click
        '*  ===============================================================
        '*  Show a graph of OD
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "OD upper (1/m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, OD_upper, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuPageSetup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPageSetup.Click
        Try
            'load page settings and display dialog box
            frmViewDocs.PageSetupDialog1.PageSettings = PrintPageSettings
            frmViewDocs.PageSetupDialog1.ShowDialog()
        Catch ex As Exception
            'display error message
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub mnuPlumeGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlumeGraph.Click
        '*  ====================================================================
        '*      Show a graph of the mass flow of air entrained in the plume
        '*      versus Time
        '*  ====================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Plume Mass Flow (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, MassPlumeFlow(), DataShift, DataMultiplier
        Graph_Data_2D_reverse(CInt(fireroom), Title, massplumeflow, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuPressure_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPressure.Click
        '*  ==========================================================
        '*  Show a graph of equilibrium pressure versus time
        '*  ==========================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Pressure (Pa)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, RoomPressure, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuPrinter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrinter.Click
        '=================================================================
        '   Display summary of input and output
        '=================================================================

        'Dim filepath As String = UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_out.txt"
        Dim filepath As String = getTempPathName()
        filepath = filepath & "temp_out.txt"

        If batch = False Then
            Call view_output(filepath)
            Dim myfilestream As New FileStream(filepath, FileMode.Open)
            frmViewDocs.RichTextBox1.LoadFile(myfilestream, RichTextBoxStreamType.PlainText)
            myfilestream.Close()
            Kill(filepath)
        Else
            Call view_output(BatchFileFolder & "\" & VB.Left(DataFile, Len(DataFile) - 4) & ".out")
            Dim myfilestream As New FileStream(filepath, FileMode.Open)
            frmViewDocs.RichTextBox1.LoadFile(myfilestream, RichTextBoxStreamType.PlainText)
            myfilestream.Close()
        End If

        Try
            frmViewDocs.RichTextBox1.Select(0, 0)
            StringToPrint = frmViewDocs.RichTextBox1.Text
            frmViewDocs.Text = "Summary of Inputs and Results"
            frmViewDocs.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Public Sub mnuPrintGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrintGraph.Click
        '***********************************************
        '*  Print current graph
        '*
        '*  Created 6 September 1996
        '***********************************************


        '    On Error GoTo handler
        '    Dim keep
        '    Dim i As Integer
        '
        '
        '    Printer.TrackDefault = True
        '    keep = Printer.Orientation
        '
        '    If PagePrinter Is Nothing Then
        '        Set PagePrinter = New PPrinter
        '    End If
        '
        '    If PagePrinter.Orientation = 0 Then
        '         PagePrinter.Orientation = Printer.Orientation
        '    End If
        '
        '    If frmGraphs.Graph1.Visible = True Then
        '        'these are the single graphs
        '        frmGraphs.Graph1.GraphStyle = 4 'print lines
        '        If Printer.ColorMode = vbPRCMColor Then
        '            frmGraphs.Graph1.PrintStyle = 1 'greyscale
        '        Else
        '            frmGraphs.Graph1.PrintStyle = 0 'monochrome
        '        End If
        '
        '        frmGraphs.Graph1.DrawMode = 5 'prints the image
        '
        '        'save a copy of the graph in a bitmap file
        '        frmGraphs.Graph1.DrawMode = 3
        '        frmGraphs.Graph1.ImageFile = App.Path + "\data\temp_graph" + str(i)
        '        frmGraphs.Graph1.DrawMode = 6
        '    Else
        '        Set Picture1.Picture = CaptureForm(frmGraphs)
        '
        '        Dim X As Printer
        '        For Each X In Printers
        '           If X.DeviceName = PagePrinter.DeviceName Then
        '              ' Set printer as system default.
        '              Set Printer = X
        '              ' Stop looking for a printer.
        '              Exit For
        '           End If
        '        Next
        '
        '
        '        PrintPictureToFitPage Printer, Picture1.Picture
        '        Printer.EndDoc
        '
        '    End If
        '
        '    Printer.Orientation = keep
        '
        '    Exit Sub
        '
        'handler:
    End Sub

    Public Sub mnuPrintout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrintout.Click
        frmprintvar.Show()
    End Sub

    Public Sub mnuPrintWordPad_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '*  =====================================================
        '*  This procedure sends model data and results to a
        '*  printer accessed through the CMDialog control
        '*
        '*  Revised 23 December 1997 Colleen Wade
        '*  =====================================================
        On Error GoTo Printerrorhandler
        Dim tmp As String

        tmp = RiskDataDirectory & "temp_out.txt"
        ' tmp = UserAppDataFolder & gcs_folder_ext & " \ data \ " & "temp_out.txt"

        Call view_output(tmp)

        Shell("notepad.exe tmp", AppWinStyle.MaximizedFocus)

        'frmBlank.Show
        Exit Sub

        'errorhandler
Printerrorhandler:
        'User pressed cancel button
        If Err.Number = 9 Then Resume Next
        If Err.Number <> 32755 And Err.Number <> 0 Then
            MsgBox(ErrorToString(Err.Number))
            'Resume Next
        End If
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = ARROW
        Exit Sub

    End Sub

    Public Sub mnuRadiationFloor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRadiationFloor.Click
        '*  ===============================================================
        '*  Show a graph of the incident radiant flux at a point on the
        '*  floor at the centre of the room.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Incident Radiant Flux on Floor (kW/sqm)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, Target, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuRadiationTarget_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRadiationTarget.Click
        '*  ===============================================================
        '*  Show a graph of the incident radiant flux at a point on a surface
        '*  at the centre of the room at the monitoring height.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Incident Radiant Flux on Target (kW/sqm)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Target(), DataShift, DataMultiplier
        'Graph_Data_2D(1, Title, SurfaceRad, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, SurfaceRad, DataShift, DataMultiplier, timeunit)

    End Sub

    Private Sub mnuRegisterAgain_Click()
        'frmLicense.ShowDialog()
    End Sub

    Public Sub mnuRoom_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRoom.Click
        '        '*  ===================================================================
        '        '*      This event displays the form for entering information about the room
        '        '*  ===================================================================

        '        Dim i As Integer
        '        Dim idr, idv, idc As Short
        '        Dim k As Integer


        '        ' frmDescribeRoom.Show()

        '        frmDescribeRoom.lstCVentID.Items.Clear()
        '        ' frmDescribeRoom.ComboBox_RoomID.Items.Clear()
        '        frmDescribeRoom.lstCVentIDlower.Items.Clear()
        '        frmDescribeRoom.lstCVentIDUpper.Items.Clear()
        '        frmDescribeRoom._lstRoomID_1.Items.Clear()
        '        frmDescribeRoom._lstRoomID_2.Items.Clear()
        '        frmDescribeRoom._lstRoomID_3.Items.Clear()

        '        On Error GoTo venterrorhandler
        '        If NumberRooms > 0 Then
        '            For i = 1 To NumberRooms
        '                frmDescribeRoom._lstRoomID_1.Items.Add(CStr(i))
        '                frmDescribeRoom._lstRoomID_2.Items.Add(CStr(i))
        '                frmDescribeRoom._lstRoomID_3.Items.Add(CStr(i))
        '                ' frmDescribeRoom.ComboBox_RoomID.Items.Add(CStr(i))
        '                frmDescribeRoom.lstCVentIDlower.Items.Add(CStr(i))
        '                frmDescribeRoom.lstCVentIDUpper.Items.Add(CStr(i))
        '            Next i
        '            frmDescribeRoom.lstCVentIDlower.Items.Add("outside")
        '            frmDescribeRoom.lstCVentIDUpper.Items.Add("outside")
        '        End If

        '        ' frmDescribeRoom.ComboBox_RoomID.SelectedIndex = 0
        '        frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
        '        frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
        '        frmDescribeRoom._lstRoomID_3.SelectedIndex = 0

        '        idc = 2
        '        idr = 1

        '        If NumberCVents(idr, idc) > 0 Then
        '            For i = 1 To NumberCVents(idr, idc)
        '                frmDescribeRoom.lstCVentID.Items.Add("Vent " & CStr(i))
        '            Next i
        '        End If
        '        idv = 1

        '        frmDescribeRoom.txtCeilingSubThickness.Text = CStr(CeilingSubThickness(idr))
        '        frmDescribeRoom.txtFloorSubThickness.Text = CStr(FloorSubThickness(idr))
        '        frmDescribeRoom.txtCeilingThickness.Text = CStr(CeilingThickness(idr))
        '        frmDescribeRoom.txtFloorThickness.Text = CStr(FloorThickness(idr))
        '        frmDescribeRoom.txtWallSubThickness.Text = CStr(WallSubThickness(idr))
        '        frmDescribeRoom.txtWallThickness.Text = CStr(WallThickness(idr))

        '        If HaveCeilingSubstrate(idr) = True Then
        '            frmDescribeRoom.chkAddSubstrate.Checked = True
        '        Else
        '            frmDescribeRoom.chkAddSubstrate.Checked = False
        '        End If
        '        If HaveWallSubstrate(idr) = True Then
        '            frmDescribeRoom.chkAddWallSubstrate.Checked = True
        '        Else
        '            frmDescribeRoom.chkAddWallSubstrate.Checked = False
        '        End If
        '        If HaveFloorSubstrate(idr) = True Then
        '            frmDescribeRoom.chkAddFloorSubstrate.Checked = True
        '        Else
        '            frmDescribeRoom.chkAddFloorSubstrate.Checked = False
        '        End If
        '        On Error Resume Next

        '        If frmDescribeRoom.lstCVentID.Items.Count > 0 Then
        '            frmDescribeRoom.lstCVentID.SelectedIndex = 0
        '        End If
        '        frmDescribeRoom.lstCVentIDlower.SelectedIndex = 0
        '        frmDescribeRoom.lstCVentIDUpper.SelectedIndex = frmDescribeRoom.lstCVentIDUpper.Items.Count - 1

        '        frmDescribeRoom._SSTab1_0.SelectedIndex = 0

        '        Exit Sub

        'venterrorhandler:
        '        Exit Sub

    End Sub
    Public Sub mnuRun_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '*  ============================================================
        '*  This subprocedure starts the fire simulation
        '*  ============================================================

        Dim j, response, vent As Short
        Dim start, simend As Double
        Dim room As Integer
        ReDim FirstTime(NumberRooms)

        On Error GoTo runerrorhandler

        flagstop = 0
        For room = 1 To NumberRooms
            FirstTime(room) = True
        Next room
        Flashover = False

        'seek confirmation from the user that they want to run the simulation
        response = MsgBox("Do you want to run the simulation now?", MB_YESNO + MB_ICONQUESTION, ProgramTitle)
        If response = IDNO Then Exit Sub

        'note start time of simulation
        start = VB.Timer()
        Timer1.Enabled = True

        'check to see that some heat release rate data has been entered by the user
        If NumberDataPoints(1) = 0 And frmOptions1.optRCNone.Checked = True Then
            MsgBox("You have not yet provided a design fire", MsgBoxStyle.Information)
            Exit Sub
        End If

        My.Computer.FileSystem.CurrentDirectory = DataDirectory

        'show the progress indicator bar at the bottom of the screen
        Me.ToolStripProgressBar1.Visible = True

        'update all variables and calculate derived variables
        Derived_Variables()
        If flagstop = 1 Then Exit Sub
        Number_TimeSteps()

        ResetWindows()

        If NumberTimeSteps > 1 Then
            Me.mnuExcel.Enabled = True
        Else
            Me.mnuExcel.Enabled = False
        End If

        'Resize the data storage arrays based on the number of timesteps needed
        Size_Arrays()

        'check for a valid vent
        Check_Vent()

        'get the plume correlation to be used
        Get_Plume()

        'call procedure to initialize endpoint strings
        Initialize_EndPointFlags()

        'solve the ordinary differential equations for layer height
        'and upper layer temperature
        Dim osprdistributions As New List(Of oDistribution)
        Dim osprinklers As New List(Of oSprinkler)
        Dim osddistributions As New List(Of oDistribution)
        Dim oSmokeDets As New List(Of oSprinkler)

        Call main_program2(1, osprinklers, osprdistributions, oSmokeDets, osddistributions)

        'reset the vent opening times
        ReDim VentBreakTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
        For room = 1 To NumberRooms
            For j = 1 To NumberRooms + 1
                If j > room Then
                    For vent = 1 To NumberVents(room, j)
                        VentBreakTime(room, j, vent) = VentOpenTime(room, j, vent)
                        VentOpenTime(room, j, vent) = oldVentOpenTime(room, j, vent)
                    Next vent
                End If
            Next j
        Next room

        'calculate any volume fractions required
        Volume_Fractions()

        'calculation incapacitation FED's for toxicity
        ReDim FEDSum(NumberRooms, NumberTimeSteps + 2)
        ReDim FEDRadSum(NumberRooms, NumberTimeSteps + 1)
        ReDim SurfaceRad(NumberRooms, NumberTimeSteps + 2)

        If FEDCO = True Then
            Call FED_CO_iso13571_multi()
        Else
            Call FED_gases_multi()
        End If

        Call FED_thermal_iso13571_multi()

        'note end time of simulation
        simend = VB.Timer()
        runtime = simend - start

        'indicate to the user that the fire simulation is complete
        MsgBox("Fire Simulation Completed in " & VB6.Format(runtime, "0.0") & " seconds.")

        ChartRuntime1.Visible = False
        ChartRuntime2.Visible = False

        Me.ToolStripProgressBar1.Visible = False
        Me.ToolStripProgressBar1.Value = 0

        Timer1.Enabled = False
        Exit Sub

runerrorhandler:

        MsgBox(ErrorToString(Err.Number))
        Exit Sub
    End Sub



    Public Sub mnuUpdateVersionInputFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUpdateVersionInputFiles.Click

        'Dim modnames(500) As String
        'Dim i As Short

        'batch = True 'we are running batch files

        ''set directory to open in
        ''DataDirectory = App.Path + "\" + "data\"
        'If BatchFileFolder = "" Then BatchFileFolder = UserAppDataFolder & gcs_folder_ext & "\data\"
        'DataDirectory = BatchFileFolder & "\"

        'i = 1
        'modnames(i) = Path.GetFileName(DataDirectory & "*.mod")

        'While modnames(i) <> "" And i < 500
        '    modnames(i) = DataDirectory & modnames(i)

        '    i = i + 1
        '    modnames(i) = Dir()

        'End While

        'i = 1
        'While modnames(i) <> ""
        '    ' DataFile = Dir(modnames(i))
        '    DataFile = Path.GetFileName(modnames(i))
        '    Call Open_File(DataDirectory & DataFile, batch)
        '    Call Save_File(DataDirectory & DataFile)
        '    i = i + 1
        'End While

    End Sub



    Public Sub mnuRunBatch_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuRunBatch.Click
        Dim frmGraphs As Object
        'folder not used
        '*  ==========================================================
        '*  This subprocedure shows a custom dialog box for
        '*  opening a file.
        '*  ==========================================================
        Dim j, response, vent As Short
        Dim start, simend As Double
        Dim room As Integer
        Dim myfile As String
        Dim modnames(500) As String
        Dim i As Short
        On Error GoTo errhandler
        batch = True 'we are running batch files

        'set directory to open in
        'DataDirectory = App.Path + "\" + "data\"
        If BatchFileFolder = "" Then BatchFileFolder = UserAppDataFolder & gcs_folder_ext & "\data"
        DataDirectory = BatchFileFolder & "\"

        i = 1
        'modnames(i) = Dir(DataDirectory & "*.mod")
        modnames(i) = Path.GetFileName(DataDirectory & "*.mod")

        While modnames(i) <> "" And i < 500
            modnames(i) = DataDirectory & modnames(i)
            i = i + 1
            modnames(i) = Dir()
        End While

        'show the progress indicator bar at the bottom of the screen
        Me.ToolStripProgressBar1.Visible = True

        i = 1

        While modnames(i) <> "" And batch = True

            'Call the open file procedure
            'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'DataFile = Dir(modnames(i))
            'Call Open_File(DataDirectory & DataFile, batch)
            Call Open_File(modnames(i), batch)
            DataFile = Path.GetFileName(modnames(i))
            ToolStripStatusLabel1.Text = DataFile
            ToolStripStatusLabel2.Text = Description

            '---------------------
            'run model
            ReDim FirstTime(NumberRooms)
            On Error GoTo runerrorhandler

            flagstop = 0
            For room = 1 To NumberRooms
                FirstTime(room) = True
            Next room
            Flashover = False

            'note start time of simulation
            start = VB.Timer()

            'check to see that some heat release rate data has been entered by the user
            If NumberDataPoints(1) = 0 And frmOptions1.optRCNone.Checked = True Then
                'MsgBox "You have not yet provided a design fire", vbInformation
                Exit Sub
            End If

            'frmBlank.Show

            'show the progress indicator bar at the bottom of the screen
            'MDIFrmMain.ProgressBar1.Visible = True

            'update all variables and calculate derived variables
            Derived_Variables()
            If flagstop = 1 Then Exit Sub
            Number_TimeSteps()

            ResetWindows()

            If NumberTimeSteps > 1 Then
                Me.mnuExcel.Enabled = True
            Else
                Me.mnuExcel.Enabled = False
            End If

            'Resize the data storage arrays based on the number of timesteps needed
            Size_Arrays()

            'check for a valid vent
            Check_Vent()

            'get the plume correlation to be used
            Get_Plume()

            'call procedure to initialize endpoint strings
            Initialize_EndPointFlags()

            'hide graphs on screen
            ChartRuntime1.Visible = False

            'solve the ordinary differential equations for layer height
            'and upper layer temperature
            Dim osprinklers As New List(Of oSprinkler)
            Dim osprdistributions As New List(Of oDistribution)
            Dim oSmokeDets As New List(Of oSmokeDet)
            Dim osddistributions As New List(Of oDistribution)

            main_program2(1, osprinklers, osprdistributions, oSmokeDets, osddistributions)

            'reset the vent opening times
            ReDim VentBreakTime(MaxNumberRooms + 1, MaxNumberRooms + 1, MaxNumberVents)
            For room = 1 To NumberRooms
                For j = 1 To NumberRooms + 1
                    If j > room Then
                        For vent = 1 To NumberVents(room, j)
                            VentBreakTime(room, j, vent) = VentOpenTime(room, j, vent)
                            VentOpenTime(room, j, vent) = oldVentOpenTime(room, j, vent)
                        Next vent
                    End If
                Next j
            Next room

            'calculate any volume fractions required
            Volume_Fractions()

            'calculation incapacitation FED's for toxicity
            ReDim FEDSum(NumberRooms, NumberTimeSteps + 2)
            ReDim FEDRadSum(NumberRooms, NumberTimeSteps + 1)
            ReDim SurfaceRad(NumberRooms, NumberTimeSteps + 2)

            If FEDCO = True Then
                Call FED_CO_iso13571_multi()
            Else
                Call FED_gases_multi()
            End If

            Call FED_thermal_iso13571_multi()

            'note end time of simulation
            simend = VB.Timer()
            runtime = simend - start

            Me.ToolStripProgressBar1.Value = 0

            Dim filepathbatch As String = BatchFileFolder & "\" & VB.Left(DataFile, Len(DataFile) - 4) & ".out"
            Call view_output(filepathbatch)

            i = i + 1
        End While

        Me.ToolStripProgressBar1.Visible = False

        batch = False
        MsgBox("Batch files run complete")

        Exit Sub

errhandler:
        'User pressed Cancel button
        FileClose(1)
        batch = False
        Exit Sub

runerrorhandler:
        MsgBox(ErrorToString(Err.Number))
        batch = False
        Exit Sub
    End Sub

    Public Sub mnuSaveFireDatabase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim SaveBox As System.Windows.Forms.Control

        On Error GoTo ErrHandler3

        ChDir(My.Application.Info.DirectoryPath)

        'default filename
        If Confirm_File(FireDatabaseName, readfile, 0) = 0 Then FireDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultFireDatabaseName
        'If Confirm_File(FireDatabaseName, readfile, 0) = 0 Then FireDatabaseName = UserAppDataFolder & gcs_folder_ext & DefaultFireDatabaseName

        Exit Sub

ErrHandler3:

        Exit Sub

    End Sub

    Public Sub mnuSaveMaterialsDataBase_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim SaveBox As System.Windows.Forms.Control
        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        'UPGRADE_ISSUE: VB.Control SaveBox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'SaveBox = frmBlank.CMDialog1

        'CancelError is True.
        On Error GoTo ErrHandler3

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = HOURGLASS
        ChDir(My.Application.Info.DirectoryPath)

        'dialogbox title
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DialogTitle = "Save Materials Database"

        'Set filters
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"

        'Specify default filter
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.FilterIndex = 2

        'default filename extension
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DefaultExt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DefaultExt = "mdb"

        'default filename
        'If Confirm_File(MaterialsDatabaseName, readfile, 0) = 0 Then MaterialsDatabaseName = App.Path + "\" + DefaultMaterialsDatabaseName
        If Confirm_File(MaterialsDatabaseName, readfile, 0) = 0 Then MaterialsDatabaseName = UserPersonalDataFolder & gcs_folder_ext & DefaultMaterialsDatabaseName
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.filename = MaterialsDatabaseName

        Dim matname As String
        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'matname = Dir(MaterialsDatabaseName)
        matname = Path.GetFileName(MaterialsDatabaseName)
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar2.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.StatusBar2.Items.Item(3).Text = matname & "   "

        'check before overwriting file
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.flags = OFN_OVERWRITEPROMPT Or OFN_NOCHANGEDIR

        'Display the Save as dialog box.
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.Action = 2

        'Call the save file procedure
        'Save_File (SaveBox.filename)
        'If Confirm_File((SaveBox.filename), replacefile) = 1 Then

        ''UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'If SaveBox.filename <> (MaterialsDatabaseName) Then
        '	On Error Resume Next
        '	'UPGRADE_ISSUE: Data property datPrimaryRS.Database was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '	frmThermal.datPrimaryRS.Database.Close()
        '	'UPGRADE_ISSUE: Data property datPrimaryRS.Database was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
        '	frmTable1.datPrimaryRS.Database.Close()
        '	'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '	FileCopy(MaterialsDatabaseName, SaveBox.filename)
        '	'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '	MaterialsDatabaseName = SaveBox.filename
        'End If

        'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'matname = Dir(MaterialsDatabaseName)
        matname = Path.GetFileName(MaterialsDatabaseName)
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar2.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.StatusBar2.Items.Item(3).Text = matname & "   "

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Sub

ErrHandler3:
        'User pressed Cancel button
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Sub


    End Sub
    Public Sub savemodeldata()

        Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()

        'dialogbox title
        SaveBox.Title = "Save Model Data as XML"

        'Set filters
        SaveBox.Filter = "All Files (*.*)|*.*|Model Files (*.xml)|*.xml"

        'Specify default filter
        SaveBox.FilterIndex = 2
        SaveBox.RestoreDirectory = True

        'default filename extension
        SaveBox.DefaultExt = "xml"

        'default filename
        'If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        If DataDirectory = "" Then DataDirectory = RiskDataDirectory
        SaveBox.InitialDirectory = DataDirectory

        If My.Computer.FileSystem.FileExists(DataFile) Then
            If DataFile.EndsWith(SaveBox.DefaultExt) Then
                SaveBox.FileName = DataFile
            End If
        End If

        'Display the Save as dialog box.
        If SaveBox.ShowDialog() = DialogResult.OK Then
            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
            End If
            myStream = SaveBox.OpenFile()

            If (myStream IsNot Nothing) Then

                DataFile = SaveBox.FileName
                myStream.Close()
                Dim oSprinklers As New List(Of oSprinkler)
                Dim oSmokeDets As New List(Of oSmokeDet)
                Dim oFans As New List(Of oFan)

                Call Save_File_xml(SaveBox.FileName, 1, oSprinklers, oSmokeDets, oFans)

            Else
                myStream.Close()
            End If
        End If
    End Sub


    Public Sub mnuSaveModelData_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSaveModelData.Click
        '*  ====================================================
        '*  This procedure saves the model input data for a run
        '*  in an external file with file extension *.MOD
        '*  ====================================================
        ProjectDirectory = DataDirectory

        Call savemodeldata()

    End Sub

    Public Sub mnuSaveResultsFile_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '*  ====================================================
        '*  This procedure saves the output results for a run
        '*  in an external file with file extension *.RES
        '*  ====================================================

        Dim SaveBox As System.Windows.Forms.Control

        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        'UPGRADE_ISSUE: VB.Control SaveBox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'SaveBox = frmBlank.CMDialog1

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = HOURGLASS

        'CancelError is True.
        On Error GoTo SaveResultsErrHandler

        'dialogbox title
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DialogTitle = "Save Simulation Results"

        'default filename
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.filename = "Demo.Res"

        'Set filters
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.Filter = "All Files (*.*)|*.*|Results Files (*.res)|*.res"

        'Specify default filter
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.FilterIndex = 2

        'default filename extension
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DefaultExt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DefaultExt = "RES"

        'check before overwriting file
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.flags. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.flags = OFN_OVERWRITEPROMPT Or OFN_NOCHANGEDIR

        'Display the Save as dialog box.
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.Action = 2

        'Call the save to file procedure
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Save_Results((SaveBox.filename))
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Sub

SaveResultsErrHandler:
        'User pressed Cancel button
        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        '	System.Windows.Forms.Cursor.Current = Default_Renamed
        Exit Sub

    End Sub

    Public Sub mnuSelectFolder_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSelectFolder.Click

        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.ShowDialog()
        BatchFileFolder = FolderBrowserDialog1.SelectedPath
        If BatchFileFolder = "" Then BatchFileFolder = UserAppDataFolder & gcs_folder_ext & "\data"
        MsgBox("Batch File Folder is " & BatchFileFolder)

    End Sub

    Public Sub mnuSimulation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSimulation.Click

        Dim room As Integer

        frmOptions1.lstRoomZone.Items.Clear()
        For room = 1 To NumberRooms
            frmOptions1.lstRoomZone.Items.Add(CStr(room))
        Next room
        frmOptions1.lstRoomZone.SelectedIndex = 0

        If TwoZones(1) = True Then
            frmOptions1.optTwoLayer.Checked = True
        Else
            frmOptions1.optOneLayer.Checked = Not True
        End If

        If cjModel = cjJET Then
            frmOptions1.optJET.Checked = True
        ElseIf cjModel = cjAlpert Then
            frmOptions1.optAlpert.Checked = True
        End If

        'update the textboxes


        frmOptions1.txtDescription.Text = Description
        frmOptions1.txtJobNumber.Text = JobNumber
        frmOptions1.txtInteriorTemp.Text = CStr(InteriorTemp - 273)
        frmOptions1.txtExteriorTemp.Text = CStr(ExteriorTemp - 273)
        frmOptions1.txtRelativeHumidity.Text = CStr(RelativeHumidity * 100)
        frmOptions1.txtMonitorHeight.Text = VB6.Format(MonitorHeight, "0.000")
        frmOptions1.txtTarget.Text = VB6.Format(TargetEndPoint, "0.0")
        frmOptions1.txtTemp.Text = VB6.Format(TempEndPoint - 273, "0.0")
        frmOptions1.txtVisibility.Text = VB6.Format(VisibilityEndPoint, "0.0")
        frmOptions1.txtConvect.Text = VB6.Format(ConvectEndPoint - 273, "0")
        frmOptions1.txtFED.Text = VB6.Format(FEDEndPoint, "0.0")
        frmOptions1.txtFlameLengthPower.Text = CStr(FlameLengthPower)
        frmOptions1.txtBurnerWidth.Text = CStr(BurnerWidth)
        frmOptions1.txtFlameAreaConstant.Text = CStr(FlameAreaConstant)
        frmOptions1.txtCeilingHeatFlux.Text = CStr(CeilingHeatFlux)
        frmOptions1.txtWallHeatFlux.Text = CStr(WallHeatFlux)
        ' frmOptions1.txtStartOccupied.Text = CStr(StartOccupied)
        ' frmOptions1.txtEndOccupied.Text = CStr(EndOccupied)
        frmOptions1.txtNodes(0).Text = CStr(Ceilingnodes)
        frmOptions1.txtNodes(1).Text = CStr(Wallnodes)
        frmOptions1.txtNodes(2).Text = CStr(Floornodes)
        frmOptions1.lstTimeStep.Text = CStr(Timestep)
        'frmOptions1.txtFLED.Text = CStr(FLED)
        frmOptions1.txtFuelThickness.Text = CStr(Fuel_Thickness)
        frmOptions1.txtStickSpacing.Text = CStr(Stick_Spacing)
        frmOptions1.txtCribheight.Text = CStr(Cribheight)
        frmOptions1.txtExcessFuelFactor.Text = CStr(ExcessFuelFactor)
        ' frmOptions1.txtHOCFuel.Text = CStr(HoC_fuel)
        frmOptions1.txtnC.Text = CStr(nC)
        frmOptions1.txtnH.Text = CStr(nH)
        frmOptions1.txtnO.Text = CStr(nO)
        frmOptions1.txtnN.Text = CStr(nN)
        frmOptions1.txtStoich.Text = CStr(StoichAFratio)
        ' frmOptions1.txtRadiantLossFraction.Text = RadiantLossFraction

        If Activity = "Rest" Then
            frmOptions1.optAtRest.Checked = True
        ElseIf Activity = "Light" Then
            frmOptions1.optLightWork.Checked = True
        Else
            frmOptions1.optHeavyWork.Checked = True
        End If

        Dim odistributions As List(Of oDistribution)
        odistributions = DistributionClass.GetDistributions()
        For Each oDistribution In odistributions
            Select Case oDistribution.varname
                Case "Relative Humidity"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtRelativeHumidity.BackColor = distbackcolor
                    Else
                        frmOptions1.txtRelativeHumidity.BackColor = distnobackcolor
                    End If
                Case "Interior Temperature"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtInteriorTemp.BackColor = distbackcolor
                    Else
                        frmOptions1.txtInteriorTemp.BackColor = distnobackcolor
                    End If
                Case "Exterior Temperature"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtExteriorTemp.BackColor = distbackcolor
                    Else
                        frmOptions1.txtExteriorTemp.BackColor = distnobackcolor
                    End If
                Case "Heat of Combustion PFO"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtHOCFuel.BackColor = distbackcolor
                    Else
                        frmOptions1.txtHOCFuel.BackColor = distnobackcolor
                    End If
                Case "Soot Preflashover Yield"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtpreSoot.BackColor = distbackcolor
                    Else
                        frmOptions1.txtpreSoot.BackColor = distnobackcolor
                    End If
                Case "CO Preflashover Yield"
                    If oDistribution.distribution <> "None" Then
                        frmOptions1.txtpreCO.BackColor = distbackcolor
                    Else
                        frmOptions1.txtpreCO.BackColor = distnobackcolor
                    End If
            End Select
        Next

        frmOptions1.SSTab2.SelectedIndex = 0
        frmOptions1.BringToFront()
        frmOptions1.Show()

    End Sub

    Public Sub mnuSmokeDetector_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)


    End Sub

    Public Sub mnuSPR_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSPR.Click
        '*  ===============================================================
        '*  Show a graph of OD in the lower layer versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "SPR (m2/s base-e) "
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        'Graph_Data_2D(1, Title, SPR, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, SPR, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuStop1.Click
        '*************************************************
        '*  Change the mouse pointer
        '*
        '*  Revised 23 October 1996 Colleen Wade
        '*************************************************

        'UPGRADE_ISSUE: MDIForm property MDIFrmMain.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'Cursor = ARROW

    End Sub

    Public Sub mnuTechRef_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTechRef.Click
        If Not CBool(ShellExecute(Me.Handle.ToInt32, "open", DocsFolder & techref_file, "", CStr(0), 1)) Then
            MsgBox("You don't have the Acrobat Reader installed.")
            Exit Sub
        End If
        'If Not CBool(ShellExecute(Me.Handle.ToInt32, "open", UserAppDataFolder & gcs_folder_ext & techref_file, "", CStr(0), 1)) Then
        '    MsgBox("You don't have the Acrobat Reader installed.")
        '    Exit Sub
        'End If

    End Sub

    Public Sub mnuTerminate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTerminate.Click
        '*******************************************************
        '*  The user wants to terminate the simulation early
        '*
        '*  Revised 19 October 1996 Colleen Wade
        '*******************************************************

        flagstop = 1
        gTimeX = SimTime
        Timer1.Enabled = False
        mnuRun.Enabled = True
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.Toolbar2.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.Toolbar2.Items.Item(8).Visible = True 'start run
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.Toolbar2.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.Toolbar2.Items.Item(9).Visible = False 'stop run
        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.Toolbar2.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        'Me.Toolbar2.Items.Item(19).Visible = True 'graphs
        'UPGRADE_ISSUE: Form property frmBlank.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'frmBlank.Cursor = Default_Renamed

    End Sub
    Private Sub mnuThermal_Click()
        '*  ==============================================================
        '*  Show form for thermal database properties.
        '*  ==============================================================

        frmMatSelect.Show()



    End Sub

    Public Sub mnuThermalEdit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub



    Public Sub mnuTUHCUpper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTUHCUpper.Click
        '*  ======================================================
        '*  Show a graph of the upper layer TUHC concentration versus time.
        '*  ======================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Upper Unburned Fuel (g/g)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_Species(1, Title, TUHC, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuTUHVLower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTUHVLower.Click
        '*  ======================================================
        '*  Show a graph of the lower layer TUHC concentration versus time.
        '*  ======================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Lower Unburned Fuel (g/g)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_Species(2, Title, TUHC, DataShift, DataMultiplier, timeunit)

    End Sub



    Public Sub mnuUsers_Guide_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        'If Not CBool(ShellExecute(Me.hWnd, "open", App.Path + userguide_file, "", 0, 1)) Then
        'If Not CBool(ShellExecute(Me.Handle.ToInt32, "open", UserAppDataFolder & gcs_folder_ext & userguide_file, "", CStr(0), 1)) Then
        '    MsgBox("You don't have the Acrobat Reader installed.")
        '    Exit Sub
        'End If
    End Sub

    Public Sub mnuVentFires_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVentFires.Click
        '*  ===============================================================
        '*  Show a graph of vent fire hrr versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Vent Fires (kW)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, ventfire, 0, 1, timeunit)
        ' graph_data_ventfire(1, Title, ventfire, DataShift, DataMultiplier, GraphStyle, MaxYValue)

    End Sub

    Public Sub mnuVerification_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVerification.Click
        If Not CBool(ShellExecute(Me.Handle.ToInt32, "open", DocsFolder & verification_file, "", CStr(0), 1)) Then
            MsgBox("You don't have the Acrobat Reader installed.")
            Exit Sub
        End If
        'If Not CBool(ShellExecute(Me.Handle.ToInt32, "open", UserAppDataFolder & gcs_folder_ext & verification_file, "", CStr(0), 1)) Then
        '    MsgBox("You don't have the Acrobat Reader installed.")
        '    Exit Sub
        'End If
    End Sub

    Public Sub mnuVersion_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVersion.Click
        '=============================================
        ' view version release notes
        '=============================================
        Try
            If My.Computer.FileSystem.FileExists(DocsFolder & "release_notes.rtf") Then
                frmViewDocs.RichTextBox1.LoadFile(DocsFolder & "release_notes.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
                frmViewDocs.SelectVariablesToolStripMenuItem.Visible = False
                frmViewDocs.Text = "B-RISK Version Release Notes"
                frmViewDocs.Show()
            End If
            'If My.Computer.FileSystem.FileExists(UserAppDataFolder & gcs_folder_ext & "\docs\" & "release_notes.rtf") Then
            '    frmViewDocs.RichTextBox1.LoadFile(UserAppDataFolder & gcs_folder_ext & "\docs\" & "release_notes.rtf", Windows.Forms.RichTextBoxStreamType.RichText)
            '    frmViewDocs.SelectVariablesToolStripMenuItem.Visible = False
            '    frmViewDocs.Text = "B-RISK Version Release Notes"
            '    frmViewDocs.Show()
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub mnuView_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuView.Click
        '============================================================
        '      If model has not been run then displaying results not
        '      allowed
        '============================================================

        If NumberTimeSteps = 0 Then
            mnuViewEndPoints.Enabled = False
            mnuViewRTB.Enabled = False
            mnuWVent.Enabled = False
            mnuView_CVent.Enabled = False
        Else
            'mnuModelResults.Enabled = True
            mnuViewEndPoints.Enabled = True
            mnuViewRTB.Enabled = True
            mnuWVent.Enabled = True
            mnuView_CVent.Enabled = True
        End If

    End Sub

    Public Sub mnuView_CVent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuView_CVent.Click
        Try
            If CVentlogfile <> "empty" Then
                frmViewDocs.RichTextBox1.LoadFile(CVentlogfile, Windows.Forms.RichTextBoxStreamType.PlainText)
                frmViewDocs.Text = "Summary of Ceiling Vent Flow Data"
                frmViewDocs.SelectVariablesToolStripMenuItem.Visible = False
                frmViewDocs.Show()
            Else
                MsgBox("No ceiling vent flow data available for this simulation.")
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub mnuViewEndPoints_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewEndPoints.Click
        '*  ==============================================================
        '*  Show form for viewing end-points.
        '*  ==============================================================

        'frmBlank.Show()
        frmEndPoints.Show()

        'get captions for labels
        Update_EndPointCaptions()

    End Sub

    Public Sub mnuViewInput_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewInput.Click
        '=================================================================
        '   Display summary of input and output
        '=================================================================

        'Dim StreamToDisplay As StreamReader
        'Dim filepath As String = UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_in.txt"
        Dim filepath As String = RiskDataDirectory & "temp_in.txt"

        Call view_input(filepath)

        Try
            Dim myfilestream As New FileStream(filepath, FileMode.Open)

            frmViewDocs.RichTextBox1.LoadFile(myfilestream, RichTextBoxStreamType.PlainText)
            myfilestream.Close()
            frmViewDocs.RichTextBox1.Select(0, 0)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        'Kill(UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_in.txt")
        Kill(RiskDataDirectory & "temp_in.txt")

        Try
            StringToPrint = frmViewDocs.RichTextBox1.Text
            frmViewDocs.Text = "Summary of Inputs"
            frmViewDocs.SelectVariablesToolStripMenuItem.Visible = True
            frmViewDocs.Show()
            Exit Sub

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub mnuViewRTB_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuViewRTB.Click
        '=================================================================
        '   Display summary of input and output
        '=================================================================
        If DataFile Is Nothing Then
            MessageBox.Show("No results To show.", "View Results", MessageBoxButtons.OK)
            Exit Sub
        End If

        ' Dim filepath As String = UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_out.txt"
        Dim filepath As String = RiskDataDirectory & "temp_out.txt"
        Dim filepathbatch As String = BatchFileFolder & "\" & VB.Left(DataFile, Len(DataFile) - 4) & ".out"
        Try
            If batch = False Then
                Call view_output(filepath)
                Dim myfilestream As New FileStream(filepath, FileMode.Open)
                frmViewDocs.RichTextBox1.LoadFile(myfilestream, RichTextBoxStreamType.PlainText)
                myfilestream.Close()
                'Kill(UserAppDataFolder & gcs_folder_ext & "\data\" & "temp_out.txt")
                Kill(RiskDataDirectory & "temp_out.txt")
                frmViewDocs.RichTextBox1.Select(0, 0)
                StringToPrint = frmViewDocs.RichTextBox1.Text

                frmViewDocs.Text = "Summary Of Inputs And Results"
                frmViewDocs.SelectVariablesToolStripMenuItem.Visible = True
                frmViewDocs.Show()
            Else
                Call view_output(filepathbatch)
            End If



        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub mnuVisibilityGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuVisibilityGraph.Click
        '*  ===============================================================
        '*  Show a graph of visibility in the upper layer versus time.
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Visibility (m) at " & CStr(MonitorHeight) & " m"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        'Graph_Data_2D(1, Title, Visibility, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, Visibility, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuWallHoG_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWallHoG.Click

        Dim Title, result As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short
        Dim room As Short

        'define variables
        Title = "Wall Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4          '2=user-defined
        MaxYValue = 0

        If NumberRooms > 1 Then
            result = InputBox("Which room Do you want?")
            If IsNumeric(result) Then
                If CInt(result) > 0 And CInt(result) <= NumberRooms Then
                    room = CInt(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        If WallConeDataFile(room) = "" Then WallConeDataFile(room) = "null.txt"
        Call Flame_Spread_Props_graph(room, 4, wallhigh, WallConeDataFile(room), Surface_Emissivity(2, room), ThermalInertiaWall(room), IgTempW(room), WallEffectiveHeatofCombustion(room), WallHeatofGasification(room), AreaUnderWallCurve(room), WallFTP(room), WallQCrit(room), Walln(room))
    End Sub

    Public Sub mnuWallIgn_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWallIgn.Click
        ''*  ======================================================
        ''*  Show a graph of the Cone HRR input curve
        ''*  14/8/98 C A Wade
        ''*  ======================================================

        Dim Title, result As String
        Dim DataShift, MaxYValue As Single
        Dim DataMultiplier As Single
        Dim GraphStyle As Short
        Dim room As Integer

        'define variables
        Title = "Wall Cone HRR (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4          '2=user-defined
        MaxYValue = 0

        If NumberRooms > 1 Then
            result = InputBox("Which room Do you want?")
            If IsNumeric(result) Then
                If CInt(result) > 0 And CInt(result) <= NumberRooms Then
                    room = CInt(result)
                Else
                    room = fireroom
                End If
            Else
                room = fireroom
            End If
        Else
            room = fireroom
        End If

        If WallConeDataFile(room) = "" Then WallConeDataFile(room) = "null.txt"
        Call Flame_Spread_Props_graph(room, 1, wallhigh, WallConeDataFile(room), Surface_Emissivity(2, room), ThermalInertiaWall(room), IgTempW(room), WallEffectiveHeatofCombustion(room), WallHeatofGasification(room), AreaUnderWallCurve(room), WallFTP(room), WallQCrit(room), Walln(room))


    End Sub

    Public Sub mnuwalllower_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuwalllower.Click
        '*  ===============================================================
        '*  Show a graph of vent flow
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Wall Flow -->L (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        'Graph_Data_2D(1, Title, WallFlowtoLower, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, WallFlowtoLower, DataShift, DataMultiplier, timeunit)
    End Sub

    Public Sub mnuWallTempGraph_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWallTempGraph.Click

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Upper Wall Temp (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, Upperwalltemp, DataShift, DataMultiplier, timeunit)

    End Sub



    Public Sub mnuwallupper_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuwallupper.Click
        '*  ===============================================================
        '*  Show a graph of vent flow
        '*  ===============================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Wall Flow -->U (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        'Graph_Data_2D(1, Title, WallFlowtoUpper, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, WallFlowtoUpper, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuWVent_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuWVent.Click
        If IsNothing(WVentlogfile) = True Then
            MsgBox("No wall vent flow data available For this simulation.")
            Exit Sub
        End If
        If WVentlogfile <> "empty" Then
            frmViewDocs.RichTextBox1.LoadFile(WVentlogfile, Windows.Forms.RichTextBoxStreamType.PlainText)
            frmViewDocs.Text = "Summary Of Vent Flow Data"
            frmViewDocs.SelectVariablesToolStripMenuItem.Visible = False
            frmViewDocs.Show()
        Else
            MsgBox("No wall vent flow data available For this simulation.")
            Exit Sub
        End If

    End Sub

    Public Sub mnuXPyrolysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuXPyrolysis.Click
        '*  ===================================================================
        '*  Show a graph of the position of the X_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "X_Pyrolysis (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, X_pyrolysis(), DataShift, DataMultiplier
        'Graph_Data_2D(1, Title, X_pyrolysis, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, X_pyrolysis, DataShift, DataMultiplier, timeunit)
    End Sub



    Public Sub mnuXvelocity_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuXvelocity.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Lateral Velocity (m/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Y_pyrolysis(), DataShift, DataMultiplier
        Graph_Data_3D(2, Title, FlameVelocity, DataShift, DataMultiplier)

    End Sub

    Public Sub mnuY_Burnout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuY_Burnout.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Y Burnout (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Y_pyrolysis(), DataShift, DataMultiplier
        'Graph_Data_2D(1, Title, Y_burnout, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, Y_burnout, DataShift, DataMultiplier, timeunit)

    End Sub

    Public Sub mnuYPyrolysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuYPyrolysis.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Y Pyrolysis (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Y_pyrolysis(), DataShift, DataMultiplier
        'Graph_Data_2D(1, Title, Y_pyrolysis, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, Y_pyrolysis, DataShift, DataMultiplier, timeunit)
    End Sub

    Private Sub Panel3D1_MouseDown(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
        '***************************************************
        '*  Show a popupmenu to terminate the simulation
        '*
        '*  Revised 22 October 1996
        '***************************************************

        If Button = 2 Then
            'If the user clicks the right mouse button
            'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
            'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            'System.Windows.Forms.Cursor.Current = ARROW
            'mnuRun.Enabled = False
            'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.Toolbar2.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            'Me.Toolbar2.Items.Item(8).Visible = False 'start run
            'UPGRADE_ISSUE: MDIForm method MDIFrmMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'PopupMenu(Me.mnuStop)
        End If

    End Sub


    Public Sub mnuYvelocity_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuYvelocity.Click
        '*  ===================================================================
        '*  Show a graph of the position of the Y_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Upward Velocity (m/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_3D(1, Title, FlameVelocity, DataShift, DataMultiplier)

    End Sub

    Public Sub mnuZPyrolysis_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuZPyrolysis.Click
        '*  ===================================================================
        '*  Show a graph of the position of the X_Pyrolysis front versus time
        '*  For use with Quintiere's room corner model
        '*  Revised 19 February 1997 Colleen Wade
        '*  ===================================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()

        'define variables
        Title = "Z_Pyrolysis (m)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data_2D(1, Title, Z_pyrolysis, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_2D(Title, Z_pyrolysis, DataShift, DataMultiplier, timeunit)
    End Sub



    'Private Sub Toolbar2_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
    '    Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
    '    '============================================================
    '    '   program toolbar button functions
    '    '
    '    '   C A Wade 25 November 1997
    '    '============================================================

    '    Select Case Button.Name
    '        Case "mechanical ventilation"
    '            'create a new model
    '            mnuMechExtract_Click(mnuMechExtract, New System.EventArgs())
    '        Case "New"
    '            'create a new model
    '            mnuNewFile_Click(mnuNewFile, New System.EventArgs())
    '        Case "open"
    '            'open an existing model
    '            mnuFileOpen_Click(mnuFileOpen, New System.EventArgs())
    '        Case "save"
    '            'save current model
    '            mnuSaveModelData_Click(mnuSaveModelData, New System.EventArgs())
    '        Case "excel"
    '            'save results to excel 97/2000
    '            mnuExcel_Click(mnuExcel, New System.EventArgs())
    '        Case "print"
    '            'print results
    '            mnuPrinter_Click(mnuPrinter, New System.EventArgs())
    '        Case "run"
    '            'start simulation
    '            mnuRun_Click(mnuRun, New System.EventArgs())
    '        Case "graph"
    '            'show graphs
    '            'UPGRADE_ISSUE: MDIForm method MDIFrmMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            'PopupMenu(mnuGraphs, 0, 0, 0)
    '        Case "Stop"
    '            'stop simulation
    '            'UPGRADE_ISSUE: MDIForm method MDIFrmMain.PopupMenu was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
    '            'PopupMenu(mnuStop, 0, 0, 0)
    '        Case "help"
    '            'show help contents
    '            mnuHelpContents_Click(mnuHelpContents, New System.EventArgs())
    '        Case "roomdim"
    '            'show room dimensions screen
    '            mnuRoom_Click(mnuRoom, New System.EventArgs())
    '        Case "simoptions"
    '            'show room dimensions screen
    '            mnuSimulation_Click(mnuSimulation, New System.EventArgs())
    '        Case "viewresults"
    '            'show room dimensions screen
    '            mnuModelResults_Click()
    '        Case "sprinkler"
    '            'show sprinkler screen
    '            mnuDetector_Click(mnuDetector, New System.EventArgs())
    '        Case "fire"
    '            'show
    '            mnuHeatRelease_Click(mnuHeatRelease, New System.EventArgs())
    '        Case "Toolbar2_Exit"
    '            mnuExit_Click(mnuExit, New System.EventArgs())
    '        Case "inputdata"
    '            mnuViewInput_Click(mnuViewInput, New System.EventArgs())
    '        Case "results"
    '            mnuViewRTB_Click(mnuViewRTB, New System.EventArgs())
    '        Case Else
    '    End Select
    'End Sub

    Public Sub Save_Registry()
        On Error Resume Next
        'SaveSetting("BRISK", "options", "DefaultRiskDataDirectory", DefaultRiskDataDirectory)

        My.Settings.riskdatafolder = DefaultRiskDataDirectory
        My.Settings.Save()

    End Sub

    Public Sub excelbatch()
        '===========================================================
        '   Saves results directly in an excel file
        '   31/1/2000 C A Wade
        ' last edited 4/7/2002
        '===========================================================

        On Error GoTo more

        Dim s, Txt As String
        Dim room As Integer
        Dim k, j, count As Integer

        Dim SaveBox As System.Windows.Forms.Control
        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        'UPGRADE_ISSUE: VB.Control SaveBox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'SaveBox = frmBlank.CMDialog1

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = HOURGLASS
        count = 0

        'dialogbox title
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DialogTitle = "Save Results To Excel File"

        'default filename
        If Len(DataFile) > 4 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'SaveBox.filename = Mid(DataFile, 1, Len(DataFile) - 3) & "xls"
        Else
            DataFile = "New.Mod"
            'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'SaveBox.filename = Mid(DataFile, 1, Len(DataFile) - 3) & "xls"
        End If

        'Set filters
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.Filter = "All Files (*.*)|*.*|Results Files (*.xls)|*.xls"

        'Specify default filter
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.FilterIndex = 2

        'default filename extension
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.DefaultExt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'SaveBox.DefaultExt = "xls"

        'check before overwriting file
        'SaveBox.Flags = OFN_OVERWRITEPROMPT Or OFN_NOCHANGEDIR

        'Display the Save as dialog box.
        'SaveBox.Action = 2

        'declare object variables for microsoft excel
        On Error Resume Next
        'Dim xlapp As excel.Application
        'Dim xlBook As excel.Workbook
        'Dim xlSheet As excel.Worksheet
        Dim xlapp As Object
        Dim xlBook As Object
        Dim xlSheet As Object

        'assign object references to the variables
        'Set xlapp = New Excel.Application
        xlapp = CreateObject("Excel.Application")
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Workbooks. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlBook = xlapp.Workbooks.Add
        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.Worksheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet = xlBook.Worksheets(1) 'new

        If Err.Number <> 0 Then
            GoTo more
        End If

        On Error GoTo excelerrorhandler
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        'xlBook.Worksheets("sheet1").Activate
        'Set xlSheet = xlBook.ActiveSheet

        'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet.Name = "Room 1"

        'assign values to excel cells
        'define a format string
        s = "Scientific"

        'LinkTemp(1 To NumberRooms, 1 To i%)
        Dim dummy() As Double
        Dim X As Object
        For room = 1 To NumberRooms
            Txt = "Time (sec)" & ", Layer(m)" & ",Upper Layer Temp (C)" & ",HRR (kW)" & ",Mass (kg/s)" & ",Plume (kg/s)"
            Txt = Txt & ",Vent Fire (kW)" & ",CO2 Upper(%)" & ",CO Upper (ppm)" & ",O2 Upper (%)" & ",CO2 Lower(%)" & ",CO Lower(ppm)"
            Txt = Txt & ",O2 Lower (%)" & ",FED gases(inc)" & ",Upper Wall Temp (C)" & ",Ceiling Temp (C)" & ",Rad On Floor (kW/m2)" & ",Lower Layer Temp (C)"
            Txt = Txt & ",Lower Wall Temp (C)" & ",Floor Temp (C)" & ",Y Pyrolysis Front (m)" & ",X Pyrolysis Front (m)" & ",Z Pyrolysis Front (m)" & ",Upward Velocity (m/s)"
            Txt = Txt & ",Lateral Velocity (m/s)" & ",Pressure (Pa)" & ",Visibility (m)" & ",Vent Flow To Upper Layer (kg/s)" & ",Vent Flow To Lower Layer (kg/s)"
            Txt = Txt & ",Rad On Target (kW/m2)" & ",FED thermal(inc)" & ",OD upper (1/m)" & ",OD lower (1/m)" & ",k upper (1/m)" & ",k lower (1/m)" & ",Vent Flow To Outside (m3/s)" & ",HCN Upper (ppm)" & ",HCN Lower (ppm)"
            Txt = Txt & ",SPR (m2/kg)" & ",Unexposed Upper Wall Temp (C)" & ",Unexposed Lower Wall Temp (C)" & ",Unexposed Ceiling Temp (C)" & ",Unexposed Floor Temp (C)"
            If room = fireroom Then Txt = Txt & ",Ceiling Jet Temp(C)" & ",Max Ceil Jet Temp(C)" & ",GER" & ",Smoke Detect OD-out (1/m)" & ",Smoke Detect OD-In (1/m)" & ",Link Temp (C)"
            'If room <> fireroom Then Txt$ = Txt$ & ",Wall Flame Height (m)" & ",Flame Temp (K)" & ",Wall Flux (kW/m2)" & Chr(10)


            'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xlSheet.Cells(1, 1).Value = Txt

            If NumberTimeSteps > 0 Then


                Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 48 > 32000 'the maximum number of data points able to be plotted in excel chart
                    ExcelInterval = ExcelInterval * 2
                Loop

                k = 2 'row
                For j = 1 To NumberTimeSteps + 1
                    count = count + 1
                    If System.Math.Round(Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval, 4) = 0 Then
                        'If tim(j, 1) = 10000 Then Stop
                        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                        Me.ToolStripStatusLabel4.Text = "Saving To Excel ... Please Wait - " & VB6.Format(count / (NumberRooms * NumberTimeSteps) * 100, "0") & "%"

                        Txt = VB6.Format(tim(j, 1), "general number") & "," & VB6.Format(layerheight(room, j), s) & "," & VB6.Format(uppertemp(room, j) - 273, s)
                        Txt = Txt & "," & VB6.Format(HeatRelease(room, j, 2), s) & "," & VB6.Format(FuelMassLossRate(j, 1), s) & "," & VB6.Format(massplumeflow(j, room), s)
                        Txt = Txt & "," & VB6.Format(ventfire(room, j), s) & "," & VB6.Format(CO2VolumeFraction(room, j, 1) * 100, s) & "," & VB6.Format(COVolumeFraction(room, j, 1) * 1000000, s)
                        Txt = Txt & "," & VB6.Format(O2VolumeFraction(room, j, 1) * 100, s) & "," & VB6.Format(CO2VolumeFraction(room, j, 2) * 100, s) & "," & VB6.Format(COVolumeFraction(room, j, 2) * 1000000, s)
                        Txt = Txt & "," & VB6.Format(O2VolumeFraction(room, j, 2) * 100, s) & "," & VB6.Format(FEDSum(room, j), s) & "," & VB6.Format(Upperwalltemp(room, j) - 273, s)
                        Txt = Txt & "," & VB6.Format(CeilingTemp(room, j) - 273, s) & "," & VB6.Format(Target(room, j), s) & "," & VB6.Format(lowertemp(room, j) - 273, s)
                        Txt = Txt & "," & VB6.Format(LowerWallTemp(room, j) - 273, s) & "," & VB6.Format(FloorTemp(room, j) - 273, s)
                        Txt = Txt & "," & VB6.Format(Y_pyrolysis(room, j), s) & "," & VB6.Format(X_pyrolysis(room, j), s) & "," & VB6.Format(Z_pyrolysis(room, j), s)
                        Txt = Txt & "," & VB6.Format(FlameVelocity(room, 1, j), s) & "," & VB6.Format(FlameVelocity(room, 2, j), s)
                        Txt = Txt & "," & VB6.Format(RoomPressure(room, j), s) & "," & VB6.Format(Visibility(room, j), s)
                        Txt = Txt & "," & VB6.Format(FlowToUpper(room, j), s) & "," & VB6.Format(FlowToLower(room, j), s)
                        Txt = Txt & "," & VB6.Format(SurfaceRad(room, j), s) & "," & VB6.Format(FEDRadSum(room, j), s)
                        Txt = Txt & "," & VB6.Format(OD_upper(room, j), s) & "," & VB6.Format(OD_lower(room, j), s) & "," & VB6.Format(2.3 * OD_upper(room, j), s) & "," & VB6.Format(2.3 * OD_lower(room, j), s)
                        Txt = Txt & "," & VB6.Format(UFlowToOutside(room, j), s) & "," & VB6.Format(HCNVolumeFraction(room, j, 1) * 1000000, s) & "," & VB6.Format(HCNVolumeFraction(room, j, 2) * 1000000, s)
                        'If room <> fireroom Then Txt$ = Txt$ & "," & Format(wallparam(2, j), s) & "," & Format(wallparam(3, j), s) & "," & Format(wallparam(1, j), s)
                        Txt = Txt & "," & VB6.Format(SPR(room, j), s) & "," & VB6.Format(UnexposedUpperwalltemp(room, j) - 273, s) & "," & VB6.Format(UnexposedLowerwalltemp(room, j) - 273, s) & "," & VB6.Format(UnexposedCeilingtemp(room, j) - 273, s) & "," & VB6.Format(UnexposedFloortemp(room, j) - 273, s)
                        If room = fireroom Then Txt = Txt & "," & VB6.Format(CJetTemp(j, 1, 0) - 273, s) & "," & VB6.Format(CJetTemp(j, 2, 0) - 273, s) & "," & VB6.Format(GlobalER(j), s) & "," & VB6.Format(OD_outside(room, j), s) & "," & VB6.Format(OD_inside(room, j), s) & "," & VB6.Format(LinkTemp(room, j) - 273, s)
                        xlSheet.Cells(k, 1).Value = Txt
                        System.Windows.Forms.Application.DoEvents()
                        k = k + 1
                    End If
                Next j
            End If
            On Error Resume Next

            'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xlSheet.Columns("A:A").Select()
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            With xlapp.Selection
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlGeneral
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .WrapText = False
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .Orientation = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .ShrinkToFit = False
                'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                .MergeCells = False

            End With

            'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Range. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xlapp.Selection.TextToColumns(Destination:=xlSheet.Range("A1"), dataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab_Renamed:=True, Semicolon:=False, Comma:=True, Space_Renamed:=False, Other:=False, FieldInfo:=New Object() {New Object() {1, 1}, New Object() {2, 1}, New Object() {3, 1}, New Object() {4, 1}, New Object() {5, 1}, New Object() {6, 1}, New Object() {7, 1}, New Object() {8, 1}, New Object() {9, 1}, New Object() {10, 1}, New Object() {11, 1}, New Object() {12, 1}, New Object() {13, 1}, New Object() {14, 1}, New Object() {15, 1}, New Object() {16, 1}, New Object() {17, 1}, New Object() {18, 1}, New Object() {19, 1}, New Object() {20, 1}, New Object() {21, 1}, New Object() {22, 1}, New Object() {23, 1}, New Object() {24, 1}, New Object() {25, 1}, New Object() {26, 1}, New Object() {27, 1}, New Object() {28, 1}, New Object() {29, 1}, New Object() {30, 1}, New Object() {31, 1}, New Object() {32, 1}, New Object() {33, 1}, New Object() {34, 1}, New Object() {35, 1}, New Object() {36, 1}, New Object() {37, 1}, New Object() {38, 1}, New Object() {39, 1}, New Object() {40, 1}, New Object() {41, 1}, New Object() {42, 1}, New Object() {43, 1}, New Object() {44, 1}, New Object() {45, 1}, New Object() {46, 1}, New Object() {47, 1}, New Object() {48, 1}, New Object() {49, 1}})

            If room < NumberRooms Then
                'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.Worksheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xlBook.Worksheets.Add()
                'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.ActiveSheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xlSheet = xlBook.ActiveSheet
                'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xlSheet.Name = "Room " & CStr(room + 1)
            End If
        Next room

        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.Worksheets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlBook.Worksheets.Add()
        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.ActiveSheet. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet = xlBook.ActiveSheet
        'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet.Name = "Outside"
        'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet.Cells(1, 1).Value = "Time (sec)" & "," & "Vent Fire (kW)"
        If NumberTimeSteps > 0 Then
            k = 2 'row
            For j = 1 To NumberTimeSteps + 1
                If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                    Txt = VB6.Format(tim(j, 1), s) & "," & VB6.Format(ventfire(room, j), s)
                    'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Cells. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xlSheet.Cells(k, 1).Value = Txt
                    k = k + 1
                End If
            Next j
        End If

        Dim rowcount As Short
        rowcount = k

        'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlSheet.Columns("A:A").Select()
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        With xlapp.Selection
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlGeneral
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .WrapText = False
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .Orientation = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .ShrinkToFit = False
            'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .MergeCells = False
        End With
        'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xlSheet.Range. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Selection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlapp.Selection.TextToColumns(Destination:=xlSheet.Range("A1"), dataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab_Renamed:=True, Semicolon:=False, Comma:=True, Space_Renamed:=False, Other:=False, FieldInfo:=New Object() {New Object() {1, 1}, New Object() {2, 1}, New Object() {3, 1}, New Object() {4, 1}, New Object() {5, 1}, New Object() {6, 1}, New Object() {7, 1}, New Object() {8, 1}, New Object() {9, 1}, New Object() {10, 1}, New Object() {11, 1}, New Object() {12, 1}, New Object() {13, 1}, New Object() {14, 1}, New Object() {15, 1}, New Object() {16, 1}, New Object() {17, 1}, New Object() {18, 1}, New Object() {19, 1}, New Object() {20, 1}, New Object() {21, 1}, New Object() {22, 1}, New Object() {23, 1}, New Object() {24, 1}, New Object() {25, 1}, New Object() {26, 1}, New Object() {27, 1}, New Object() {28, 1}, New Object() {29, 1}, New Object() {30, 1}, New Object() {31, 1}, New Object() {32, 1}, New Object() {33, 1}, New Object() {34, 1}, New Object() {35, 1}, New Object() {36, 1}, New Object() {37, 1}, New Object() {38, 1}, New Object() {39, 1}, New Object() {40, 1}, New Object() {41, 1}, New Object() {42, 1}, New Object() {43, 1}, New Object() {44, 1}, New Object() {45, 1}, New Object() {46, 1}, New Object() {47, 1}, New Object() {48, 1}, New Object() {49, 1}})

        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        Me.ToolStripStatusLabel4.Text = "Saving Excel Charts... Please Wait"

        If frmprintvar.chkLH.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,B:B", "Layer Height (m)", "B2", "A2:A" & CStr(rowcount), "B2:B" & CStr(rowcount))
        If frmprintvar.chkUT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,C:C", "Upper Layer Temp (C)", "C2", "A2:A" & CStr(rowcount), "C2:C" & CStr(rowcount))
        If frmprintvar.chkLT.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,R:R", "Lower Layer Temp (C)", "R2", "A2:A" & CStr(rowcount), "R2:R" & CStr(rowcount))
        If frmprintvar.chkHRR.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,D:D", "Heat Release (kW)", "D2", "A2:A" & CStr(rowcount), "D2:D" & CStr(rowcount))
        If frmprintvar.chkMassLoss.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,E:E", "Mass Loss Rate (kg.s-1)", "E2", "A2:A" & CStr(rowcount), "E2:E" & CStr(rowcount))
        If frmprintvar.chkPlume.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,F:F", "Plume Flow (kg.s-1)", "F2", "A2:A" & CStr(rowcount), "F2:F" & CStr(rowcount))
        If frmprintvar.chkFED.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,N:N", "FED gases", "N2", "A2:A" & CStr(rowcount), "N2:N" & CStr(rowcount))
        If frmprintvar.chkVisi.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AA:AA", "Visibility (m)", "AA2", "A2:A" & CStr(rowcount), "AA2:AA" & CStr(rowcount))
        If frmprintvar.chkFEDrad.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AE:AE", "FED thermal", "AE2", "A2:A" & CStr(rowcount), "AE2:AE" & CStr(rowcount))
        If frmprintvar.chkODupper.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AF:AF", "Upper Layer OD (m-1)", "AF2", "A2:A" & CStr(rowcount), "AF2:AF" & CStr(rowcount))
        If frmprintvar.chkODlower.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AG:AG", "Lower Layer OD (m-1)", "AG2", "A2:A" & CStr(rowcount), "AG2:AG" & CStr(rowcount))
        If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AU:AU", "Smoke Detector OD-OUTSIDE (m-1)", "AU2", "A2:A" & CStr(rowcount), "AU2:AU" & CStr(rowcount))
        If frmprintvar.chkODsmoke.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,AV:AV", "Smoke Detector OD-INSIDE (m-1)", "AV2", "A2:A" & CStr(rowcount), "AV2:AV" & CStr(rowcount))
        If frmprintvar.chkLCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,L:L", "Lower Layer CO (%)", "L2", "A2:A" & CStr(rowcount), "L2:L" & CStr(rowcount))
        If frmprintvar.chkUCO.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,I:I", "Upper Layer CO (%)", "I2", "A2:A" & CStr(rowcount), "I2:I" & CStr(rowcount))
        If frmprintvar.chkPressure.CheckState = System.Windows.Forms.CheckState.Checked Then Call Add_ExcelChart(xlapp, "Room 1", "A:A,Z:Z", "Pressure (Pa)", "Z2", "A2:A" & CStr(rowcount), "Z2:Z" & CStr(rowcount))

        'save the worksheet
        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.DisplayAlerts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlapp.DisplayAlerts = False
        'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.SaveAs. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'xlBook.SaveAs(DataDirectory & SaveBox.filename)
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.DisplayAlerts. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlapp.DisplayAlerts = True

        If Err.Number = 1004 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object SaveBox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'MsgBox(SaveBox.filename & " is already Open. Please close and then try again.", MsgBoxStyle.OKOnly)
            Err.Clear()
        Else
            'MsgBox "Data saved in " & SaveBox.FileTitle, vbInformation + vbOKOnly
        End If
        'close excel
        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlBook.Close(SaveChanges:=False)
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Application. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlapp.Application.Quit()
        'release the objects
        'UPGRADE_NOTE: Object xlapp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlapp = Nothing
        'UPGRADE_NOTE: Object xlBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlBook = Nothing
        'UPGRADE_NOTE: Object xlSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlSheet = Nothing

        'UPGRADE_WARNING: Lower bound of collection MDIFrmMain.StatusBar1.Panels has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        Me.ToolStripStatusLabel4.Text = " "

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = ARROW
        Exit Sub

excelerrorhandler:
        'close excel
        'xlApp.[_quit]
        On Error Resume Next
        'UPGRADE_WARNING: Couldn't resolve default property of object xlBook.Close. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlBook.Close()
        'UPGRADE_WARNING: Couldn't resolve default property of object xlapp.Application. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xlapp.Application.Quit()

        'release the objects
        'UPGRADE_NOTE: Object xlapp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlapp = Nothing
        'UPGRADE_NOTE: Object xlBook may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlBook = Nothing
        'UPGRADE_NOTE: Object xlSheet may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        xlSheet = Nothing
        Me.ToolStripStatusLabel4.Text = " "

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        'User pressed Cancel button
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        Exit Sub
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Call save_output_xml()
    End Sub

    Private Sub MaterialsDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub FuelSourcesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub mnuFireSpec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFireSpec.Click

    End Sub

    Private Sub OpenModelXMLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        OpenFileDialog1.InitialDirectory = DataDirectory
        OpenFileDialog1.Title = "Open Data XML File"
        OpenFileDialog1.Filter = "All Files (*.*)|*.*|Model Files (*.xml)|*.xml"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.FileName = ""

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DataFile = OpenFileDialog1.FileName
        End If

        Call Read_File_xml(DataFile, False)

        DataDirectory = VB.Left(DataFile, Len(DataFile) - Len(Path.GetFileName(DataFile)))

        ChDir((My.Application.Info.DirectoryPath))

    End Sub



    Private Sub Inputs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuInputs.Click

        ' If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"
        frmInputs.Chart1.Titles(1).Text = ProgramTitle & " (" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision & ")"
        frmInputs.Chart2.Titles(3).Text = frmInputs.Chart1.Titles(1).Text
        frmInputs.Chart3.Titles(2).Text = frmInputs.Chart1.Titles(1).Text
        frmInputs.NumericUpDown_Bins.Hide()
        frmInputs.Label45.Hide()
        frmInputs.NumericUpDown1.Hide()
        frmInputs.NumericUpDownRoom.Hide()
        frmInputs.Label18.Hide()
        frmInputs.Label19.Hide()
        frmInputs.Chart1.Hide()
        frmInputs.Chart2.Hide()

        frmInputs.txtSimTime.Text = CStr(SimTime)
        frmInputs.txtExcelInterval.Text = CStr(ExcelInterval)
        frmInputs.txtDisplayInterval.Text = CStr(DisplayInterval)
        frmInputs.txtSeed.Text = CStr(seed)

        If ignitetargets = True Then
            frmInputs.chkIgniteTargets.Checked = CHECKED
        Else
            frmInputs.chkIgniteTargets.Checked = UNCHECKED
        End If
        'frmInputs.BringToFront()
        frmInputs.Show()
    End Sub

    Private Sub NewToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripButton.Click
        mnuNewFile.PerformClick()
    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click
        'mnuFileOpen.PerformClick()
        OpenBaseModelToolStripMenuItem.PerformClick()
    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click
        Try
            SaveBaseToolStripMenuItem.PerformClick()
            'SaveBaseModelAsAsToolStripMenuItem.PerformClick()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb SaveToolStripButton_Click")
        End Try
    End Sub

    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripButton.Click
        mnuPrinter.PerformClick()
    End Sub


    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If gTimeX > 0 And gTimeX < SimTime Then
            ToolStripStatusLabel4.Text = FormatNumber(gTimeX, 1) & " sec"
        ElseIf gTimeX >= SimTime Then
            ToolStripStatusLabel4.Text = "finished"
        Else
            ToolStripStatusLabel4.Text = ""
        End If
    End Sub

    Private Sub SprinklersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmSprinklerList.Show()
    End Sub

    Private Sub CreateSmokeviewFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()

        'Try

        '    If IsNothing(basefile) = True Then
        '        MsgBox("Please first load or save your model.")
        '        Exit Sub
        '    End If

        '    'dialogbox title
        '    SaveBox.Title = "Create Smokeview Geometry File "

        '    'Set filters
        '    SaveBox.Filter = "All Files (*.*)|*.*|Model Files (*.smv)|*.smv"

        '    'Specify default filter
        '    SaveBox.FilterIndex = 2
        '    SaveBox.RestoreDirectory = True

        '    'default filename extension
        '    SaveBox.DefaultExt = "smv"

        '    ' Split the string on the period character
        '    Dim parts As String() = basefile.Split(New Char() {"."c})
        '    DataFile = parts(0) & ".smv"

        '    'default filename
        '    ' If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        '    If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        '    SaveBox.InitialDirectory = RiskDataDirectory
        '    SaveBox.FileName = DataFile

        '    'Display the Save as dialog box.
        '    If SaveBox.ShowDialog() = DialogResult.OK Then
        '        If SaveBox.CheckFileExists = False Then
        '            'create the file
        '            My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
        '        End If
        '        myStream = SaveBox.OpenFile()

        '        If (myStream IsNot Nothing) Then

        '            DataFile = SaveBox.FileName
        '            myStream.Close()
        '            Call Save_Smokeview(SaveBox.FileName)
        '        Else
        '            myStream.Close()
        '        End If
        '    End If
        '    Exit Sub

        'Catch ex As Exception
        '    If (myStream IsNot Nothing) Then
        '        myStream.Close()
        '    End If
        '    MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb CreateSmokeviewFileToolStripMenuItem")
        'End Try
    End Sub


    Private Sub NZBCVM2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NZBCVM2ToolStripMenuItem.Click

        Dim action As Integer
        If NZBCVM2ToolStripMenuItem.Checked = False Then
            action = MsgBox("Changing this setting will change a number of settings and input values in your model to be in accordance with the NZBC Verification Method VM2. Do you want to proceed?", MsgBoxStyle.YesNo)
            If action = MsgBoxResult.Yes Then
                NZBCVM2ToolStripMenuItem.Checked = True
                RiskSimulatorToolStripMenuItem.Checked = False
                Call VM2setup()
            End If
        End If

    End Sub

    Private Sub RiskSimulatorToolStripMenuItem_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RiskSimulatorToolStripMenuItem.CheckStateChanged

    End Sub

    Private Sub RiskSimulatorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RiskSimulatorToolStripMenuItem.Click
        NZBCVM2ToolStripMenuItem.Checked = UNCHECKED
        RiskSimulatorToolStripMenuItem.Checked = CHECKED
        Call VM2setup()
    End Sub

    Public Sub VM2setup()

        'Me.Text = ProgramTitle & " (" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision & ")"
        Me.Text = ProgramTitle & " " & Version

        If RiskSimulatorToolStripMenuItem.Checked = True Then
            VM2 = False
            Me.Text = Me.Text & " - RISK SIMULATOR MODE"

            frmOptions1.Frame26.Enabled = True
            frmOptions1.Frame27.Enabled = True

            frmOptions1.GB_flashovercriteria.Enabled = True

            frmOptions1.chkModGER.Visible = True
            frmOptions1.txtHOCFuel.Enabled = True
            frmOptions1.txtFuelThickness.Enabled = True
            frmOptions1.txtStickSpacing.Enabled = True
            frmOptions1.GB_gasmodel.Enabled = True
            frmOptions1.GB_activity.Enabled = True
            frmOptions1.cmdQuintiere.Enabled = True
            frmOptions1.Frame12.Enabled = True
            frmOptions1.Frame13.Enabled = True
            frmOptions1.Frame17.Enabled = True
            frmOptions1.Frame18.Enabled = True
            frmOptions1.Frame20.Enabled = True
            frmOptions1.Frame22.Enabled = True
            frmOptions1.OptPreFlashover.Enabled = True
            frmOptions1.optPostFlashover.Enabled = True
            frmOptions1.chkHCNcalc.Enabled = True
            frmOptions1.optRCNone.Enabled = True
            frmNewVents.optGlassManualOpen.Enabled = True
            frmNewVents.optGlassAutoBreak.Enabled = True
            frmSmokeDetector.Frame3.Enabled = True
            frmSmokeDetector.chkUseODinside.Enabled = True
            frmSmokeDetector.txtOpticalDensity.Enabled = True
            frmSmokeDetector.txtSDdelay.Enabled = True

            frmOptions1.txtnC.Text = 1
            frmOptions1.txtnH.Text = 2
            frmOptions1.txtnO.Text = 0.5
            frmOptions1.txtnN.Text = 0

            frmItemList.chkAlphaTFire.Enabled = True
            frmItemList.cmdRoomPopulate.Enabled = True
            frmItemList.cmdImport.Enabled = True
            frmItemList.cmdAdd.Enabled = True
            frmItemList.cmdRemove.Enabled = True
            frmItemList.cmdIgniteFirst.Enabled = True
            frmItemList.cmdRandomIgnite.Enabled = True

        Else

            VM2 = True
            Me.Text = Me.Text & " - C/VM2 MODE"
            FEDCO = True
            autopopulate = False
            usepowerlawdesignfire = True
            'useT2fire = True
            'PeakHRR = 20000
            'AlphaT = 0.0469
            'StoreHeight = 5
            preSoot = 0.07
            postSoot = 0.14
            preCO = 0.04
            postCO = 0.4
            EnergyYield(1) = 20
            ObjectRLF(1) = 0.35
            CO2Yield(1) = 1.5

            FOFluxCriteria = False
            'HCNcalc = False
            sootmode = True
            comode = True


            frmItemList.chkAlphaTFire.Enabled = False
            frmItemList.cmdRoomPopulate.Enabled = False
            frmItemList.cmdImport.Enabled = False
            frmItemList.cmdCopy.Enabled = False
            frmItemList.cmdAdd.Enabled = False
            frmItemList.cmdRemove.Enabled = False
            frmItemList.cmdRandomIgnite.Enabled = False
            frmItemList.cmdIgniteFirst.Enabled = False

            frmOptions1.GB_gasmodel.Enabled = False
            frmOptions1.GB_activity.Enabled = False
            frmOptions1.optFEDCO.Checked = True
            frmOptions1.chkModGER.Checked = False
            frmOptions1.chkModGER.Visible = False
            frmOptions1.optFOtemp.Checked = True
            frmOptions1.optCOman.Checked = True
            frmOptions1.optSootman.Checked = True
            frmOptions1.txtpreCO.Text = CStr(preCO) 'preflashover CO yield g/g
            frmOptions1.txtpostCO.Text = CStr(postCO) 'postflashover CO yield g/g
            frmOptions1.txtpreSoot.Text = CStr(preSoot) 'preflashover soot yield g/g
            frmOptions1.txtPostSoot.Text = CStr(postSoot) 'postflashover soot yield g/g
            frmOptions1.txtHOCFuel.Text = CStr(EnergyYield(1))
            frmOptions1.txtHOCFuel.Enabled = False
            frmOptions1.txtFuelThickness.Enabled = False
            frmOptions1.txtStickSpacing.Enabled = False
            frmOptions1.Frame26.Enabled = False
            frmOptions1.Frame27.Enabled = False
            frmOptions1.GB_flashovercriteria.Enabled = False
            frmOptions1.OptPreFlashover.Checked = CHECKED
            frmOptions1.Frame17.Enabled = False
            frmOptions1.Frame18.Enabled = False
            frmOptions1.Frame20.Enabled = False
            frmOptions1.optEnhanceOff.Checked = True
            frmOptions1.OptPreFlashover.Checked = True 'turn postflashover model off
            frmOptions1.OptPreFlashover.Enabled = False
            frmOptions1.optPostFlashover.Enabled = False
            frmOptions1.optJET.Checked = True
            frmOptions1.Frame22.Enabled = False
            frmOptions1.txtnC.Text = 1
            frmOptions1.txtnH.Text = 2
            frmOptions1.txtnO.Text = 0.5
            frmOptions1.txtnN.Text = 0
            frmOptions1.cboABSCoeff.Text = "VM2"
            fueltype = "VM2"
            frmOptions1.chkHCNcalc.Checked = False
            frmOptions1.chkHCNcalc.Enabled = False
            frmOptions1.optRCNone.Checked = False
            frmOptions1.Frame12.Enabled = False
            frmOptions1.Frame13.Enabled = False
            frmOptions1.cmdQuintiere.Enabled = False
            frmOptions1.optRCNone.Enabled = False
            frmNewVents.optGlassManualOpen.Checked = True
            frmNewVents.optGlassManualOpen.Enabled = False
            frmNewVents.optGlassAutoBreak.Enabled = False

            frmSmokeDetector.optSpecifyOD.Checked = True
            frmSmokeDetector.Frame3.Enabled = False
            frmSmokeDetector.txtOpticalDensity.Text = 0.097
            frmSmokeDetector.chkUseODinside.Checked = False
            frmSmokeDetector.chkUseODinside.Enabled = False
            frmSmokeDetector.txtSDRadialDist.Text = 7
            frmSmokeDetector.txtOpticalDensity.Enabled = False

            Dim oSmokedets As List(Of oSmokeDet)
            Dim osddistributions As List(Of oDistribution)
            oSmokedets = SmokeDetDB.GetSmokDets()
            osddistributions = SmokeDetDB.GetSDDistributions
            For Each oSmokeDet In oSmokedets
                oSmokeDet.sdinside = False
            Next
            SmokeDetDB.SaveSmokeDets(oSmokedets, osddistributions)

            Dim oItems As List(Of oItem)
            Dim oItemDistributions As List(Of oDistribution)
            oItems = ItemDB.GetItemsv2
            oItemDistributions = ItemDB.GetItemDistributions
            For Each oitem In oItems
                ObjectRLF(1) = 0.35
                oitem.radlossfrac = ObjectRLF(1)
                'ObjectLHoG(1) = 3
                'oitem.LHoG = ObjectLHoG(1)
                EnergyYield(1) = 20
                oitem.hoc = EnergyYield(1)
                'ObjectMLUA(2, 1) = 250
                'oitem.HRRUA = ObjectMLUA(2, 1)
                'oitem.elevation = 0.3
                CO2Yield(1) = 1.5
                oitem.co2 = CO2Yield(1)
                SootYield(1) = 0.07
                oitem.soot = SootYield(1)
                oitem.xleft = RoomLength(fireroom) / 2
                oitem.ybottom = RoomWidth(fireroom) / 2
            Next

            Dim oVents As List(Of oVent)
            Dim oVentDistributions As List(Of oDistribution)
            oVents = VentDB.GetVents
            oVentDistributions = VentDB.GetVentDistributions
            For Each ovent In oVents
                ovent.probventclosed = 0
                ovent.horeliability = 1
            Next

            For Each oDistribution In oItemDistributions
                If oDistribution.varname = "Radiant Loss Fraction" And oDistribution.id = 1 Then
                    oDistribution.varvalue = 0.35
                End If
                'If oDistribution.varname = "HRRUA" And oDistribution.id = 1 Then
                '    oDistribution.varvalue = 250
                'End If
                'If oDistribution.varname = "Latent Heat of Gasification" And oDistribution.id = 1 Then
                '    oDistribution.varvalue = 3
                'End If
                If oDistribution.varname = "co2" And oDistribution.id = 1 Then
                    oDistribution.varvalue = 1.5
                End If
                If oDistribution.varname = "soot yield" And oDistribution.id = 1 Then
                    oDistribution.varvalue = 0.07
                End If
                If oDistribution.varname = "heat of combustion" And oDistribution.id = 1 Then
                    oDistribution.varvalue = 20
                End If
            Next
            ItemDB.SaveItemsv2(oItems, oItemDistributions)

            nC = 1
            nH = 2
            nO = 0.5
            nN = 0

        End If


    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call mnuRoom.PerformClick()
    End Sub


    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mnuGraphs.PerformClick()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'mnuRun.PerformClick()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        mnuSimulation.PerformClick()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        'mnuExcel.PerformClick()
        Call mnuExcel_Click(sender, e)
    End Sub

    Private Sub TempToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TempToolStripMenuItem.Click
        '===========================================================
        '   Saves glass temperatures directly in an excel file
        '===========================================================

        Dim s As String
        Dim room As Integer
        Dim k, j, count As Integer

        Dim SaveBox As New SaveFileDialog()

        'dialogbox title
        SaveBox.Title = "Save Glass Temperatures to Excel File"

        'Set filters
        SaveBox.Filter = "All Files (*.*)|*.*|xlsx Files (*.xlsx)|*.xlsx|xls Files (*.xls)|*.xls"

        'Specify default filter
        SaveBox.FilterIndex = 2
        SaveBox.RestoreDirectory = True

        'default filename extension
        'SaveBox.DefaultExt = "xlsx"

        'default filename
        'If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        If DataDirectory = "" Then DataDirectory = RiskDataDirectory
        SaveBox.InitialDirectory = DataDirectory

        'default filename
        If Len(DataFile) > 4 Then
            'SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 4) & "_glass.xlsx"
            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 4) & "_glass"
        Else
            DataFile = "new"
            SaveBox.FileName = Mid(DataFile, 1, Len(DataFile) - 4) & "_glass"
        End If

        count = 0

        'Create an array
        Dim DataArray(0 To (NumberTimeSteps * Timestep / ExcelInterval + 1), 0 To 48) As Object

        'On Error GoTo excelerrorhandler
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        Dim oExcel As Object
        Dim oBook As Object
        Dim oSheet As Object

        'Start a new workbook in Excel
        oExcel = CreateObject("Excel.Application")

        oBook = oExcel.Workbooks.Add

        'Add headers to the worksheet on row 1
        oSheet = oBook.Worksheets(1)

        oSheet.name = "Room 1"

        'assign values to excel cells
        'define a format string
        's = "Scientific"
        s = "0.000E+00"

        '==========
        Dim dummy() As Double, vcounter As Integer = 6
        Dim X As Object
        Dim rowcount As Short

        For room = 1 To NumberRooms

            DataArray(0, 0) = "Time (sec)"
            DataArray(0, 1) = "Layer (m)"
            DataArray(0, 2) = "Upper Layer Temp (C)"
            DataArray(0, 3) = "Upper Layer Temp (C)"
            DataArray(0, 4) = "Lower Layer Temp (C)"
            DataArray(0, 5) = "HRR (kW)"

            'Txt = "Time (sec)" & ",Layer (m)" & ",Upper Layer Temp (C)" & ",Lower Layer Temp (C)" & ",HRR (kW)"
            For P = 2 To NumberRooms + 1
                For vent = 1 To NumberVents(room, P)
                    If AutoBreakGlass(room, P, vent) = True Then
                        'Txt = Txt & ",Exposure Temp (C)" & ",Emissivity"
                        DataArray(0, vcounter) = "Exposure Temp (C)"
                        vcounter = vcounter + 1
                        DataArray(0, vcounter) = "Emissivity"
                        vcounter = vcounter + 1

                        For node = 1 To UBound(GLASSTempHistory, 4)
                            'Txt = Txt & ",Node " & CStr(node) & " (" & CStr(room) & CStr(P) & CStr(vent) & ")"
                            DataArray(0, vcounter) = "Node " & CStr(node) & " (" & CStr(room) & CStr(P) & CStr(vent) & ")"
                            vcounter = vcounter + 1

                        Next node
                    End If
                Next vent
            Next P

            'oSheet.Cells._Default(1, 1).Value = Txt

            If NumberTimeSteps > 0 Then

                Do While NumberTimeSteps * Timestep / ExcelInterval * NumberRooms * 37 > 32000
                    ExcelInterval = ExcelInterval * 2
                Loop

                k = 2 'row
                For j = 1 To NumberTimeSteps

                    count = count + 1
                    If System.Math.Round(Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval, 4) = 0 Then
                        Me.ToolStripStatusLabel4.Text = "Saving to Excel ... Please Wait - " & VB6.Format(count / (NumberRooms * NumberTimeSteps) * 100, "0") & "%"

                        DataArray(k - 1, 0) = VB6.Format(tim(j, 1), "general number")
                        DataArray(k - 1, 1) = VB6.Format(layerheight(room, j), s)
                        DataArray(k - 1, 2) = VB6.Format(uppertemp(room, j) - 273, s)
                        DataArray(k - 1, 3) = VB6.Format(lowertemp(room, j) - 273, s)
                        DataArray(k - 1, 4) = VB6.Format(HeatRelease(room, j, 2), s)

                        vcounter = 5

                        'Txt = VB6.Format(tim(j, 1), s) & "," & VB6.Format(layerheight(room, j), s) & "," & VB6.Format(uppertemp(room, j) - 273, s)
                        'Txt = Txt & "," & VB6.Format(lowertemp(room, j) - 273, s) & "," & VB6.Format(HeatRelease(room, j, 2), s)
                        For P = 2 To NumberRooms + 1
                            For vent = 1 To NumberVents(room, P)
                                If AutoBreakGlass(room, P, vent) = True Then
                                    If tim(j, 1) > VentBreakTime(room, P, vent) Then

                                        DataArray(k - 1, vcounter) = ""
                                        vcounter = vcounter + 1
                                        DataArray(k - 1, vcounter) = ""
                                        vcounter = vcounter + 1

                                        'Txt = Txt & "," & ","
                                    Else
                                        DataArray(k - 1, vcounter) = VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 1, j) - 273, s)
                                        vcounter = vcounter + 1
                                        DataArray(k - 1, vcounter) = VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 2, j), s)
                                        vcounter = vcounter + 1

                                        ' Txt = Txt & "," & VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 1, j) - 273, s) & "," & VB6.Format(GLASSOtherHistory(room, NumberRooms + 1, vent, 2, j), s)
                                        For node = 1 To UBound(GLASSTempHistory, 4)
                                            DataArray(k - 1, vcounter) = VB6.Format(GLASSTempHistory(room, NumberRooms + 1, vent, node, j) - 273, s)
                                            vcounter = vcounter + 1

                                            'Txt = Txt & "," & VB6.Format(GLASSTempHistory(room, NumberRooms + 1, vent, node, j) - 273, s)
                                        Next node
                                    End If
                                End If
                            Next vent
                        Next P
                        'xlSheet.Cells._Default(k, 1).Value = Txt
                        System.Windows.Forms.Application.DoEvents()
                        k = k + 1
                    End If
                Next j
                rowcount = k
            End If

            'Transfer the array to the worksheet starting at cell A1
            oSheet.Range("A1").Resize(rowcount, vcounter - 1).Value = DataArray

            If room < NumberRooms Then
                oBook.Worksheets.add()
                oSheet = oBook.ActiveSheet
                oSheet.Name = "Room " & CStr(room + 1)
            End If

            On Error Resume Next

            oSheet.Columns._Default("A:A").Select()

            With oExcel.Selection
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlGeneral
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlBottom
                .WrapText = False
                .Orientation = 0
                .ShrinkToFit = False
                .MergeCells = False

            End With

            oExcel.Selection.TextToColumns(Destination:=oSheet.Range("A1"), dataType:=Microsoft.Office.Interop.Excel.XlTextParsingType.xlDelimited, TextQualifier:=Microsoft.Office.Interop.Excel.Constants.xlDoubleQuote, ConsecutiveDelimiter:=False, Tab_Renamed:=True, Semicolon:=False, Comma:=True, Space_Renamed:=False, Other:=False)

            If room < NumberRooms Then
                oBook.Worksheets.add()
                oSheet = oBook.ActiveSheet
                oSheet.Name = "Room " & CStr(room + 1)
            End If


        Next room

        'Save the Workbook and Quit Excel
        oBook.SaveAs(SaveBox.FileName)
        oExcel.Quit()

        If Err.Number = 1004 Then
            MsgBox(SaveBox.FileName & " is already Open. Please close and then try again.", MsgBoxStyle.OkOnly)
            Err.Clear()
        Else
            MsgBox("Data saved in " & SaveBox.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        Exit Sub

excelerrorhandler:
        'close excel
        On Error Resume Next
        oBook.Close()
        oExcel.Application.Quit()

        'release the objects
        oExcel = Nothing
        oBook = Nothing
        oSheet = Nothing

more:
        If Err.Number <> 32755 Then MsgBox(ErrorToString(Err.Number))
        'User pressed Cancel button
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Arrow
        Exit Sub

    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub SelectFireToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ''*  ========================================================
        ''*  This procedure updates the text boxes when the form
        ''*  gets the focus.
        ''*  ========================================================

        'Dim i As Integer
        'Dim id As Short

        'frmFire.Show()

        ''Set Room = frmFire.lstFireRoom2

        ''On Error Resume Next
        ''frmBlank.Show()
        ''frmFireSpecification.Show

        ''update list box
        ''frmFireSpecification.lstObjectID.Clear
        'frmFire.lstObjectID.Items.Clear()
        'If NumberObjects > 0 Then
        '    id = 1
        '    For i = 1 To NumberObjects
        '        frmFire.lstObjectID.Items.Add("Fire " & CStr(i) & " of " & CStr(NumberObjects))
        '    Next i
        'Else
        '    frmFire.lstObjectID.Items.Add("None")
        'End If
        'frmFire.lstObjectID.SelectedIndex = 0
        'frmFire.lstFireRoom2.Items.Clear()
        'For i = 1 To NumberRooms
        '    frmFire.lstFireRoom2.Items.Add(CStr(i))
        'Next i
        'frmFire.lstFireRoom2.SelectedIndex = fireroom - 1

        'frmFire.lstFireRoom2.Items.Clear()
        'For i = 1 To NumberRooms
        '    frmFire.lstFireRoom2.Items.Add(CStr(i))
        'Next i
        'frmFire.lstFireRoom2.SelectedIndex = fireroom - 1
        ''frmFire.lstFireRoom2_Click
        ''frmFire.lstFireRoom2.Refresh
        ''frmFire.lstFireRoom2.SetFocus

        ''If id = 0 Then id = 1

        'Dim fireobject As Object
        'fireobject = frmFire

        'fireobject.txtEnergyYield.Text = VB6.Format(EnergyYield(id), "0.0")
        'fireobject.txtCO2Yield.Text = VB6.Format(CO2Yield(id), "0.000")
        'fireobject.txtSootYield.Text = VB6.Format(SootYield(id), "0.000")
        'fireobject.txtHCNYield.Text = VB6.Format(HCNuserYield(id), "0.000")

        ''frmFire.txtWaterVaporYield.text = Format(WaterVaporYield(id), "0.000")
        'fireobject.lblObjectDescription.Text = ObjectDescription(id)
        'fireobject.txtFireHeight.Text = VB6.Format(FireHeight(id), "0.000")

        'frmFire.Show()
    End Sub

    Private Sub SelectItemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"

        frmItemList.Show()
        frmItemList.BringToFront()
    End Sub

    Private Sub PowerLawGrowthToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"
        frmPowerlaw.Show()
        frmPowerlaw.BringToFront()
    End Sub
    Private Sub Read_dumpfile(ByVal thisiteration As Integer)

        Try

            Dim s As String
            Dim k, j, count As Integer
            Dim dummystring As String
            Dim dummyint As Integer
            Dim room As Integer
            Dim ver As String
            Dim binfile As String = RiskDataDirectory & "dumpdata.dat"

            If My.Computer.FileSystem.FileExists(binfile) = False Then Exit Sub
            Dim binaryin As New BinaryReader(New FileStream(binfile, FileMode.Open, FileAccess.Read))

            s = "0.000E+00"

            Dim rowcount As Short

            'Me.ToolStripStatusLabel1.Text = "Saving to Excel ... Please Wait. "
            For count = 1 To thisiteration

                'read in data from binary file
                dummystring = binaryin.ReadString

                If dummystring <> "START" Then
                    binaryin.Close()
                    MsgBox("Iteration " & thisiteration.ToString & ": encountered error reading dump.dat file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                    Exit Sub
                End If


                ver = binaryin.ReadString

                If InStrRev(ver, "201") = 0 Then
                    'old dump file with version data
                Else
                    dummyint = binaryin.ReadInt32
                End If

                'dummyint = binaryin.ReadInt32
                NumSprinklers = binaryin.ReadInt32
                NumberTimeSteps = binaryin.ReadInt32

                Dim DataArray(0 To NumberTimeSteps + 2, 0 To 50) As Object

                NumberRooms = binaryin.ReadInt32
                fireroom = binaryin.ReadInt32

                Size_Arrays()

                If NumberTimeSteps > 0 Then

                    For room = 1 To NumberRooms
                        k = 2 'row
                        For j = 1 To NumberTimeSteps + 1

                            If k > DataArray.GetLength(0) Then Exit For


                            If count = thisiteration Then

                                tim(j, 1) = binaryin.ReadDouble
                                layerheight(room, j) = binaryin.ReadDouble
                                uppertemp(room, j) = binaryin.ReadDouble
                                If CSng(ver) > CSng("2013.16") Then
                                    HeatRelease(room, j, 1) = binaryin.ReadDouble
                                End If
                                HeatRelease(room, j, 2) = binaryin.ReadDouble
                                FuelMassLossRate(j, 1) = binaryin.ReadDouble
                                massplumeflow(j, room) = binaryin.ReadDouble
                                ventfire(room, j) = binaryin.ReadDouble
                                CO2MassFraction(room, j, 1) = binaryin.ReadDouble
                                COMassFraction(room, j, 1) = binaryin.ReadDouble
                                O2MassFraction(room, j, 1) = binaryin.ReadDouble
                                CO2MassFraction(room, j, 2) = binaryin.ReadDouble
                                COMassFraction(room, j, 2) = binaryin.ReadDouble
                                O2MassFraction(room, j, 2) = binaryin.ReadDouble
                                FEDSum(room, j) = binaryin.ReadDouble
                                Upperwalltemp(room, j) = binaryin.ReadDouble
                                CeilingTemp(room, j) = binaryin.ReadDouble
                                Target(room, j) = binaryin.ReadDouble
                                lowertemp(room, j) = binaryin.ReadDouble
                                LowerWallTemp(room, j) = binaryin.ReadDouble
                                FloorTemp(room, j) = binaryin.ReadDouble
                                Y_pyrolysis(room, j) = binaryin.ReadDouble
                                X_pyrolysis(room, j) = binaryin.ReadDouble
                                Z_pyrolysis(room, j) = binaryin.ReadDouble
                                FlameVelocity(room, 1, j) = binaryin.ReadDouble
                                FlameVelocity(room, 2, j) = binaryin.ReadDouble
                                RoomPressure(room, j) = binaryin.ReadDouble
                                Visibility(room, j) = binaryin.ReadDouble
                                FlowToUpper(room, j) = binaryin.ReadDouble
                                FlowToLower(room, j) = binaryin.ReadDouble
                                SurfaceRad(room, j) = binaryin.ReadDouble
                                FEDRadSum(room, j) = binaryin.ReadDouble
                                OD_upper(room, j) = binaryin.ReadDouble
                                OD_lower(room, j) = binaryin.ReadDouble
                                UFlowToOutside(room, j) = binaryin.ReadDouble
                                HCNMassFraction(room, j, 1) = binaryin.ReadDouble
                                HCNMassFraction(room, j, 2) = binaryin.ReadDouble
                                SPR(room, j) = binaryin.ReadDouble
                                UnexposedUpperwalltemp(room, j) = binaryin.ReadDouble
                                UnexposedLowerwalltemp(room, j) = binaryin.ReadDouble
                                UnexposedCeilingtemp(room, j) = binaryin.ReadDouble
                                UnexposedFloortemp(room, j) = binaryin.ReadDouble
                                upperemissivity(room, j) = binaryin.ReadDouble
                                If CSng(ver) > 2014.13 Then
                                    NHL(1, room, j) = binaryin.ReadDouble
                                End If

                                If room = fireroom Then
                                    If CSng(ver) > 2012.51 Then
                                        CJetTemp(j, 0, 0) = binaryin.ReadDouble
                                    End If

                                    CJetTemp(j, 1, 0) = binaryin.ReadDouble
                                    CJetTemp(j, 2, 0) = binaryin.ReadDouble
                                    GlobalER(j) = binaryin.ReadDouble
                                    OD_outside(room, j) = binaryin.ReadDouble
                                    OD_inside(room, j) = binaryin.ReadDouble
                                    If NumSprinklers > 0 Then
                                        SprinkTemp(1, j) = binaryin.ReadDouble 'only reads data fir spr 1
                                    Else
                                        binaryin.ReadDouble()
                                    End If


                                End If

                                ventfire(NumberRooms + 1, j) = binaryin.ReadDouble
                            Else
                                'discard data
                                For m = 1 To 42
                                    binaryin.ReadDouble()
                                Next
                                Dim dummy As Double
                                If room = fireroom Then
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                    dummy = binaryin.ReadDouble()
                                End If

                                dummy = binaryin.ReadDouble()

                            End If

                            k = k + 1

                        Next

                    Next

                    rowcount = k

                End If


            Next count

            binaryin.Close()

            Volume_Fractions()

            MsgBox("Iteration " & thisiteration.ToString & " loaded.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)

        Catch ex As Exception

            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb Read_dumpfile")

        End Try

    End Sub

    Private Sub LoadIterationToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadIterationToolStripMenuItem.Click
        Try

            OpenFileDialog1.InitialDirectory = RiskDataDirectory
            OpenFileDialog1.Title = "Open Input XML File for requested iteration"
            OpenFileDialog1.Filter = "Input Files (input*.xml)|input*.xml"
            OpenFileDialog1.FilterIndex = 2
            OpenFileDialog1.FileName = ""

            If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                DataFile = OpenFileDialog1.FileName
            End If

            Call Read_File_xml(DataFile, False)

            DataDirectory = VB.Left(DataFile, Len(DataFile) - Len(Path.GetFileName(DataFile)))

            ChDir((My.Application.Info.DirectoryPath))
            Dim inputname As String = Path.GetFileName(DataFile)

            'extract the iteration number from the end of the filename
            Dim thisiteration As Integer
            Dim it As String = inputname.Replace("input", "")
            it = it.Replace(".xml", "")
            thisiteration = CInt(it)

            Read_dumpfile(thisiteration)
            Number_TimeSteps()
            If NumberTimeSteps > 1 Then
                mnuExcel.Enabled = True
            Else
                mnuExcel.Enabled = False
            End If
            Call createsmokeviewdata()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb LoadIterationToolStripMenuItem")
        End Try
    End Sub

    Private Sub VentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmVentList.Show()
    End Sub

    Private Sub VentsWallToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        frmVentList.Show()
    End Sub

    Private Sub OtherParsmetersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub OpenBaseModelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenBaseModelToolStripMenuItem.Click

        Try
            OpenFileDialog1.CheckPathExists = True
            OpenFileDialog1.InitialDirectory = DefaultRiskDataDirectory

            OpenFileDialog1.Title = "Open Base File"
            OpenFileDialog1.Filter = "All Files (*.*)|*.*|Base Files (base*.xml)|base*.xml"
            OpenFileDialog1.FilterIndex = 2
            OpenFileDialog1.FileName = ""
            OpenFileDialog1.RestoreDirectory = True

            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                basefile = OpenFileDialog1.FileName
            Else
                Exit Sub
            End If

            frmInputs.rtb_log.Clear() 'clear previous content in log screen

            RiskDataDirectory = Path.GetDirectoryName(basefile) & "\"
            Dim remember As String = basefile
            Call frmInputs.Read_BaseFile_xml(basefile, False)
            basefile = remember

            Erase mc_item_hoc
            Erase mc_item_co2
            Erase mc_item_lhog
            Erase mc_item_RLF
            Erase mc_item_soot
            Erase mc_item_hrrua

            'read inputs with distributions, not included in basemodel file
            Call DistributionClass.Read_Distributions()

            'check to see if input files exist in data folder
            If upgrade = False Then 'only load previous run data if it was created using the same B-RISK version
                If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                    If My.Computer.FileSystem.FileExists(RiskDataDirectory & "sampledata.dat") = True Then
                        'read data

                        Call frmInputs.Read_OutputFile_xml(basefile)
                        Call frmInputs.Read_InputFile_dat()
                    End If
                End If
            End If

            ProjectDirectory = RiskDataDirectory

            frmInputs.ToolStripStatusLabel1.Text = OpenFileDialog1.SafeFileName & " loaded"
            ToolStripStatusLabel1.Text = OpenFileDialog1.SafeFileName & " loaded"
            frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory

            ChDir((My.Application.Info.DirectoryPath))

            Dim datafile As String = ProjectDirectory & "input1.xml"


            'THIS OPENS PREVIOUS RUN, INCLUDING PREVIOUS INPUT DATA
            'If File.Exists(datafile) Then

            '    Call Read_File_xml(datafile, True)

            '    Dim inputname As String = Path.GetFileName(datafile)

            '    'extract the iteration number from the end of the filename
            '    Dim thisiteration As Integer
            '    Dim it As String = inputname.Replace("input", "")
            '    it = it.Replace(".xml", "")
            '    thisiteration = CInt(it)

            '    Read_dumpfile(thisiteration)
            '    Number_TimeSteps()
            '    If NumberTimeSteps > 1 Then
            '        mnuExcel.Enabled = True
            '    Else
            '        mnuExcel.Enabled = False
            '    End If
            '    Call createsmokeviewdata()
            'End If

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb OpenBaseModelToolStripMenuItem_Click")
        End Try

    End Sub

    Private Sub OpenModFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WallVentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WallVentsToolStripMenuItem.Click
        frmVentList.Show()
        frmVentList.BringToFront()

    End Sub

    Private Sub CeilingVentsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeilingVentsToolStripMenuItem.Click

        mnuRoom.PerformClick()
        'frmDescribeRoom._SSTab1_0.SelectedIndex = 3
        'frmDescribeRoom.Show()

    End Sub

    Private Sub MechanicalFansToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MechanicalFansToolStripMenuItem.Click
        '*  ==============================================================
        '*  Show form for data input
        '*  Created 19 March 1998
        '*  ==============================================================
        'Dim room As Integer

        'frmExtract.Show()

        'If fanon(1) = True Then
        '    frmExtract.optFanOn.Checked = True
        'Else
        '    frmExtract.optFanOn.Checked = False
        'End If

        'frmExtract.lstRoomID.Items.Clear()
        'For room = 1 To NumberRooms
        '    frmExtract.lstRoomID.Items.Add(CStr(room))
        'Next room

        'frmExtract.txtOperationTime.Text = VB6.Format(ExtractStartTime(1), "0.0")
        'frmExtract.txtFanElevation.Text = VB6.Format(FanElevation(1), "0.00")
        'frmExtract.txtCeilingExtractRate.Text = VB6.Format(ExtractRate(1), "0.00")
        'frmExtract.txtNumberFans.Text = VB6.Format(NumberFans(1), "0")
        'frmExtract.txtMaxPressure.Text = VB6.Format(MaxPressure(1), "0")
        'frmExtract.optExtract.Checked = Extract(1)
        'frmExtract.optPressurise.Checked = Not Extract(1)
        'If UseFanCurve(1) = True Then
        '    frmExtract.chkUseFanCurve.CheckState = System.Windows.Forms.CheckState.Checked
        'Else
        '    frmExtract.chkUseFanCurve.CheckState = System.Windows.Forms.CheckState.Unchecked
        'End If
        'If fanon(1) = True Then
        '    frmExtract.optFanOn.Checked = True
        'Else
        '    frmExtract.optFanOn.Checked = False
        'End If

        'frmExtract.lstRoomID.SelectedIndex = 0
    End Sub

    Private Sub SmokeDetectorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmokeDetectorsToolStripMenuItem.Click
        '*  ==============================================================
        '*  Show form for data input of smoke detectors.
        '*  ==============================================================
        Dim room As Integer
        'frmBlank.Show()
        frmSmokeDetector.Show()
        frmSmokeDetector.lstRoomSD.Items.Clear()
        For room = 1 To NumberRooms
            frmSmokeDetector.lstRoomSD.Items.Add(CStr(room))
        Next room
        frmSmokeDetector.lstRoomSD.SelectedIndex = 0
        frmSmokeDetector.txtOpticalDensity.Text = VB6.Format(SmokeOD(1), "0.000")
        frmSmokeDetector.txtSDdelay.Text = VB6.Format(SDdelay(1), "0.0")
        frmSmokeDetector.txtDetSensitivity.Text = VB6.Format(DetSensitivity(1), "0.0")
        frmSmokeDetector.txtSDRadialDist.Text = VB6.Format(SDRadialDist(1), "0.000")
        frmSmokeDetector.txtSDdepth.Text = VB6.Format(SDdepth(1), "0.000")
        If SDinside(1) = True Then
            frmSmokeDetector.chkUseODinside.CheckState = System.Windows.Forms.CheckState.Checked
        Else
            frmSmokeDetector.chkUseODinside.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
        If HaveSD(1) = True Then
            frmSmokeDetector.chkSmokeDetector.CheckState = System.Windows.Forms.CheckState.Checked
            frmSmokeDetector.Frame2.Visible = True
            frmSmokeDetector.Frame3.Visible = True
            frmSmokeDetector.Frame4.Visible = True
        Else
            frmSmokeDetector.chkSmokeDetector.CheckState = System.Windows.Forms.CheckState.Unchecked
            frmSmokeDetector.Frame2.Visible = False
            frmSmokeDetector.Frame3.Visible = False
            frmSmokeDetector.Frame4.Visible = False
        End If
        If SpecifyOD(1) = True Then
            If SmokeOD(1) = SDNormalDefault Then frmSmokeDetector.optSDnormal.Checked = True
            If SmokeOD(1) = SDhighDefault Then frmSmokeDetector.optSDhigh.Checked = True
            If SmokeOD(1) = SDvhighDefault Then frmSmokeDetector.optSDvhigh.Checked = True
            frmSmokeDetector.optSpecifyOD.Checked = True
            frmSmokeDetector.optSDnormal.Enabled = False
            frmSmokeDetector.optSDhigh.Enabled = False
            frmSmokeDetector.optSDvhigh.Enabled = False
            frmSmokeDetector.txtOpticalDensity.Enabled = True
            frmSmokeDetector.txtDetSensitivity.Enabled = False
        Else
            frmSmokeDetector.optAS1603.Checked = True
            frmSmokeDetector.optSDnormal.Enabled = True
            frmSmokeDetector.optSDhigh.Enabled = True
            frmSmokeDetector.optSDvhigh.Enabled = True
            frmSmokeDetector.txtOpticalDensity.Enabled = False
            frmSmokeDetector.txtDetSensitivity.Enabled = False
        End If
    End Sub

    Private Sub SprinklersHeatDetectorsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SprinklersHeatDetectorsToolStripMenuItem.Click

        ProjectDirectory = RiskDataDirectory

        oDistributions = DistributionClass.GetDistributions()
        For Each oDistribution In oDistributions
            If oDistribution.varname = "Sprinkler Reliability" Then
                frmSprinklerList.txtSprReliability.Text = oDistribution.varvalue
                If oDistribution.distribution <> "None" Then
                    frmSprinklerList.txtSprReliability.BackColor = distbackcolor
                Else
                    frmSprinklerList.txtSprReliability.BackColor = distnobackcolor
                End If
            End If
            If oDistribution.varname = "Sprinkler Suppression Probability" Then
                frmSprinklerList.txtSprSuppressProb.Text = oDistribution.varvalue
                If oDistribution.distribution <> "None" Then
                    frmSprinklerList.txtSprSuppressProb.BackColor = distbackcolor
                Else
                    frmSprinklerList.txtSprSuppressProb.BackColor = distnobackcolor
                End If
            End If
            If oDistribution.varname = "Sprinkler Cooling Coefficient" Then
                frmSprinklerList.txtSprCoolingCoeff.Text = oDistribution.varvalue
                If oDistribution.distribution <> "None" Then
                    frmSprinklerList.txtSprCoolingCoeff.BackColor = distbackcolor
                Else
                    frmSprinklerList.txtSprCoolingCoeff.BackColor = distnobackcolor
                End If
            End If
        Next

        If calc_sprdist = True Then
            frmSprinklerList.chkCalcSprinkRadialDist.Checked = True
        Else
            frmSprinklerList.chkCalcSprinkRadialDist.Checked = False
        End If

        frmSprinklerList.Show()
        frmSprinklerList.BringToFront()

    End Sub

    Private Sub ToolStripButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRun.Click

        frmInputs.Show()
        frmInputs.StartToolStripLabel1.PerformClick()

    End Sub

    Private Sub ToolStripButton2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuStop.Click

        frmInputs.StopToolStripButton1.PerformClick()

    End Sub

    Private Sub RoomDimensionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        mnuRoom.PerformClick()
        'frmDescribeRoom._SSTab1_0.SelectedIndex = 0
        'frmDescribeRoom.Show()


    End Sub



    Private Sub ViewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ViewToolStripMenuItem.Click
        'Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()

        'Try
        '    ' If IsNothing(DataFile) = True Then
        '    If IsNothing(basefile) = True Then
        '        MsgBox("Please first load or save your model.")
        '        Exit Sub
        '    End If
        '    'dialogbox title
        '    SaveBox.Title = "Create Smokeview Geometry File "

        '    'Set filters
        '    SaveBox.Filter = "All Files (*.*)|*.*|Model Files (*.smv)|*.smv"

        '    'Specify default filter
        '    SaveBox.FilterIndex = 2
        '    SaveBox.RestoreDirectory = True

        '    'default filename extension
        '    SaveBox.DefaultExt = "smv"

        '    ' Split the string on the period character
        '    Dim parts As String() = basefile.Split(New Char() {"."c})
        '    DataFile = parts(0) & ".smv"

        '    'default filename
        '    ' If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
        '    If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        '    SaveBox.InitialDirectory = RiskDataDirectory
        '    SaveBox.FileName = DataFile

        '    'Display the Save as dialog box.
        '    If SaveBox.ShowDialog() = DialogResult.OK Then
        '        If SaveBox.CheckFileExists = False Then
        '            'create the file
        '            My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
        '        End If
        '        myStream = SaveBox.OpenFile()

        '        If (myStream IsNot Nothing) Then

        '            DataFile = SaveBox.FileName
        '            myStream.Close()
        '            Call Save_Smokeview(SaveBox.FileName)
        '        Else
        '            myStream.Close()
        '        End If
        '    End If

        'Catch ex As Exception
        '    MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb CreateSmokeviewFileToolStripMenuItem")
        'End Try
    End Sub

    Private Sub CreateSmokeViewDataFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateSmokeViewDataFileToolStripMenuItem.Click
        '===========================================================
        '   Saves results in a .csv file to be read by smokeview
        '===========================================================

        Dim s, Txt As String
        Dim room As Integer
        Dim k, j As Integer
        Dim count As Integer = 0
        Dim SaveBox As New SaveFileDialog()
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter

        Try

            DataFile = Mid(basefile, 1, Len(basefile) - 4) & "_zone.csv"

            SaveBox.FileName = DataFile
            'Display the Save as dialog box.

            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", True)
            End If


            oWrite = oFile.CreateText(DataFile)

            'define a format string
            's = "Scientific"
            s = "General Number"

            Txt = "s,"
            For i = 1 To NumberRooms
                Txt = Txt & "C, C, m, Pa, 1 / m, 1 / m,"
            Next
            For i = 1 To NumberObjects
                Txt = Txt & " kW, m, m, m ^ 2,"
            Next
            For i = 1 To number_vents
                Txt = Txt & " m ^ 2,"
            Next

            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            Txt = Txt & " m ^ 2,"
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            Txt = "Time,"
            For i = 1 To NumberRooms
                Txt = Txt & "ULT_" & i.ToString & ",LLT_" & i.ToString & ",HGT_" & i.ToString & ",PRS_" & i.ToString & ",ULOD_" & i.ToString & ",LLOD_" & i.ToString & ","
            Next
            For i = 1 To NumberObjects
                Txt = Txt & "HRR_" & i.ToString & ",FLHGT_" & i.ToString & ",FBASE_" & i.ToString & ",FAREA_" & i.ToString & ","
            Next
            For i = 1 To number_vents
                Txt = Txt & "HVENT_" & i.ToString & ","
            Next

            Dim m As Integer = 0
            For room = 1 To NumberRooms + 1
                For i = 1 To NumberRooms + 1
                    'If room < i Then
                    If NumberCVents(i, room) > 0 Then
                        For k = 1 To NumberCVents(i, room)
                            m = m + 1
                            Txt = Txt & "VVENT_" & m.ToString & ","
                        Next
                    End If
                    'End If
                Next
            Next

            oWrite.WriteLine("{0,1}", Txt)

            For j = 1 To NumberTimeSteps
                If Int(tim(j, 1) / ExcelInterval) - tim(j, 1) / ExcelInterval = 0 Then
                    Txt = VB6.Format(tim(j, 1), s) & ","
                    For i = 1 To NumberRooms
                        Txt = Txt & VB6.Format(uppertemp(i, j) - 273, s) & ","
                        Txt = Txt & VB6.Format(lowertemp(i, j) - 273, s) & ","
                        Txt = Txt & VB6.Format(layerheight(i, j), s) & ","
                        Txt = Txt & VB6.Format(RoomPressure(i, j), s) & ","
                        Txt = Txt & VB6.Format(OD_upper(i, j), s) & ","
                        Txt = Txt & VB6.Format(OD_lower(i, j), s) & ","
                    Next

                    For i = 1 To NumberObjects
                        'If i = 1 Then
                        Txt = Txt & VB6.Format(HeatRelease(fireroom, j, 2), s) & ","
                        Txt = Txt & VB6.Format(0.235 * HeatRelease(fireroom, j, 2) ^ 0.4 - 1.2 * System.Math.Sqrt(4 * ObjLength(i) * ObjWidth(i) / PI), s) & "," 'FLHGT_
                        Txt = Txt & VB6.Format(FireHeight(i), s) & "," 'FBASE_
                        Txt = Txt & VB6.Format(ObjLength(i) * ObjWidth(i), s) & "," 'FAREA_
                        'Else
                        'Txt = Txt & VB6.Format(HeatRelease(fireroom, j, 2), s) & ","
                        'Txt = Txt & VB6.Format(1, s) & ","
                        'Txt = Txt & VB6.Format(FireHeight(i), s) & ","
                        'Txt = Txt & VB6.Format(1, s) & ","
                        'End If
                    Next

                    For room = 1 To NumberRooms
                        For i = 2 To NumberRooms + 1
                            If room < i Then
                                If NumberVents(room, i) > 0 Then
                                    For k = 1 To NumberVents(room, i)
                                        Txt = Txt & VB6.Format(VentHeight(room, i, k) * VentWidth(room, i, k), s) & ","
                                    Next
                                End If
                            End If
                        Next
                    Next

                    For room = 1 To NumberRooms + 1
                        For i = 1 To NumberRooms + 1
                            'If room < i Then
                            If NumberCVents(i, room) > 0 Then
                                For k = 1 To NumberCVents(i, room)
                                    Txt = Txt & CVentArea(i, room, k) & ","
                                Next
                            End If
                            'End If
                        Next
                    Next

                    oWrite.WriteLine("{0,1}", Txt)

                End If
            Next

            oWrite.Close()
            oWrite.Dispose()

            MsgBox("Data saved in " & SaveBox.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly)

            Exit Sub

        Catch ex As Exception
            oWrite.Close()
            oWrite.Dispose()
        End Try

    End Sub

    Private Sub TenabilityParametersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TenabilityParametersToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 4
        frmOptions1.BringToFront()
        frmOptions1.Show()

    End Sub

    Private Sub FEDEgressPathToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FEDEgressPathToolStripMenuItem.Click
        frmFEDpath.Show()
    End Sub

    Private Sub SelectFireToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectFireToolStripMenuItem.Click
        'If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
        If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"

        oDistributions = DistributionClass.GetDistributions()
        For Each oDistribution In oDistributions
            If oDistribution.varname = "Fire Load Energy Density" Then
                frmItemList.txtFLED.Text = CSng(oDistribution.varvalue)
                If oDistribution.distribution <> "None" Then
                    frmItemList.txtFLED.BackColor = distbackcolor
                Else
                    frmItemList.txtFLED.BackColor = distnobackcolor
                End If
            End If
        Next

        frmItemList.lstFireRoom2.Items.Clear()
        For room = 1 To NumberRooms
            frmItemList.lstFireRoom2.Items.Add(CStr(room))
        Next room
        frmItemList.lstFireRoom2.SelectedIndex = fireroom - 1
        frmItemList.BringToFront()

        frmItemList.Show()
        frmItemList.BringToFront()

    End Sub

    Private Sub AdditionalCombustionParametersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AdditionalCombustionParametersToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 2
        frmOptions1.Show()
    End Sub

    Private Sub COSootToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles COSootToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 8
        frmOptions1.Show()
    End Sub

    Private Sub FireObjectDatabaseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FireObjectDatabaseToolStripMenuItem1.Click
        frmFireObjectDB.ToolStripButton2.Visible = False
        frmFireObjectDB.ToolStripLabel1.Visible = True
        frmFireObjectDB.Show()
    End Sub

    Private Sub SurfaceMaterialsDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SurfaceMaterialsDatabaseToolStripMenuItem.Click
        frmMaterials.Show()
    End Sub


    Private Sub SetingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SetingsToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 5
        frmOptions1.Show()
    End Sub

    Private Sub MaterialConeFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaterialConeFileToolStripMenuItem.Click
        '====================================================================
        '   read data from a cone .VEC file and create a material .txt file
        '====================================================================

        Dim Openbox As System.Windows.Forms.Control
        Dim numfiles As Integer
        Dim str_Renamed As String
        Dim i As Short
        Dim vectorfile As String = ""
        Dim A As Short
        'UPGRADE_ISSUE: MSComDlg.CommonDialog control CMDialog1 was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E047632-2D91-44D6-B2A3-0801707AF686"'
        'UPGRADE_ISSUE: VB.Control Openbox was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'Openbox = frmBlank.CMDialog1

        'CancelError is True
        On Error GoTo errhandler
10:
        str_Renamed = InputBox("How many cone calorimeter data files do you want to enter?", "Add files in order of increasing exernal heat flux", CStr(1), VB6.TwipsToPixelsX(500))
        If IsNumeric(str_Renamed) = True Then
            numfiles = CInt(str_Renamed)
        Else
            If str_Renamed = "" Then
                'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
                Err.Number = DialogResult.Cancel : GoTo errhandler
            End If 'user pressed cancel
            Err.Number = 13
            GoTo errhandler 'generate error type mismatch
        End If

        'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        'System.Windows.Forms.Cursor.Current = HOURGLASS

        'set directory to open in
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.InitDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.InitDir = UserAppDataFolder & gcs_folder_ext & "\cone\"

        'Set filters
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.Filter = "All Files (*.*)|*.*|Vector Files (*.vec)|*.vec|FDMS Files(*.exp)|*.exp|"

        'Specify default filter
        'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.FilterIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Openbox.FilterIndex = 2

        For i = 1 To numfiles
            'dialogbox title
            'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.DialogTitle. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Openbox.DialogTitle = "Open Vector File " & CStr(i) & " : Add files in order of increasing exernal heat flux"

            'Display the Open dialog box
            If i > 1 Then
                'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
                'A = Len(vectorfile) - Len(Dir(vectorfile))
                A = Len(vectorfile) - Len(Path.GetFileName(vectorfile))
                'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.InitDir. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'Openbox.InitDir = VB.Left(vectorfile, A)
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Openbox.filename = ""
            'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.Action. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'Openbox.Action = 1
            If Err.Number > 0 Then GoTo errhandler
            'Call the open file procedure
            'UPGRADE_WARNING: Couldn't resolve default property of object Openbox.filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'vectorfile = Openbox.filename
            Call Open_Vector_File(vectorfile, i, numfiles)
        Next i
        ChDir((My.Application.Info.DirectoryPath))

        Exit Sub

errhandler:
        'User pressed Cancel button
        If Err.Number = 13 Then 'type mismatch
            MsgBox("Please enter a number greater than zero.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
            Err.Clear()
            GoTo 10
            'UPGRADE_WARNING: The CommonDialog CancelError property is not supported in Visual Basic .NET. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8B377936-3DF7-4745-AA26-DD00FA5B9BE1"'
        ElseIf Err.Number = DialogResult.Cancel Then
            Err.Clear()
            Exit Sub
        Else
            MsgBox("Error - operation cancelled", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly)
        End If
        FileClose(1)
        Err.Clear()
        Exit Sub
    End Sub

    Private Sub mnuSmokeView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSmokeView.Click
        Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()

        Try
            ' If IsNothing(DataFile) = True Then
            If IsNothing(basefile) = True Or basefile = "" Then
                MsgBox("Please load or save your model first.")
                Exit Sub
            End If
            'dialogbox title
            SaveBox.Title = "Create Smokeview Geometry File "

            'Set filters
            SaveBox.Filter = "All Files (*.*)|*.*|Smokeview Files (*.smv)|*.smv"

            'Specify default filter
            SaveBox.FilterIndex = 2
            SaveBox.RestoreDirectory = True

            'default filename extension
            SaveBox.DefaultExt = "smv"

            ' Split the string on the period character
            'Dim parts As String() = basefile.Split(New Char() {"."c})
            'DataFile = parts(0) & ".smv"

            DataFile = basefile.Replace(".xml", ".smv")

            If basefile.Length > 0 Then
                RiskDataDirectory = FileIO.FileSystem.GetParentPath(basefile) & "\"
            End If

            'default filename
            ' If DataDirectory = "" Then DataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "data\"
            'If RiskDataDirectory = "" Then RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\basemodel_default\"
            If RiskDataDirectory = "" Then RiskDataDirectory = DefaultRiskDataDirectory & "basemodel_default\"
            SaveBox.InitialDirectory = RiskDataDirectory
            SaveBox.FileName = DataFile

            'Display the Save as dialog box.
            'If SaveBox.ShowDialog() = DialogResult.OK Then
            If SaveBox.CheckFileExists = False Then
                'create the file
                My.Computer.FileSystem.WriteAllText(SaveBox.FileName, "", False)
            End If

            myStream = SaveBox.OpenFile()

            If (myStream IsNot Nothing) Then

                'DataFile = SaveBox.FileName
                myStream.Close()
                Call Save_Smokeview(SaveBox.FileName)
            Else
                myStream.Close()
            End If

        Catch ex As Exception
            myStream.Close()
            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb mnuSmokeView")
        End Try
    End Sub

    Private Sub AmbientConditionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AmbientConditionsToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 3
        frmOptions1.BringToFront()
        frmOptions1.Show()
    End Sub

    Private Sub SolversToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SolversToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 6
        frmOptions1.BringToFront()
        frmOptions1.Show()

    End Sub

    Private Sub CompartmentEffectsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CompartmentEffectsToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 9
        frmOptions1.Show()
    End Sub

    Private Sub PostflashoverBehaviourToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PostflashoverBehaviourToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 7
        frmOptions1.Show()
    End Sub

    Private Sub ModelPhysicsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModelPhysicsToolStripMenuItem.Click
        mnuSimulation.PerformClick()
        frmOptions1.SSTab2.SelectedIndex = 1
        frmOptions1.Show()
    End Sub

    Public Sub SaveBaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBaseToolStripMenuItem.Click

        Dim getfolder As String = "basemodel_default"
        Dim oldfolder As String = RiskDataDirectory
        Dim answer As String
        Dim msg As String = "Would you like to copy this project to the default riskdata folder?"

        If frmInputs.txtBaseName.Text = "" Then
            MsgBox("Please enter a file base name for your project", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim filename As String = "basemodel_" & frmInputs.txtBaseName.Text
        filename = filename.Replace(" ", "_")

        If DefaultRiskDataDirectory & filename & "\" = RiskDataDirectory Then
            RiskDataDirectory = DefaultRiskDataDirectory & filename & "\"
            basefile = DefaultRiskDataDirectory & filename & "\" & filename & ".xml"
        Else
            answer = MsgBox(msg, MB_ICONQUESTION + MB_YESNO, ProgramTitle)
            If answer = IDYES Then
                RiskDataDirectory = DefaultRiskDataDirectory & filename & "\"
                basefile = DefaultRiskDataDirectory & filename & "\" & filename & ".xml"

                If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                    My.Computer.FileSystem.DeleteDirectory(RiskDataDirectory, FileIO.DeleteDirectoryOption.DeleteAllContents)
                End If

            Else

            End If
        End If


        'RiskDataDirectory = DefaultRiskDataDirectory & filename & "\"
        'basefile = DefaultRiskDataDirectory & filename & "\" & filename & ".xml"


        Try

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = True Then
                'Stop
                'End If
                ' If basefile <> "" Then
                Call frmInputs.Save_BaseFile_xml(basefile)
                If mc_go = False Then
                    frmInputs.ToolStripStatusLabel1.Text = basefile & " saved"
                    frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory
                End If
                ToolStripStatusLabel1.Text = ""
                Exit Sub
            End If


            'Dim newfolder As String = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\" & getfolder & "\"
            'RiskDataDirectory = newfolder

            'getfolder = "basemodel_" & Convert.ToString(frmInputs.txtBaseName.Text)
            getfolder = filename
            getfolder = getfolder.Replace(" ", "_")

            If getfolder.Length > 0 Then
                'RiskDataDirectory = UserAppDataFolder & gcs_folder_ext & "\" & "riskdata\" & getfolder & "\"
                RiskDataDirectory = DefaultRiskDataDirectory & getfolder & "\"
            End If

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = False Then
                'create folder
                My.Computer.FileSystem.CreateDirectory(RiskDataDirectory)
            End If

            If RiskDataDirectory <> oldfolder And basefile <> "" Then
                ' My.Computer.FileSystem.CopyFile(basefile, RiskDataDirectory & getfolder & ".xml", True)
                If My.Computer.FileSystem.FileExists(oldfolder & filename & ".xml") Then
                    My.Computer.FileSystem.CopyFile(oldfolder & filename & ".xml", basefile, True)
                Else
                    'My.Computer.FileSystem.CopyFile(oldbasefile, basefile, True)
                    My.Computer.FileSystem.CopyFile(DataFolder & "basemodel_default.xml", basefile, True)
                End If

                My.Computer.FileSystem.CopyFile(oldfolder & "items.xml", RiskDataDirectory & "items.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "sprinklers.xml", RiskDataDirectory & "sprinklers.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "distributions.xml", RiskDataDirectory & "distributions.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "vents.xml", RiskDataDirectory & "vents.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "cvents.xml", RiskDataDirectory & "cvents.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "smokedets.xml", RiskDataDirectory & "smokedets.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "fans.xml", RiskDataDirectory & "fans.xml", True)
                If My.Computer.FileSystem.FileExists(oldfolder & "rooms.xml") Then
                    My.Computer.FileSystem.CopyFile(oldfolder & "rooms.xml", RiskDataDirectory & "rooms.xml", True)
                End If
            End If

            Dim stemp As String = ""
            stemp = frmInputs.txtBaseName.Text
            stemp = stemp.Replace(" ", "_")
            frmInputs.txtBaseName.Text = stemp

            ' basefile = RiskDataDirectory & getfolder & ".xml"
            Call frmInputs.Save_BaseFile_xml(basefile)

            frmInputs.ToolStripStatusLabel1.Text = basefile & " saved"
            ToolStripStatusLabel1.Text = ""
            frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory


        Catch ex As Exception

            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb SaveToolStripButton")

        End Try
    End Sub

    Private Sub SmokeDetectorsnewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmokeDetectorsnewToolStripMenuItem.Click
        ProjectDirectory = RiskDataDirectory

        oDistributions = DistributionClass.GetDistributions()
        For Each oDistribution In oDistributions
            If oDistribution.varname = "Smoke Detector Reliability" Then
                frmSmokeDetList.txtSmokeDetReliability.Text = oDistribution.varvalue
                If oDistribution.distribution <> "None" Then
                    frmSmokeDetList.txtSmokeDetReliability.BackColor = distbackcolor
                Else
                    frmSmokeDetList.txtSmokeDetReliability.BackColor = distnobackcolor
                End If
            End If

        Next

        If calc_sddist = True Then
            frmSmokeDetList.chkCalcSRadialDist.Checked = True
        Else
            frmSmokeDetList.chkCalcSRadialDist.Checked = False
        End If

        frmSmokeDetList.Show()
        frmSmokeDetector.BringToFront()

    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        ProjectDirectory = RiskDataDirectory

        oDistributions = DistributionClass.GetDistributions()
        For Each oDistribution In oDistributions
            If oDistribution.varname = "Mechanical Ventilation Reliability" Then
                frmFanList.txtFanReliability.Text = oDistribution.varvalue
                If oDistribution.distribution <> "None" Then
                    frmFanList.txtFanReliability.BackColor = distbackcolor
                Else
                    frmFanList.txtFanReliability.BackColor = distnobackcolor
                End If
            End If

        Next
        frmFanList.Show()
        frmFanList.BringToFront()

    End Sub

    Private Sub OpenBRANZFIREModFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenBRANZFIREModFileToolStripMenuItem.Click
        '*  ==========================================================
        '*  This subprocedure shows a custom dialog box for
        '*  opening a file.
        '*  ==========================================================

        OpenFileDialog1.InitialDirectory = DataDirectory
        OpenFileDialog1.Title = "Open Data MOD File"
        OpenFileDialog1.Filter = "All Files (*.*)|*.*|Model Files (*.mod)|*.mod"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.FileName = ""

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DataFile = OpenFileDialog1.FileName
        End If

        Call Open_File(DataFile, False)

        DataDirectory = VB.Left(DataFile, Len(DataFile) - Len(Path.GetFileName(DataFile)))

        'ChDir((My.Application.Info.DirectoryPath))
    End Sub

    Private Sub mnuOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOptions.Click

    End Sub

    Private Sub SaveBaseModelAsAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveBaseModelAsAsToolStripMenuItem.Click
        'Dim myStream As Stream
        Dim SaveBox As New SaveFileDialog()
        Dim getfolder As String = "basemodel_default"

        Try

            'dialogbox title
            SaveBox.Title = "Save Base Model as XML"
            SaveBox.CheckPathExists = True

            'Set filters
            SaveBox.Filter = "All Files (*.*)|*.*|Base Files (base*.xml)|base*.xml"

            'Specify default filter
            SaveBox.FilterIndex = 2
            SaveBox.RestoreDirectory = True

            'default filename extension
            SaveBox.DefaultExt = "xml"

            Dim oldfolder As String = RiskDataDirectory

            getfolder = "basemodel_" & Convert.ToString(frmInputs.txtBaseName.Text)
            'Dim oldfile As String = "basemodel_" & Convert.ToString(frmInputs.txtBaseName.Text) & ".xml"
            Dim newfilename As String = ""

            SaveBox.FileName = getfolder

            If SaveBox.ShowDialog() = Windows.Forms.DialogResult.OK Then
                newfilename = SaveBox.FileName
            Else
                Exit Sub
            End If

            If newfilename = "" Then newfilename = RiskDataDirectory & "basemodel_new"
            Dim strname = My.Computer.FileSystem.GetName(newfilename)
            Dim path = My.Computer.FileSystem.GetParentPath(newfilename)
            Dim strname1 = ""
            If Mid(path, path.Length) = "\" Then
                strname1 = strname.Replace(".xml", "\")
            Else
                strname1 = "\" & strname.Replace(".xml", "\")
            End If

            newfilename = path & strname1 & strname
            ' newfilename = newfilename.Replace("basemodel", "basemodel")

            If newfilename.Length > 0 Then
                RiskDataDirectory = FileIO.FileSystem.GetParentPath(newfilename) & "\"
            End If

            If My.Computer.FileSystem.DirectoryExists(RiskDataDirectory) = False Then
                'create folder
                My.Computer.FileSystem.CreateDirectory(RiskDataDirectory)
            End If

            Dim str As Object = My.Computer.FileSystem.GetFileInfo(newfilename).Name
            str = str.Replace("basemodel_", "")
            str = str.Replace(".xml", "")
            frmInputs.txtBaseName.Text = str

            If (RiskDataDirectory <> oldfolder) And (newfilename <> "") Then

                If basefile <> "" Then My.Computer.FileSystem.CopyFile(basefile, newfilename, True)
                My.Computer.FileSystem.CopyFile(oldfolder & "items.xml", RiskDataDirectory & "items.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "sprinklers.xml", RiskDataDirectory & "sprinklers.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "distributions.xml", RiskDataDirectory & "distributions.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "vents.xml", RiskDataDirectory & "vents.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "cvents.xml", RiskDataDirectory & "cvents.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "smokedets.xml", RiskDataDirectory & "smokedets.xml", True)
                My.Computer.FileSystem.CopyFile(oldfolder & "fans.xml", RiskDataDirectory & "fans.xml", True)
                If My.Computer.FileSystem.FileExists(oldfolder & "rooms.xml") Then
                    My.Computer.FileSystem.CopyFile(oldfolder & "rooms.xml", RiskDataDirectory & "rooms.xml", True)
                End If

            Else
                'Exit Sub
            End If



            Dim stemp As String = ""
            stemp = frmInputs.txtBaseName.Text
            stemp = stemp.Replace(" ", "_")
            frmInputs.txtBaseName.Text = stemp

            Call frmInputs.Save_BaseFile_xml(newfilename)
            basefile = newfilename
            frmInputs.Label5.Text = "Current project folder is: " & RiskDataDirectory
            frmInputs.ToolStripStatusLabel1.Text = basefile & " saved"
            ToolStripStatusLabel1.Text = ""

        Catch ex As Exception

            MsgBox(Err.Description, MsgBoxStyle.Exclamation, "Exception in MDIFRMMA.vb SaveToolStripButton")

        End Try
    End Sub


    Private Sub ToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem3.Click

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        If DetectorType(fireroom) = 0 Then
            MsgBox("You need to have specified a thermal detector or sprinkler in the fire room to record ceiling jet temperatures")
            Exit Sub
        End If
        'define variables
        Title = "Ceiling Jet Temp at Link (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data(1, Title, CJetTemp, DataShift, DataMultiplier, GraphStyle, MaxYValue)
        Graph_Data_3D_CJET(1, Title, CJetTemp, DataShift, DataMultiplier)

    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        If DetectorType(fireroom) = 0 Then
            MsgBox("You need to have specified a thermal detector or sprinkler in the fire room to record ceiling jet temperatures")
            Exit Sub
        End If
        'define variables
        Title = "Ceiling Jet Velocity at Link (m/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data

        Graph_Data_3D_CJET(0, Title, CJetTemp, DataShift, DataMultiplier)
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        '*  ==========================================================
        '*  Show a graph ceiling jet temp at the location of the link versus time
        '*  ==========================================================

        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'frmBlank.Show()
        If DetectorType(fireroom) = 0 Then
            MsgBox("You need to have specified a thermal detector or sprinkler in the fire room to record ceiling jet temperatures")
            Exit Sub
        End If

        'define variables
        Title = "Maximum Ceiling Jet Temp at Radial Position of Link (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_3D_CJET(2, Title, CJetTemp, DataShift, DataMultiplier)

    End Sub

    Private Sub ChangeFireDatabaseFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeFireDatabaseFileToolStripMenuItem.Click

        DBDirectory = UserPersonalDataFolder & gcs_folder_ext & "\dbases"

        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.InitialDirectory = DBDirectory

        OpenFileDialog1.Title = "Open Fire Database File"
        OpenFileDialog1.Filter = "All Files (*.*)|*.*| Database Files (*.mdb)|*.mdb"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.RestoreDirectory = True

        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            FireDatabaseName = OpenFileDialog1.FileName
        ElseIf Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        My.Settings("fireConnectionString") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FireDatabaseName



    End Sub

    Private Sub RoomsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoomsToolStripMenuItem.Click
        ProjectDirectory = RiskDataDirectory
        frmRoomList.Show()
    End Sub

    'Private Sub WallToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    '    frmDescribeRoom._SSTab1_0.SelectedIndex = 0
    '    If frmDescribeRoom._lstRoomID_1.Items.Count > 0 Then
    '        If frmDescribeRoom._lstRoomID_1.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_2.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_3.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        End If
    '    End If

    '    frmDescribeRoom.Show()
    'End Sub

    'Private Sub CeilingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    frmDescribeRoom._SSTab1_0.SelectedIndex = 1
    '    If frmDescribeRoom._lstRoomID_1.Items.Count > 0 Then
    '        If frmDescribeRoom._lstRoomID_1.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_2.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_3.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        End If
    '    End If

    '    frmDescribeRoom.Show()
    'End Sub

    'Private Sub FloorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    frmDescribeRoom._SSTab1_0.SelectedIndex = 2
    '    If frmDescribeRoom._lstRoomID_1.Items.Count > 0 Then
    '        If frmDescribeRoom._lstRoomID_1.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_1.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_2.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_2.SelectedIndex = 0
    '        End If
    '        If frmDescribeRoom._lstRoomID_3.SelectedIndex = -1 Then
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        Else
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = -1
    '            frmDescribeRoom._lstRoomID_3.SelectedIndex = 0
    '        End If
    '    End If

    '    frmDescribeRoom.Show()
    'End Sub

    Private Sub CeilingVentsNewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeilingVentsNewToolStripMenuItem.Click
        frmCVentList.Show()
        frmCVentList.BringToFront()
    End Sub

    Private Sub mnuNormalisedHeatLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNormalisedHeatLoad.Click
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier, XMultiplier, sum As Double
        Dim GraphStyle As Short
        Dim data(0 To 4, 0 To fireroom, 0 To NumberTimeSteps + 1) As Double
        Dim room As Integer = fireroom

        'define variables
        Title = "Normalised Heat Load (s1/2 K)"
        DataShift = 0
        DataMultiplier = 1
        XMultiplier = 1 / 60
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        sum = 0
        For i = 1 To NumberTimeSteps
            If IsNothing(QCeiling) Then Exit For
            'calculate NHL by integrating the flux
            If QCeiling(room, i) < 0 Then
                sum = sum - QCeiling(room, i) * Timestep / Sqrt(ThermalInertiaCeiling(room))
            End If
            data(1, room, i) = sum
        Next
        NHLmax(1) = data(1, room, NumberTimeSteps)
        sum = 0
        For i = 1 To NumberTimeSteps
            If IsNothing(QUpperWall) Then Exit For
            'calculate NHL by integrating the flux
            If QUpperWall(room, i) < 0 Then
                sum = sum - QUpperWall(room, i) * Timestep / Sqrt(ThermalInertiaWall(room))
            End If
            data(2, room, i) = sum
        Next
        NHLmax(2) = data(2, room, NumberTimeSteps)
        sum = 0
        For i = 1 To NumberTimeSteps
            If IsNothing(QLowerWall) Then Exit For
            'calculate NHL by integrating the flux
            If QLowerWall(room, i) < 0 Then
                sum = sum - QLowerWall(room, i) * Timestep / Sqrt(ThermalInertiaWall(room))
            End If
            data(3, room, i) = sum
        Next
        NHLmax(3) = data(3, room, NumberTimeSteps)
        sum = 0
        For i = 1 To NumberTimeSteps
            If IsNothing(QFloor) Then Exit For
            'calculate NHL by integrating the flux
            If QFloor(room, i) < 0 Then
                sum = sum - QFloor(room, i) * Timestep / Sqrt(ThermalInertiaFloor(room))
            End If
            data(4, room, i) = sum
        Next
        NHLmax(4) = data(4, room, NumberTimeSteps)
        sum = 0

        If IsNothing(NHL) Then Exit Sub

        For i = 1 To NumberTimeSteps
            data(0, room, i) = NHL(1, room, i)
        Next
        'calculate a max weighted average NHL
        'NHLmax(0) = NHLmax(1) * RoomFloorArea(room) + NHLmax(4) * RoomFloorArea(room) + NHLmax(2) * 2 * (RoomWidth(room) + RoomLength(room)) * (RoomHeight(room) + MinStudHeight(room)) / 2
        'NHLmax(0) = NHLmax(0) / (2 * RoomFloorArea(room) + 2 * (RoomWidth(room) + RoomLength(room)) * (RoomHeight(room) + MinStudHeight(room)) / 2)


        Graph_Data_NHLmulti(0, Title, data, DataShift, DataMultiplier, XMultiplier)

    End Sub

    Private Sub mnuUpperWallFlux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Upper Wall Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QUpperWall, DataShift, DataMultiplier, timeunit)

    End Sub

    Private Sub mnuCeilingFlux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub mnuLowerWallFlux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Lower Wall Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QLowerWall, DataShift, DataMultiplier, timeunit)
    End Sub

    Private Sub mnuFloorFlux_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Floor Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QFloor, DataShift, DataMultiplier, timeunit)
    End Sub



    Private Sub mnuSetRiskdataFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSetRiskdataFolder.Click
        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.ShowDialog()
        ProjectDirectory = FolderBrowserDialog1.SelectedPath

        If ProjectDirectory = "" Then Exit Sub

        DefaultRiskDataDirectory = ProjectDirectory & "\riskdata\"
        Call Save_Registry()
        MsgBox("New default B-RISK riskdata folder is " & DefaultRiskDataDirectory)
        frmInputs.Label4.Text = "Default riskdata folder is: " & DefaultRiskDataDirectory
    End Sub

    Private Sub CeilingNodeTemperaturesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub mnuCeilingAST_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '*  =======================================================
        '*  Show a graph of the ceiling AST versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Ceiling AST (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QCeilingAST(i, 2, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click
        'ceiling net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Ceiling Total Penetrating Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = -1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QCeiling, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click
        'upper wall net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Upper Wall Total Penetrating Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = -1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QUpperWall, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click
        'lower wall net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Lower Wall Total Penetrating Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = -1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QLowerWall, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem9.Click
        'floor net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Floor Total Penetrating Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = -1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        Graph_Data_2D(Title, QFloor, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem10.Click
        'ceiling net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Ceiling Incident Radiant Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QCeilingAST(i, 0, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem11.Click
        'upper wall net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Upper Wall Incident Radiant Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QUpperWallAST(i, 0, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem12.Click
        'lower wall net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Lower Wall Incident Radiant Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QLowerWallAST(i, 0, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        'floor net heat flux
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Floor Incident Radiant Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QFloorAST(i, 0, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        '*  =======================================================
        '*  Show a graph of the ceiling AST versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Ceiling AST (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QCeilingAST(i, 2, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        '*  =======================================================
        '*  Show a graph of the upper wall AST versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Upper Wall AST (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QUpperWallAST(i, 2, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        '*  =======================================================
        '*  Show a graph of the lower wall AST versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Lower Wall AST (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QLowerWallAST(i, 2, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        '*  =======================================================
        '*  Show a graph of the floor AST versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Floor AST (C)"
        DataShift = -273
        DataMultiplier = 1
        '    MaxYValue = 0
        'call procedure to plot data

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QFloorAST(i, 2, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub
    Private Sub ToolStripMenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        '*  ===============================================================
        '*  Show a graph of Nodal Ceiling Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Ceiling Temperatures (C)"
            DataShift = -273
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            frmPlot.Chart1.Series.Clear()
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberCeilingNodes As Integer
            NumberCeilingNodes = UBound(CeilingNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberCeilingNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = NumberCeilingNodes

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1

            '-----------------------------
            'Start a new workbook in Excel
            'Dim oExcel As Object
            'Dim oBook As Object
            'Dim oSheet As Object
            'Dim getname As String
            'Dim outputinterval As Integer = CInt(frmInputs.txtExcelInterval.Text)
            'Dim tcounter As Integer = Ceiling(SimTime / outputinterval)
            'Dim colcounter As Integer
            'Dim DataArray2(0 To tcounter + 1, 0 To 2 * numsets - 1) As Object
            'oExcel = CreateObject("Excel.Application")
            'oBook = oExcel.Workbooks.Add
            'getname = RiskDataDirectory & "excel_Cnodetemps_" & Convert.ToString(frmInputs.txtBaseName.Text)
            ''Add headers to the worksheet on row 1
            'oSheet = oBook.Worksheets(1)
            'oSheet.Range("A1").Value = "time (sec)"
            'oSheet.Range("B1").Value = "node depth (mm)--->"
            'colcounter = 2
            '-----------------------------

            For i = 0 To 2 * numsets - 1 Step 2

                If HaveCeilingSubstrate(idr) = True Then
                    depth = (curve - 1) * CeilingThickness(idr) / (Ceilingnodes - 1)
                    If depth > CeilingThickness(idr) Then
                        depth = CeilingThickness(idr) + (curve - Ceilingnodes) * CeilingSubThickness(idr) / (Ceilingnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * CeilingThickness(idr) / (NumberCeilingNodes - 1)
                End If

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                frmPlot.Chart1.Series.Add(chdata(0, i))
                frmPlot.Chart1.Series(chdata(0, i)).ChartType = SeriesChartType.FastLine
                maxtemp = 0


                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * CeilingNode(idr, curve, j) + DataShift) 'data to be plotted
                    frmPlot.Chart1.Series(chdata(0, i)).Points.AddXY(tim(j, 1), chdata(j, i + 1))
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)

                    '-----------------------------
                    'If j = 1 Then
                    '    DataArray2(1, 0) = 0
                    '    DataArray2(1, colcounter) = chdata(j, i + 1)
                    'End If
                    'If j Mod outputinterval = 0 Then
                    '    DataArray2(j / outputinterval + 1, 0) = Timestep * j
                    '    DataArray2(j / outputinterval + 1, colcounter) = chdata(j, i + 1)
                    'End If
                    '-----------------------------
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1

                '-----------------------------
                'DataArray2(0, colcounter) = depth
                'colcounter = colcounter + 1
                '-----------------------------
            Next i

            '-----------------------------
            'oSheet.Range("A2").Resize(tcounter + 2, 2 * numsets).Value = DataArray2



            'If Not System.String.IsNullOrEmpty(getname) Then oBook.SaveAs(getname)
            'oBook.Close(SaveChanges:=False)
            'oExcel.Quit()
            'oExcel = Nothing
            'oBook = Nothing
            'oSheet = Nothing
            'MsgBox("File " & getname & " saved.", MsgBoxStyle.Information)

            '-----------------------------


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ToolStripMenuItem19_Click")
        End Try
    End Sub
    Private Sub uwallchardepth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles uwallchardepth.Click
        '*  ===============================================================
        '*  Show a graph of the char depth in upper wall versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim j As Integer
        Dim outputinterval As Integer = CInt(frmInputs.txtExcelInterval.Text)
        Dim Message, dDefault, MyValue As String
        Dim crittemp As Double
        Dim datatobeplotted(0 To NumberRooms + 1, 0 To NumberTimeSteps + 1) As Double

        Try
            Title = "Select data for which Room?"   ' Set title.
            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            Dim chardepth As Double
            DataShift = 0
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            Dim depth As Single = 0
            Description = "Room " + CStr(idr)
            If NumberTimeSteps < 2 Then Exit Sub

            If IntegralModel = False And KineticModel = False Then

                Title = "Enter char or isotherm temperature?"   ' Set title.
                Message = "Enter the temperature (C)"   ' Set prompt.
                dDefault = "300"   ' Set default.
                ' Display message, title, and default value.
                MyValue = InputBox(Message, Title, dDefault)
                If Not IsNumeric(MyValue) Then
                    Exit Sub
                End If

                crittemp = CDbl(MyValue) + 273
                Dim maxtemp As Double = 0
                Dim NumberwallNodes As Integer
                NumberwallNodes = UBound(UWallNode, 2)
                Dim X(0 To NumberwallNodes) As Double
                Dim Y(0 To NumberwallNodes) As Double
                Dim T(0 To NumberwallNodes) As Double

                'define variables
                Title = "Upper wall char depth (mm) - based on " & crittemp.ToString & " C isotherm"
                'Title = "Upper wall char depth (mm)"

                For j = 0 To NumberTimeSteps
                    'If j Mod outputinterval = 0 Then
                    For curve = 1 To NumberwallNodes
                        X(NumberwallNodes - curve + 1) = UWallNode(idr, curve, j) 'descending order
                        If HaveWallSubstrate(idr) = True Then
                            depth = (curve - 1) * WallThickness(idr) / (Wallnodes - 1)
                            If depth > WallThickness(idr) Then
                                depth = WallThickness(idr) + (curve - Wallnodes) * WallSubThickness(idr) / (Wallnodes - 1)
                            End If
                        Else
                            depth = (curve - 1) * WallThickness(idr) / (NumberwallNodes - 1)
                        End If
                        Y(NumberwallNodes - curve + 1) = depth
                        T(NumberwallNodes - curve + 1) = j * Timestep

                    Next
                    'char depth by interpolation
                    Interpolate_D(X, Y, NumberwallNodes, crittemp, chardepth)
                    ' If chardepth > 25 Then Stop
                    datatobeplotted(idr, j) = chardepth
                    If j > 1 Then
                        If datatobeplotted(idr, j) < datatobeplotted(idr, j - 1) Then
                            datatobeplotted(idr, j) = datatobeplotted(idr, j - 1)
                        End If
                    End If


                Next
            ElseIf IntegralModel = True Then
                'define variables
                Title = "Upper wall char depth (mm) - based on integral model for burning rate"
                'Title = "Upper wall char depth (mm)"

                For j = 0 To NumberTimeSteps
                    datatobeplotted(idr, j) = wall_char(j, 2)
                Next
            ElseIf kineticModel = True Then
                'new
                Stop
            End If


            'define variables
            DataShift = 0
            DataMultiplier = 1
            'call procedure to plot data
            ' Graph_Data_2D(Title, datatobeplotted, DataShift, DataMultiplier, timeunit)
            Graph_Data_CharDepth((crittemp - 273), Title, datatobeplotted, DataShift, DataMultiplier, timeunit)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in uwallchardepth_Click")
        End Try
    End Sub

    Private Sub ceilingdepth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ceilingchardepth.Click
        '*  ===============================================================
        '*  Show a graph of the char depth in ceiling versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim j As Integer
        Dim outputinterval As Integer = CInt(frmInputs.txtExcelInterval.Text)
        Dim Message, dDefault, MyValue As String
        Dim crittemp As Double
        Dim datatobeplotted(0 To NumberRooms + 1, 0 To NumberTimeSteps + 1) As Double

        Try
            Title = "Select data for which Room?"   ' Set title.
            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            Dim chardepth As Double
            DataShift = 0
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            Dim depth As Single = 0
            Description = "Room " + CStr(idr)
            If NumberTimeSteps < 2 Then Exit Sub

            If IntegralModel = False And KineticModel = False Then

                Title = "Enter char temperature?"   ' Set title.
                Message = "Enter the temperature (C)"   ' Set prompt.
                dDefault = "300"   ' Set default.
                ' Display message, title, and default value.
                MyValue = InputBox(Message, Title, dDefault)
                If Not IsNumeric(MyValue) Then
                    Exit Sub
                End If

                crittemp = CDbl(MyValue) + 273
                Dim maxtemp As Double = 0
                Dim NumberceilingNodes As Integer
                NumberceilingNodes = UBound(CeilingNode, 2)
                Dim X(0 To NumberceilingNodes) As Double
                Dim Y(0 To NumberceilingNodes) As Double
                Dim T(0 To NumberceilingNodes) As Double

                'define variables
                Title = "Ceiling char depth (mm) - based on " & crittemp.ToString & " C isotherm"
                'Title = "Upper wall char depth (mm)"

                For j = 0 To NumberTimeSteps
                    'If j Mod outputinterval = 0 Then
                    For curve = 1 To NumberceilingNodes
                        X(NumberceilingNodes - curve + 1) = CeilingNode(idr, curve, j) 'descending order
                        If HaveCeilingSubstrate(idr) = True Then
                            depth = (curve - 1) * CeilingThickness(idr) / (Ceilingnodes - 1)
                            If depth > CeilingThickness(idr) Then
                                depth = CeilingThickness(idr) + (curve - Ceilingnodes) * CeilingSubThickness(idr) / (Ceilingnodes - 1)
                            End If
                        Else
                            depth = (curve - 1) * CeilingThickness(idr) / (NumberceilingNodes - 1)
                        End If
                        Y(NumberceilingNodes - curve + 1) = depth
                        T(NumberceilingNodes - curve + 1) = j * Timestep

                    Next
                    'char depth by interpolation
                    Interpolate_D(X, Y, NumberceilingNodes, crittemp, chardepth)
                    ' If chardepth > 25 Then Stop
                    datatobeplotted(idr, j) = chardepth
                    If j > 1 Then
                        If datatobeplotted(idr, j) < datatobeplotted(idr, j - 1) Then
                            datatobeplotted(idr, j) = datatobeplotted(idr, j - 1)
                        End If
                    End If


                Next
            ElseIf IntegralModel = True Then

                'define variables
                Title = "Ceiling char depth (mm) - based on integral model for burning rate"
                'Title = "Upper wall char depth (mm)"

                For j = 0 To NumberTimeSteps
                    datatobeplotted(idr, j) = ceil_char(j, 2)
                Next
            ElseIf KineticModel = True Then
                'new
                Stop
            End If


            'define variables
            DataShift = 0
            DataMultiplier = 1
            'call procedure to plot data
            Graph_Data_2D(Title, datatobeplotted, DataShift, DataMultiplier, timeunit)

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ceilingchardepth_click")
        End Try
    End Sub

    Private Sub ToolStripMenuItem20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        '*  ===============================================================
        '*  Show a graph of Nodal upper wall Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Upper Wall Temperatures (C)"
            DataShift = -273
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            frmPlot.Chart1.Series.Clear()
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberNodes As Integer
            NumberNodes = UBound(UWallNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = NumberNodes

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1
            For i = 0 To 2 * numsets - 1 Step 2


                If HaveWallSubstrate(idr) = True Then
                    depth = (curve - 1) * WallThickness(idr) / (Wallnodes - 1)
                    If depth > WallThickness(idr) Then
                        depth = WallThickness(idr) + (curve - Wallnodes) * WallSubThickness(idr) / (Wallnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * WallThickness(idr) / (NumberNodes - 1)
                End If

                'depth = (curve - 1) * WallThickness(idr) / (NumberNodes - 1)

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                frmPlot.Chart1.Series.Add(chdata(0, i))
                frmPlot.Chart1.Series(chdata(0, i)).ChartType = SeriesChartType.FastLine
                maxtemp = 0
                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * UWallNode(idr, curve, j) + DataShift) 'data to be plotted
                    frmPlot.Chart1.Series(chdata(0, i)).Points.AddXY(tim(j, 1), chdata(j, i + 1))
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1
            Next i

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ToolStripMenuItem20_Click")
        End Try
    End Sub
    Private Sub ToolStripMenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        '*  ===============================================================
        '*  Show a graph of Nodal lower wall Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Lower Wall Temperatures (C)"
            DataShift = -273
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            frmPlot.Chart1.Series.Clear()
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberNodes As Integer
            NumberNodes = UBound(LWallNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = NumberNodes

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1
            For i = 0 To 2 * numsets - 1 Step 2

                If HaveWallSubstrate(idr) = True Then
                    depth = (curve - 1) * WallThickness(idr) / (Wallnodes - 1)
                    If depth > WallThickness(idr) Then
                        depth = WallThickness(idr) + (curve - Wallnodes) * WallSubThickness(idr) / (Wallnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * WallThickness(idr) / (NumberNodes - 1)
                End If

                'depth = (curve - 1) * WallThickness(idr) / (NumberNodes - 1)

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                frmPlot.Chart1.Series.Add(chdata(0, i))
                frmPlot.Chart1.Series(chdata(0, i)).ChartType = SeriesChartType.FastLine
                maxtemp = 0
                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * LWallNode(idr, curve, j) + DataShift) 'data to be plotted
                    frmPlot.Chart1.Series(chdata(0, i)).Points.AddXY(tim(j, 1), chdata(j, i + 1))
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1
            Next i

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ToolStripMenuItem21_Click")
        End Try
    End Sub
    Private Sub ToolStripMenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem22.Click
        '*  ===============================================================
        '*  Show a graph of Nodal floor Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Floor Temperatures (C)"
            DataShift = -273
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            frmPlot.Chart1.Series.Clear()
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberNodes As Integer
            NumberNodes = UBound(FloorNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = NumberNodes

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1
            For i = 0 To 2 * numsets - 1 Step 2

                If HaveFloorSubstrate(idr) = True Then
                    depth = (curve - 1) * FloorThickness(idr) / (Floornodes - 1)
                    If depth > FloorThickness(idr) Then
                        depth = FloorThickness(idr) + (curve - Floornodes) * FloorSubThickness(idr) / (Floornodes - 1)
                    End If
                Else
                    depth = (curve - 1) * FloorThickness(idr) / (NumberNodes - 1)
                End If
                'depth = (curve - 1) * FloorThickness(idr) / (NumberNodes - 1)

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                frmPlot.Chart1.Series.Add(chdata(0, i))
                frmPlot.Chart1.Series(chdata(0, i)).ChartType = SeriesChartType.FastLine
                maxtemp = 0
                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * FloorNode(idr, curve, j) + DataShift) 'data to be plotted
                    frmPlot.Chart1.Series(chdata(0, i)).Points.AddXY(tim(j, 1), chdata(j, i + 1))
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1
            Next i

            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in ToolStripMenuItem22_Click")
        End Try
    End Sub
    Private Sub NetHeatFluxesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NetHeatFluxesToolStripMenuItem.Click

    End Sub

    Private Sub ToolStripMenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem24.Click
        '*  =======================================================
        '*  Show a graph of the floor convective heat transfer coefficient versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Ceiling Convective Heat Transfer Coefficient (W/m2K)"
        DataShift = 0
        DataMultiplier = 1

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QCeilingAST(i, 1, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub

    Private Sub ToolStripMenuItem25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem25.Click
        '*  =======================================================
        '*  Show a graph of the floor convective heat transfer coefficient versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Upper Wall Convective Heat Transfer Coefficient (W/m2K)"
        DataShift = 0
        DataMultiplier = 1

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QUpperWallAST(i, 1, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub

    Private Sub ToolStripMenuItem26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem26.Click
        '*  =======================================================
        '*  Show a graph of the lower wall convective heat transfer coefficient versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Lower Wall Convective Heat Transfer Coefficient (W/m2K)"
        DataShift = 0
        DataMultiplier = 1

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QLowerWallAST(i, 1, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub

    Private Sub ToolStripMenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem27.Click
        '*  =======================================================
        '*  Show a graph of the floor convective heat transfer coefficient versus time
        '*  =======================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double
        '
        'define variables
        Title = "Floor Convective Heat Transfer Coefficient (W/m2K)"
        DataShift = 0
        DataMultiplier = 1

        Dim datatosend(0 To NumberRooms, 0 To NumberTimeSteps) As Double
        For i = 1 To NumberRooms
            For j = 1 To NumberTimeSteps
                datatosend(i, j) = QFloorAST(i, 1, j)
            Next
        Next

        Graph_Data_2D(Title, datatosend, DataShift, DataMultiplier, timeunit)
    End Sub


    Private Sub ToolStripMenuItem23_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem23.Click
        Dim Title As String
        Dim DataShift, MaxXValue, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short
        'If IsNothing(furnaceNHL) Then
        Call surfacetempcalc(1, fireroom, HaveCeilingSubstrate(fireroom))

        'calculate NHL in furnace for the  wall material
        Call surfacetempcalc(2, fireroom, HaveWallSubstrate(fireroom))

        'calculate NHL in furnace for the floor material
        Call surfacetempcalc(4, fireroom, HaveFloorSubstrate(fireroom))
        'End If
        'Exit Sub

        'define variables
        Title = "ISO 834 - Normalised Heat Load (s1/2 K)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        '  MaxYValue = NHL(1, fireroom, stepcount)
        ' If MaxYValue < NHLmax Then MaxYValue = NHLmax
        'MaxYValue = 250000
        MaxXValue = 6 * 3600
        Dim i As Integer
        'Dim steps = CInt(SimTime / Timestep)
        Dim steps = 6 * 3600 / Timestep

        Dim data(0 To 2, 0 To steps) As Double

        For j = 0 To 2
            For i = 1 To steps
                data(j, i) = furnaceNHL(j, i) 'ceiling material in furnace 
                'data(i) = furnaceNHL(1, i) 'wall material in furnace 
                'data(i) = furnaceNHL(2, i) 'floor material in furnace 
            Next
        Next

        Graph_Data_furnaceNHL(1, Title, data, DataShift, DataMultiplier, GraphStyle, MaxYValue, MaxXValue)

    End Sub

    'Private Sub CeilingToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeilingToolStripMenuItem1.Click
    '    Dim Title As String
    '    Dim DataShift, MaxYValue As Double
    '    Dim DataMultiplier, sum As Double
    '    Dim GraphStyle As Short
    '    Dim data(0 To 4, 0 To fireroom, 0 To NumberTimeSteps + 1) As Double
    '    Dim room As Integer = fireroom

    '    'define variables
    '    Title = "Normalised Heat Load (s1/2 K)"
    '    DataShift = 0
    '    DataMultiplier = 1
    '    GraphStyle = 4 '2=user-defined
    '    MaxYValue = 0

    '    sum = 0
    '    For i = 1 To stepcount
    '        'calculate NHL by integrating the flux
    '        If QCeiling(room, i) < 0 Then
    '            sum = sum - QCeiling(room, i) / Sqrt(ThermalInertiaCeiling(room))
    '        End If
    '        data(1, room, i) = sum
    '    Next
    '    sum = 0
    '    For i = 1 To stepcount
    '        'calculate NHL by integrating the flux
    '        If QUpperWall(room, i) < 0 Then
    '            sum = sum - QUpperWall(room, i) / Sqrt(ThermalInertiaWall(room))
    '        End If
    '        data(2, room, i) = sum
    '    Next
    '    sum = 0
    '    For i = 1 To stepcount
    '        'calculate NHL by integrating the flux
    '        If QLowerWall(room, i) < 0 Then
    '            sum = sum - QLowerWall(room, i) / Sqrt(ThermalInertiaWall(room))
    '        End If
    '        data(3, room, i) = sum
    '    Next
    '    sum = 0
    '    For i = 1 To stepcount
    '        'calculate NHL by integrating the flux
    '        If QFloor(room, i) < 0 Then
    '            sum = sum - QFloor(room, i) / Sqrt(ThermalInertiaFloor(room))
    '        End If
    '        data(4, room, i) = sum
    '    Next
    '    sum = 0
    '    For i = 1 To stepcount
    '        data(0, room, i) = NHL(1, room, i)
    '    Next

    '    Graph_Data_NHLmulti(0, Title, data, DataShift, DataMultiplier)

    'End Sub

    'Private Sub AverageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AverageToolStripMenuItem.Click
    '    Dim Title As String
    '    Dim DataShift, MaxYValue As Double
    '    Dim DataMultiplier As Double
    '    Dim GraphStyle As Short

    '    'define variables
    '    Title = "Normalised Heat Load (s1/2 K) - weighted average"
    '    DataShift = 0
    '    DataMultiplier = 1
    '    GraphStyle = 4 '2=user-defined
    '    MaxYValue = 0

    '    Graph_Data_NHL(1, Title, NHL, DataShift, DataMultiplier)
    'End Sub

    Private Sub mnuUnburnedFuel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUnburnedFuel.Click

    End Sub

    Private Sub ToolStripMenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem28.Click
        Dim Title As String
        Dim DataShift, MaxXValue, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'calculate NHL in furnace for the wall material
        'index 1 ias ceiling
        'index 2 or 3 is wall
        'index other is floor
        Call surfacetempcalc(1, fireroom, HaveCeilingSubstrate(fireroom))

        'calculate NHL in furnace for the  wall material
        Call surfacetempcalc(2, fireroom, HaveWallSubstrate(fireroom))

        'calculate NHL in furnace for the floor material
        Call surfacetempcalc(4, fireroom, HaveFloorSubstrate(fireroom))

        'define variables
        Title = "ISO 834 - Surface Temperature (C)"
        DataShift = -273
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxXValue = SimTime
        Dim i As Integer
        Dim steps = CInt(SimTime / Timestep)

        Dim data(0 To 2, 0 To steps) As Double

        'Dim j As Integer = 1
        For j = 0 To 2
            For i = 1 To steps
                data(j, i) = furnaceST(j, i) 'ceiling material in furnace 
                'data(i) = furnaceNHL(1, i) 'wall material in furnace 
                'data(i) = furnaceNHL(2, i) 'floor material in furnace 
            Next
        Next

        Graph_Data_furnaceST(1, 1, Title, data, DataShift, DataMultiplier, GraphStyle, MaxYValue, MaxXValue)
    End Sub
    Private Sub ToolStripMenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem29.Click
        Dim Title As String
        Dim DataShift, MaxXValue, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'calculate NHL in furnace for the ceiling material
        Call surfacetempcalc(1, fireroom, HaveCeilingSubstrate(fireroom))

        'calculate NHL in furnace for the  wall material
        Call surfacetempcalc(2, fireroom, HaveWallSubstrate(fireroom))

        'calculate NHL in furnace for the floor material
        Call surfacetempcalc(4, fireroom, HaveFloorSubstrate(fireroom))

        'define variables
        Title = "ISO 834 - Net Heat Flux (kW/m2)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxXValue = SimTime
        Dim i As Integer
        Dim steps = CInt(SimTime / Timestep)

        Dim data(0 To 2, 0 To steps) As Double

        'Dim j As Integer = 1
        For j = 0 To 2
            For i = 1 To steps
                'data(j, i) = furnaceqnet(j, i) 'ceiling material in furnace 
                data(j, i) = furnaceqnet(j, i) 'wall material in furnace 
                'data(i) = furnaceqnet(2, i) 'floor material in furnace 
            Next
        Next

        Graph_Data_furnaceST(1, 1, Title, data, DataShift, DataMultiplier, GraphStyle, MaxYValue, MaxXValue)
    End Sub

    Private Sub IncidentHeatFluxesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IncidentHeatFluxesToolStripMenuItem.Click

    End Sub

    Private Sub ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChangeMaterialsDatabaseFilethermalmdbToolStripMenuItem.Click

        frmMaterials.Close()
        DBDirectory = UserPersonalDataFolder & gcs_folder_ext & "\dbases"

        OpenFileDialog1.CheckPathExists = True
        OpenFileDialog1.InitialDirectory = DBDirectory

        OpenFileDialog1.Title = "Open Material Database File"
        OpenFileDialog1.Filter = "All Files (*.*)|*.*| Database Files (*.mdb)|*.mdb"
        OpenFileDialog1.FilterIndex = 2
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.RestoreDirectory = True

        If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            MaterialsDatabaseName = OpenFileDialog1.FileName
        ElseIf Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If

        My.Settings("thermalConnectionString") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MaterialsDatabaseName

    End Sub

    Private Sub DevelopmentKeyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DevelopmentKeyToolStripMenuItem.Click
        Dim Title As String = "Development Key" ' Set title.
        Dim Message As String = "Enter a development key code" ' Set prompt.
        Dim Default_Renamed As Integer = 0 ' Set default.
        Dim MyValue As String

        MyValue = InputBox(Message, Title, Default_Renamed)

        If MyValue = "0" Or MyValue = "" Then
            Exit Sub

        ElseIf MyValue = DevKeyCode Then
            DevKey = True
            MsgBox("Thanks, you're good to go.", MsgBoxStyle.OkOnly)

        Else
            MsgBox("You have entered an incorrect development key. Try again.", MsgBoxStyle.OkOnly)
            DevKey = False

        End If

        My.Settings.devkeycode = MyValue
        My.Settings.Save()

    End Sub

    Private Sub UserModeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UserModeToolStripMenuItem.Click
        If DevKey = True Then
            TalkToEVACNZToolStripMenuItem.Visible = True
            TalkToEVACNZToolStripMenuItem.Checked = TalkToEVACNZ
            'CLT_ToolStripMenuItem31.Visible = True
            'uwallchardepth.Visible = True
        Else
            TalkToEVACNZToolStripMenuItem.Visible = False
            'CLT_ToolStripMenuItem31.Visible = False
            'uwallchardepth.Visible = False
        End If

    End Sub

    Private Sub UserModeToolStripMenuItem_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles UserModeToolStripMenuItem.MouseHover
        If DevKey = True Then
            TalkToEVACNZToolStripMenuItem.Visible = True
            TalkToEVACNZToolStripMenuItem.Checked = TalkToEVACNZ
        Else
            TalkToEVACNZToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub TalkToEVACNZToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TalkToEVACNZToolStripMenuItem.Click
        'manual start
        If TalkToEVACNZToolStripMenuItem.Checked = True Then
            TalkToEVACNZ = True
        Else
            TalkToEVACNZ = False
        End If
    End Sub

    Private Sub UtilitiesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UtilitiesToolStripMenuItem.Click
        If DevKey = True Then
            EvacuatioNZSettingsToolStripMenuItem.Visible = True
        Else
            EvacuatioNZSettingsToolStripMenuItem.Visible = False
        End If
    End Sub
    Private Sub UtilitiesToolStripMenuItem_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles UtilitiesToolStripMenuItem.MouseHover
        If DevKey = True Then
            EvacuatioNZSettingsToolStripMenuItem.Visible = True
        Else
            EvacuatioNZSettingsToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub mnuExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExport.Click
        If DevKey = True Then
            BoundaryNodeTemperaturesToolStripMenuItem.Visible = True
        Else
            BoundaryNodeTemperaturesToolStripMenuItem.Visible = False
        End If
    End Sub

    Private Sub UpperWallToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpperWallToolStripMenuItem.Click
        '*  ===============================================================
        '*  Show a graph of Nodal upper wall Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Wall Temperatures (C)"

            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberwallNodes As Integer
            NumberwallNodes = UBound(UWallNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberwallNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub
            DataShift = -273
            DataMultiplier = 1
            numpoints = maxpoints
            numsets = NumberwallNodes

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1

            '-----------------------------
            'Start a new workbook in Excel
            Dim oExcel As Object
            Dim oBook As Object
            Dim oSheet As Object
            Dim getname As String
            Dim outputinterval As Integer = CInt(frmInputs.txtExcelInterval.Text)
            Dim tcounter As Integer = Ceiling(SimTime / outputinterval)
            Dim colcounter As Integer
            Dim DataArray2(0 To tcounter + 1, 0 To 2 * numsets - 1) As Object
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            getname = RiskDataDirectory & "excel_UWnodetemps_" & Convert.ToString(frmInputs.txtBaseName.Text)
            'Add headers to the worksheet on row 1
            oSheet = oBook.Worksheets(1)
            oSheet.Range("A1").Value = "time (sec)"
            oSheet.Range("B1").Value = "node depth (mm)--->"

            '-----------------------------
            colcounter = 2
            For i = 0 To 2 * numsets - 1 Step 2

                If HaveCeilingSubstrate(idr) = True Then
                    depth = (curve - 1) * WallThickness(idr) / (Wallnodes - 1)
                    If depth > WallThickness(idr) Then
                        depth = WallThickness(idr) + (curve - Wallnodes) * WallSubThickness(idr) / (Wallnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * WallThickness(idr) / (NumberwallNodes - 1)
                End If

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"

                maxtemp = 0


                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * UWallNode(idr, curve, j) + DataShift) 'data to be plotted
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)
                    If j = 1 Then
                        DataArray2(1, 0) = 0
                        DataArray2(1, colcounter) = chdata(j, i + 1)
                    End If
                    If j Mod outputinterval = 0 Then
                        DataArray2(j / outputinterval + 1, 0) = Timestep * j
                        DataArray2(j / outputinterval + 1, colcounter) = chdata(j, i + 1)
                    End If
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1

                '-----------------------------
                DataArray2(0, colcounter) = depth
                colcounter = colcounter + 1
                '-----------------------------
            Next i

            oSheet.Range("A2").Resize(tcounter + 2, 2 * numsets).Value = DataArray2

            If Not System.String.IsNullOrEmpty(getname) Then oBook.SaveAs(getname)
            oBook.Close(SaveChanges:=False)
            oExcel.Quit()
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing

            MsgBox("File " & getname & " saved.", MsgBoxStyle.Information)


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in UpperWallToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub CeilingToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CeilingToolStripMenuItem2.Click
        '*  ===============================================================
        '*  Show a graph of Nodal Ceiling Temperatures versus time.
        '*  ===============================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Display Nodal Temperatures for Which Room?"   ' Set title.

            Message = "Enter the room"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue)

            'define variables
            Title = "Nodal Ceiling Temperatures (C)"
            Dim depth As Single = 0
            Dim maxtemp As Double = 0
            Dim NumberCeilingNodes As Integer
            NumberCeilingNodes = UBound(CeilingNode, 2)

            Description = "Room " + CStr(idr)
            maxpoints = NumberCeilingNodes * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = NumberCeilingNodes
            DataShift = -273
            DataMultiplier = 1
            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1

            '-----------------------------
            'Start a new workbook in Excel
            Dim oExcel As Object
            Dim oBook As Object
            Dim oSheet As Object
            Dim getname As String
            Dim outputinterval As Integer = CInt(frmInputs.txtExcelInterval.Text)
            Dim tcounter As Integer = Ceiling(SimTime / outputinterval)
            Dim colcounter As Integer
            Dim DataArray2(0 To tcounter + 1, 0 To 2 * numsets - 1) As Object
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            getname = RiskDataDirectory & "excel_Cnodetemps_" & Convert.ToString(frmInputs.txtBaseName.Text)
            'Add headers to the worksheet on row 1
            oSheet = oBook.Worksheets(1)
            oSheet.Range("A1").Value = "time (sec)"
            oSheet.Range("B1").Value = "node depth (mm)--->"

            colcounter = 2
            For i = 0 To 2 * numsets - 1 Step 2

                If HaveCeilingSubstrate(idr) = True Then
                    depth = (curve - 1) * CeilingThickness(idr) / (Ceilingnodes - 1)
                    If depth > CeilingThickness(idr) Then
                        depth = CeilingThickness(idr) + (curve - Ceilingnodes) * CeilingSubThickness(idr) / (Ceilingnodes - 1)
                    End If
                Else
                    depth = (curve - 1) * CeilingThickness(idr) / (NumberCeilingNodes - 1)
                End If

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                maxtemp = 0


                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    chdata(j, i + 1) = (DataMultiplier * CeilingNode(idr, curve, j) + DataShift) 'data to be plotted
                    If chdata(j, i + 1) > maxtemp Then maxtemp = chdata(j, i + 1)

                    If j = 1 Then
                        DataArray2(1, 0) = 0
                        DataArray2(1, colcounter) = chdata(j, i + 1)
                    End If
                    If j Mod outputinterval = 0 Then
                        DataArray2(j / outputinterval + 1, 0) = Timestep * j
                        DataArray2(j / outputinterval + 1, colcounter) = chdata(j, i + 1)
                    End If
                Next

                chdata(0, i) = "N" & CStr(curve) & " " & Format(depth, "0.0") & " mm"
                curve = curve + 1

                '-----------------------------
                DataArray2(0, colcounter) = depth
                colcounter = colcounter + 1
                '-----------------------------
            Next i

            oSheet.Range("A2").Resize(tcounter + 2, 2 * numsets).Value = DataArray2

            If Not System.String.IsNullOrEmpty(getname) Then oBook.SaveAs(getname)
            oBook.Close(SaveChanges:=False)
            oExcel.Quit()
            oExcel = Nothing
            oBook = Nothing
            oSheet = Nothing

            MsgBox("File " & getname & " saved.", MsgBoxStyle.Information)


        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, "Exception in CeilingToolStripMenuItem_Click")
        End Try
    End Sub

    Private Sub mnuPlotWoodBurningRate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPlotWoodBurningRate.Click
        Dim Title As String
        Dim DataShift, MaxYValue As Double
        Dim DataMultiplier As Double
        Dim GraphStyle As Short

        'define variables
        Title = "Wood burning rate (kg/s)"
        DataShift = 0
        DataMultiplier = 1
        GraphStyle = 4 '2=user-defined
        MaxYValue = 0

        'call procedure to plot data
        'Graph_Data3 frmGraphs.MSChart1, Title, Visibility(), DataShift, DataMultiplier, GraphStyle
        Graph_GER(1, Title, WoodBurningRate, DataShift, DataMultiplier, GraphStyle, MaxYValue)
    End Sub

    Private Sub burningrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles burningrate.Click
        '*  =============================================================
        '*  Show a graph of the fuel mass loss rate and burning rate versus time
        '*  =============================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double

        'define variables
        Title = "Mass Loss and Burning Rate(g/s)"
        DataShift = 0
        DataMultiplier = 1

        If FuelResponseEffects = True Then
            Title = "Mass Loss and Burning Rate(g/s)"
            DataShift = 0
            DataMultiplier = 1000
            Graph_Data_fuelburningrate2(1, Title, FuelBurningRate, DataShift, DataMultiplier, timeunit)
        End If
    End Sub

    Private Sub areashrinkage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles areashrinkage.Click
        '*  =============================================================
        '*  Show a graph of the Fuel surface area shrinkage ratio versus time
        '*  =============================================================

        Dim Title As String, DataShift As Double
        Dim DataMultiplier As Double

        'define variables
        Title = "Fuel surface area shrinkage ratio (-)"
        DataShift = 0
        DataMultiplier = 1

        If FuelResponseEffects = True Then
            Title = "Fuel surface area shrinkage ratio (-)"
            DataShift = 0
            DataMultiplier = 1
            Graph_Data_areashrinkage(1, Title, FuelBurningRate, DataShift, DataMultiplier, timeunit)
        End If
    End Sub



    Private Sub BatchFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BatchFilesToolStripMenuItem.Click

        'select the master folder
        FolderBrowserDialog1.Description = "Select the Batch File Root Folder"
        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.ShowDialog()
        MasterDirectory = FolderBrowserDialog1.SelectedPath

        If MasterDirectory = "" Then Exit Sub

        'seek confirmation from the user that they want to run the simulation
        Dim response As Short = MsgBox("Do you want to run all simulations now?", MB_YESNO + MB_ICONQUESTION, ProgramTitle)
        If response = IDNO Then Exit Sub
        batch = True
        'loop through all the subfolders
        For Each Dir As String In System.IO.Directory.GetDirectories(MasterDirectory)
            Dim dirInfo As New System.IO.DirectoryInfo(Dir)
            RiskDataDirectory = MasterDirectory & "\" & dirInfo.Name & "\"

            If InStr(dirInfo.Name, "basemodel_") = 1 Then

                opendatafile = RiskDataDirectory & dirInfo.Name & ".xml"

                'open the base model in this folder
                frmInputs.Read_BaseFile_xml(opendatafile, True)

                basefile = opendatafile

                'go away and run this project
                Call frmInputs.readytorun()

            Else
                'skip this folder
                Continue For
            End If
        Next
        batch = False

        MsgBox("Batch file simulations complete.", MsgBoxStyle.Information)

    End Sub



    Private Sub ResidualMassFractions2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResidualMassFractions2ToolStripMenuItem.Click


    End Sub

    Private Sub ResidualMassFractionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResidualMassFractionsToolStripMenuItem.Click

        'kinetic wood pyrolysis model
        'plot the residual mass fractions for water, cellulose, hemicellulose and lignin

        Dim Title As String, DataShift As Double, MaxYValue As Double
        Dim DataMultiplier As Double, GraphStyle As Integer, idr As Integer
        Dim i As Integer, j As Integer, maxpoints As Long

        Dim Message, dDefault, MyValue As String
        Dim numpoints, numsets As Integer

        Try
            Title = "Select data for which finite difference element (exposed surface = 1)?"   ' Set title.
            Message = "Enter the element"   ' Set prompt.
            dDefault = "1"   ' Set default.
            ' Display message, title, and default value.
            MyValue = InputBox(Message, Title, dDefault)
            If Not IsNumeric(MyValue) Then
                Exit Sub
            End If
            idr = CInt(MyValue) 'store the element number

            'define variables
            Title = "Residual mass fractions"
            DataShift = 0
            DataMultiplier = 1
            GraphStyle = 4            '2=user-defined
            MaxYValue = 0
            frmPlot.Chart1.Series.Clear()
            Dim depth As Single = 0
            Dim maxtemp As Double = 0

            Description = "Component " + CStr(idr)
            maxpoints = 5 * NumberTimeSteps
            If NumberTimeSteps < 2 Then Exit Sub

            numpoints = maxpoints
            numsets = 5

            Dim chdata(0 To numpoints - 1, 0 To 2 * numsets - 1) As Object
            Dim curve As Integer

            curve = 1

            For i = 0 To 4

                If i = 0 Then chdata(0, i) = "Water"
                If i = 1 Then chdata(0, i) = "Cellulose"
                If i = 2 Then chdata(0, i) = "Hemicellulose"
                If i = 3 Then chdata(0, i) = "Lignin"
                If i = 4 Then chdata(0, i) = "Char residue"

                frmPlot.Chart1.Series.Add(chdata(0, i))
                frmPlot.Chart1.Series(chdata(0, i)).ChartType = SeriesChartType.FastLine

                For j = 1 To stepcount
                    chdata(j, i) = tim(j, 1)
                    If i < 4 Then
                        chdata(j, i + 1) = (DataMultiplier * CeilingElementMF(idr, i, j) + DataShift) 'data to be plotted
                    Else
                        chdata(j, i + 1) = (DataMultiplier * CeilingCharResidue(idr, j) + DataShift) 'data to be plotted
                    End If
                    frmPlot.Chart1.Series(chdata(0, i)).Points.AddXY(tim(j, 1), chdata(j, i + 1))
                Next

            Next i


            frmPlot.Chart1.BackColor = Color.AliceBlue
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderWidth = 1
            frmPlot.Chart1.ChartAreas("ChartArea1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.Title = Title
            'frmPlot.Chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Format = "0.0"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Maximum = [Double].NaN
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.Title = "Time (sec)"
            frmPlot.Chart1.ChartAreas("ChartArea1").AxisX.IsMarginVisible = False
            frmPlot.Chart1.Legends("Legend1").BorderWidth = 1
            frmPlot.Chart1.Legends("Legend1").BackColor = Color.White
            frmPlot.Chart1.Legends("Legend1").BorderDashStyle = ChartDashStyle.Solid
            frmPlot.Chart1.Legends("Legend1").Docking = Docking.Right
            frmPlot.Chart1.Titles("Title1").Text = Title

            frmPlot.Chart1.Visible = True
            frmPlot.BringToFront()
            frmPlot.Show()

        Catch ex As Exception
            MsgBox(Err.Description, MsgBoxStyle.OkOnly, Err.Source & "Line " & Err.Erl)
        End Try
    End Sub
End Class