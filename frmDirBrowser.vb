Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmDirBrowser
	Inherits System.Windows.Forms.Form
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
	Private Declare Function lstrcat Lib "kernel32"  Alias "lstrcatA"(ByVal lpString1 As String, ByVal lpString2 As String) As Integer
	
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
	
	Private Structure SHITEMID
		Dim cb As Integer
		Dim abID As Byte
	End Structure
	
	Private Structure ITEMIDLIST
		Dim mkid As SHITEMID
	End Structure

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
	
	
	
	Public Function fBrowseForFolder(ByRef hwndOwner As Integer, ByRef sPrompt As String) As String
		'
		' Opens the system dialog for browsing for a folder.
		'
		Dim iNull As Short
		Dim lpIDList As Integer
		Dim lResult As Integer
        Dim sPath As String = ""
		Dim sPath1 As String
		Dim udtBI As BrowseInfo
		Dim IDL As ITEMIDLIST
		
		'
		' Get the ID of the folder to use as the root
		' in the directory box. Change the "CSIDL_"
		' constant to any of the defined values as
		' shown in the declarations section of this form.
		'
		sPath1 = fGetSpecialFolder(CSIDL_DESKTOP, IDL)
		
		'Initialize Drag & Drop capabilities in the dialog.
        'Call OleInitialize(0)
		
		With udtBI
			'    .pIDLRoot = 0  'Display the entire namespace hierarchy
			'                   'starting with the desktop folder.
			
			.pIDLRoot = IDL.mkid.cb 'Use the desired starting folder.
			
			.hwndOwner = hwndOwner
			.lpszTitle = lstrcat(sPrompt, "")
			
			If Option1(0).Checked Then
				.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
			ElseIf Option1(1).Checked Then 
				.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_BROWSEINCLUDEFILES
			ElseIf Option1(2).Checked Then 
				.ulFlags = BIF_BROWSEFORCOMPUTER
			Else
				.ulFlags = BIF_BROWSEFORPRINTER
			End If
		End With
		
		lpIDList = SHBrowseForFolder(udtBI)
		
		If lpIDList Then
			sPath = New String(Chr(0), MAX_PATH)
			lResult = SHGetPathFromIDList(lpIDList, sPath)
			Call CoTaskMemFree(lpIDList)
			
			iNull = InStr(sPath, vbNullChar)
			If iNull Then sPath = VB.Left(sPath, iNull - 1)
		End If
		
        'Call OleUninitialize()
		
		fBrowseForFolder = sPath
	End Function
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		Dim sStr As String
		
		sStr = fBrowseForFolder(Handle.ToInt32, "Click on an entry to select it.")
		'If sStr <> "" Then
		'    MsgBox sStr, vbInformation, "Directory Browser"
		'End If
		BatchFileFolder = sStr
	End Sub


    Private Sub cmdQuit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdQuit.Click
		Me.Close()
	End Sub
	
	Private Sub frmDirBrowser_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Option1(0).Checked = True
        'frmBlank.Hide()
	End Sub
End Class