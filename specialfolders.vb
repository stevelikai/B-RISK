Option Strict Off
Option Explicit On
Module folders
	
	Private Const CSIDL_ADMINTOOLS As Integer = &H30 '{user}\Start Menu _                                                        '\Programs\Administrative Tools
	Private Const CSIDL_COMMON_ADMINTOOLS As Integer = &H2F '(all users)\Start Menu\Programs\Administrative Tools
	Private Const CSIDL_APPDATA As Integer = &H1A '{user}\Application Data
	Private Const CSIDL_COMMON_APPDATA As Integer = &H23 '(all users)\Application Data
	Private Const CSIDL_COMMON_DOCUMENTS As Integer = &H2E '(all users)\Documents
	Private Const CSIDL_COOKIES As Integer = &H21
	Private Const CSIDL_HISTORY As Integer = &H22
	Private Const CSIDL_INTERNET_CACHE As Integer = &H20 'Internet Cache folder
	Private Const CSIDL_LOCAL_APPDATA As Integer = &H1C '{user}\Local Settings\Application Data (non roaming)
	Private Const CSIDL_MYPICTURES As Integer = &H27 'C:\Program Files\My Pictures
	Private Const CSIDL_PERSONAL As Integer = &H5 'My Documents
	Private Const CSIDL_PROGRAM_FILES As Integer = &H26 'Program Files folder
	Private Const CSIDL_PROGRAM_FILES_COMMON As Integer = &H2B 'Program Files\Common
	Private Const CSIDL_SYSTEM As Integer = &H25 'system folder
	Private Const CSIDL_WINDOWS As Integer = &H24 'Windows directory or SYSROOT()
	Private Const CSIDL_FLAG_CREATE As Integer = &H8000 'combine with CSIDL_ value to force
	Private Const MAX_PATH As Short = 260
	
	Private Const CSIDL_FLAG_MASK As Integer = &HFF00 'mask for all possible flag values
	Private Const SHGFP_TYPE_CURRENT As Integer = &H0 'current value for user, verify it exists
	Private Const SHGFP_TYPE_DEFAULT As Integer = &H1
	Private Const S_OK As Short = 0
	Private Const S_FALSE As Short = 1
	Private Const E_INVALIDARG As Integer = &H80070057 ' Invalid CSIDL Value
	
	Private Declare Function SHGetFolderPath Lib "shfolder"  Alias "SHGetFolderPathA"(ByVal hwndOwner As Integer, ByVal nFolder As Integer, ByVal hToken As Integer, ByVal dwFlags As Integer, ByVal pszPath As String) As Integer
	
    Public Sub get_folders()

        'Dim strBuffer As String
        Dim strPath As String
        Dim lngReturn As Integer
        Dim lngCSIDL As Integer

        '
        ' When a listbox item is clicked, get the
        ' associated folder's path.
        '
        'lngCSIDL = CSIDL_COMMON_APPDATA '= folder required eg CSIDL_APPDATA
        lngCSIDL = CSIDL_PERSONAL 'my documents
        strPath = New String(Chr(0), MAX_PATH)

        ' Get the folder's path. If the
        ' "Create" flag is used, the folder will be created
        ' if it does not exist.
        '
        lngReturn = SHGetFolderPath(0, lngCSIDL, 0, SHGFP_TYPE_CURRENT, strPath)
        'lngReturn = SHGetFolderPath(0, lngCSIDL Or CSIDL_FLAG_CREATE, 0, SHGFP_TYPE_CURRENT, strPath)

        Select Case lngReturn
            Case S_OK
                UserAppDataFolder = Left(strPath, InStr(1, strPath, Chr(0)) - 1)
                'lblPath.caption = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)

            Case S_FALSE
                'lblPath.caption = "Folder Does Not Exist"

            Case E_INVALIDARG
                'lblPath.caption = "Folder Not Valid on this OS"

            Case Else
                'lblPath.caption = "Folder Not Valid on this OS"
        End Select

        DocsFolder = Application.StartupPath & "\docs\"
        DataFolder = Application.StartupPath & "\data\"

    End Sub
End Module