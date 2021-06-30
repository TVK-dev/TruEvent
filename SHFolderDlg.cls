VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHFolderDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===========
'SHFolderDlg version 1.0
'===========
'
'A class with one method: BrowseForFolder().
'
'Raises a dialog for browsing to a folder.  You can specify the "browse root" as
'well as an optional StartPath under that for the dialog to show selected.
'
'The Instructions property only displays as ANSI characters but everything else
'should be fully Unicode.
'
'The Flags property can be used to produce many variations.
'
'Some features require Windows 2000/Me/IE 5 or later, others Windows XP or later.
'
'Depends on: ShFolderDlgStatic.bas
'

Private Const WIN32_FALSE As Long = 0
Private Const WIN32_TRUE As Long = Not WIN32_FALSE
Private Const WIN32_NULL As Long = 0

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Type BROWSEINFO
    hWndOwner As Long
    PIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As BIF_OPTIONS
    lpfn As Long
    lParam As SHFolderDlg
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (ByRef BROWSEINFO As BROWSEINFO) As Long

Private Enum HRESULT
    S_OK = 0
    S_FALSE = 1
    ERROR_FILE_NOT_FOUND = 2
    E_INVALIDARG = &H80070057
End Enum

Private Declare Function SHGetFolderLocation Lib "shell32" ( _
    ByVal hWndOwnerReserved As Long, _
    ByVal nFolder As CSIDL, _
    ByVal hToken As Long, _
    ByVal dwReserved As Long, _
    ByRef PIDL As Long) As HRESULT

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListW" ( _
    ByVal PIDL As Long, _
    ByVal pszPath As Long) As Long

Private Declare Function SHILCreateFromPath Lib "shell32" ( _
    ByVal pszPath As Long, _
    ByRef PIDL As Long, _
    ByVal prgflnOut As Long) As HRESULT

Private mStartDir As String

Public Enum BIF_OPTIONS
    'Some options require shell32.dll versions later than 4.71:
    BIF_RETURNONLYFSDIRS = &H1&
    BIF_DONTGOBELOWDOMAIN = &H2&
    BIF_RETURNFSANCESTORS = &H8&
    BIF_EDITBOX = &H10&
    BIF_VALIDATE = &H11&
    BIF_NEWDIALOGSTYLE = &H40&
    BIF_BROWSEINCLUDEURLS = &H80&
    BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
    BIF_UAHINT = &H100&
    BIF_NONEWFOLDERBUTTON = &H200&
    BIF_NOTRANSLATETARGETS = &H400&
    BIF_BROWSEFORCOMPUTER = &H1000&
    BIF_BROWSEFORPRINTER = &H2000&
    BIF_BROWSEINCLUDEFILES = &H4000&
    BIF_SHAREABLE = &H8000&
    BIF_BROWSEFILEJUNCTIONS = &H100000 'Win7 and later: Allow folder junctions such
                                       'as a library or a compressed file with a .zip
                                       'file name extension to be browsed.
End Enum

Public Enum CSIDL
    'Values and comments taken from ShlObj.h in the Windows SDK:
    CSIDL_DESKTOP = &H0&    '<desktop>
    CSIDL_INTERNET = &H1&    'Internet Explorer (icon on desktop)
    CSIDL_PROGRAMS = &H2&    'Start Menu\Programs
    CSIDL_CONTROLS = &H3&    'My Computer\Control Panel
    CSIDL_PRINTERS = &H4&    'My Computer\Printers
    CSIDL_PERSONAL = &H5&    'My Documents
    CSIDL_FAVORITES = &H6&    '<user name>\Favorites
    CSIDL_STARTUP = &H7&    'Start Menu\Programs\Startup
    CSIDL_RECENT = &H8&    '<user name>\Recent
    CSIDL_SENDTO = &H9&    '<user name>\SendTo
    CSIDL_BITBUCKET = &HA&    '<desktop>\Recycle Bin
    CSIDL_STARTMENU = &HB&    '<user name>\Start Menu
    CSIDL_MYDOCUMENTS = CSIDL_PERSONAL 'Personal was just a silly name for My Documents
    CSIDL_MYMUSIC = &HD&    '"My Music" folder
    CSIDL_MYVIDEO = &HE&    '"My Videos" folder
    CSIDL_DESKTOPDIRECTORY = &H10&   '<user name>\Desktop
    CSIDL_DRIVES = &H11&   'My Computer
    CSIDL_NETWORK = &H12&   'Network Neighborhood (My Network Places)
    CSIDL_NETHOOD = &H13&   '<user name>\nethood
    CSIDL_FONTS = &H14&   'windows\fonts
    CSIDL_TEMPLATES = &H15&
    CSIDL_COMMON_STARTMENU = &H16&   'All Users\Start Menu
    CSIDL_COMMON_PROGRAMS = &H17&   'All Users\Start Menu\Programs
    CSIDL_COMMON_STARTUP = &H18&   'All Users\Startup
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19&   'All Users\Desktop
    CSIDL_APPDATA = &H1A&   '<user name>\Application Data
    CSIDL_PRINTHOOD = &H1B&   '<user name>\PrintHood
    CSIDL_LOCAL_APPDATA = &H1C&   '<user name>\Local Settings\Applicaiton Data (non roaming)
    CSIDL_ALTSTARTUP = &H1D&   'non localized startup
    CSIDL_COMMON_ALTSTARTUP = &H1E&   'non localized common startup
    CSIDL_COMMON_FAVORITES = &H1F&
    CSIDL_INTERNET_CACHE = &H20&
    CSIDL_COOKIES = &H21&
    CSIDL_HISTORY = &H22&
    CSIDL_COMMON_APPDATA = &H23&   'All Users\Application Data
    CSIDL_WINDOWS = &H24&   'GetWindowsDirectory()
    CSIDL_SYSTEM = &H25&   'GetSystemDirectory()
    CSIDL_PROGRAM_FILES = &H26&   'C:\Program Files
    CSIDL_MYPICTURES = &H27&   'C:\Program Files\My Pictures
    CSIDL_PROFILE = &H28&   'USERPROFILE
    CSIDL_SYSTEMX86 = &H29&   'x86 system directory on RISC
    CSIDL_PROGRAM_FILESX86 = &H2A&   'x86 C:\Program Files on RISC
    CSIDL_PROGRAM_FILES_COMMON = &H2B&   'C:\Program Files\Common
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C&   'x86 Program Files\Common on RISC
    CSIDL_COMMON_TEMPLATES = &H2D&   'All Users\Templates
    CSIDL_COMMON_DOCUMENTS = &H2E&   'All Users\Documents
    CSIDL_COMMON_ADMINTOOLS = &H2F&   'All Users\Start Menu\Programs\Administrative Tools
    CSIDL_ADMINTOOLS = &H30&   '<user name>\Start Menu\Programs\Administrative Tools
    CSIDL_CONNECTIONS = &H31&   'Network and Dial-up Connections
    CSIDL_COMMON_MUSIC = &H35&   'All Users\My Music
    CSIDL_COMMON_PICTURES = &H36&   'All Users\My Pictures
    CSIDL_COMMON_VIDEO = &H37&   'All Users\My Video
    CSIDL_RESOURCES = &H38&   'Resource Direcotry
    CSIDL_RESOURCES_LOCALIZED = &H39&   'Localized Resource Direcotry
    CSIDL_COMMON_OEM_LINKS = &H3A&   'Links to All Users OEM specific apps
    CSIDL_CDBURN_AREA = &H3B&   'USERPROFILE\Local Settings\Application Data\Microsoft\CD Burning
    CSIDL_COMPUTERSNEARME = &H3D&   'Computers Near Me (computered from Workgroup membership)
End Enum

Public CancelValidation As Boolean 'Applies when we have both BIF_EDITBOX and BIF_VALIDATE
                                   'Flag bits set.
Public ExpandStartPath As Boolean  'Requires shell32.dll version 6.0 (XP) or later.

Public Flags As BIF_OPTIONS

Public Instructions As String      'I don't call this "Title" because it doesn't change the dialog
                                   'title (caption) in the title bar.
Public OkCaption As String         'Set to non-empty value to replace "Ok" button text.
                                   '    Requires shell32.dll version 6.0 (XP) or later.
Public Root As CSIDL               'Normally this is used to set the desired browse root....

Public RootPath As String          '... but if this is non-empty then it will be used instead of
                                   'Root.
                                   '    Requires shell32.dll version 5.0 (Windows 2000 or Me, or
                                   '    IE 5) or later.
Public StartPath As String

Public Function BrowseForFolder(ByVal Owner As Object) As String
    'Owner should be a Form, MDIForm, PictureBox, UserControl with a window, etc.
    Dim Result As HRESULT
    Dim PIDLRoot As Long
    Dim Buffer As String
    Dim BROWSEINFO As BROWSEINFO
    Dim PIDL As Long
    Dim NulPos As Long
    Dim Success As Long
    
    If Len(RootPath) > 0 And SHFolderDlgStatic.Shell32VersionOrLater(5) Then
        Result = SHILCreateFromPath(StrPtr(RootPath), PIDLRoot, WIN32_NULL)
    Else
        Result = SHGetFolderLocation(0, Root, 0, 0, PIDLRoot)
    End If
    Select Case Result
        Case S_FALSE, ERROR_FILE_NOT_FOUND 'Different versions of the MSDN Library show
                                           'different return values here.
            Err.Raise 53, TypeName(Me) 'File not found.
        Case E_INVALIDARG
            Err.Raise 5, TypeName(Me) 'Invalid procedure call or argument.
        Case S_OK
            'Continue.
        Case Else
            Err.Raise &H8004EE00, _
                      TypeName(Me), _
                      "SHGetFolderLocation error " & Right$(String$(7, "0") & Hex$(Result), 8)
    End Select
    Buffer = StrConv(Instructions, vbFromUnicode) 'Even though we use the Unicode SHBrowseForFolder()
                                                  'it still wants the Title in ANSI!
    With BROWSEINFO
        .hWndOwner = Owner.hWnd
        .PIDLRoot = PIDLRoot
        .lpszTitle = StrPtr(Buffer)
        .ulFlags = Flags
        .lpfn = SHFolderDlgStatic.AddressOfBrowseCallbackProc
        Set .lParam = Me
    End With
    PIDL = SHBrowseForFolder(BROWSEINFO)
    CoTaskMemFree PIDLRoot
    If PIDL Then
        Buffer = Space$(32767) 'Max Unicode path.
        Success = SHGetPathFromIDList(PIDL, StrPtr(Buffer))
        CoTaskMemFree PIDL
        If Success = WIN32_FALSE Then
            Err.Raise &H8004EE02, _
                      TypeName(Me), _
                      "SHGetPathFromIDList failed"
        End If
        NulPos = InStr(Buffer, vbNullChar)
        If NulPos = 0 Then
            BrowseForFolder = Buffer
        Else
            BrowseForFolder = Left$(Buffer, NulPos - 1)
        End If
    'Else
        'Return empty String value (i.e. user canceled).
    End If
End Function

Private Sub Class_Initialize()
    'Set defaults:
    CancelValidation = True
    Root = CSIDL_DESKTOP
End Sub
