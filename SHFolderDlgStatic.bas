Attribute VB_Name = "SHFolderDlgStatic"
Option Explicit
'=================
'SHFolderDlgStatic version 1.0
'=================
'
'Used by: SHFolderDlg.cls
'

Private Const WIN32_FALSE As Long = 0
Private Const WIN32_TRUE As Long = Not WIN32_FALSE

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SELCHANGED As Long = 2
Private Const BFFM_VALIDATEFAILED As Long = 4
Private Const WM_USER As Long = &H400&
Private Const BFFM_SETSELECTION As Long = WM_USER + 103
Private Const BFFM_SETSTATUSTEXT As Long = WM_USER + 104
Private Const BFFM_SETOKTEXT As Long = WM_USER + 105 'Requires shell32.dll version 6.0 (XP) or later.
Private Const BFFM_SETEXPANDED As Long = WM_USER + 106 'Requires shell32.dll version 6.0 (XP) or later.

Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

Private Declare Function DllGetVersion Lib "shell32" (ByRef dvi As DLLVERSIONINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Property Get AddressOfBrowseCallbackProc() As Long
    AddressOfBrowseCallbackProc = GetAddressOf(AddressOf BrowseCallbackProc)
End Property

Public Function Shell32VersionOrLater(ByVal MajorVersion As Long) As Boolean
    Dim DLLVERSIONINFO As DLLVERSIONINFO

    On Error Resume Next
    DLLVERSIONINFO.cbSize = LenB(DLLVERSIONINFO)
    DllGetVersion DLLVERSIONINFO
    Shell32VersionOrLater = DLLVERSIONINFO.dwMajorVersion >= MajorVersion
End Function

Private Function BrowseCallbackProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal lParam As Long, _
    ByVal lpData As SHFolderDlg) As Long
    
    On Error Resume Next
    With lpData
        Select Case uMsg
            Case BFFM_INITIALIZED
                SendMessage hWnd, BFFM_SETSELECTION, WIN32_TRUE, StrPtr(.StartPath)
                If Shell32VersionOrLater(6) Then
                    If .ExpandStartPath Then
                        SendMessage hWnd, BFFM_SETEXPANDED, WIN32_TRUE, StrPtr(.StartPath)
                    End If
                    If Len(.OkCaption) Then
                        SendMessage hWnd, BFFM_SETOKTEXT, 0, StrPtr(.OkCaption)
                    End If
                End If
            Case BFFM_SELCHANGED
                SendMessage hWnd, BFFM_SETSTATUSTEXT, WIN32_FALSE, lParam
            Case BFFM_VALIDATEFAILED
                'Returning WIN32_TRUE keeps the dialog displayed instead of canceling:
                BrowseCallbackProc = IIf(.CancelValidation, WIN32_FALSE, WIN32_TRUE)
        End Select
    End With
End Function

Private Function GetAddressOf(ByVal TheAddressOf As Long)
    GetAddressOf = TheAddressOf
End Function

