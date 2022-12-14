Attribute VB_Name = "myGoldLib"
'Все полезные константы, функции и декларации

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long


Declare Function SetFocusAPI Lib "user32" Alias _
    "SetFocus" (ByVal hWnd As Long) As Long
Declare Function WindowFromPointXY Lib "user32" Alias _
    "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) _
    As Long
Declare Function GetWindowText Lib "user32" Alias _
    "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, _
    ByVal cch As Long) As Long
    
'Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal Flags As Long) As Long

' Переключаем на русский
'Call ActivateKeyboardLayout(68748313, 0)
' Переключаем на английский
'Call ActivateKeyboardLayout(67699721, 0)
 Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Public Type RECT1
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
    value As Integer
    lDist As Single
    rDist As Single
    arrelem As Integer
    eventType As String
End Type


Public Type cEvent
    Left As Single
    center As Single
    Right As Single
End Type


'Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    
    
 'Интернет
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Declare Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" _
     (ByVal pCaller As Long, _
      ByVal szURL As String, _
      ByVal szFileName As String, _
      ByVal dwReserved As Long, _
      ByVal lpfnCB As Long) As Long

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, _
ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long) As Long

Private Const MB_PRECOMPOSED = &H1 ' use precomposed chars
Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
    
    

Public Const pi = 3.14159265358979




'ждём не подвешивая систему
Public Sub wait(s As Single)
  Dim t As Long, t1 As Long, s1 As Single, l As Long
  
     t = GetTickCount
      
    Do
      DoEvents
      t1 = GetTickCount
      
    Loop While t1 - t < s
End Sub


'строка в число с точкой 6 зн после зап
Public Function StrNum(s As Single) As String
  Dim s1 As String
  s1 = Format(s, "0.0#####")
  StrNum = Replace(s1, ",", ".")
End Function



Public Function Bin$(a)
'Преобразование десятичного в шестнадцатиричный код
Dim s As String, b, i As Integer, C, v
    C = ""
    b = Hex$(a)
    v = Len(b)
    For i = 1 To v
    s = Mid$(b, i, 1)
    Select Case s
    Case "0"
    s = "0000"
    Case "1"
    s = "0001"
    Case "2"
    s = "0010"
    Case "3"
    s = "0011"
    Case "4"
    s = "0100"
    Case "5"
    s = "0101"
    Case "6"
    s = "0110"
    Case "7"
    s = "0111"
    Case "8"
    s = "1000"
    Case "9"
    s = "1001"
    Case "A"
    s = "1010"
    Case "B"
    s = "1011"
    Case "C"
    s = "1100"
    Case "D"
    s = "1101"
    Case "E"
    s = "1110"
    Case "F"
    s = "1111"
    Case Else
    s = "Error"
    End Select
    If i = 1 Then C = s Else C = C + " " + s
    Next i
    Bin$ = C
End Function


Public Function HeDec(a$)
'Преобразование шестноадц. в десятичный код
Dim s As String, b, i As Integer, C, v
    C = 0
    b = "&h" + a$
    s = Val(b)
    HeDec = s
End Function


Public Sub WriteGur(s As String)
     Open App.Path + "\" + Format(Now, "ddmmyy") + ".txt" For Append As #22
       Print #22, s
     Close #22
End Sub


Public Sub CheckPrev()
    If App.PrevInstance Then
      MsgBox "Программа уже запущена"
      End
    End If
End Sub






Public Sub clearDebug(hWnd As Long)
Dim x As Long
Dim y As Long
Dim xStep As Long
Dim yStep As Long
Dim hWndOver As Long
Dim hWndLast As Long
Dim blnExit As Long
Dim sWindowText As String * 100

    xStep = Screen.TwipsPerPixelX
    yStep = Screen.TwipsPerPixelY
    
    For x = 0 To Screen.Width Step xStep
        For y = Screen.Height To 0 Step -yStep
            
            hWndOver = WindowFromPointXY(x, y)
            
            If hWndOver <> hWndLast Then
                hWndLast = hWndOver
                
                If Left(sWindowText, GetWindowText(hWndOver, sWindowText, 100)) = "Immediate" Then
                    blnExit = True
                    Exit For
                End If
            End If
        Next y
        If blnExit Then
            Exit For
        End If
    Next x
    
    If blnExit Then
        SetFocusAPI hWndOver
        Sendkeys "^{HOME}+^{END}", True
        'SendKeys "{F5}{DEL}{F5}", True
    
    End If
    
  Dim i As Integer
   Sendkeys "^{A}", True
   Sendkeys "1", True
   For i = 1 To 100
      Sendkeys "{DOWN}", True
   Next
   
   For i = 1 To 100
     Debug.Print " " + vbCrLf
   Next
    'Wait 1000
    
    ''Return focus back to the app
    'SetFocusAPI hwnd
End Sub



Public Function Translit(txt As String) As String
     Dim i As Integer, j As Integer, с As String, flag As Integer, outchr As String, outstr As String
     Dim Rus As Variant
     Rus = Array("а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", _
     "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", _
     "щ", "ъ", "ы", "ь", "э", "ю", "я", "А", "Б", "В", "Г", "Д", "Е", _
     "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", _
     "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я")
     
     Dim Eng As Variant
     Eng = Array("a", "b", "v", "g", "d", "e", "jo", "zh", "z", "i", "j", _
     "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", _
     "sh", "sch", "''", "y", "'", "e", "yu", "ya", "A", "B", "V", "G", "D", _
     "E", "JO", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", _
     "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SCH", "''", "Y", "'", "E", "YU", "YA")
     
         For i = 1 To Len(txt)
         с = Mid(txt, i, 1)
         
         flag = 0
         For j = 0 To 65
         If Rus(j) = с Then
         outchr = Eng(j)
         flag = 1
         Exit For
         End If
         Next j
        If flag Then outstr = outstr & outchr Else outstr = outstr & с
         Next i
         
         Translit = outstr
End Function



' =====================================================
' ======================= electronic
' =====================================================

'Adjustable Regulators
Private Function CalcR()
  Dim Vref As Single, r1 As Single, r2 As Single
  'LM2991 Negative Low-Dropout Adjustable Regulator
  'VOUT = VREF (1 + R2/R1)
  ' -4V  R1=220 R2=507 '510
  Vref = -1.21: r1 = 300: Vout = -4: r2 = (Vout / Vref - 1) * r1: Debug.Print r2
  'LM2931-N Series Low Dropout Regulators
  ' 4V  R1=220 R2=513  '510
  Vref = 1.2: r1 = 300: Vout = 4: r2 = (Vout / Vref - 1) * r1: Debug.Print r2
  'LM317
  Vref = 1.25: r1 = 300: Vout = 4: r2 = (Vout / Vref - 1) * r1: Debug.Print r2
  
CalcR = Vref
End Function


Private Function CalcRDobPar()
  Dim Vref As Single, r1 As Single, r2 As Single, Rzad As Single
  
  Rzad = 525:  r1 = 680:  r2 = (r1 * Rzad / (r1 - Rzad)): Debug.Print r2
  
CalcRDobPar = r2
End Function


Public Function ToSerial_r(w As Single, r As Single, C As Single) As Single
  ToSerial_r = r / (w * w * r * r * C * C / (10 ^ 24) + 1)
End Function

Public Function ToSerial_c(w As Single, r As Single, C As Single) As Single
  If w = 0 Or r = 0 Or C = 0 Then Exit Function
  ToSerial_c = C + (10 ^ 24) / (w * w * r * r * C)
End Function


Public Function ToColor(r As Byte, g As Byte, b As Byte) As Long
  Dim l As Long
       l = b * 65536 + g * 256 + r
       ToColor = l
End Function

'тихое копирование
Public Sub FileCopy1(s As String, s1 As String)
  On Error GoTo err
  FileCopy s, s1
err:
End Sub
'тихое удаление
Public Sub Kill1(s As String)
  On Error GoTo err
  Kill s
err:
End Sub
'тихое пернеименование
Public Sub Rename1(sFileName As String, sNewFileName As String)
    'If Dir(sFileName, 16) = "" Then MsgBox "Нет такого файла", vbCritical, "Ошибка": Exit Sub
  On Error GoTo err
 
    Name sFileName As sNewFileName 'переименовываем файл
err:
 End Sub


'Public Sub TestDesktopDC(s As String)
'    Dim hdc As Long
'    Dim tR As RECT
'    Dim lCol As Long
'    hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
'    'SetBkMode hdc, 1
'    tR.Left = 60
'    tR.Top = 300
'    tR.Right = 640
'    tR.Bottom = 332
'    lCol = GetTextColor(hdc)
'    SetTextColor hdc, &HFF&
'    DrawText hdc, s, Len(s), tR, 2
'    'DrawText hdc, s, 10, tR, 2
'    SetTextColor hdc, lCol
'    DeleteDC hdc
'End Sub


Public Sub OpenFolder(s As String)
  Dim l As Long
    
    l = ShellExecute(0, "open", s, "", "", 1)

End Sub

'доработать!!!!!!!!!!!!!!!!!

Public Sub Copy_File(sFileName As String, sNewFileName As String)
    If Dir(sFileName, 16) = "" Then MsgBox "Нет такого файла", vbCritical, "Ошибка": Exit Sub
 
    FileCopy sFileName, sNewFileName 'копируем файл
    MsgBox "Файл скопирован", vbInformation, "www.excel-vba.ru"
End Sub
'ПЕРЕМЕЩЕНИЕ:
Public Sub Move_File()
    Dim sFileName As String, sNewFileName As String
 
    sFileName = "C:\WWW.xls"    'имя исходного файла
    sNewFileName = "D:\WWW.xls"    'имя файла для перемещения. Директория(в данном случае диск D) должна существовать
    If Dir(sFileName, 16) = "" Then MsgBox "Нет такого файла", vbCritical, "Ошибка": Exit Sub
 
    Name sFileName As sNewFileName 'перемещаем файл
    MsgBox "Файл перемещен", vbInformation, "www.excel-vba.ru"
End Sub
'ПЕРЕИМЕНОВАНИЕ:
Public Sub Rename_File(sFileName As String, sNewFileName As String)
    If Dir(sFileName, 16) = "" Then MsgBox "Нет такого файла", vbCritical, "Ошибка": Exit Sub
 
    Name sFileName As sNewFileName 'переименовываем файл
 
    'MsgBox "Файл переименован", vbInformation, "www.excel-vba.ru"
End Sub

'УДАЛЕНИЕ ФАЙЛА:
Public Sub Delete_File()
    Dim sFileName As String
 
    sFileName = "C:\WWW.xls"    'имя файла для удаления
 
    If Dir(sFileName, 16) = "" Then MsgBox "Нет такого файла", vbCritical, "Ошибка": Exit Sub
 
    Kill sFileName 'удаляем файл
    MsgBox "Файл удален", vbInformation, "www.excel-vba.ru"
End Sub


'обмен значениями двух переменных
Public Sub Swap(ByRef a As Variant, ByRef b As Variant)
    Dim tmp As Variant
    tmp = a
    a = b
    b = tmp
End Sub


'деление с бесконечностью
Public Function DivInf(a As Single, b As Single) As String
 If b <> 0 Then
     DivInf = (Format(a / b, "0.0####"))
 Else
     DivInf = "inf"
 End If
End Function


'функции работы с клипбордом, не портящие текст
Public Function ClipboardText() ' чтение
   With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .GetFromClipboard
        ClipboardText = .GetText
    End With
End Function
 
Sub SetClipboardText(ByVal txt$) ' Запись
   With GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText txt$
        .PutInClipboard
    End With
End Sub




'For Windows 7: Change the UAC settings to never notify.
'
'For Windows 8 and 10:
'Add this method to any module:
'
Public Sub Sendkeys(Text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), wait
   Set WshShell = Nothing
End Sub
'It 's worked fine for me in windows 10.


'округление до .05
Public Function Okrugl(a As Single) As String
   Dim s As String, b As Single, s1 As String, i As Integer
   
   s = Trim(Str(a))
   
   i = InStr(1, s, ".")
   b = Val(Mid(s, i + 2, 1))
   
   a = Val(Mid(s, 1, i + 1))
    s = Trim(Str(a)) + s1
   
   If b < 3 Then
      s1 = "0"
      If InStr(1, s, ".") = 0 Then s1 = ".0"
      GoTo r
   End If
   If b < 8 Then
      s1 = "5"
      If InStr(1, s, ".") = 0 Then s1 = ".05"
      GoTo r
   End If
   a = a + 0.1
   s = Trim(Str(a)) + s1
   s1 = "0"
   If InStr(1, s, ".") = 0 Then s1 = ".0"
   
r:
    s = Trim(Str(a)) + s1

    If Left(s, 1) = "." Then s = "0" + s
    Okrugl = s
End Function



Public Function PathToFile(s As String) As String
  Dim i As Integer
  For i = Len(s) To 1 Step -1
    If Mid(s, i, 1) = "\" Then
      PathToFile = Right(s, Len(s) - i)
      Exit For
    End If
  Next
End Function




Public Function DownloadFile(sSourceUrl As String, _
                                sLocalFile As String) As Boolean

     'Download the file. BINDF_GETNEWESTVERSION forces
     'the API to download from the specified source.
     'Passing 0& as dwReserved causes the locally-cached
     'copy to be downloaded, if available. If the API
     'returns ERROR_SUCCESS (0), DownloadFile returns True.
      DownloadFile = URLDownloadToFile(0&, _
                                       sSourceUrl, _
                                       sLocalFile, _
                                       BINDF_GETNEWESTVERSION, _
                                       0&) = ERROR_SUCCESS

End Function





Private Function OpenURL(ByVal sUrl As String) As String
Dim hOpen As Long
Dim hOpenUrl As Long
Dim bDoLoop As Boolean
Dim bRet As Boolean
Dim sReadBuffer As String * 2048
Dim lNumberOfBytesRead As Long
Dim sBuffer As String
Dim ccc As Long



    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    DoEvents
    hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
    DoEvents
    bDoLoop = True
    
    ccc = 0
    
    While bDoLoop
        sReadBuffer = vbNullString
        ccc = ccc + 1
        'Label3.Caption = ccc
        DoEvents
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
    Wend
    
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer
End Function


Public Function Convert(ByVal strSrc As String, ByVal nFromCP As Long, ByVal nToCP As Long) As String
Dim nLen As Long
Dim strDst As String
Dim strRet As String
Dim nRet As Long
nLen = Len(strSrc)
strDst = String(nLen * 2, Chr(0))
strRet = String(nLen * 2, Chr(0))
nRet = MultiByteToWideChar(nFromCP, MB_PRECOMPOSED, strSrc, nLen, strDst, nLen)
nRet = WideCharToMultiByte(nToCP, 0, strDst, nRet, strRet, nLen * 2, ByVal 0, 0)
Convert = Left(strRet, nRet)
End Function
'Пример
'Имеем, допустим, TextBox. В нем текст в кодировке KOI, надо получить в Windows. Запускаешь:
'TextBox = StringConvert(TextBox, 20866, 1251)



Public Function ExtractNumber(s As String) As Long
  Dim i As Integer, k As Integer
  
  j = -1
  For i = 1 To Len(s)
     k = Asc(Mid(s, i, 1))
     If k > 47 And k < 58 Then
        For j = i To Len(s)
            k = Asc(Mid(s, j, 1))
            If k < 48 Or k > 57 Then GoTo 1
        Next
        GoTo 1
     End If
  Next
1
   If j > 0 Then
      ExtractNumber = Val(Mid(s, i, j - i))
   Else
      ExtractNumber = -1
   End If
End Function












