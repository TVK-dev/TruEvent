Attribute VB_Name = "main_sub"
Option Explicit

Public NumRect As Integer
Public Rects(5000) As RECT1
Public NumRectL As Integer
Public RectsLow(1000) As RECT1
Public CurrentFN As String, CurrentFNind As Long
Public minm(20) As Single, maxm(20) As Single
Public numSer As Long, LastEvent As Integer, iLastRect As Integer
Public GrafF(1000) As String, GrafF1(1000) As String, GrafF2(1000) As Long
Public flRun As Boolean, flChg As Boolean
Public sCurPath As String
Public point1 As Single, point2 As Single
Public koefMA As Single
Public iShGr As Integer, iDefEvent As Integer, iLastEvent As Integer
Public wOld As Long, hOld As Long

Public MasDat(1000, 1) As Long, fl As Boolean, fqdiv As Integer, fqdiv1 As Integer, flzap As Boolean
Public MasDat3(5000, 12, 12, 12) As Single

Public MethodsTxt(100) As String, NumMethods As Integer
            
Public indMas As Integer
Public NG As Integer, selg As Integer
Public ftime As Date, flft As Boolean
Public ProcSpeed As Integer
Public numPoint As Integer
Public masg() As Single, masg1() As Single, masgtmp(10000) As Single, max1 As Single, max2 As Single

Public Grafn(100, 10) As String, Grafn1(100, 10) As String
Public numGraf(100) As Integer ', numGraf1 As Integer, numGraf2 As Integer, numGraf3 As Integer

Public Indexes(20, 20) As Integer
Public indMas1 As Integer

Public Const numSeries = 3
Public sEvents(10) As String












