VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Расчёт параметров findPeaks"
   ClientHeight    =   2385
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "закрыть"
      Height          =   360
      Left            =   2025
      TabIndex        =   2
      Top             =   1935
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   1785
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   45
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "копировать"
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   1935
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Command1_Click()
   Clipboard.Clear
   Clipboard.SetText Text1.Text
End Sub


Private Sub Command2_Click()
  Me.Hide
End Sub
