VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   195
      Left            =   3045
      TabIndex        =   13
      Top             =   2055
      Width           =   195
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   3000
      TabIndex        =   11
      Text            =   "10"
      Top             =   1590
      Width           =   825
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1185
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   375
      Left            =   3015
      TabIndex        =   7
      Top             =   825
      Value           =   1  'Checked
      Width           =   270
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   375
      Left            =   3015
      TabIndex        =   5
      Top             =   480
      Value           =   1  'Checked
      Width           =   270
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���������"
      Height          =   360
      Left            =   810
      TabIndex        =   3
      Top             =   2775
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������"
      Height          =   360
      Left            =   2475
      TabIndex        =   2
      Top             =   2775
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3015
      TabIndex        =   1
      Text            =   "0.95"
      Top             =   150
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����������� ������� "
      Height          =   270
      Index           =   5
      Left            =   105
      TabIndex        =   12
      Top             =   2040
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "������� �� ���������"
      Height          =   270
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1665
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����� �� ���������"
      Height          =   270
      Index           =   3
      Left            =   105
      TabIndex        =   8
      Top             =   1275
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���������� ������� �����"
      Height          =   270
      Index           =   2
      Left            =   105
      TabIndex        =   6
      Top             =   930
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "���������� ���������� �������"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   585
      Width           =   2835
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����������� ����������"
      Height          =   450
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   2835
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
  Dim i As Integer
  
  koefMA = Val(Text1.Text)
  
  iShGr = iShGr And 255 - 2 - 4
  iShGr = iShGr Or 2 * Sgn(Check1.value)
  iShGr = iShGr Or 4 * Sgn(Check2.value)
  
  Select Case Combo1.Text
    Case "������. �������"
     iDefEvent = 4
    Case "���"
     iDefEvent = 1
    Case "��������� ������"
     iDefEvent = 2
    Case "�������� ������"
     iDefEvent = 3
    Case "��������"
     iDefEvent = 5
  End Select
  iLastEvent = iDefEvent
  
  With CommForm
    .MA
    .GlobalMaxMin
    .FindMin
    ' ���� ��������� ��������
    If .CheckRazm Then
       .GetRects
    Else
       .FindMax
    End If
    
    .FindBaseline
    .grafik
    .DrawRects
  End With
End Sub


Private Sub Command2_Click()
  Me.Hide
End Sub


Private Sub Form_Load()
  Combo1.Clear
  Combo1.Text = "������. �������"
  Combo1.AddItem "������. �������"
  Combo1.AddItem "���"
  Combo1.AddItem "��������� ������"
  Combo1.AddItem "�������� ������"
  Combo1.AddItem "��������"
End Sub







