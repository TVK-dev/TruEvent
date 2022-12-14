VERSION 5.00
Begin VB.Form CommForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Кальциевые события"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14115
   Icon            =   "CommForm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   941
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "задать каталог"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   64
      Top             =   45
      Width           =   1350
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "параметры события"
      Height          =   930
      Left            =   1500
      TabIndex        =   56
      Top             =   -15
      Width           =   5970
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "перед"
         Height          =   195
         Index           =   5
         Left            =   5355
         TabIndex        =   71
         Top             =   345
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "дв. клеток"
         Height          =   195
         Index           =   3
         Left            =   4740
         TabIndex        =   70
         Top             =   690
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "фокус"
         Height          =   195
         Index           =   2
         Left            =   4740
         TabIndex        =   69
         Top             =   510
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "шум"
         Height          =   195
         Index           =   1
         Left            =   4740
         TabIndex        =   68
         Top             =   315
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "событие"
         Height          =   195
         Index           =   4
         Left            =   4740
         TabIndex        =   67
         Top             =   135
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         LargeChange     =   10
         Left            =   615
         Max             =   100
         TabIndex        =   60
         Top             =   255
         Width           =   3570
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   270
         LargeChange     =   10
         Left            =   615
         Max             =   100
         TabIndex        =   59
         Top             =   585
         Width           =   3570
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4215
         TabIndex        =   58
         Text            =   "30"
         Top             =   255
         Width           =   510
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4215
         TabIndex        =   57
         Text            =   "30"
         Top             =   585
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "конец"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   45
         TabIndex        =   62
         Top             =   615
         Width           =   600
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "начало"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   61
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Измерения"
      Height          =   930
      Left            =   7485
      TabIndex        =   47
      Top             =   -15
      Width           =   4455
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3270
         TabIndex        =   65
         Text            =   "0"
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1725
         TabIndex        =   51
         Text            =   "0"
         Top             =   300
         Width           =   660
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2430
         TabIndex        =   50
         Text            =   "0"
         Top             =   300
         Width           =   660
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2430
         TabIndex        =   49
         Text            =   "0"
         Top             =   615
         Width           =   675
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1725
         TabIndex        =   48
         Text            =   "0"
         Top             =   615
         Width           =   675
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "абс    длительность"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   2625
         TabIndex        =   66
         Top             =   105
         Width           =   1950
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "между точками"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   55
         Top             =   615
         Width           =   1560
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "от базовой линии "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   54
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2010
         TabIndex        =   53
         Top             =   105
         Width           =   555
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "абс"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2625
         TabIndex        =   52
         Top             =   105
         Width           =   600
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "перейти"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13335
      TabIndex        =   41
      Top             =   15
      Width           =   690
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   11010
      TabIndex        =   40
      Text            =   "Combo1"
      Top             =   7725
      Width           =   945
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      TabIndex        =   39
      Top             =   7725
      Width           =   10860
   End
   Begin VB.CommandButton Command18 
      Caption         =   "копировать"
      Height          =   525
      Left            =   14895
      TabIndex        =   38
      Top             =   900
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   180
      TabIndex        =   37
      Text            =   "1"
      Top             =   11700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   840
      TabIndex        =   36
      Text            =   "10"
      Top             =   11715
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Text12 
      Height          =   300
      Left            =   1440
      TabIndex        =   35
      Text            =   "1"
      Top             =   11700
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2040
      TabIndex        =   34
      Text            =   "2"
      Top             =   11700
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Text15 
      Height          =   885
      Left            =   14205
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Расчёт"
      Height          =   345
      Left            =   14895
      TabIndex        =   33
      Top             =   1425
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton Command19 
      Caption         =   "сохранить"
      Height          =   255
      Left            =   90
      TabIndex        =   32
      Top             =   330
      Width           =   1365
   End
   Begin VB.CommandButton Command17 
      Caption         =   "загрузить"
      Height          =   300
      Left            =   9135
      TabIndex        =   31
      Top             =   -240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Копировать рис"
      Height          =   270
      Left            =   90
      TabIndex        =   30
      Top             =   600
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   120
      ScaleHeight     =   6675
      ScaleWidth      =   11775
      TabIndex        =   29
      Top             =   930
      Width           =   11835
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   11535
         TabIndex        =   42
         Top             =   45
         Width           =   11595
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Базовая линия"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   4
            Left            =   9600
            TabIndex        =   63
            Top             =   0
            Width           =   2145
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   4
            Left            =   9420
            Top             =   30
            Width           =   165
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   3645
            X2              =   3645
            Y1              =   265
            Y2              =   0
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   135
            X2              =   135
            Y1              =   250
            Y2              =   -15
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   3
            Left            =   7470
            Top             =   30
            Width           =   165
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Исходные данные"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   7650
            TabIndex        =   46
            Top             =   0
            Width           =   2145
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00000070&
            FillStyle       =   0  'Solid
            Height          =   165
            Index           =   2
            Left            =   5865
            Top             =   30
            Width           =   165
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Границы соб."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   6045
            TabIndex        =   45
            Top             =   0
            Width           =   2145
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Подтверждённые соб."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   3750
            TabIndex        =   44
            Top             =   0
            Width           =   2145
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00FF0000&
            Height          =   165
            Index           =   1
            Left            =   3570
            Top             =   30
            Width           =   165
         End
         Begin VB.Line Line1 
            X1              =   60
            X2              =   240
            Y1              =   180
            Y2              =   15
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Предложенные события, не сохр."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   43
            Top             =   15
            Width           =   3795
         End
         Begin VB.Shape Shape1 
            Height          =   165
            Index           =   0
            Left            =   60
            Top             =   30
            Width           =   165
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "параметры метода"
      Height          =   600
      Index           =   3
      Left            =   14220
      TabIndex        =   16
      Top             =   -90
      Visible         =   0   'False
      Width           =   1275
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Text            =   "25"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3060
         TabIndex        =   17
         Text            =   "1"
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Число точек"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Фильтрация"
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   255
         Width           =   1695
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "CommForm.frx":1CFA
      Left            =   14220
      List            =   "CommForm.frx":1CFC
      TabIndex        =   15
      Text            =   "Расст. между год.+k-мин"
      Top             =   2175
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   255
      Left            =   13380
      TabIndex        =   12
      Top             =   30
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command10 
      Caption         =   "задать каталог"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12105
      TabIndex        =   14
      Top             =   15
      Width           =   1230
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4605
      TabIndex        =   13
      Text            =   "1"
      Top             =   -225
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "серия1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14145
      MaskColor       =   &H008080FF&
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   7665
      Left            =   12090
      TabIndex        =   10
      Top             =   315
      Width           =   1950
   End
   Begin VB.Frame Frame3 
      Caption         =   "Прием информации"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   -4440
      TabIndex        =   0
      Top             =   0
      Width           =   4410
      Begin VB.CommandButton Command13 
         DisabledPicture =   "CommForm.frx":1CFE
         DownPicture     =   "CommForm.frx":2460
         Height          =   375
         Left            =   1800
         Picture         =   "CommForm.frx":2B1A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   6240
         Width           =   375
      End
      Begin VB.CommandButton cmdClearBuffer 
         Caption         =   "Буфер"
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   6240
         Width           =   945
      End
      Begin VB.TextBox txtTerm 
         Height          =   855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "CommForm.frx":327C
         Top             =   5040
         Width           =   4215
      End
      Begin VB.CommandButton cmdClearTable 
         Caption         =   "Таблица"
         Height          =   375
         Left            =   3330
         TabIndex        =   4
         Top             =   6240
         Width           =   945
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Прочитать"
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   6255
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Очистка:"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bufer Len"
         Height          =   300
         Left            =   2235
         TabIndex        =   2
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Длинна буфера приема"
         Height          =   330
         Left            =   195
         TabIndex        =   1
         Top             =   315
         Width           =   1890
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   11160
      TabIndex        =   28
      Top             =   3015
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   11400
      TabIndex        =   27
      Top             =   5535
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   11940
      TabIndex        =   26
      Top             =   705
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   12090
      TabIndex        =   25
      Top             =   5535
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   12690
      TabIndex        =   24
      Top             =   6375
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   12690
      TabIndex        =   23
      Top             =   7740
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   12090
      TabIndex        =   22
      Top             =   6375
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   12090
      TabIndex        =   21
      Top             =   7740
      Width           =   915
   End
   Begin VB.Menu mf1 
      Caption         =   "Файл"
      Begin VB.Menu mf13 
         Caption         =   "Сохранить разметку"
      End
      Begin VB.Menu mf14 
         Caption         =   "Сохранить события"
      End
      Begin VB.Menu mf15 
         Caption         =   "Выгрузить все"
      End
      Begin VB.Menu mf11 
         Caption         =   "Задать каталог"
      End
      Begin VB.Menu mf12 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu ms1 
      Caption         =   "Расчёт"
      Begin VB.Menu ms11 
         Caption         =   "Расчёт параметров"
      End
   End
   Begin VB.Menu mn1 
      Caption         =   "Настройки"
   End
End
Attribute VB_Name = "CommForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Buffer to hold input string
Dim Instring() As Byte
Dim intPortNumber As Integer
Dim hLogFile As Integer ' Handle of open log file.
Private shlShell As shell32.Shell
Private shlFolder As shell32.Folder
Private Const BIF_RETURNONLYFSDIRS = &H1


Private Sub cmbCommand_Click()
  '  lblCommandDescription = CommandDescription(Val("&H" & cmbCommand.List(cmbCommand.ListIndex)))
End Sub


Private Sub Check1_Click(Index As Integer)
  Dim cl As Long
   selg = Index
   Select Case selg
      Case 0
        cl = vbRed
      Case 1
        cl = vbGreen
      Case 2
        cl = vbYellow
      Case 3
        cl = vbBlue
   End Select
End Sub


Private Sub Combo1_Click()
   Dim i As Integer
    indMas1 = Val(Combo1.Text) * 2
    i = numPoint - indMas1
    If i < 0 Then i = 0
    HScroll3.max = i
    MA
    grafik
  
    FindMin
    FindBaseline
  
    If CheckRazm Then
      GetRects
    Else
      FindMax
    End If
    
    'FindMin
    'FindBaseline
    
    grafik
    DrawRects
End Sub


Private Sub Command5_Click()
   Dim sFiles As String
    If shlShell Is Nothing Then
        Set shlShell = New shell32.Shell
    End If
    ' посл парам 0 - c раб стола. "" - c обзора дисков
    'Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, "выберите каталог данных", 0, 0)
    ' для отладки
    Set shlFolder = shlShell.BrowseForFolder(Me.hWnd, "выберите каталог данных", 0, App.Path + "\data\")
    If shlFolder Is Nothing Then
        Exit Sub
    End If
    'Debug.Print shlFolder.Self.Path
    sFiles = shlFolder.Self.Path + "\"
    sCurPath = sFiles
    FindFiles (sFiles)
End Sub


Private Sub Command16_Click()
   Clipboard.Clear
   Clipboard.SetData Picture1.Image
End Sub

Private Sub Command17_Click()
 Dim i As Integer, st As String, j As Integer, s As String
    ClearVals
    CurrentFN = st
   
    
    For i = 1 To Len(st)
      If Mid$(st, i, 1) = "\" Then j = i
    Next
'    CommonDialog1.InitDir = Mid$(st, 1, j)
    
    If st = "" Then Exit Sub
      
    numPoint = LoadCSV(st)
    MA
    
    'FindMaxMin
    
    FindBaseline
    grafik
    DrawRects
    
End Sub



'save
Private Sub Command19_Click()
    SaveCSV (CurrentFN)
    flChg = False
    FindFiles (sCurPath)
End Sub


Private Sub Command20_Click()
   FindParamPeaks
End Sub


Private Sub Command3_Click()
    Dim s As String, l As Long
        l = ShellExecute(0, "open", sCurPath, "", "", 1)
End Sub


Private Sub Command10_Click()
    Dim FolderPath As String, k As Long
    
    FolderPath = App.Path + "\data\"
    With New SHFolderDlg
        .Instructions = "выберите каталог данных"
        .OkCaption = "Задать"
        .ExpandStartPath = True
        'It can even (optionally) return files as well as folders:
        .Flags = BIF_RETURNONLYFSDIRS _
              Or BIF_NONEWFOLDERBUTTON _
              Or BIF_NEWDIALOGSTYLE _
              Or BIF_UAHINT _
              Or BIF_SHAREABLE
        .Root = CSIDL_DRIVES 'CSIDL_DESKTOP is the default.
        'We'll repeat this until the user cancels:
        Do
            'When we begin this demo we'll be at the Root we selected because
            'FolderPath starts as empty:
            .StartPath = FolderPath
            FolderPath = .BrowseForFolder(Me)
            If Len(FolderPath) Then
                sCurPath = FolderPath + "\"
                k = FindFiles(FolderPath + "\")
            Else
                'Optional:
                Exit Do
            End If
        Loop While k < 1
    End With
    
    'sFiles = shlFolder.Self.Path + "\"


End Sub

Private Sub Form_Load()
'ИНИЦИАЛИЗАЦИЯ ПРОГРАММЫ
'=========================
Dim i As Integer, s As Integer, j As Integer, g As Integer
Dim intCounter%
Dim dd As Integer
Dim k1 As Single, k2 As Single, k3 As Single, k4 As Single

    Me.Top = 0
    
    numGraf(0) = 1
    j = 0
    s = 0
    g = 0
    
    'harmonic amp
    k1 = Rnd - 0.5
    k2 = Rnd - 0.5
    k3 = 2 * Rnd - 1
    'noise amp
    k4 = Rnd - 0.5
    
    For i = 0 To 1000
       MasDat3(i, j, g, s) = 50 + k1 * Cos(i / 7) + k2 * Sin(i / 3) + k3 * Sin(i / 250) + k4 * Rnd
    Next
    
    Combo1.Clear
    Combo1.Text = "400": indMas1 = 800
    Combo1.AddItem "100":    Combo1.AddItem "200"
    Combo1.AddItem "300":    Combo1.AddItem "400"
    Combo1.AddItem "500":    Combo1.AddItem "600"
    
    sEvents(4) = "event": sEvents(1) = "noise": sEvents(2) = "focus": sEvents(3) = "move"
    sEvents(5) = "transd"
    iDefEvent = 4
      iLastEvent = iDefEvent
    
    numPoint = 1000
    'CurrentFN = "D:\Магистратура ИББМ\наука\разметка\SplitDropCmp\data\1.csv"
    CurrentFN = App.Path + "\data\1.csv"
    'ClearVals
'    numPoint = LoadCSV(CurrentFN)
    i = numPoint - indMas1
    If i < 0 Then i = 0
    HScroll3.max = i
    FindFiles (App.Path + "\data\")
    sCurPath = App.Path + "\data\"
    
    koefMA = 0.95
    iShGr = 255

'    MA
'    grafik
'    GlobalMaxMin
'    FindMax
'    FindMin
'    FindBaseline
'    grafik
'    DrawRects

End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close
  End
End Sub


Private Sub Form_Resize()
   Dim l As Long, h As Long, w As Long, h2 As Long, w1 As Long, h3 As Long
    
    '  не вызывается при сворачивании и разворачивании из панели задач
    If Me.Height < 600 Or hOld < 600 Then
        wOld = Me.Width
        hOld = Me.Height
        Exit Sub
    End If
    
    wOld = Me.Width
    hOld = Me.Height
    
    l = Me.Width / 15 - List1.Width - 10
    h = Me.Height / 15 - List1.Top - 45
    w = Me.Width / 15 - List1.Width - 25
    h2 = Me.Height / 15 - Text15.Height - 60 - HScroll3.Height
    w1 = w - Combo1.Width
    h3 = Me.Height / 15 - HScroll3.Height - 50
    
    If l < 1 Or h < 1 Or w < 1 Or h2 < 1 Or w1 < 1 Or h3 < 1 Then Exit Sub
    
    List1.Left = l
    List1.Height = h
    Picture1.Width = w
    Picture1.Height = h2
    HScroll3.Top = h3
    Combo1.Top = h3
    HScroll3.Width = w1
    Command10.Left = l
    Command3.Left = l + Command10.Width
    Combo1.Left = w1 + 10
     
    If CheckRazm Then
      GetRects
    Else
      FindMax
    End If
    FindMin
    grafik
    DrawRects
    DrawEvent
End Sub


Private Sub HScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     HScroll2.SetFocus
  End If
End Sub


' слева
Private Sub HScroll1_Change()
   Dim l As Integer
   l = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
   If l < 0 Then l = 0
   If MasDat3(l, 0, 0, 10) = 2 Then MasDat3(l, 0, 0, 10) = 0

   Rects(iLastRect).lDist = HScroll1.value
   l = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
   If l < 1 Then l = 1: Rects(iLastRect).lDist = Rects(iLastRect).arrelem - 1
   Do
     If MasDat3(l, 0, 0, 10) = 0 Then Exit Do
     l = l + 1
     Rects(iLastRect).lDist = Rects(iLastRect).lDist - 1
   Loop
   If Rects(iLastRect).lDist > 1 Then
      MasDat3(l, 0, 0, 10) = 2
   Else
      Rects(iLastRect).lDist = 0
   End If
    
   flChg = True
   Text2.Text = Str(Rects(iLastRect).lDist)
   grafik
   DrawRects
   DrawEvent
   Text14.Text = Str((Rects(iLastRect).lDist + Rects(iLastRect).rDist) * 0.5)
   'HScroll2.SetFocus
End Sub

'Private Sub HScroll1_LostFocus()
'   Dim l As Integer
'   l = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
'   If l < 0 Then l = 0
'   If MasDat3(l, 0, 0, 10) = 2 Then MasDat3(l, 0, 0, 10) = 0
'
'   Rects(iLastRect).lDist = HScroll1.value
'   l = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
'   If l < 1 Then l = 1: Rects(iLastRect).lDist = Rects(iLastRect).arrelem - 1
'   If MasDat3(l, 0, 0, 10) = 3 Then l = l + 1: Rects(iLastRect).lDist = Rects(iLastRect).lDist - 1
'   MasDat3(l, 0, 0, 10) = 2
'
'   Text2.Text = Str(Rects(iLastRect).lDist)
'   grafik
'   DrawRects
'   DrawEvent
'End Sub

' справа
Private Sub HScroll2_Change()
   Dim l As Integer
   l = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
   If MasDat3(l, 0, 0, 10) = 3 Then MasDat3(l, 0, 0, 10) = 0
   Rects(iLastRect).rDist = HScroll2.value
   l = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
   If l > numPoint Then l = numPoint: Rects(iLastRect).rDist = l - Rects(iLastRect).arrelem
   
   Do
     If MasDat3(l, 0, 0, 10) = 0 Then Exit Do
     l = l - 1
     Rects(iLastRect).rDist = Rects(iLastRect).rDist - 1
   Loop
   If Rects(iLastRect).rDist > 1 Then
      MasDat3(l, 0, 0, 10) = 3
   Else
      Rects(iLastRect).rDist = 0
   End If
    
   flChg = True
   Text3.Text = Str(Rects(iLastRect).rDist)
   grafik
   DrawRects
   DrawEvent
   Text14.Text = Str((Rects(iLastRect).lDist + Rects(iLastRect).rDist) * 0.5)
End Sub

'Private Sub HScroll2_LostFocus()
'   Dim l As Integer
'   l = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
'   If MasDat3(l, 0, 0, 10) = 3 Then MasDat3(l, 0, 0, 10) = 0
'   Rects(iLastRect).rDist = HScroll2.value
'   l = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
'   If l > numPoint Then l = numPoint: Rects(iLastRect).rDist = l - Rects(iLastRect).arrelem
'   If MasDat3(l, 0, 0, 10) = 2 Then l = l - 1: Rects(iLastRect).rDist = Rects(iLastRect).rDist - 1
'   MasDat3(l, 0, 0, 10) = 3
'   Text3.Text = Str(Rects(iLastRect).rDist)
'   grafik
'   DrawRects
'   DrawEvent
'End Sub


Private Sub HScroll3_Change()
  grafik
  DrawRects
  DrawEvent
End Sub


Private Sub List1_Click()
  Dim i As Integer, j As Integer, s As String
    
  If flChg = True Then
       i = MsgBox("Перейти к новым данным ? Изменения не сохранятся.", vbYesNo, "Загрузка")
       If i <> vbYes Then Exit Sub
  End If
  flChg = False
    
  For i = 0 To numSer - 1
     If List1.Selected(i) Then Exit For
  Next
  
  s = GrafF(i)
  If s = "" Then Exit Sub
  Me.Caption = "Кальциевые события  " + GrafF1(i)
  
  Erase Rects
      
  ClearVals
  CurrentFN = s
  CurrentFNind = i
  numPoint = LoadCSV(s)
  i = numPoint - indMas1
  If i < 0 Then i = 0
  HScroll3.max = i
  
  MA
  GlobalMaxMin
  FindMin
  FindBaseline
  
  ' если сохранена разметка
  If CheckRazm Then
     GetRects
  Else
     FindMax
  End If
  
  
  grafik
  DrawRects

End Sub


Public Sub grafik()
  
  Dim min(20, 20) As Single, max(20, 20) As Single
  Dim i As Integer, j As Integer, k As Integer, g As Integer, s As Integer, jj As Integer, CColor As Long
  Dim ss As String, ss1 As String
  Dim xp1 As Single, xp2 As Single, yp1 As Single, yp2 As Single
  Dim x1 As Integer, x2 As Integer
  Dim C As Single, w As Integer, fps As Single
  Dim hsv As Integer, lv As Integer, hv As Integer
  Dim pw As Long, ph As Long, kx As Single, ky As Single
  
   Me.Cls
          
    For k = 0 To numSeries
        For g = 0 To numGraf(k) - 1
            max(0, g) = 0
            min(0, g) = 10000000#
        Next
    Next
    
    
    For s = 0 To numSeries
        For g = 0 To numGraf(s) - 1
            For j = 0 To 4
                For i = 1 To k
                   If max(j, g) < MasDat3(i, j, g, s) Then max(j, g) = MasDat3(i, j, g, s)
                   If min(j, g) > MasDat3(i, j, g, s) Then min(j, g) = MasDat3(i, j, g, s)
                Next
            Next
        Next
    Next
    
    Axis
    
    With Picture1
    
    hsv = HScroll3.value
    lv = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
    hv = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
    pw = .Width * 15
    ph = .Height * 15
    kx = (pw - 1000) / indMas1
    If (maxm(0) - minm(0)) <> 0 Then ky = (ph - 1000) / (maxm(0) - minm(0))
        
    For k = 0 To numSeries
     
         If iShGr And 2 ^ k Then
        
                 Select Case k
                   Case 0
                     CColor = vbRed
                   Case 1
                     CColor = vbGreen
                   Case 2
                     CColor = vbBlue
                   Case 3
                     CColor = vbYellow
                   Case 4
                     CColor = &HC0E0FF
                   Case 5
                     CColor = &HFFC0FF
                 End Select
                
            For g = 0 To numGraf(k) - 1
               
                  Me.ForeColor = &HFF&
                         
                 'цикл задания толщины линии
                w = 3
                For jj = 0 To w * 15 - 10 Step 15
                     For i = hsv + 1 To hsv + indMas1
                        If k = 0 Then
                            If i > lv And i < hv Then
                              CColor = &H70
                            Else
                              CColor = vbRed
                            End If
                        End If
            
                        xp1 = 500 + kx * (i - 1 - hsv)
                        yp1 = jj - 500 + ph - ky * (MasDat3(i - 1, 0, g, k) - minm(0))
                        xp2 = 500 + kx * (i - hsv)
                        yp2 = jj - 500 + ph - ky * (MasDat3(i, 0, g, k) - minm(0))
                        
                        Picture1.Line (xp1, yp1)-(xp2, yp2), CColor
                        
                     Next
              Next jj
              'DoEvents
             
            Next g
         
         End If
     
    Next k
     
    End With
      
       Dim vAxTxt As String, hAxTxt As String
       Dim nZn As Integer, mulV As Single, mulH As Single
       
       'подписи осей
         mulH = 1: mulV = 1
       
       'выводим метки вертикальная ось
        Picture1.FontSize = Val(Text9.Text) + 2 '12
        Picture1.FontBold = True
        Picture1.CurrentY = 100
        Picture1.CurrentX = 100
        'Picture1.Print vAxTxt
        Picture1.FontSize = Val(Text9.Text)  '12
        
       
       For i = 0 To 10 Step Val(Text12.Text)  '2
            
            Picture1.CurrentY = Picture1.Height * 15 - 600 - ((Picture1.Height) * 15 - 1000) * i / 10 - i
            Picture1.CurrentX = 5
            Picture1.Print Format(minm(x2) * mulV + i * (maxm(x2) - minm(x2)) * mulV / 10, "0.0#")
            'Picture1.Print Okrugl(minm(x2) + i * (maxm(x2) - minm(x2)) / 10)
       Next
      
        'Picture1.FontSize = 12
        Picture1.CurrentY = Picture1.Height * 15 - 700
        Picture1.CurrentX = Picture1.Width * 15 - 400
        Picture1.Print "c"
        'Picture1.FontSize = 12
      
       
    'горизонтальная ось время
    ' на деление 900/15/2 = 20c  2 - fps
     fps = 2
     C = indMas1 / 15 / fps
    
     For i = 0 To 15
        Picture1.CurrentY = Picture1.Height * 15 - 300
        Picture1.CurrentX = (Picture1.Width * 15 - 1000) * i / 15 + 400
        'If i < 2 Or Not (ss1 = "0" And ss = "0") Then
        Picture1.Print Format(i * C + HScroll3.value / fps, "#0")
     Next

3
  Picture1.FontBold = False
End Sub

  
Private Function spc1(i As Integer) As String
   Dim k As Integer, s As String

   For k = 1 To i
    s = s + " "
   Next
    spc1 = s
End Function


Public Sub Axis()
  Dim i As Integer, j As Integer, n As Integer, k As Integer
  Dim h As Integer, w As Integer
  
  'With Picture1
  
  Picture1.Cls
  
  n = 0
  w = Picture1.Width * 15 - 1000
  h = Picture1.Height * 15 - 1000
  
  For i = 500 To w + 500 Step w / 15
     Picture1.Line (i, 500)-(i, h + 500), &HA0A0A0
  Next
    
  For i = 500 To h + 520 Step h / 10
     Picture1.Line (500, i)-(w + 500, i), &HA0A0A0
  Next
 
End Sub


Private Sub mf11_Click()
   Command10_Click
End Sub

Private Sub mf12_Click()
  Close
  End
End Sub

Private Sub mf13_Click()
  Command19_Click
End Sub

Private Sub mf14_Click()
  SaveEvents
End Sub

Private Sub mf15_Click()
  SaveEventsAll
End Sub

Private Sub mn1_Click()
  Form2.Show
End Sub

Private Sub ms11_Click()
  Form1.Show
  Form1.Text1.Text = FindParamPeaks
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
     Case 4
        Rects(iLastRect).eventType = "event"
     Case 1
        Rects(iLastRect).eventType = "noise"
     Case 2
        Rects(iLastRect).eventType = "focus"
     Case 3
        Rects(iLastRect).eventType = "move"
     Case 5
        Rects(iLastRect).eventType = "transd"
   End Select
   iLastEvent = Index
   MasDat3(Rects(iLastRect).arrelem, 0, 0, 9) = iLastEvent
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, l As Integer, k As Integer, xShft As Single
  Dim flRect As Boolean, h As Integer, ll As Integer
  Dim Perc As Single, iDist As Integer
  
'  Select Case Val(Combo1.Text)
'    Case 100, 200
'       iDist = 10
'    Case 300, 400
'       iDist = 20
'    Case 500, 600
'       iDist = 30
'  End Select
  iDist = Val(Form2.Text2.Text)
   
  xShft = HScroll3.value * 15 * ((Picture1.Width - 1000 / 15) / indMas1)
  flRect = False
  
  If Button = 1 Then
    For i = 0 To NumRect
      If Rects(i).Left < x + xShft And Rects(i).Right > x + xShft And Rects(i).Top < y And Rects(i).Bottom > y Then
         l = Rects(i).arrelem
         ' Расчёт параметров
         If MasDat3(l, 0, 0, 2) <> 0 Then
           Text4.Text = Format((100 * (MasDat3(l, 0, 0, 0) - MasDat3(l, 0, 0, 2)) / MasDat3(l, 0, 0, 2)), "0.0##")
         End If
         Text5.Text = Format(MasDat3(l, 0, 0, 0) - MasDat3(l, 0, 0, 2), "0.0##")
         Rects(i).value = 3 - Rects(i).value
         
         ' если разотмечаем, то убираем метки начали и конца из общего массива
         If MasDat3(l, 0, 0, 10) = 1 Then
            If Rects(i).lDist > 0 Then
              k = Rects(i).arrelem - Rects(i).lDist
              If k < 0 Then k = 0
              MasDat3(k, 0, 0, 10) = 0
            End If
            If Rects(i).rDist > 0 Then MasDat3(Rects(i).arrelem + Rects(i).rDist, 0, 0, 10) = 0
            MasDat3(l, 0, 0, 10) = 0
            MasDat3(l, 0, 0, 9) = 0
              
         ElseIf MasDat3(l, 0, 0, 10) = 0 Then
              If Rects(i).lDist = 0 Then
                  Rects(i).lDist = iDist
              End If
              If Rects(i).rDist = 0 Then
                  Rects(i).rDist = iDist
              End If
             
             iLastRect = i
             
             Select Case Rects(iLastRect).eventType
              Case "event"
                 iLastEvent = 4
              Case "noise"
                 iLastEvent = 1
              Case "focus"
                 iLastEvent = 2
              Case "move"
                 iLastEvent = 3
              Case "transd"
                 iLastEvent = 5
              Case Else
                 iLastEvent = iDefEvent
                 Rects(iLastRect).eventType = sEvents(iDefEvent)
             End Select
             Option1(iLastEvent).value = 1
             
             Rects(i).arrelem = l
             ll = Rects(i).arrelem - Rects(i).lDist
             If ll < 1 Then ll = 1: Rects(i).lDist = Rects(i).arrelem - 1
             If MasDat3(ll, 0, 0, 10) = 3 Then ll = ll + 1: Rects(i).lDist = Rects(i).lDist - 1
             MasDat3(ll, 0, 0, 10) = 2
             ll = Rects(i).arrelem + Rects(i).rDist
             If ll > numPoint Then ll = numPoint: Rects(i).rDist = ll - Rects(i).arrelem
             If MasDat3(ll, 0, 0, 10) = 2 Then ll = ll - 1: Rects(i).rDist = Rects(i).rDist - 1
             MasDat3(ll, 0, 0, 10) = 3
      
             HScroll1.value = CheckMax(Rects(i).lDist)
             HScroll2.value = CheckMax(Rects(i).rDist)
             MasDat3(l, 0, 0, 10) = 1
             MasDat3(l, 0, 0, 9) = iLastEvent
         Else
             'если наткнулись на метку начала или конца
             Do
               If MasDat3(l, 0, 0, 10) = 0 Then Exit Do
               l = l + 1
             Loop
             
             If Rects(i).lDist = 0 Then
                Rects(i).lDist = iDist
             End If
             If Rects(i).rDist = 0 Then
                 Rects(i).rDist = iDist
             End If
             
             iLastRect = i
             
           
             Select Case Rects(iLastRect).eventType
              Case "event"
                 iLastEvent = 4
              Case "noise"
                 iLastEvent = 1
              Case "focus"
                 iLastEvent = 2
              Case "move"
                 iLastEvent = 3
              Case "transd"
                 iLastEvent = 5
              Case Else
                 iLastEvent = iDefEvent
                 Rects(iLastRect).eventType = sEvents(iDefEvent)
             End Select
             Option1(iLastEvent).value = 1

             
             Rects(i).arrelem = l
             ll = Rects(i).arrelem - Rects(i).lDist
             If ll < 1 Then ll = 1: Rects(i).lDist = Rects(i).arrelem - 1
             If MasDat3(ll, 0, 0, 10) = 3 Then ll = ll + 1: Rects(i).lDist = Rects(i).lDist - 1
             MasDat3(ll, 0, 0, 10) = 2
             ll = Rects(i).arrelem + Rects(i).rDist
             If ll > numPoint Then ll = numPoint: Rects(i).rDist = ll - Rects(i).arrelem
             If MasDat3(ll, 0, 0, 10) = 2 Then ll = ll - 1: Rects(i).rDist = Rects(i).rDist - 1
             MasDat3(ll, 0, 0, 10) = 3
      
             HScroll1.value = CheckMax(Rects(i).lDist)
             HScroll2.value = CheckMax(Rects(i).rDist)
             MasDat3(l, 0, 0, 10) = 1
             MasDat3(l, 0, 0, 9) = iLastEvent
         End If
        
         grafik
         DrawRects
         DrawEvent
         flRect = True
         HScroll1.SetFocus
      End If
    Next
    
    If flRect = False Then
       i = MsgBox("Добавить событие в этом месте?", vbYesNo, "Новое событие")
       If i <> vbYes Then Exit Sub
     
       h = 6
       Rects(NumRect).Left = x - 15 * h + HScroll3.value * 15 * ((Picture1.Width - 1000 / 15) / indMas1)
       Rects(NumRect).Right = x + 15 * h + HScroll3.value * 15 * ((Picture1.Width - 1000 / 15) / indMas1)
       Rects(NumRect).Top = y - 15 * h
       Rects(NumRect).Bottom = y + 15 * h
       If x < 500 Then x = 500
       i = (x - 500) * indMas1 / (Picture1.Width * 15 - 1000) + HScroll3.value
       Rects(NumRect).arrelem = i
       MasDat3(i, 0, 0, 10) = 1 '- MasDat3(i, 0, 0, 10)
       MasDat3(l, 0, 0, 9) = iDefEvent
       Rects(NumRect).value = 3
       
       If Rects(NumRect).lDist = 0 Then Rects(NumRect).lDist = iDist
       If Rects(NumRect).rDist = 0 Then Rects(NumRect).rDist = iDist
        
       ll = Rects(NumRect).arrelem - Rects(NumRect).lDist
       If ll < 1 Then ll = 1: Rects(NumRect).lDist = Rects(NumRect).arrelem - 1
       If MasDat3(ll, 0, 0, 10) = 3 Then ll = ll + 1: Rects(NumRect).lDist = Rects(NumRect).lDist - 1
       MasDat3(ll, 0, 0, 10) = 2
       ll = Rects(NumRect).arrelem + Rects(NumRect).rDist
       If ll > numPoint Then ll = numPoint: Rects(NumRect).rDist = ll - Rects(NumRect).arrelem
       If MasDat3(ll, 0, 0, 10) = 2 Then ll = ll - 1: Rects(NumRect).rDist = Rects(NumRect).rDist - 1
       MasDat3(ll, 0, 0, 10) = 3
        
       HScroll1.value = CheckMax(Rects(NumRect).lDist)
       HScroll2.value = CheckMax(Rects(NumRect).rDist)
        
       iLastRect = NumRect
       NumRect = NumRect + 1
       grafik
       DrawRects
       DrawEvent
    End If
 Else
    
    i = (x - 500) * indMas1 / (Picture1.Width * 15 - 1000) + HScroll3.value
    If point1 = 0 Then
       point1 = MasDat3(i, 0, 0, 0)
       Text6.Text = "Укажите"
       Text7.Text = "2 точку"
    Else
       point2 = MasDat3(i, 0, 0, 0)
       Perc = 100 * (point2 - point1) / (point1)
       Text6.Text = Format(Perc, "0.0##")
       Text7.Text = Format(Abs(point2 - point1), "0.0##")
       point1 = 0
    End If
 End If
  
End Sub



Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, xShft As Single
 
  xShft = HScroll3.value * 15 * ((Picture1.Width - 1000 / 15) / indMas1)
  For i = 0 To NumRect
    If Rects(i).value < 2 Then
        Rects(i).value = 0
        If Rects(i).Left < x + xShft And Rects(i).Right > x + xShft And Rects(i).Top < y And Rects(i).Bottom > y Then
           Rects(i).value = 1 '- Rects(i).value
           DrawRects
        End If
    End If
  Next
  'Picture1.ToolTipText = "Нажмите левую клавишу чтобы добавить событие, правую - для измерения расстояния"
End Sub


Public Sub DrawRects()
  Dim i As Integer, w As Integer, h As Integer, ccol As Long, ccol1 As Long
  Dim x As Single, y As Single, xShft As Single
  With CommForm.Picture1
     
  For i = 0 To NumRect - 1
        
        If Rects(i).value = 0 Then
            ccol = 0
            ccol1 = ccol
        ElseIf Rects(i).value = 1 Then
            ccol = vbGreen
            ccol1 = ccol
        Else
            ccol = vbBlue
            ccol1 = vbWhite
            Picture1.CurrentX = Rects(i).Left + 200 - xShft
            Picture1.CurrentY = Rects(i).Top
            Picture1.Print Left(Rects(i).eventType, 1)
        End If
        xShft = HScroll3.value * 15 * ((Picture1.Width - 1000 / 15) / indMas1)
        ' пересчёт из Rects(i).arrelem
        
        x = (Rects(i).Right + Rects(i).Left) / 2 - xShft
        y = (Rects(i).Bottom + Rects(i).Top) / 2
        Picture1.Line (x, y - 15 * 12)-(x, y + 15 * 12), ccol
        'Picture1.Circle (x - Rects(i).lDist * 15, y), 50, ccol
        'Picture1.Circle (x + Rects(i).rDist * 15, y), 50, ccol
        Picture1.Line (Rects(i).Left - xShft, Rects(i).Top)-(Rects(i).Left - xShft, Rects(i).Bottom), ccol
        Picture1.Line (Rects(i).Right - xShft, Rects(i).Top)-(Rects(i).Right - xShft, Rects(i).Bottom), ccol
        Picture1.Line (Rects(i).Left - xShft, Rects(i).Top)-(Rects(i).Right - xShft, Rects(i).Top), ccol
        Picture1.Line (Rects(i).Left - xShft, Rects(i).Bottom)-(Rects(i).Right - xShft, Rects(i).Bottom), ccol
        ' диаг
        Picture1.Line (Rects(i).Left - xShft, Rects(i).Bottom)-(Rects(i).Right - xShft, Rects(i).Top), ccol1

  Next
  End With

End Sub



Public Function LoadCSV(fn As String) As Integer
   Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
   Dim s As String
   
        Open fn For Input As #1
        'For i = 1 To 1000
        i = 1
        Do
            If Not EOF(1) Then Line Input #1, s Else Exit Do 'For
            
            If InStr(1, s, ";") > 0 Then i1 = InStr(1, s, ";") Else i1 = Len(s)
            If InStr(i1 + 1, s, ";") > 0 Then i2 = InStr(i1 + 1, s, ";") Else i2 = Len(s)
            If InStr(i2 + 1, s, ";") > 0 Then i3 = InStr(i2 + 1, s, ";") Else i3 = Len(s)
            
            i4 = InStr(i3 + 1, s, ";")
            i5 = InStr(i4 + 1, s, ";")

            MasDat3(i, 0, 0, 0) = Val(Replace(Mid$(s, 1, i1), ",", "."))
            MasDat3(i, 0, 0, 10) = Val(Replace(Mid$(s, i1 + 1, i2 - i1), ",", "."))
            If MasDat3(i, 0, 0, 10) = 1 Then
               i = i
            End If
            If i3 <> 0 Then MasDat3(i, 0, 0, 9) = Val(Replace(Mid$(s, i2 + 1, i3 - i2), ",", "."))
            'MasDat3(i, 0, 0, k) = Val(Replace(Mid$(s, i3 + 1, i4 - i3 - 1), ",", "."))
            'MasDat3(i, 0, 0, k) = Val(Replace(Mid$(s, i4 + 1, i5 - i4 - 1), ",", "."))
            'MasDat3(i, 0, 0, k) = Val(Replace(Mid$(s, i5 + 1, Len(s) - i5), ",", "."))
            i = i + 1
        Loop While Not EOF(1)
        'Next
        
cc:        Close #1
          
   LoadCSV = i - 1

End Function


Public Function SaveCSV(s As String) As Integer
   Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
        If s = "" Then Exit Function
        Open s For Output As #1
        For i = 1 To numPoint
          
            Print #1, Format(MasDat3(i, 0, 0, 0), "0.0000"), ";", _
            Str(MasDat3(i, 0, 0, 10)), ";", Str(MasDat3(i, 0, 0, 9))
            
        Next
        
        Close #1
End Function


Public Sub SaveEvents()
   Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
   Dim s As String
   Dim sAbsVal As String, sPerc As String, sDur As String
   
      's = Left(sCurPath, Len(sCurPath) - 1)
      j = 0
      For i = Len(sCurPath) To 1 Step -1
        If Mid(sCurPath, i, 1) = "\" Then
           j = j + 1
           If j >= 4 Then Exit For
        End If
      Next
      s = Mid(sCurPath, i, Len(sCurPath) - i)
      s = Replace(s, "\", "")
  
      Open sCurPath + s + ".csv" For Append As #2
        
      If FileLen(sCurPath + s + ".csv") < 3 Then
            Print #2, "NumCell; AbsVal; Perc; Duration"
      End If
        
      For i = 0 To NumRect - 1
         
        If MasDat3(Rects(i).arrelem, 0, 0, 2) <> 0 Then
          sPerc = Format((100 * (MasDat3(Rects(i).arrelem, 0, 0, 0) - MasDat3(Rects(i).arrelem, 0, 0, 2)) / MasDat3(Rects(i).arrelem, 0, 0, 2)), "0.0##")
        End If
        sAbsVal = Format(MasDat3(Rects(i).arrelem, 0, 0, 0) - MasDat3(Rects(i).arrelem, 0, 0, 2), "0.0##")
        sDur = Format((Rects(i).lDist + Rects(i).rDist) * 0.5, "0.0##")
        Print #2, GrafF2(CurrentFNind), ";", sAbsVal, ";", sPerc, ";", sDur
         
      Next

      Close #2
End Sub



Public Sub SaveEventsAll()
   Dim i As Integer, j As Integer, i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer
   Dim s As String
   Dim sAbsVal As String, sPerc As String, sDur As String
   Dim sBaseCat As String
   
      Close
      's = Left(sCurPath, Len(sCurPath) - 1)
      j = 0
      For i = Len(sCurPath) To 1 Step -1
        If Mid(sCurPath, i, 1) = "!" Then Exit For
        If Mid(sCurPath, i, 1) = "\" Then
           j = j + 1
           
           If j >= 4 Then Exit For
        End If
      Next
      s = Mid(sCurPath, i, Len(sCurPath) - i)
      s = Replace(s, "\", "")
      sBaseCat = App.Path + "\data\"
  
      Open sBaseCat + s + ".csv" For Output As #2
      Open sBaseCat + "events\" + s + ".csv" For Output As #3
      Open sBaseCat + "noise\" + s + ".csv" For Output As #4
      Open sBaseCat + "focus\" + s + ".csv" For Output As #5
      Open sBaseCat + "move\" + s + ".csv" For Output As #6
        
      If FileLen(sBaseCat + s + ".csv") < 3 Then
            Print #2, "NumCell; AbsVal; Perc; Duration; Value; Type"
            Print #3, "NumCell; AbsVal; Perc; Duration; Value"
            Print #4, "NumCell; AbsVal; Perc; Duration; Value"
            Print #5, "NumCell; AbsVal; Perc; Duration; Value"
            Print #6, "NumCell; AbsVal; Perc; Duration; Value"
      End If
      
    For j = 0 To numSer - 1
        ' если сохранена разметка
        If FileLen(GrafF(j)) > 25000 Then
            numPoint = LoadCSV(GrafF(j))
            
            MA
            GlobalMaxMin
            FindMin
            GetRects
            FindBaseline
            
  
            For i = 0 To NumRect - 1
               
              If MasDat3(Rects(i).arrelem, 0, 0, 2) <> 0 Then
                sPerc = Format((100 * (MasDat3(Rects(i).arrelem, 0, 0, 0) - MasDat3(Rects(i).arrelem, 0, 0, 2)) / MasDat3(Rects(i).arrelem, 0, 0, 2)), "0.0##")
              End If
              sAbsVal = Format(MasDat3(Rects(i).arrelem, 0, 0, 0) - MasDat3(Rects(i).arrelem, 0, 0, 2), "0.0##")
              sDur = Format((Rects(i).lDist + Rects(i).rDist) * 0.5, "0.0##")
              If sPerc > 0 And sAbsVal > 0 And sDur > 0 Then
                    Print #2, GrafF2(j), ";", sAbsVal, ";", sPerc, ";", sDur, ";", Format(MasDat3(Rects(i).arrelem, 0, 0, 0)), ";", Rects(i).eventType
                    Select Case Rects(i).eventType
                     Case "event"
                        Print #3, GrafF2(j), ";", sAbsVal, ";", sPerc, ";", sDur, ";", Format(MasDat3(Rects(i).arrelem, 0, 0, 0))
                     Case "noise"
                        Print #4, GrafF2(j), ";", sAbsVal, ";", sPerc, ";", sDur, ";", Format(MasDat3(Rects(i).arrelem, 0, 0, 0))
                     Case "focus"
                        Print #5, GrafF2(j), ";", sAbsVal, ";", sPerc, ";", sDur, ";", Format(MasDat3(Rects(i).arrelem, 0, 0, 0))
                     Case "move"
                        Print #6, GrafF2(j), ";", sAbsVal, ";", sPerc, ";", sDur, ";", Format(MasDat3(Rects(i).arrelem, 0, 0, 0))
                    
                    End Select
              End If
            Next i
            
        End If
        DoEvents
    Next j
    MsgBox "Данные выгружены"
    Close
End Sub



'предварительный отбор максимумов
'========== метод возвышение посерёдке ===================
Public Sub FindMaxMin_()
  Dim x As Single, y As Single, w As Integer, h As Single, i As Integer
  With Picture1
  'отметки максимумов
  w = 1
  h = 6
  NumRect = 0
  'indMas1 = 900

    For i = w To numPoint - w
        If MasDat3(i - w, 0, 0, 0) < MasDat3(i, 0, 0, 0) And MasDat3(i + w, 0, 0, 0) < MasDat3(i, 0, 0, 0) Then
           x = 500 + (.Width * 15 - 1000) * (i - 1) / indMas1
           y = -500 + .Height * 15 - (.Height * 15 - 1000) * (MasDat3(i, 0, 0, 0) - minm(0)) / (maxm(0) - minm(0))
           Rects(NumRect).Left = x - 15 * h
           Rects(NumRect).Right = x + 15 * h
           Rects(NumRect).Top = y - 15 * h
           Rects(NumRect).Bottom = y + 15 * h
           NumRect = NumRect + 1
        End If
    Next
    
  NumRectL = 0
    For i = w To numPoint - w
        If MasDat3(i - w, 0, 0, 0) > MasDat3(i, 0, 0, 0) And MasDat3(i + w, 0, 0, 0) > MasDat3(i, 0, 0, 0) Then
           RectsLow(NumRectL).Left = i
           RectsLow(NumRectL).Top = MasDat3(i, 0, 0, 0)
           NumRectL = NumRectL + 1
        End If
    Next
    
  End With
End Sub


Public Sub GlobalMaxMin()
  Dim i As Integer, j As Integer
  
     minm(0) = 10000000#
     maxm(0) = 0
 
     For i = 1 To numPoint
        If maxm(0) < MasDat3(i, 0, 0, 0) Then maxm(0) = MasDat3(i, 0, 0, 0)
        If minm(0) > MasDat3(i, 0, 0, 0) Then minm(0) = MasDat3(i, 0, 0, 0)
     Next
End Sub


' метод - максимум участка выше базовой линии
Public Sub FindMax()
  Dim x As Single, y As Single, w As Integer, h As Single, i As Integer, j As Integer
  Dim k As Integer, max As Single, kk As Integer, min As Single
   
  NumRect = 0
  h = 6
  'indMas1 = 900
   With Picture1
 
  j = 0
  Do
    For i = j + 1 To numPoint - 1
        If MasDat3(i, 0, 0, 0) > MasDat3(i, 0, 0, 1) Then Exit For
    Next
    For j = i + 1 To numPoint - 1
        If MasDat3(j, 0, 0, 0) < MasDat3(j, 0, 0, 1) Then Exit For
    Next
    max = -10000000#
    For k = i To j
        If MasDat3(k, 0, 0, 0) > max Then
            max = MasDat3(k, 0, 0, 0)
            kk = k
        End If
    Next
        
    ' отбор максимумов по уравнению из python
    If Form2.Check3 = 1 Then
        y = MasDat3(kk, 0, 0, 0) - MasDat3(kk, 0, 0, 2)
        If MasDat3(kk, 0, 0, 2) <> 0 Then x = 100 * (MasDat3(kk, 0, 0, 0) - MasDat3(kk, 0, 0, 2)) / MasDat3(kk, 0, 0, 2)
        If y >= 3.5063 - 0.05295 * x Then
            
            If indMas1 <> 0 Then x = 500 + (.Width * 15 - 1000) * (kk - 1) / indMas1
            If (maxm(0) - minm(0)) <> 0 Then y = -500 + .Height * 15 - (.Height * 15 - 1000) * (max - minm(0)) / (maxm(0) - minm(0))
            Rects(NumRect).Left = x - 15 * h
            Rects(NumRect).Right = x + 15 * h
            Rects(NumRect).Top = y - 15 * h
            Rects(NumRect).Bottom = y + 15 * h
            Rects(NumRect).arrelem = kk
            NumRect = NumRect + 1
        End If
    Else
            If indMas1 <> 0 Then x = 500 + (.Width * 15 - 1000) * (kk - 1) / indMas1
            If (maxm(0) - minm(0)) <> 0 Then y = -500 + .Height * 15 - (.Height * 15 - 1000) * (max - minm(0)) / (maxm(0) - minm(0))
            Rects(NumRect).Left = x - 15 * h
            Rects(NumRect).Right = x + 15 * h
            Rects(NumRect).Top = y - 15 * h
            Rects(NumRect).Bottom = y + 15 * h
            Rects(NumRect).arrelem = kk
            NumRect = NumRect + 1
    End If
    
  Loop While j < numPoint - 1
    
  End With
End Sub


Public Sub FindMin()
  Dim x As Single, y As Single, w As Integer, h As Single, i As Integer, j As Integer
  Dim k As Integer, max As Single, kk As Integer, min As Single
       
  NumRectL = 1
        RectsLow(0).Left = 0
        RectsLow(0).Top = MasDat3(0, 0, 0, 0)
        
  j = 0
  Do
    For i = j + 1 To numPoint - 1
        If MasDat3(i, 0, 0, 0) < MasDat3(i, 0, 0, 1) Then Exit For
    Next
    For j = i + 1 To numPoint - 1
        If MasDat3(j, 0, 0, 0) > MasDat3(j, 0, 0, 1) Then Exit For
    Next
    min = 10000000#
    For k = i To j
        If MasDat3(k, 0, 0, 0) < min Then
            min = MasDat3(k, 0, 0, 0)
            kk = k
        End If
    Next
        
        RectsLow(NumRectL).Left = kk
        RectsLow(NumRectL).Top = MasDat3(kk, 0, 0, 0)
        NumRectL = NumRectL + 1
        
  Loop While j < numPoint - 1
    
    RectsLow(NumRectL - 1).Left = numPoint
    RectsLow(NumRectL - 1).Top = MasDat3(numPoint, 0, 0, 0)
   ' NumRectL = NumRectL + 1

End Sub


Public Sub FindBaseline()
    Dim a As Single, ss As Single, i As Integer, j As Single
    Dim x1 As Single, x2 As Single, y1 As Single, y2 As Single
    
    numGraf(2) = 1
    
    For i = 0 To 5000
       MasDat3(i, 0, 0, 2) = 0
    Next
    
    MasDat3(0, 0, 0, 2) = MasDat3(5, 0, 0, 0)
    MasDat3(numPoint, 0, 0, 2) = MasDat3(numPoint, 0, 0, 0)
'    For i = 0 To numPoint
'       MasDat3(i, 0, 0, 2) = MasDat3(5, 0, 0, 0)
'    Next
    For i = 1 To NumRectL
       x1 = RectsLow(i - 1).Left
       x2 = RectsLow(i).Left
       y1 = RectsLow(i - 1).Top
       y2 = RectsLow(i).Top
       
       For j = x1 To x2
         If x2 - x1 <> 0 Then MasDat3(j, 0, 0, 2) = y1 + (j - x1) * (y2 - y1) / (x2 - x1)
       Next
     Next
    
End Sub


Public Sub MA()
    Dim a As Single, ss As Single, i As Integer
    
    a = koefMA '0.95
    numGraf(1) = 1
    ss = MasDat3(2, 0, 0, 0)
        For i = 1 To numPoint   '3-1000
           ss = ss * a + (1 - a) * MasDat3(i, 0, 0, 0)
           MasDat3(i, 0, 0, 1) = ss '- dispersion/2
        Next

End Sub


Public Function FindParamPeaks() As String
    Dim a As Single, ss As Single, i As Integer, min As Single
        
    min = 10000000#
    For i = 1 To numPoint
       If MasDat3(i, 0, 0, 10) = 1 Then
           ss = (MasDat3(i, 0, 0, 0) - MasDat3(i, 0, 0, 2))
           If (ss > 0) And (ss < min) Then
              min = ss
           End If
       End If
    Next
    'Text15.Text = "'MinPeakProminence', " + Str(min)
    FindParamPeaks = "'MinPeakProminence', " + Str(min)
End Function


Public Sub ClearVals()
  Dim i As Integer
  For i = 1 To numPoint
     MasDat3(i, 0, 0, 10) = 0
     MasDat3(i, 0, 0, 1) = 0
     Rects(i).value = 0
     Rects(i).arrelem = 0
     Rects(i).lDist = 0
     Rects(i).rDist = 0
  Next
End Sub




Public Function FindFiles(s As String) As Long
   Dim i As Integer, sFiles As String, j As Integer, ff As Integer

    List1.Clear
    sFiles = Dir(s + "*.csv")
    numSer = 0
    ff = 0
     
    Do While sFiles <> ""
    
       If sFiles <> "." And sFiles <> ".." Then
         GrafF1(numSer) = sFiles
         GrafF(numSer) = s + sFiles
         GrafF2(numSer) = ExtractNumber(sFiles)
         If GrafF2(numSer) < 0 Then ff = 1
                  
         numSer = numSer + 1
       End If
        sFiles = Dir()
    Loop
    
    ' если всё с числами - сортируем
    If ff = 0 Then SortFiles GrafF1(), GrafF(), GrafF2()
    
    For i = 0 To numSer - 1
        ' если сохранена разметка - добавляем плюсик
        If FileLen(s + GrafF1(i)) > 25000 Then
           List1.AddItem "+ " + GrafF1(i)
        Else
           List1.AddItem "  " + GrafF1(i)
        End If
    Next
    
    FindFiles = numSer
End Function


Private Sub Text2_Change()
    Rects(iLastRect).lDist = Val(Text2.Text)
    HScroll1.value = CheckMax(Rects(iLastRect).lDist)
    grafik
    DrawRects
End Sub


Private Sub Text3_Change()
    Rects(iLastRect).rDist = Val(Text3.Text)
    HScroll2.value = CheckMax(Rects(iLastRect).rDist)
    grafik
    DrawRects
End Sub


Public Function CheckMax(i As Single) As Integer
    If i < 0 Then i = 0
    If i > HScroll1.max Then CheckMax = HScroll1.max Else CheckMax = i
End Function


' файл уже размечен?
Public Function CheckRazm() As Boolean
   Dim i As Integer
   CheckRazm = False
   For i = 1 To numPoint
      If MasDat3(i, 0, 0, 10) <> 0 Then
         CheckRazm = True
         Exit For
      End If
   Next
End Function


Public Function GetCRect(n As Integer) As Integer
   Dim i As Integer
   
   For i = 1 To NumRect
      If n = Rects(i).arrelem Then
         GetCRect = i
         Exit For
      End If
   Next
End Function


' разметка из файла
Public Sub GetRects()
  Dim i As Integer, j As Integer, k As Integer
  Dim x As Single, y As Single, h As Integer
  
   With Picture1
   NumRect = 0
   k = 0
   h = 6
   
   Erase Rects
   
   For i = 1 To numPoint
      If MasDat3(i, 0, 0, 10) = 1 Then
         'k = GetCRect(i)
         Rects(NumRect).arrelem = i
         Rects(NumRect).value = 3
         Rects(NumRect).eventType = sEvents(MasDat3(i, 0, 0, 9))
         
         If indMas1 <> 0 Then x = 500 + (.Width * 15 - 1000) * (i - 1) / indMas1
         If maxm(0) - minm(0) <> 0 Then
             y = -500 + .Height * 15 - (.Height * 15 - 1000) * (MasDat3(i, 0, 0, 0) - minm(0)) / (maxm(0) - minm(0))
         End If
         Rects(NumRect).Left = x - 15 * h
         Rects(NumRect).Right = x + 15 * h
         Rects(NumRect).Top = y - 15 * h
         Rects(NumRect).Bottom = y + 15 * h
         
         For j = i To 1 Step -1
           If MasDat3(j, 0, 0, 10) = 2 Then
               Rects(NumRect).lDist = i - j
               Exit For
           End If
         Next
         For j = i To numPoint
           If MasDat3(j, 0, 0, 10) = 3 Then
               Rects(NumRect).rDist = j - i
               Exit For
           End If
         Next
         NumRect = NumRect + 1
      End If
   Next
   End With
   iLastRect = 1
End Sub


Public Sub DrawEvent()
  Dim i As Integer, j As Integer
  Dim xp1 As Single, xp2 As Single, yp1 As Single, yp2 As Single
  Dim hsv As Integer, lv As Integer, hv As Integer
  Dim pw As Long, ph As Long, kx As Single, ky As Single
  
  hsv = HScroll3.value
  With Picture1
    lv = Rects(iLastRect).arrelem - Rects(iLastRect).lDist
    hv = Rects(iLastRect).arrelem + Rects(iLastRect).rDist
    pw = .Width * 15
    ph = .Height * 15
    kx = (pw - 1000) / indMas1
    If (maxm(0) - minm(0)) <> 0 Then ky = (ph - 1000) / (maxm(0) - minm(0))
        
    xp1 = 500 + kx * (lv - hsv)
    xp2 = 500 + kx * (hv - hsv)
    
    Picture1.Line (xp1, 500)-(xp1, .Height * 15 - 500), vbBlack
    Picture1.Line (xp2, 500)-(xp2, .Height * 15 - 500), vbBlack
    
   End With
End Sub


Public Sub SortFiles(a() As String, b() As String, l() As Long)    '
  Dim i As Integer, n As Integer, j1 As Integer, j As Integer, m As Long
  
  n = numSer - 1
  For i = 0 To n
      m = l(0)
      j1 = 0
      For j = 0 To n - i
         If l(j) > m Then
            m = l(j)
            j1 = j
         End If
      Next j
      Swap l(j1), l(n - i)
      Swap a(j1), a(n - i)
      Swap b(j1), b(n - i)
  Next i
  
End Sub




