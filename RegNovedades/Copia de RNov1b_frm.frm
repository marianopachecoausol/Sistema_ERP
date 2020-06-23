VERSION 5.00
Begin VB.Form RNov1b_frm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   19605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   19605
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   19170
      Top             =   750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2865
      Left            =   15000
      TabIndex        =   18
      Top             =   120
      Width           =   80
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   870
      Left            =   -1950
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   2130
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Novedades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   225
         TabIndex        =   37
         Top             =   450
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registro de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   150
         TabIndex        =   14
         Top             =   150
         Width           =   1485
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   350
      Left            =   18120
      TabIndex        =   19
      Top             =   0
      Width           =   1635
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   650
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   30
         Width           =   500
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   480
         TabIndex        =   20
         Top             =   120
         Width           =   2475
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   15120
      TabIndex        =   15
      Top             =   0
      Width           =   4500
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   400
         Width           =   75
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   400
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label10 
         BackColor       =   &H00808080&
         Caption         =   " Registro  de Novedades   -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   350
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   4035
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 90"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   8
      Left            =   12360
      TabIndex        =   44
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 89"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   7
      Left            =   10920
      TabIndex        =   43
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 88"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   6
      Left            =   9480
      TabIndex        =   42
      Top             =   2050
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   8
      Left            =   12360
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G090"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   7
      Left            =   10920
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G089"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   6
      Left            =   9480
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G088"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   12
      Left            =   12600
      MouseIcon       =   "RNov1b_frm.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":0152
      Stretch         =   -1  'True
      Tag             =   "M015"
      Top             =   60
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   12
      Left            =   12600
      TabIndex        =   41
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   11
      Left            =   11600
      TabIndex        =   39
      Top             =   730
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   11
      Left            =   11600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "M012"
      Top             =   60
      Width           =   950
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Amb. AU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   4
      Left            =   18525
      TabIndex        =   38
      Top             =   2840
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image MovExternos 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   4
      Left            =   18525
      Stretch         =   -1  'True
      Tag             =   "AMB1"
      Top             =   2380
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   45
      X2              =   14760
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   75
      X2              =   14760
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   10
      Left            =   10600
      MouseIcon       =   "RNov1b_frm.frx":427F
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":43D1
      Stretch         =   -1  'True
      Tag             =   "M011"
      Top             =   60
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   10
      Left            =   10600
      TabIndex        =   36
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   9
      Left            =   9600
      TabIndex        =   35
      Top             =   730
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   9
      Left            =   9600
      MouseIcon       =   "RNov1b_frm.frx":84FE
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":8650
      Stretch         =   -1  'True
      Tag             =   "M010"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   8
      Left            =   8600
      MouseIcon       =   "RNov1b_frm.frx":C77D
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":C8CF
      Stretch         =   -1  'True
      Tag             =   "M009"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   7
      Left            =   7600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "M008"
      Top             =   60
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 09"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   8
      Left            =   8600
      TabIndex        =   34
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   7
      Left            =   7600
      TabIndex        =   33
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Gend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   3
      Left            =   18600
      TabIndex        =   30
      Top             =   2120
      Width           =   600
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Policia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   17640
      TabIndex        =   29
      Top             =   2120
      Width           =   600
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Bomb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   16665
      TabIndex        =   28
      Top             =   2120
      Width           =   600
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ambu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   15720
      TabIndex        =   27
      Top             =   2120
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   5
      Left            =   16665
      TabIndex        =   26
      Top             =   2840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   4
      Left            =   15720
      TabIndex        =   25
      Top             =   2840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   3
      Left            =   18600
      TabIndex        =   24
      Top             =   1215
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   17640
      TabIndex        =   23
      Top             =   1215
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   16650
      TabIndex        =   22
      Top             =   1215
      Width           =   600
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "GP 01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   15705
      TabIndex        =   21
      Top             =   1215
      Width           =   600
   End
   Begin VB.Image MovExternos 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   3
      Left            =   18525
      Stretch         =   -1  'True
      Tag             =   "GEND"
      Top             =   1600
      Width           =   735
   End
   Begin VB.Image MovExternos 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   2
      Left            =   17565
      Stretch         =   -1  'True
      Tag             =   "POLI"
      Top             =   1600
      Width           =   735
   End
   Begin VB.Image MovExternos 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   1
      Left            =   16605
      Stretch         =   -1  'True
      Tag             =   "BOMB"
      Top             =   1600
      Width           =   735
   End
   Begin VB.Image MovExternos 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   15645
      Stretch         =   -1  'True
      Tag             =   "AMBU"
      Top             =   1600
      Width           =   735
   End
   Begin VB.Image GPesadax 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   5
      Left            =   16605
      Stretch         =   -1  'True
      Tag             =   "GP06"
      Top             =   2380
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image GPesadax 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   4
      Left            =   15645
      Stretch         =   -1  'True
      Tag             =   "GP05"
      Top             =   2380
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image GPesadax 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   3
      Left            =   18525
      Stretch         =   -1  'True
      Tag             =   "GP04"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image GPesadax 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   2
      Left            =   17565
      Stretch         =   -1  'True
      Tag             =   "GP03"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image GPesada 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   1
      Left            =   16605
      Stretch         =   -1  'True
      Tag             =   "GP02"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image GPesada 
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   15645
      Stretch         =   -1  'True
      Tag             =   "GP01"
      Top             =   720
      Width           =   735
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   0
      Left            =   600
      MouseIcon       =   "RNov1b_frm.frx":109FC
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":10B4E
      Stretch         =   -1  'True
      Tag             =   "M001"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   1
      Left            =   1600
      MouseIcon       =   "RNov1b_frm.frx":14C7B
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":14DCD
      Stretch         =   -1  'True
      Tag             =   "M002"
      Top             =   60
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   2
      Left            =   2600
      MouseIcon       =   "RNov1b_frm.frx":18EFA
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":1904C
      Stretch         =   -1  'True
      Tag             =   "M003"
      Top             =   60
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   3
      Left            =   3600
      MouseIcon       =   "RNov1b_frm.frx":1D179
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":1D2CB
      Stretch         =   -1  'True
      Tag             =   "M004"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   4
      Left            =   4600
      MouseIcon       =   "RNov1b_frm.frx":213F8
      MousePointer    =   99  'Custom
      Picture         =   "RNov1b_frm.frx":2154A
      Stretch         =   -1  'True
      Tag             =   "M005"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   5
      Left            =   5600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "M006"
      Top             =   60
      Width           =   950
   End
   Begin VB.Image Pat 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   6
      Left            =   6600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "M007"
      Top             =   60
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   1
      Left            =   1600
      TabIndex        =   11
      Top             =   730
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   2
      Left            =   2600
      TabIndex        =   10
      Top             =   730
      Visible         =   0   'False
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 04"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   4
      Left            =   4600
      TabIndex        =   8
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   5
      Left            =   5600
      TabIndex        =   7
      Top             =   730
      Width           =   950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "M 07"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   6
      Left            =   6600
      TabIndex        =   6
      Top             =   730
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   0
      Left            =   360
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G000"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2050
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   1
      Left            =   2040
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G082"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   2
      Left            =   3600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G083"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   3
      Left            =   5160
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G084"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   4
      Left            =   6600
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G086"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Image Grua 
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Index           =   5
      Left            =   8040
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Tag             =   "G087"
      Top             =   1400
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 82"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   4
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 83"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 84"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   3
      Left            =   5160
      TabIndex        =   2
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 86"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   4
      Left            =   6600
      TabIndex        =   1
      Top             =   2050
      Width           =   950
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "G 87"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   240
      Index           =   5
      Left            =   8040
      TabIndex        =   0
      Top             =   2050
      Width           =   950
   End
   Begin VB.Menu MnuPpal 
      Caption         =   "Patrullas"
      Visible         =   0   'False
      Begin VB.Menu Patr 
         Caption         =   "Liberar"
         Index           =   0
      End
      Begin VB.Menu Patr 
         Caption         =   "Novedad"
         Index           =   1
      End
      Begin VB.Menu Patr 
         Caption         =   "QTH"
         Index           =   2
      End
      Begin VB.Menu Patr 
         Caption         =   "Liberar Tarea"
         Index           =   3
      End
      Begin VB.Menu Patr 
         Caption         =   "Retomar"
         Index           =   4
      End
      Begin VB.Menu Patr 
         Caption         =   "Carga Combustible"
         Index           =   5
      End
      Begin VB.Menu Patr 
         Caption         =   "Fuera de Servicio"
         Index           =   6
      End
      Begin VB.Menu Patr 
         Caption         =   "Fin de Rutina"
         Index           =   7
      End
   End
   Begin VB.Menu MenuGruas 
      Caption         =   "Grúas"
      Visible         =   0   'False
      Begin VB.Menu MnuGrua 
         Caption         =   "Liberar"
         Index           =   0
      End
      Begin VB.Menu MnuGrua 
         Caption         =   "Novedad"
         Index           =   1
      End
      Begin VB.Menu MnuGrua 
         Caption         =   "QTH"
         Index           =   2
      End
      Begin VB.Menu MnuGrua 
         Caption         =   "Fuera de Servicio"
         Index           =   3
      End
   End
   Begin VB.Menu mnuPMovil 
      Caption         =   "Moviles"
      Visible         =   0   'False
      Begin VB.Menu mnuMovSub 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Arribo"
         Index           =   2
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Pedido"
         Index           =   3
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Cancelar"
         Index           =   4
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Operativo"
         Index           =   5
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Novedad"
         Index           =   6
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "Liberar"
         Index           =   7
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "QTH"
         Index           =   8
      End
      Begin VB.Menu mnuMovSub 
         Caption         =   "No arribó"
         Index           =   9
      End
   End
   Begin VB.Menu MenuAmbu 
      Caption         =   "Ambulancias"
      Visible         =   0   'False
      Begin VB.Menu MnuAmbu 
         Caption         =   "Arribo"
         Index           =   0
      End
      Begin VB.Menu MnuAmbu 
         Caption         =   "Pedido"
         Index           =   1
      End
      Begin VB.Menu MnuAmbu 
         Caption         =   "QTH"
         Index           =   2
      End
      Begin VB.Menu MnuAmbu 
         Caption         =   "Retomar"
         Index           =   3
      End
      Begin VB.Menu MnuAmbu 
         Caption         =   "Fuera de Servicio"
         Index           =   4
      End
   End
End
Attribute VB_Name = "RNov1b_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mRec As New ADODB.Recordset
Public mIndexMov As String
Dim mi As Integer
Dim xKm As String
Dim xDescr As String
Dim xCodAlfa As String
Dim xSent As String
Dim xCodRamal As String
Dim xCodReferencia As String
Dim xObjGrua As Object
Dim xObjLbl As Object
Dim xImgMov As String
Dim xClima As String
Dim mPc As String
Dim mResp As Boolean

Private Sub Form_Load()
   Me.Height = 2400
   Me.Width = 19695
   Me.Top = 0
   Me.Left = 0
   RNov1a_frm.Top = RNov1b_frm.Height + 20
   InitForm
   mPc = Mid(MDI.mPCname, 1, Len(MDI.mPCname) - 1)
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   MDI.ERP_VenMini_Click
   MDI.ERP_Vent.Visible = True
   MDI.ERP_Vacio.Visible = True
Else
   If MsgBox("¿Desea salir de Sistema?", vbYesNo, sMessage & " - Atención!!") = vbYes Then
      Unload RNov1a_frm
      Unload RNov1b_frm
      Unload RNov1d_frm
      MDI.mRNovFlag = True  'var para que era
      ShowMenu 1, True, False
   End If
End If
End Sub

Private Sub GPesada_Click(Index As Integer)
   fAddToList GPesada, Label7, Index
End Sub

Private Sub Grua_Click(Index As Integer)
   fAddToList Grua, Label4, Index
End Sub

Private Sub Pat_Click(Index As Integer)
   fAddToList Pat, Label3, Index
End Sub

Private Sub Grua_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   mIndexMov = Index
   Select Case Label4(Index).Tag
      Case "g1" 'En recorrida
          MnuGrua(0).Visible = False
      Case "g2"  'Trabajando
          MnuGrua(2).Visible = False
      Case "g3"  ' Carga de Combustible
          MnuGrua(1).Visible = False
          MnuGrua(2).Visible = False
      Case "g4"  'QTH
          MnuGrua(1).Visible = False
          MnuGrua(2).Visible = False
      Case "g5"   'Fuera de Servicio
          MnuGrua(1).Visible = False
          MnuGrua(2).Visible = False
          MnuGrua(3).Visible = False
      Case "g6"  'Trabajando
          MnuGrua(2).Visible = False
  End Select
  Set xObjGrua = Grua
  Set xObjLbl = Label4
  PopupMenu MenuGruas
End If
For mi = 0 To MnuGrua.UBound
   MnuGrua(mi).Visible = True
Next
End Sub

Private Sub GPesada_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      mIndexMov = Index
      Select Case Label7(Index).Tag
         Case "x1" 'En recorrida //libre
             MnuGrua(0).Visible = False
             MnuGrua(2).Visible = False
         Case "x2"  'Llamada de pedido
             MnuGrua(2).Visible = False
         Case "x3"  'Trabajando
             MnuGrua(2).Visible = False
         Case "x4"   'Fuera de Servicio
             MnuGrua(1).Visible = False
             MnuGrua(2).Visible = False
             MnuGrua(3).Visible = False
     End Select
     Set xObjGrua = GPesada
     Set xObjLbl = Label7
     PopupMenu MenuGruas
   End If
   For mi = 0 To MnuGrua.UBound
      MnuGrua(mi).Visible = True
   Next
End Sub

Private Sub Pat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     mIndexMov = Index
     Select Case Label3(Index).Tag
       Case "p1" 'en QTH
           sMenuPatrulla True, True, False, False, False, False, True, False
       Case "p2" 'Realizando una Tarea
           sMenuPatrulla False, True, False, True, True, False, True, False
       Case "p3" 'Fuera de Servicio
           sMenuPatrulla True, False, False, False, False, False, False, False
       Case "15" 'Carga de Combustible
           sMenuPatrulla True, False, False, False, False, False, True, False
       Case "p5" 'Trabajando
           sMenuPatrulla True, True, False, False, True, False, True, False
       Case "p6" 'en Recorrida
           sMenuPatrulla False, True, True, False, True, True, True, False
       Case "p7" 'en Rutina
           sMenuPatrulla False, True, False, False, True, False, True, True
       Case "p8" 'En camino
           sMenuPatrulla True, True, False, False, True, False, True, False
     End Select
     PopupMenu MnuPpal
   End If
   For mi = 0 To Patr.UBound
      Patr(mi).Visible = True
   Next
End Sub

'Private Sub AmbuGCO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim mI As Integer
'   If Button = 2 Then
'   '  mIndexMov = Index
'      'a1=
'      mnuMovSub(0).Caption = "---- " & AmbuGCO(Index).Tag & " ----"
'      Select Case Label10(Index).Tag
'         Case "a0" 'en QTH
'            sMenuMovExt False, True, False, False, True
'            mnuMovSub(6).Visible = True
'         Case "a1" 'en Recorrida
'           sMenuMovExt False, True, False, False, True
'      End Select
'      PopupMenu mnuPMovil
'   End If
'   For mI = 0 To Patr.UBound
'      Patr(mI).Visible = True
'   Next
'End Sub


Private Sub MovExternos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      mnuMovSub(6).Visible = False
      mnuMovSub(7).Visible = False
      mnuMovSub(8).Visible = False
      mnuMovSub(9).Visible = False
      mIndexMov = Index
      mnuMovSub(0).Caption = "---- " & MovExternos(Index).Tag & " ----"
      Select Case Label8(Index).Tag
         Case "7", "3" 'AMBU, BOMB Libre
             sMenuMovExt False, True, False, False, False
         Case "8", "4"   'AMBU, BOMB Ocupado
            If MovExternos(Index).Tag = "AMB1" Then
               sMenuMovExt True, False, True, False, True
            Else
               sMenuMovExt True, True, True, False, False
            End If
         Case "1", "5" 'POLI, GEND Libre
             sMenuMovExt False, True, False, True, False
         Case "2", "6" 'POLI, GEND Ocupado
             sMenuMovExt True, True, True, True, False
             mnuMovSub(9).Visible = True
             'sMenuMovExt True, false, True, True
         Case "a0" 'AMBU GCO QTH
            sMenuMovExt False, True, False, False, True
            mnuMovSub(6).Visible = True
            mnuMovSub(7).Visible = True 'liberar
         Case "a1" 'AMBU GCO
            sMenuMovExt False, True, False, False, True
            mnuMovSub(6).Visible = True 'novedad
            mnuMovSub(8).Visible = True 'qth
         Case "a2" 'AMBU GCO TRABAJANDO
            sMenuMovExt False, False, False, False, True
            mnuMovSub(6).Visible = True
            mnuMovSub(7).Visible = True 'liberar
     End Select
     PopupMenu mnuPMovil
   End If
   sMenuMovExt True, True, True, True, False
End Sub

'///////////////////////////////////////////////
'        ACCIONES DE LOS ITEMS DE LOS MENUES
'///////////////////////////////////////////////
Private Sub MnuGrua_Click(Index As Integer)
Dim mObj As New clRNov
Dim mOrigen As String
   xClima = ClimaOK(25.92)
   xCodAlfa = NVL(Mid(xObjGrua(mIndexMov).ToolTipText, 2, 7), "")
   Select Case Index
      Case 0  'liberar
         If xObjLbl(mIndexMov).Tag = "g5" Or xObjLbl(mIndexMov).Tag = "x4" Or xObjLbl(mIndexMov).Tag = "g4" Then
            If xObjLbl(mIndexMov).Tag = "g4" Then
               xDescr = "Móvil " & xObjGrua(mIndexMov).Tag & " Liberado."
               xImgMov = "g1"
            Else
               xDescr = "Móvil " & xObjGrua(mIndexMov).Tag & " Liberado de F/S."
               If Left(xObjGrua(mIndexMov).Tag, 2) = "GP" Then
                  xImgMov = "x1"
               Else
                  xImgMov = "g1"
               End If
            End If
            xObjGrua(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\iconos\" & xImgMov & ".gif")
            Set mRec = mObj.oTabla("novedades2", "where Mov1='" & xObjGrua(mIndexMov).Tag & "' and codigo = '' order by fecha desc limit 1")
            If Not mRec.EOF Then
               'xKm = Left(mRec!km, 2) & "." & Right(mRec!km, 2)  'mp20160523
               xKm = mRec!km
               xSent = mRec!sent
               xCodRamal = mRec!codramal
               xCodReferencia = mRec!codreferencia
            Else
               xKm = "0"
               xSent = "0"
               xCodRamal = "0"
               xCodReferencia = "0"
            End If
            mRec.Close
            xCodAlfa = ""
            mObj.xUpEstMoviles xObjGrua(mIndexMov).Tag, "L", "", xImgMov
            mObj.xInsNovedades "", xCodAlfa, Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", xDescr, "L", "0", xObjGrua(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia
            mObj.xUpActualizarNot mPc, 1
            xObjGrua(mIndexMov).ToolTipText = ""
            xObjLbl(mIndexMov).Tag = xImgMov
            Unload RNov1d_frm
            RNov1d_frm.Show
         Else
            RNov3_frm.Show
            RNov3_frm.sInitFreeGrua
            RNov3_frm.Frame1.Caption = "Liberar Móvil " & xObjGrua(mIndexMov).Tag
            RNov3_frm.Frame1.Tag = "LIB"
            RNov3_frm.Label2.Caption = xCodAlfa
            Set RNov3_frm.xObjGrua = xObjGrua
            'hay que mostrar la ventana con las opciones de tareas realizadas por la grúa
         End If
                
      Case 1 'Novedad
         RNov1c_frm.Show   'muestro form novedades
         If xCodAlfa <> "" Then
            RNov1c_frm.lCodAlfa = xCodAlfa
            RNov1c_frm.lCodAlfa.Visible = True
            
       
            xSent = Left(Right(xObjGrua(mIndexMov).ToolTipText, 7), 2)
            xCodRamal = Right(xObjGrua(mIndexMov).ToolTipText, 4)
            xCodReferencia = fGetCodigoReferencia(xCodAlfa)
            For mi = 0 To RNov1c_frm.Combo1(4).ListCount - 1
               If Left(Right(RNov1c_frm.Combo1(4).List(mi), 7), 4) = xCodRamal Then
                  RNov1c_frm.Combo1(4).ListIndex = mi
                  mi = 999
               End If
            Next
            For mi = 0 To RNov1c_frm.Combo1(0).ListCount - 1
               If Left(RNov1c_frm.Combo1(0).List(mi), 2) = xSent Then
                  RNov1c_frm.Combo1(0).ListIndex = mi
                  mi = 999
               End If
            Next
            For mi = 0 To RNov1c_frm.Combo1(5).ListCount - 1
               If Trim(Right(RNov1c_frm.Combo1(5).List(mi), 3)) = xCodReferencia Then
                  RNov1c_frm.Combo1(5).ListIndex = mi
                  mi = 999
               End If
            Next
         'RNov1c_frm.Text1(1).Text = Mid(Left(xObjGrua(mIndexMov).ToolTipText, InStr(1, xObjGrua(mIndexMov).ToolTipText, " ") - 1), 11) 'km
         RNov1c_frm.Text1(1).Text = fGetKm(xCodAlfa)
         End If
         
        
         
         RNov1c_frm.Label2 = xObjGrua(mIndexMov).Tag
         RNov1c_frm.Label2.Visible = True
         RNov1c_frm.Check1.Visible = False
         sObtOrigen RNov1c_frm.Label2.Caption, xCodAlfa, RNov1c_frm.Combo1(1)
        
      Case 2 'QTH
         RNov3_frm.Show
         RNov3_frm.Frame1.Caption = "Ingreso de QTH - Móvil " & xObjGrua(mIndexMov).Tag
         RNov3_frm.Frame1.Tag = "QTH"
         RNov3_frm.sInitQTH
        
      Case 3 'Fuera de Servicio
         If MsgBox("¿Seguro de Cambiar a Estado Fuera de Servicio el Móvil " & xObjGrua(mIndexMov).Tag & "?", vbYesNo, sMessage & " - Atención!!") = vbYes Then
            If Mid(xObjGrua(mIndexMov).ToolTipText, 11, 2) = "" Then
               xKm = "0"
               xSent = "0"
               xCodRamal = "0"
               xCodReferencia = "0" 'mp20160523
            Else
               'xKm = Mid(xObjGrua(mIndexMov).ToolTipText, 11, 2) '20160315
               xKm = Mid(Left(xObjGrua(mIndexMov).ToolTipText, InStr(1, xObjGrua(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
               xClima = ClimaOK(xKm)
               'xSent = Mid(xObjGrua(mIndexMov).ToolTipText, 17, 2) '20160315
               xSent = Left(Right(xObjGrua(mIndexMov).ToolTipText, 7), 2)
               'xCodRamal = Mid(xObjGrua(mIndexMov).ToolTipText, 20, 4) '20160315
               xCodRamal = Right(xObjGrua(mIndexMov).ToolTipText, 4)
               xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
               xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
               
               xCodReferencia = fGetCodigoReferencia(Mid(xObjGrua(mIndexMov).ToolTipText, 2, 7)) 'mp20160523
            End If
            mResp = mObj.xInsNovedades("", NVL(Mid(xObjGrua(mIndexMov).ToolTipText, 2, 7), ""), Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", "Movil " & xObjGrua(mIndexMov).Tag & " Fuera de Servicio", "RF", xClima, xObjGrua(mIndexMov).Tag, "", "0", "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia)
            mObj.xUpActualizarNot mPc, 1
            xObjGrua(mIndexMov).ToolTipText = ""
            If Left(xObjGrua(mIndexMov).Tag, 2) = "GP" Then
               xImgMov = "x4"
               Label7(mIndexMov).Tag = xImgMov
            Else
               xImgMov = "g5"
               Label4(mIndexMov).Tag = xImgMov
            End If
            xObjGrua(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\iconos\" & xImgMov & ".gif")
            mObj.xUpEstMoviles xObjGrua(mIndexMov).Tag, "L", "", xImgMov
            Unload RNov1d_frm
            RNov1d_frm.Show
         End If
        
      Case 4 'Cancelar
               
   End Select
   Set xObjGrua = Nothing
   Set mObj = Nothing
End Sub

Private Sub Patr_Click(Index As Integer)
Dim mObj As New clRNov
Dim mResp As Boolean

   xClima = ClimaOK(25.92)
   xCodAlfa = NVL(Mid(Pat(mIndexMov).ToolTipText, 2, 7), "")
   Select Case Index
      Case 0  '"liberar movil"
         Pat(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\iconos\p6.gif")
         If Mid(Pat(mIndexMov).ToolTipText, 11, 2) = "" Then
            Set mRec = mObj.oTabla("novedades2", "WHERE Mov1='" & Pat(mIndexMov).Tag & "' and codigo = '' ORDER BY fecha DESC LIMIT 1")
            If Not mRec.EOF Then
              xKm = Format(mRec!km, "00.00")
              xSent = mRec!sent
              xCodRamal = mRec!codramal
              xCodReferencia = mRec!codreferencia 'mp20160523
            Else
              xKm = "0"
              xSent = "0"
              xCodRamal = "0"
              xCodReferencia = "0" 'mp20160523
            End If
            mRec.Close
         Else
            xKm = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
            xSent = Left(Right(Pat(mIndexMov).ToolTipText, 7), 2)
            xCodRamal = Right(Pat(mIndexMov).ToolTipText, 4)
            xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
            xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
            xClima = ClimaOK(xKm)
            xCodReferencia = fGetCodigoReferencia(Mid(Pat(mIndexMov).ToolTipText, 2, 7))
         End If
         If Label3(mIndexMov).Tag = "p3" Then 'Si sale de un F/Servicio
            xDescr = "Móvil " & Pat(mIndexMov).Tag & " Liberado de F/S."
         Else
            xDescr = "Móvil " & Pat(mIndexMov).Tag & " Liberado."
         End If
         mObj.xUpEstMoviles Pat(mIndexMov).Tag, "L", "", "p6"

         mResp = mObj.xInsNovedades("", xCodAlfa, Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", xDescr, "L", xClima, Pat(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia) 'mp20160523
         mObj.xUpActualizarNot mPc, 1
         Pat(mIndexMov).ToolTipText = ""
         Label3(mIndexMov).Tag = "p6"
         Unload RNov1d_frm
         RNov1d_frm.Show
          
         Set mRec = mObj.waze_liberar_movil(xCodAlfa)
         If Not mRec.EOF Then
            If mRec!estadowaze = 2 Then
               RNov12.Show
               RNov1a_frm.Enabled = False
               RNov1b_frm.Enabled = False
               RNov1d_frm.Enabled = False
            End If
         End If
          
          
      Case 1 'Novedad desde Móvil
         RNov1c_frm.Show   'muestro form novedades
         If Label3(mIndexMov).Tag = "p7" Or Label3(mIndexMov).Tag = "p2" Then ''ver cual va
            RNov1c_frm.sNovTareas
         End If
         If xCodAlfa <> "" Then
            RNov1c_frm.lCodAlfa = xCodAlfa
            RNov1c_frm.lCodAlfa.Visible = True
            'RNov1c_frm.Text1(1).Text = Mid(Pat(mIndexMov).ToolTipText, 11, 5) 'Km
            RNov1c_frm.Text1(1).Text = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
           
            For mi = 0 To RNov1c_frm.Combo1(4).ListCount - 1
               If Left(Right(RNov1c_frm.Combo1(4).List(mi), 7), 4) = Right(Pat(mIndexMov).ToolTipText, 4) Then
                  RNov1c_frm.Combo1(4).ListIndex = mi
                  mi = 999
               End If
            Next
            
            For mi = 0 To RNov1c_frm.Combo1(0).ListCount - 1
              'If Left(RNov1c_frm.Combo1(0).List(mI), 2) = Mid(Pat(mIndexMov).ToolTipText, 17, 2) Then
              If Left(RNov1c_frm.Combo1(0).List(mi), 2) = Left(Right(Pat(mIndexMov).ToolTipText, 7), 2) Then
                 RNov1c_frm.Combo1(0).ListIndex = mi
                 mi = 999
              End If
            Next
            
            For mi = 0 To RNov1c_frm.Combo1(5).ListCount - 1
               If Trim(Right(RNov1c_frm.Combo1(5).List(mi), 3)) = fGetCodigoReferencia(xCodAlfa) Then
                  RNov1c_frm.Combo1(5).ListIndex = mi
               End If
            Next
            
            RNov1c_frm.Text1(1).Text = fGetKm(xCodAlfa)
            
            
         End If
         RNov1c_frm.Label2 = Pat(mIndexMov).Tag
         RNov1c_frm.Label2.Visible = True
         RNov1c_frm.Check1.Visible = False
         sObtOrigen RNov1c_frm.Label2.Caption, xCodAlfa, RNov1c_frm.Combo1(1)
     
      Case 2 'QTH
         RNov3_frm.Show
         RNov3_frm.Frame1.Caption = "Ingreso de QTH - Móvil " & Pat(mIndexMov).Tag
         RNov3_frm.sInitQTH
         RNov1a_frm.Enabled = False
         RNov1b_frm.Enabled = False
         RNov1d_frm.Enabled = False
          
      Case 3 ' "Liberar Tarea"
         MsgBox "Liberar Tarea"
         xCodAlfa = ""
         xKm = "0"
         xSent = "0"
         xCodRamal = "0"
         xCodReferencia = "0" 'mp20160523
         If Pat(mIndexMov).ToolTipText <> "" Then
            xCodAlfa = Mid(Pat(mIndexMov).ToolTipText, 2, 7)
            'xKm = Mid(Pat(mIndexMov).ToolTipText, 11, 5) 'mp 20160315
            'xSent = Mid(Pat(mIndexMov).ToolTipText, 17, 2) 'mp 20160315
            'xCodRamal = Mid(Pat(mIndexMov).ToolTipText, 20, 4) 'mp 20160315
            xKm = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
            xSent = Left(Right(Pat(mIndexMov).ToolTipText, 7), 2)
            xCodRamal = Right(Pat(mIndexMov).ToolTipText, 4)
            xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
            xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
            xCodReferencia = fGetCodigoReferencia(xCodAlfa) 'mp20160523
            xClima = ClimaOK(xKm)
         End If
         xDescr = "Móvil " & Pat(mIndexMov).Tag & " Liberado de Tarea."
         mResp = mObj.xInsNovedades("", xCodAlfa, Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", xDescr, "S", xClima, Pat(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia)
         mObj.xUpActualizarNot mPc, 1
         Set mRec = mObj.oTabla("novedades2", "WHERE CodNov in ('MM','A','R') AND Codigo='" & xCodAlfa & "'")
         If Not mRec.EOF Then
            mObj.xUpMovilesCodNov Pat(mIndexMov).Tag, "p5"
            Pat(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\p5.gif")
            Label3(mIndexMov).Tag = "p5"
         Else
            mObj.xUpEstMoviles Pat(mIndexMov).Tag, "L", "", "p6"
            Pat(mIndexMov).ToolTipText = ""
            Pat(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\p6.gif")
            Label3(mIndexMov).Tag = "p6"
         End If
         mRec.Close
         Unload RNov1d_frm
         RNov1d_frm.Show
     
      Case 4 'Retome de Móvil
         RNov1c_frm.Show
         xCodAlfa = ""
         If Trim(Pat(mIndexMov).ToolTipText) <> "" Then
            xCodAlfa = Mid(Pat(mIndexMov).ToolTipText, 2, 7)
            RNov1c_frm.lCodAlfa = xCodAlfa
            RNov1c_frm.lCodAlfa.Visible = True
         End If
         RNov1c_frm.Label2.Caption = Pat(mIndexMov).Tag
         RNov1c_frm.sRetomeMov
          
      Case 5 'Carga de Combustible
         RNov3_frm.Show
         RNov3_frm.Frame1.Caption = "Carga de Combustible - Móvil " & Pat(mIndexMov).Tag
         RNov3_frm.sInitCombustible
         
      Case 6 'Fuera de Servicio
         If MsgBox("¿Seguro de Cambiar a Estado Fuera de Servicio el Móvil " & Pat(mIndexMov).Tag & "?", vbYesNo, sMessage & " - Atención!!") = vbYes Then
            Pat(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\iconos\p3.gif")
            mObj.xUpEstMoviles Pat(mIndexMov).Tag, "L", "", "p3"
            If Mid(Pat(mIndexMov).ToolTipText, 11, 2) = "" Then
               xKm = "19.5"
               xSent = "1"
               xCodRamal = "1"
               xCodReferencia = "26" 'SEDE AUSOL - KM 19.5 'mp20160523
            Else
               'xSent = Mid(Pat(mIndexMov).ToolTipText, 17, 2)'mp 20160315
               'xCodRamal = Mid(Pat(mIndexMov).ToolTipText, 20, 4)'mp 20160315
               xSent = Left(Right(Pat(mIndexMov).ToolTipText, 7), 2)
               xCodRamal = Right(Pat(mIndexMov).ToolTipText, 4)
               xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
               xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
               'xKm = Mid(Pat(mIndexMov).ToolTipText, 11, 5) 'mp 20160315
               xKm = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
               xCodReferencia = fGetCodigoReferencia(Mid(Pat(mIndexMov).ToolTipText, 2, 7)) 'mp20160523
            End If
            mResp = mObj.xInsNovedades("", NVL(Mid(Pat(mIndexMov).ToolTipText, 2, 7), ""), Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", "Móvil Fuera de Servicio", "KO", xClima, Pat(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia) 'mp20160523
            mObj.xUpActualizarNot mPc, 1
            Pat(mIndexMov).ToolTipText = ""
            Label3(mIndexMov).Tag = "p3"
            Unload RNov1d_frm
            RNov1d_frm.Show
         End If
          
      Case 7  'Fin de Rutina
         Pat(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\iconos\p6.gif")
         Pat(mIndexMov).ToolTipText = ""
         Label3(mIndexMov).Tag = "p6"
         mObj.xUpEstMoviles Pat(mIndexMov).Tag, "L", "", "p6"
         xCodAlfa = ""
         xKm = "19.5"
         xSent = "1"
         xCodRamal = "1"
         xCodReferencia = "26" 'SEDE AUSOL - KM 19.5 'mp20160523
         If Pat(mIndexMov).ToolTipText <> "" Then
            xCodAlfa = Mid(Pat(mIndexMov).ToolTipText, 2, 7)
            'xKm = Mid(Pat(mIndexMov).ToolTipText, 11, 5)'mp 20160315
            'xSent = Mid(Pat(mIndexMov).ToolTipText, 17, 2)'mp 20160315
            'xCodRamal = Mid(Pat(mIndexMov).ToolTipText, 20, 4)'mp 20160315
            xKm = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
            xSent = Left(Right(Pat(mIndexMov).ToolTipText, 7), 2)
            xCodRamal = Right(Pat(mIndexMov).ToolTipText, 4)
            xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
            xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
            xCodReferencia = fGetCodigoReferencia(xCodAlfa) 'mp20160523
         End If
         xClima = ClimaOK(xKm)
         xDescr = "Móvil " & Pat(mIndexMov).Tag & " - Fin de Rutina."
         mResp = mObj.xInsNovedades("", xCodAlfa, Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", xDescr, "L", xClima, Pat(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia)
         mObj.xUpActualizarNot mPc, 1
         Unload RNov1d_frm
         RNov1d_frm.Show
   End Select
   Set mObj = Nothing
End Sub

Private Sub mnuMovSub_Click(Index As Integer)
Dim mObj As New clRNov
   Select Case Index
      Case 2 'Arribo
         RNov1c_frm.Show
         RNov1c_frm.sInitArriboMovExt (MovExternos(mIndexMov).Tag)
           
      Case 3 'Pedido de Móvil AMBU, GEND, BOMB o POLI
         RNov1c_frm.Show
         RNov1c_frm.Label2.Caption = MovExternos(mIndexMov).Tag
         RNov1c_frm.sInitMovExternos
         RNov1c_frm.Frame1.Caption = "Pedido de Móviles"
         RNov1c_frm.Check1.Visible = False
           
      Case 4 'Cancelar
         RNov1c_frm.Show
         RNov1c_frm.Label2.Caption = MovExternos(mIndexMov).Tag
         RNov1c_frm.sInitArriboMovExt (MovExternos(mIndexMov).Tag)
         RNov1c_frm.Frame1.Caption = "Cancelar Pedido de Móvil"
         RNov1c_frm.Label1(0).Visible = False
         RNov1c_frm.Text1(0).Visible = False
         RNov1c_frm.Check1.Visible = False
           
      Case 5 'Operativo
         RNov3_frm.Show
         RNov3_frm.sInitQTH
         RNov3_frm.Frame1.Tag = "OPER"
         RNov3_frm.Frame1.Caption = "Operativo Móvil " & MovExternos(mIndexMov).Tag
      
      Case 6 'Novedades
         RNov1c_frm.Show   'muestro form novedades
         
         If Len(MovExternos(mIndexMov).ToolTipText) > 10 Then
            xCodAlfa = Mid(MovExternos(mIndexMov).ToolTipText, 2, 7)
            RNov1c_frm.lCodAlfa = xCodAlfa
            RNov1c_frm.lCodAlfa.Visible = True
            'RNov1c_frm.Text1(1).Text = Mid(MovExternos(mIndexMov).ToolTipText, 11, 5) 'Km 'mp 20160315
            RNov1c_frm.Text1(1).Text = Mid(Left(MovExternos(mIndexMov).ToolTipText, InStr(1, Pat(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
            For mi = 0 To RNov1c_frm.Combo1(4).ListCount - 1
               'If Left(Right(RNov1c_frm.Combo1(4).List(mI), 7), 4) = Mid(MovExternos(mIndexMov).ToolTipText, 20, 4) Then 'mp 20160315
               If Left(Right(RNov1c_frm.Combo1(4).List(mi), 7), 4) = Right(MovExternos(mIndexMov).ToolTipText, 4) Then
                  RNov1c_frm.Combo1(4).ListIndex = mi
                  mi = 999
               End If
            Next
            
            For mi = 0 To RNov1c_frm.Combo1(0).ListCount - 1
              'If Trim(Left(RNov1c_frm.Combo1(0).List(mI), 2)) = Trim(Mid(MovExternos(mIndexMov).ToolTipText, 17, 2)) Then 'mp 20160315
              If Trim(Left(RNov1c_frm.Combo1(0).List(mi), 2)) = Left(Right(MovExternos(mIndexMov).ToolTipText, 7), 2) Then
                 RNov1c_frm.Combo1(0).ListIndex = mi
                 mi = 999
              End If
            Next
         End If
         RNov1c_frm.Label2 = MovExternos(4).Tag 'MovExternos(4).ToolTipText
         RNov1c_frm.Label2.Visible = True
         RNov1c_frm.Check1.Visible = False
'         sObtOrigen RNov1c_frm.Label2.Caption, xCodAlfa, RNov1c_frm.Combo1(1)

      Case 7 'Liberar AMB1 GCO
         xCodAlfa = ""
         xKm = "0"
         xSent = "0"
         xCodRamal = "0"
         xCodReferencia = "0" 'mp20160523
         If MovExternos(mIndexMov).ToolTipText <> "" Then
            xCodAlfa = Mid(MovExternos(mIndexMov).ToolTipText, 2, 7)
            'xKm = Mid(MovExternos(mIndexMov).ToolTipText, 11, 5) 'mp 20160315
            'xSent = Mid(MovExternos(mIndexMov).ToolTipText, 17, 2) 'mp 20160315
            'xCodRamal = Mid(MovExternos(mIndexMov).ToolTipText, 20, 4) 'mp 20160315
            xKm = Mid(Left(Pat(mIndexMov).ToolTipText, InStr(1, MovExternos(mIndexMov).ToolTipText, " ") - 1), 11) 'Km
            xSent = Left(Right(MovExternos(mIndexMov).ToolTipText, 7), 2)
            xCodRamal = Right(MovExternos(mIndexMov).ToolTipText, 4)
            xCodRamal = mObj.sTablaDescr("ramales", "abrevia='" & xCodRamal & "'", 0) 'obtengo código de tabla
            xSent = mObj.sTablaDescr("sentidos", "left(descripcion,2)='" & xSent & "' and codramal=" & xCodRamal, 0) 'obtengo código de tabla
            xCodReferencia = fGetCodigoReferencia(xCodAlfa) 'mp20160523
            xClima = ClimaOK(xKm)
         End If
         xDescr = MovExternos(mIndexMov).Tag & " Liberada."
         mResp = mObj.xInsNovedades("", xCodAlfa, Trim(Right(MDI.mUser, 15)), xKm, xSent, "SSV", xDescr, "S", xClima, MovExternos(mIndexMov).Tag, "", 0, "", "", 0, "", "", 0, "", "", xCodRamal, xCodReferencia)
         mObj.xUpActualizarNot mPc, 1
         mObj.xUpEstMoviles MovExternos(mIndexMov).Tag, "L", "", "a1"
         mObj.xUpMovilesCodNov MovExternos(mIndexMov).Tag, "a1"
         MovExternos(mIndexMov).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\a1.gif")
         Label8(mIndexMov).Tag = "a1"
         Unload RNov1d_frm
         RNov1d_frm.Show
         MovExternos(mIndexMov).ToolTipText = ""
         
      Case 8 'QTH
         RNov3_frm.Show
         RNov3_frm.Frame1.Caption = "AMBU QTH - " & MovExternos(4).Tag
         RNov3_frm.sInitQTH
         RNov1a_frm.Enabled = False
         RNov1b_frm.Enabled = False
         RNov1d_frm.Enabled = False
      
      Case 9 'No Arribó
         RNov1c_frm.Show
         RNov1c_frm.Label2.Caption = MovExternos(mIndexMov).Tag
         RNov1c_frm.sInitArriboMovExt (MovExternos(mIndexMov).Tag)
         RNov1c_frm.Frame1.Caption = "Móvil No Arribó"
         RNov1c_frm.Label1(0).Visible = False
         RNov1c_frm.Text1(0).Visible = False
         RNov1c_frm.Check1.Visible = False
      
      
   End Select
   Set mObj = Nothing
End Sub

Private Function fControlList(ByVal pMovil As String) As Boolean
   fControlList = True
   If RNov1a_frm.List1.ListCount > 0 Then
      For mi = 0 To (RNov1a_frm.List1.ListCount - 1)
        RNov1a_frm.List1.ListIndex = mi
        If RNov1a_frm.List1.Text = pMovil Then
           fControlList = False
           MsgBox "Móvil Ya Seleccionado", vbCritical, sMessage
        End If
      Next
   End If
   If RNov1a_frm.List1.ListCount = 3 And fControlList Then
      fControlList = False
      MsgBox "Lista de Móviles a Asignar llena", vbCritical, sMessage
   End If
End Function

Private Sub InitForm()
Dim mObj As New clRNov
   mAlinearObj Grua, 9, Frame3, 0     'gruas
   mAlinearObj Pat, 13, Frame3, 0     'patrullas
  ' mAlinearObj Pat, 10, Frame3, 6     'patrullas
   mAlinearObj Label4, 9, Frame3, 0   'descr gruas
   mAlinearObj Label3, 13, Frame3, 0  'descr patrullas
   'mAlinearObj Label3, 10, Frame3, 6  'descr patrullas
   Label5.Caption = Trim(Left(MDI.mUser, 50))
   Set xObjGrua = Nothing
   Set mRec = mObj.oTabla("moviles", "where CodTipoMov='PAT' order by 1")
   Do While Not mRec.EOF
      For mi = 0 To Pat.UBound
         If Pat(mi).Tag = mRec!Codigo Then
            If mRec.Fields(3) <> "L" Then
               Pat(mi).ToolTipText = mRec.Fields(4)
            End If
            Pat(mi).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & NVL(mRec.Fields(5), "p1") & ".gif")
            Label3(mi).Tag = NVL(mRec.Fields(5), "p1")
            mi = 99
         End If
      Next
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTabla("moviles", "where CodTipoMov='GRU' order by 1")
   Do While Not mRec.EOF
      For mi = 0 To Grua.UBound
        If Grua(mi).Tag = mRec!Codigo Then
           If mRec.Fields(3) <> "L" Then
              Grua(mi).ToolTipText = mRec.Fields(4)
           End If
           Grua(mi).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & NVL(mRec.Fields(5), "g1") & ".gif")
           Label4(mi).Tag = NVL(mRec.Fields(5), "g1")
           mi = 99
        End If
      Next
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTabla("moviles", "where CodTipoMov='GRP' order by 1")
   Do While Not mRec.EOF
      For mi = 0 To GPesada.UBound
        If GPesada(mi).Tag = mRec!Codigo Then
           If mRec.Fields(3) <> "L" Then
              GPesada(mi).ToolTipText = mRec.Fields(4)
           End If
           GPesada(mi).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & mRec.Fields(5) & ".gif")
           Label7(mi).Tag = NVL(mRec.Fields(5), "x1")
           mi = 99
        End If
      Next
      mRec.MoveNext
   Loop
   mRec.Close
   sInitMovExternos MovExternos(0).Tag, "7", 0 'AMBU
   sInitMovExternos MovExternos(1).Tag, "3", 1 'BOMB
   sInitMovExternos MovExternos(2).Tag, "5", 2 'POLI
   sInitMovExternos MovExternos(3).Tag, "1", 3 'GEND
   sInitMovExternos MovExternos(4).Tag, "1", 4 'AMB1
   
   'inicializar Ambu GCO
   Set mRec = mObj.oTabla("moviles", "where CodTipoMov='AMB' and codigo<>'AMBU' order by 1")
   If Not mRec.EOF Then
      If mRec.Fields(3) <> "L" Then
         MovExternos(4).ToolTipText = mRec.Fields(4)
      End If
      MovExternos(4).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & mRec.Fields(5) & ".gif")
      Label8(4).Tag = NVL(mRec.Fields(5), "x1")
   End If
   mRec.Close
   
   Set mObj = Nothing
End Sub

Private Function mAlinearObj(ByVal mObj As Object, ByVal mCantMovil As Integer, mFrame As Object, ByVal pStartMovil As Integer)
Dim mLeft As Integer
Dim mLeft2 As Integer
Dim mi As Integer
Dim mCantMovilVisibles As Integer
   
   mCantMovilVisibles = 0
   For mi = pStartMovil To mCantMovil - 1
      If mObj(mi).Visible Then
         mCantMovilVisibles = mCantMovilVisibles + 1
      End If
   Next

 
 
 
  ' mLeft = (mFrame.Left - (mCantMovil - pStartMovil) * mObj(pStartMovil).Width) / (mCantMovil - pStartMovil + 1)
   mLeft = (mFrame.Left - (mCantMovilVisibles - pStartMovil) * mObj(pStartMovil).Width) / (mCantMovilVisibles - pStartMovil + 1)
   mLeft2 = mLeft
   For mi = pStartMovil To mCantMovil - 1
      If mObj(mi).Visible Then
         mObj(mi).Left = mLeft2
         mLeft2 = mLeft + mObj(mi).Width + mLeft2
      End If
   Next
End Function

Private Function fAddToList(ByRef pObjMov As Object, ByRef pObjLab As Object, pIndex As Integer)
Dim pFlag As Boolean
   pFlag = False
   Select Case pObjMov(pIndex).Name
      Case "Grua"
          pFlag = (pObjLab(pIndex).Tag = "g1" Or pObjLab(pIndex).Tag = "g4")
      Case "Pat"
          pFlag = (pObjLab(pIndex).Tag = "p1" Or pObjLab(pIndex).Tag = "p6")
      Case "GPesada"
          pFlag = (pObjLab(pIndex).Tag = "x1")
   End Select
   If pFlag Then
      If fControlList(pObjMov(pIndex).Tag) Then
         RNov1a_frm.Label3.Visible = True
         RNov1a_frm.List1.Visible = True
         RNov1a_frm.List1.AddItem pObjMov(pIndex).Tag
      End If
   Else
      MsgBox "Móvil Ya Asignado con Evento", vbInformation, sMessage
   End If
End Function

Private Sub sMenuPatrulla(pItem1 As Boolean, pItem2 As Boolean, pItem3 As Boolean, _
                          pItem4 As Boolean, pItem5 As Boolean, pItem6 As Boolean, _
                          pItem7 As Boolean, pItem8 As Boolean)
Patr(0).Visible = pItem1
Patr(1).Visible = pItem2
Patr(2).Visible = pItem3
Patr(3).Visible = pItem4
Patr(4).Visible = pItem5
Patr(5).Visible = pItem6
Patr(6).Visible = pItem7
Patr(7).Visible = pItem8
End Sub

Private Sub sMenuMovExt(pItem1 As Boolean, pItem2 As Boolean, _
                       pItem3 As Boolean, pItem4 As Boolean, ByVal pItem5 As Boolean)
mnuMovSub(2).Visible = pItem1
mnuMovSub(3).Visible = pItem2
mnuMovSub(4).Visible = pItem3
mnuMovSub(5).Visible = pItem4
mnuMovSub(6).Visible = pItem5
End Sub

Private Sub sMenuAmbu(pItem1 As Boolean, pItem2 As Boolean, pItem3 As Boolean, _
                          pItem4 As Boolean, pItem5 As Boolean, pItem6 As Boolean, _
                          pItem7 As Boolean, pItem8 As Boolean)
Patr(0).Visible = pItem1
Patr(1).Visible = pItem2
Patr(2).Visible = pItem3
Patr(3).Visible = pItem4
Patr(4).Visible = pItem5
Patr(5).Visible = pItem6
Patr(6).Visible = pItem7
Patr(7).Visible = pItem8
End Sub

Private Sub sInitMovExternos(pCodMov As String, pImg As String, pIndex As Integer)
Dim mObj As New clRNov
Set mRec = mObj.oTabla("moviles", "where codigo='" & pCodMov & "'")
If Not mRec.EOF Then
   If mRec.Fields(3) <> "L" Then
      MovExternos(pIndex).ToolTipText = mRec.Fields(4)
   End If
   MovExternos(pIndex).Picture = LoadPicture(App.Path & "\RegNovedades\Image\Iconos\" & mRec.Fields(5) & ".gif")
   Label8(pIndex).Tag = NVL(mRec.Fields(5), pImg)
End If
mRec.Close
Set mObj = Nothing
End Sub

