VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RAcc1beta 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10275
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   14580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   14580
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   6
      Left            =   11880
      MaxLength       =   5
      TabIndex        =   191
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00CECECE&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   11700
      TabIndex        =   189
      Top             =   540
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00CECECE&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   1
      Left            =   11700
      TabIndex        =   188
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo9 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   12720
      Style           =   2  'Dropdown List
      TabIndex        =   187
      Top             =   540
      Width           =   1815
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   175
      Top             =   540
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E8E8E3&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Index           =   5
      Left            =   9540
      MaxLength       =   5
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   8280
      MaxLength       =   5
      TabIndex        =   6
      Top             =   540
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   8280
      MaxLength       =   5
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   7050
      MaxLength       =   5
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Width           =   2855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2340
      MaxLength       =   7
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   720
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CECECE&
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
      Height          =   9315
      Left            =   50
      TabIndex        =   93
      Top             =   900
      Width           =   14475
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C3"
         Height          =   195
         Index           =   2
         Left            =   1460
         TabIndex        =   10
         Top             =   660
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "B.Ext"
         Height          =   195
         Index           =   10
         Left            =   2780
         TabIndex        =   18
         Top             =   900
         Width           =   700
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "B.Int"
         Height          =   195
         Index           =   9
         Left            =   2120
         TabIndex        =   17
         Top             =   900
         Width           =   625
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C9"
         Height          =   195
         Index           =   8
         Left            =   1460
         TabIndex        =   16
         Top             =   900
         Width           =   500
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   4065
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   660
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C6"
         Height          =   195
         Index           =   5
         Left            =   3440
         TabIndex        =   13
         Top             =   660
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C5"
         Height          =   195
         Index           =   4
         Left            =   2780
         TabIndex        =   12
         Top             =   660
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C4"
         Height          =   195
         Index           =   3
         Left            =   2120
         TabIndex        =   11
         Top             =   660
         Width           =   500
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Index           =   2
         Left            =   9600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   184
         Top             =   5760
         Width           =   1515
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Index           =   1
         Left            =   3240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   183
         Top             =   5745
         Width           =   2655
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Index           =   0
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   182
         Top             =   5745
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   21
         Left            =   10140
         MaxLength       =   200
         TabIndex        =   180
         Top             =   3660
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Index           =   1
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   179
         Top             =   4560
         Width           =   3435
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Index           =   0
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   178
         Top             =   4200
         Width           =   3435
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   1275
         Index           =   2
         Left            =   300
         TabIndex        =   51
         Top             =   6060
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   2249
         _Version        =   327680
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   11780556
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   17
         Left            =   540
         MaxLength       =   200
         TabIndex        =   38
         Top             =   4920
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   12
         Left            =   6600
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Tag             =   "21"
         Top             =   3660
         Width           =   3375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   11
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Tag             =   "17"
         Top             =   3840
         Width           =   3435
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   13
         Left            =   300
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   7860
         Width           =   1995
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   20
         Left            =   6660
         MaxLength       =   40
         TabIndex        =   55
         Top             =   7860
         Width           =   1995
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   19
         Left            =   3540
         MaxLength       =   40
         TabIndex        =   54
         Top             =   7860
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   18
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   53
         Top             =   7860
         Width           =   975
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   15
         Left            =   6300
         MaxLength       =   40
         TabIndex        =   50
         Top             =   5760
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   12
         Left            =   13740
         MaxLength       =   3
         TabIndex        =   49
         Top             =   4980
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   13740
         MaxLength       =   3
         TabIndex        =   48
         Top             =   4620
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   12240
         MaxLength       =   3
         TabIndex        =   47
         Top             =   4980
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   12240
         MaxLength       =   3
         TabIndex        =   46
         Top             =   4620
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   10200
         MaxLength       =   3
         TabIndex        =   45
         Top             =   4980
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   10200
         MaxLength       =   3
         TabIndex        =   44
         Top             =   4620
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   43
         Top             =   4980
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   8520
         MaxLength       =   3
         TabIndex        =   42
         Top             =   4620
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   40
         Top             =   4980
         Width           =   555
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   41
         Top             =   4620
         Width           =   555
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   10
         Left            =   10740
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3000
         Width           =   3195
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   5700
         MaxLength       =   200
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   9
         Left            =   3180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Tag             =   "2"
         Top             =   3000
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   8
         Left            =   480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CECECE&
         Caption         =   "Obra / Mant."
         Height          =   195
         Index           =   2
         Left            =   8940
         TabIndex        =   30
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CECECE&
         Caption         =   "Horizontal"
         Height          =   195
         Index           =   1
         Left            =   9060
         TabIndex        =   29
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CECECE&
         Caption         =   "Vertical"
         Height          =   195
         Index           =   0
         Left            =   9240
         TabIndex        =   28
         Top             =   1680
         Width           =   915
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   10920
         MaxLength       =   200
         TabIndex        =   32
         Top             =   2040
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.TextBox Text3 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   4680
         MaxLength       =   200
         TabIndex        =   26
         Top             =   2100
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   7
         Left            =   10920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "1"
         Top             =   1680
         Width           =   1995
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   6
         Left            =   6480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1680
         Width           =   1995
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   5
         Left            =   4680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Tag             =   "0"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   4
         Left            =   2640
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1680
         Width           =   1635
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   3
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   9600
         MaxLength       =   250
         TabIndex        =   22
         Top             =   660
         Width           =   4695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   2
         Left            =   7620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   660
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   5685
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   660
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C2"
         Height          =   195
         Index           =   1
         Left            =   800
         TabIndex        =   9
         Top             =   660
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C8"
         Height          =   195
         Index           =   7
         Left            =   800
         TabIndex        =   15
         Top             =   900
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C7"
         Height          =   195
         Index           =   6
         Left            =   140
         TabIndex        =   14
         Top             =   900
         Width           =   500
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CECECE&
         Caption         =   "C1"
         Height          =   195
         Index           =   0
         Left            =   140
         TabIndex        =   8
         Top             =   660
         Width           =   500
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   1095
         Index           =   3
         Left            =   240
         TabIndex        =   56
         Top             =   8160
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   1931
         _Version        =   327680
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   14673886
      End
      Begin VB.Image Image3 
         Height          =   210
         Index           =   0
         Left            =   1320
         MouseIcon       =   "RAcc1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":0152
         Stretch         =   -1  'True
         ToolTipText     =   "Agregar Patrulleros"
         Top             =   5520
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   10140
         TabIndex        =   181
         Top             =   3420
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   4
         Left            =   10140
         MouseIcon       =   "RAcc1.frx":0395
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":04E7
         Stretch         =   -1  'True
         ToolTipText     =   "Borrar"
         Top             =   7620
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   1
         Left            =   12300
         MouseIcon       =   "RAcc1.frx":071F
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":0871
         Stretch         =   -1  'True
         ToolTipText     =   "Borrar"
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   5
         Left            =   10920
         MouseIcon       =   "RAcc1.frx":0AA9
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":0BFB
         Stretch         =   -1  'True
         ToolTipText     =   "Volver"
         Top             =   7620
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   3
         Left            =   9300
         MouseIcon       =   "RAcc1.frx":10F0
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":1242
         Stretch         =   -1  'True
         Tag             =   "G"
         ToolTipText     =   "Agregar"
         Top             =   7620
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   2
         Left            =   13080
         MouseIcon       =   "RAcc1.frx":1773
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":18C5
         Stretch         =   -1  'True
         ToolTipText     =   "Volver"
         Top             =   5520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   0
         Left            =   11580
         MouseIcon       =   "RAcc1.frx":1DBA
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":1F0C
         Stretch         =   -1  'True
         Tag             =   "G"
         ToolTipText     =   "Agregar"
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs"
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   169
         Top             =   4980
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Causas del Vehículo"
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
         Index           =   28
         Left            =   4680
         TabIndex        =   168
         Top             =   3720
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Causas Prob. Cond"
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
         Index           =   27
         Left            =   240
         TabIndex        =   167
         Top             =   3600
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   26
         Left            =   300
         TabIndex        =   166
         Top             =   7680
         Width           =   390
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         Index           =   7
         X1              =   2340
         X2              =   14700
         Y1              =   7510
         Y2              =   7510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   6
         X1              =   2340
         X2              =   14460
         Y1              =   7500
         Y2              =   7500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   5
         X1              =   3480
         X2              =   14460
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         Index           =   4
         X1              =   3480
         X2              =   14460
         Y1              =   5470
         Y2              =   5470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dependencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   36
         Left            =   6660
         TabIndex        =   140
         Top             =   7680
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal a cargo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   35
         Left            =   3540
         TabIndex        =   139
         Top             =   7680
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Móvil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   34
         Left            =   2400
         TabIndex        =   138
         Top             =   7680
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervención de Terceros"
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
         Index           =   25
         Left            =   120
         TabIndex        =   137
         Top             =   7380
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Móvil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   24
         Left            =   9600
         TabIndex        =   136
         Top             =   5580
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Polad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   23
         Left            =   6300
         TabIndex        =   135
         Top             =   5580
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patrullero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   22
         Left            =   3240
         TabIndex        =   134
         Top             =   5565
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patrullero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   360
         MouseIcon       =   "RAcc1.frx":243D
         MousePointer    =   99  'Custom
         TabIndex        =   133
         Top             =   5520
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intervención de Personal de Autopista"
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
         Index           =   20
         Left            =   120
         TabIndex        =   132
         Top             =   5280
         Width           =   3270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otros"
         Height          =   195
         Index           =   9
         Left            =   13320
         TabIndex        =   131
         Top             =   5040
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Poste SOS"
         Height          =   195
         Index           =   8
         Left            =   12900
         TabIndex        =   130
         Top             =   4680
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Señalam. Vert."
         Height          =   195
         Index           =   7
         Left            =   11160
         TabIndex        =   129
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Col. de  Iluminación"
         Height          =   195
         Index           =   6
         Left            =   10800
         TabIndex        =   128
         Top             =   4680
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peaje Barrera"
         Height          =   195
         Index           =   5
         Left            =   9180
         TabIndex        =   127
         Top             =   5040
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peaje Cabina"
         Height          =   195
         Index           =   4
         Left            =   9180
         TabIndex        =   126
         Top             =   4680
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Jersey Barrera"
         Height          =   195
         Index           =   3
         Left            =   6960
         TabIndex        =   125
         Top             =   5040
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Jersey Cabina"
         Height          =   195
         Index           =   2
         Left            =   7020
         TabIndex        =   124
         Top             =   4680
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guardarrail-Poste"
         Height          =   195
         Index           =   1
         Left            =   4920
         TabIndex        =   123
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guardarrail"
         Height          =   195
         Index           =   0
         Left            =   5340
         TabIndex        =   122
         Top             =   4680
         Width           =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   3
         X1              =   7200
         X2              =   14520
         Y1              =   4260
         Y2              =   4260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         Index           =   2
         X1              =   7200
         X2              =   14520
         Y1              =   4270
         Y2              =   4270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Daños a la autopista (cant.)"
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
         Index           =   19
         Left            =   4680
         TabIndex        =   121
         Top             =   4140
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colisión contra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   10740
         TabIndex        =   120
         Top             =   2760
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   5700
         TabIndex        =   119
         Top             =   2760
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   3180
         TabIndex        =   118
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Con otro vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   480
         TabIndex        =   117
         Top             =   2760
         Width           =   1545
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         Index           =   1
         X1              =   1740
         X2              =   14460
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         Index           =   0
         X1              =   1740
         X2              =   14460
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Accidente"
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
         Index           =   14
         Left            =   120
         TabIndex        =   116
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   13920
         MouseIcon       =   "RAcc1.frx":258F
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":26E1
         Stretch         =   -1  'True
         ToolTipText     =   "Siguiente"
         Top             =   8760
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   10440
         TabIndex        =   114
         Top             =   2100
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   4200
         TabIndex        =   113
         Top             =   2160
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inconvenientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   10920
         TabIndex        =   112
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   10
         Left            =   9960
         TabIndex        =   111
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Señalización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   56
         Left            =   8700
         TabIndex        =   110
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iluminación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   6480
         TabIndex        =   109
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   77
         Left            =   4680
         TabIndex        =   108
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pavimento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   2640
         TabIndex        =   107
         Top             =   1440
         Width           =   1605
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Calzada Secundaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   55
         Left            =   180
         TabIndex        =   106
         Top             =   1440
         Width           =   1710
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   9600
         TabIndex        =   105
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visibilidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   7620
         TabIndex        =   104
         Top             =   420
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Est. Rodadura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   5685
         TabIndex        =   103
         Top             =   420
         Width           =   1230
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configuración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4065
         TabIndex        =   102
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carriles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   101
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CECECE&
      Height          =   9315
      Left            =   50
      TabIndex        =   115
      Top             =   900
      Width           =   14475
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   5280
         TabIndex        =   176
         Text            =   "Combo6"
         Top             =   3840
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   12180
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   8700
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CECECE&
         Caption         =   "Datos de Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2055
         Left            =   60
         TabIndex        =   148
         Top             =   60
         Visible         =   0   'False
         Width           =   14355
         Begin VB.TextBox Text5 
            Height          =   315
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   73
            Top             =   960
            Width           =   2175
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CECECE&
            Caption         =   "RUTA"
            Height          =   435
            Index           =   3
            Left            =   12060
            TabIndex        =   82
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CECECE&
            Caption         =   "VTV"
            Height          =   435
            Index           =   2
            Left            =   10920
            TabIndex        =   81
            Top             =   1440
            Width           =   795
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CECECE&
            Caption         =   "Balizas Encendidas"
            Height          =   435
            Index           =   1
            Left            =   9300
            TabIndex        =   80
            Top             =   1440
            Width           =   1155
         End
         Begin VB.CheckBox Check6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CECECE&
            Caption         =   "Luces Encendidas"
            Height          =   435
            Index           =   0
            Left            =   7800
            TabIndex        =   79
            Top             =   1440
            Width           =   1155
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   6
            Left            =   6660
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   1500
            Width           =   615
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   5
            Left            =   4260
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1500
            Width           =   615
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   1
            Left            =   7860
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   420
            Width           =   1935
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   0
            Left            =   2580
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   420
            Width           =   2415
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   4
            Left            =   7500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   75
            Top             =   960
            Width           =   2355
         End
         Begin VB.ComboBox Combo4 
            ForeColor       =   &H80000002&
            Height          =   315
            Index           =   3
            Left            =   4080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   2
            Left            =   720
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   6
            Left            =   11040
            MaxLength       =   20
            TabIndex        =   76
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   5
            Left            =   5880
            MaxLength       =   8
            TabIndex        =   70
            Top             =   420
            Width           =   1335
         End
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00CECECE&
            Caption         =   "Titular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   300
            TabIndex        =   68
            Top             =   480
            Width           =   915
         End
         Begin VB.Image Image3 
            Height          =   210
            Index           =   2
            Left            =   9960
            MouseIcon       =   "RAcc1.frx":2C2B
            MousePointer    =   99  'Custom
            Picture         =   "RAcc1.frx":2D7D
            Stretch         =   -1  'True
            ToolTipText     =   "Agregar Cía. Seguros"
            Top             =   1020
            Width           =   210
         End
         Begin VB.Image Image3 
            Height          =   210
            Index           =   1
            Left            =   3000
            MouseIcon       =   "RAcc1.frx":2FC0
            MousePointer    =   99  'Custom
            Picture         =   "RAcc1.frx":3112
            Stretch         =   -1  'True
            ToolTipText     =   "Agregar Marcas"
            Top             =   1020
            Width           =   210
         End
         Begin VB.Image Image2 
            Height          =   495
            Index           =   11
            Left            =   13560
            MouseIcon       =   "RAcc1.frx":3355
            MousePointer    =   99  'Custom
            Picture         =   "RAcc1.frx":34A7
            Stretch         =   -1  'True
            ToolTipText     =   "Cancelar"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Image Image2 
            Height          =   495
            Index           =   10
            Left            =   13560
            MouseIcon       =   "RAcc1.frx":399C
            MousePointer    =   99  'Custom
            Picture         =   "RAcc1.frx":3AEE
            Stretch         =   -1  'True
            ToolTipText     =   "Borrar"
            Top             =   900
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Image Image2 
            Height          =   495
            Index           =   9
            Left            =   13560
            MouseIcon       =   "RAcc1.frx":3D26
            MousePointer    =   99  'Custom
            Picture         =   "RAcc1.frx":3E78
            Stretch         =   -1  'True
            Tag             =   "G"
            ToolTipText     =   "Agregar"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Est. de Neumáticos"
            Height          =   195
            Index           =   7
            Left            =   5220
            TabIndex        =   172
            Top             =   1560
            Width           =   1380
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado gral."
            Height          =   195
            Index           =   6
            Left            =   3240
            TabIndex        =   171
            Top             =   1560
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condiciones del Vehículo"
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
            Left            =   780
            TabIndex        =   170
            Top             =   1560
            Width           =   2190
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            Height          =   195
            Index           =   10
            Left            =   7440
            TabIndex        =   155
            Top             =   480
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Vehíc."
            Height          =   195
            Index           =   9
            Left            =   1560
            TabIndex        =   154
            Top             =   480
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Póliza"
            Height          =   195
            Index           =   5
            Left            =   10320
            TabIndex        =   153
            Top             =   1020
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cía. Seguro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   6540
            MouseIcon       =   "RAcc1.frx":43A9
            MousePointer    =   99  'Custom
            TabIndex        =   152
            Top             =   1020
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
            Height          =   195
            Index           =   3
            Left            =   3420
            TabIndex        =   151
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   180
            MouseIcon       =   "RAcc1.frx":44FB
            MousePointer    =   99  'Custom
            TabIndex        =   150
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dominio"
            Height          =   195
            Index           =   0
            Left            =   5220
            TabIndex        =   149
            Top             =   480
            Width           =   570
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cargar Vehículo..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00CECECE&
         Caption         =   "Informe Pericial"
         Height          =   195
         Index           =   4
         Left            =   6420
         TabIndex        =   92
         Top             =   8820
         Width           =   1515
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00CECECE&
         Caption         =   "ITS"
         Height          =   195
         Index           =   3
         Left            =   5460
         TabIndex        =   91
         Top             =   8820
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00CECECE&
         Caption         =   "Fotos papel"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   9200
         TabIndex        =   90
         Top             =   8820
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00CECECE&
         Caption         =   "Fotos digital"
         Height          =   195
         Index           =   1
         Left            =   3840
         TabIndex        =   89
         Top             =   8820
         Width           =   1275
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00CECECE&
         Caption         =   "Filmación"
         Height          =   195
         Index           =   0
         Left            =   2460
         TabIndex        =   88
         Top             =   8820
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   1125
         Index           =   10
         Left            =   7320
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Top             =   7500
         Width           =   7035
      End
      Begin VB.TextBox Text4 
         Height          =   1125
         Index           =   9
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   7500
         Width           =   7035
      End
      Begin VB.TextBox Text4 
         Height          =   1125
         Index           =   8
         Left            =   7320
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   6120
         Width           =   7095
      End
      Begin VB.TextBox Text4 
         Height          =   1125
         Index           =   7
         Left            =   120
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Top             =   6120
         Width           =   7095
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CECECE&
         Caption         =   "Cinturón o casco"
         Height          =   375
         Left            =   12900
         TabIndex        =   59
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   4
         Left            =   4920
         MaxLength       =   15
         TabIndex        =   63
         Top             =   1080
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   1635
         Index           =   0
         Left            =   120
         TabIndex        =   67
         Top             =   2160
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2884
         _Version        =   327680
         Cols            =   15
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   13420716
      End
      Begin VB.ComboBox Combo3 
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   2
         Left            =   10140
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   1080
         Width           =   2115
      End
      Begin VB.ComboBox Combo3 
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   1
         Left            =   7380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   1080
         Width           =   2115
      End
      Begin VB.ComboBox Combo3 
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   0
         Left            =   540
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   3
         Left            =   4020
         MaxLength       =   3
         TabIndex        =   62
         Top             =   1080
         Width           =   435
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   2
         Left            =   1860
         MaxLength       =   9
         TabIndex        =   61
         Top             =   1080
         Width           =   1515
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   6900
         MaxLength       =   50
         TabIndex        =   58
         Top             =   600
         Width           =   5595
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   57
         Top             =   600
         Width           =   4815
      End
      Begin MSFlexGridLib.MSFlexGrid Flex1 
         Height          =   1695
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   4140
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   2990
         _Version        =   327680
         Cols            =   12
         FixedCols       =   0
         BackColorFixed  =   11780556
      End
      Begin VB.ComboBox Combo3 
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   3
         Left            =   7380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   185
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   14265
         TabIndex        =   196
         Top             =   7290
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   7080
         TabIndex        =   195
         Top             =   7290
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   14265
         TabIndex        =   194
         Top             =   5865
         Width           =   45
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   7080
         TabIndex        =   193
         Top             =   5865
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medio Trasl."
         Height          =   195
         Index           =   10
         Left            =   6360
         TabIndex        =   186
         Top             =   1500
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   4380
         TabIndex        =   177
         Top             =   3900
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Image Image2 
         Height          =   315
         Index           =   13
         Left            =   3780
         MouseIcon       =   "RAcc1.frx":464D
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":479F
         Stretch         =   -1  'True
         ToolTipText     =   "Agregar Ocupante"
         Top             =   3810
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   8
         Left            =   13860
         MouseIcon       =   "RAcc1.frx":4D1B
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":4E6D
         Stretch         =   -1  'True
         ToolTipText     =   "Cancelar"
         Top             =   1320
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   7
         Left            =   13260
         MouseIcon       =   "RAcc1.frx":5362
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":54B4
         Stretch         =   -1  'True
         ToolTipText     =   "Borrar"
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   495
         Index           =   6
         Left            =   12600
         MouseIcon       =   "RAcc1.frx":56EC
         MousePointer    =   99  'Custom
         Picture         =   "RAcc1.frx":583E
         Stretch         =   -1  'True
         Tag             =   "G"
         ToolTipText     =   "Agregar"
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00B3C1CC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " D a t o s   d e   V í c t i m a s . . . "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   42
         Left            =   180
         TabIndex        =   173
         Top             =   3900
         Width           =   3075
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   350
         Index           =   9
         Left            =   4080
         TabIndex        =   165
         Tag             =   "65"
         Top             =   180
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3180
         TabIndex        =   164
         Top             =   300
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo"
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
         Index           =   15
         Left            =   1380
         TabIndex        =   163
         Top             =   8820
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones del patrullero."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   7320
         TabIndex        =   162
         Top             =   7260
         Width           =   2505
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mecánica del accidente."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   161
         Top             =   7260
         Width           =   2115
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios de los involucrados o testigos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   7320
         TabIndex        =   160
         Top             =   5880
         Width           =   3705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción al arribo al lugar."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   159
         Top             =   5880
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   540
         Index           =   1
         Left            =   240
         Picture         =   "RAcc1.frx":5D6F
         Stretch         =   -1  'True
         Top             =   8640
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00CCC8AC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " D a t o s   d e   V e h í c u l o s . . . "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   41
         Left            =   180
         TabIndex        =   158
         Top             =   1860
         Width           =   3255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel."
         Height          =   195
         Index           =   7
         Left            =   4620
         TabIndex        =   157
         Top             =   1140
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
         Height          =   195
         Index           =   4
         Left            =   3600
         TabIndex        =   156
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   195
         Index           =   6
         Left            =   9600
         TabIndex        =   147
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Traslado a"
         Height          =   195
         Index           =   5
         Left            =   6540
         TabIndex        =   146
         Top             =   1140
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N°"
         Height          =   195
         Index           =   3
         Left            =   1620
         TabIndex        =   145
         Top             =   1140
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   144
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         Height          =   195
         Index           =   1
         Left            =   6120
         TabIndex        =   143
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apell. y Nomb."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   142
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datos del"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   40
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ramal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4080
      TabIndex        =   197
      Top             =   240
      Width           =   540
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   12180
      MouseIcon       =   "RAcc1.frx":62CD
      MousePointer    =   99  'Custom
      Picture         =   "RAcc1.frx":641F
      Stretch         =   -1  'True
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registro de Accidentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   180
      TabIndex        =   192
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lib."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   11520
      TabIndex        =   190
      Top             =   120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   12
      Left            =   13980
      MouseIcon       =   "RAcc1.frx":6A12
      MousePointer    =   99  'Custom
      Picture         =   "RAcc1.frx":6B64
      Stretch         =   -1  'True
      ToolTipText     =   "Salir del sistema"
      Top             =   60
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   9120
      TabIndex        =   100
      Top             =   660
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Arr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   7800
      TabIndex        =   99
      Top             =   660
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Av."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7800
      TabIndex        =   98
      Top             =   240
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6600
      TabIndex        =   97
      Top             =   240
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sentido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4080
      TabIndex        =   96
      Top             =   660
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   95
      Top             =   660
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   94
      Top             =   660
      Width           =   540
   End
End
Attribute VB_Name = "RAcc1beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mi As Integer
Dim mj As Integer
Dim mColores(2)
Dim mTablasError As String

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
     Case 1
         If Combo1(1).ListIndex >= 0 Then
            sLlenoSentido
         End If

   End Select
End Sub

Private Sub Form_Load()
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'nada
End Sub

Private Sub Check7_Click(Index As Integer)
Dim mObj As New clRAcc
Dim mRec3 As New ADODB.Recordset
   Combo5.Clear
   If Check7(0).Value = 0 And Check7(1).Value = 0 Then
      Check7(0).Value = 1
   End If
   If Check7(0).Value = 1 And Check7(1).Value = 0 Then
      Set mRec3 = mObj.oTabla("Ficha", " where codtipoficha='01' order by 1")
   Else
      If Check7(0).Value = 0 And Check7(1).Value = 1 Then
         Set mRec3 = mObj.oTabla("Ficha", " where codtipoficha='02' order by 1")
      Else
         Set mRec3 = mObj.oTabla("Ficha", "order by 1")
      End If
   End If
   Do While Not mRec3.EOF
      Combo5.AddItem mRec3.Fields(0) & " - " & mRec3!Fecha
      mRec3.MoveNext
   Loop
   mRec3.Close
   Set mObj = Nothing
End Sub

Private Sub Combo2_Click(Index As Integer)
   Select Case Index
      Case 5, 7, 9, 11
         If Left(Combo2(Index).Text, 4) = "Otro" Or Left(Combo2(Index).Text, 5) = "Otros" Then
            Label3(Index).Visible = True
            Text3(Combo2(Index).Tag).Visible = True
            Text3(Combo2(Index).Tag).Text = ""
         Else
            Label3(Index).Visible = False
            Text3(Combo2(Index).Tag).Visible = False
         End If
      Case 12
         If Left(Combo2(Index).Text, 4) = "Otra" Then
            Label3(13).Visible = True
            Text3(Combo2(Index).Tag).Visible = True
            Text3(Combo2(Index).Tag).Text = ""
         Else
            Label3(13).Visible = False
            Text3(Combo2(Index).Tag).Visible = False
         End If
   End Select
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
   sEscape Combo2(Index), KeyAscii
End Sub

Private Sub Combo3_KeyPress(Index As Integer, KeyAscii As Integer)
   sEscape Combo3(Index), KeyAscii
End Sub

Private Sub Combo4_KeyPress(Index As Integer, KeyAscii As Integer)
   sEscape Combo4(Index), KeyAscii
End Sub

Private Sub Combo4_Click(Index As Integer)
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset
   Select Case Index
      Case 0
         If Combo4(Index).ListIndex > -1 Then
            Combo4(2).Clear
            Set mRec = mObj.oMarcasVehic(Right(Combo4(0).Text, 2))
            Do While Not mRec.EOF
               Combo4(2).AddItem mRec!descripcion & Space(50) & mRec!CodMarca
               mRec.MoveNext
            Loop
            mRec.Close
         End If
   End Select
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub Combo5_Click()
Dim mObj As New clRAcc
   If Combo5.ListIndex > -1 Then
      sBorrarTodo
      If mObj.oFichasViejas(Trim(Left(Combo5.Text, 5))) Then
         sDatosModifViejos Trim(Left(Combo5.Text, 5))
      Else
         sDatosModif Trim(Left(Combo5.Text, 5))
      End If
      Frame1.Enabled = True
   End If
   Set mObj = Nothing
End Sub

Private Sub Combo7_KeyPress(Index As Integer, KeyAscii As Integer)
   sEscape Combo7(Index), KeyAscii
End Sub

Private Sub Combo9_Click()
   sChgColorForm Right(Combo9.Text, 2)
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mObj As New clRAcc
Dim mFlagEst As Boolean
Dim mError2 As Boolean
Dim mTextos As String
Dim mi As Integer
Dim mDaniosLeg As Integer

   Select Case Index
      Case 0 'Cargar Vehículo
         If Command1(0).Caption <> "Fin Carga" Then
            Frame3.Visible = True
            Label5(9).Caption = Chr(Label5(9).Tag)
            Check4.Visible = False
            sFlexEnabled False
            Image2(11).Visible = True
         Else 'Fin de Carga
            mError2 = False
            If Trim(Text4(0).Text) <> "" Or Trim(Text4(1).Text) <> "" Or Trim(Text4(2).Text) <> "" Or Trim(Text4(4).Text) <> "" Or Combo3(0).ListIndex > -1 Or Combo3(1).ListIndex > -1 Or Combo3(2).ListIndex > -1 Then
               mError2 = True
               MsgBox "Datos de Ocupantes SIN CARGAR", vbExclamation, sMessage
            End If
            If Trim(Text4(5).Text) <> "" Or Trim(Text4(6).Text) <> "" Or Combo4(0).ListIndex > -1 Or Combo4(1).ListIndex > -1 Or Combo4(2).ListIndex > -1 Or Combo4(3).ListIndex > -1 Or Combo4(4).ListIndex > -1 Or Combo4(5).ListIndex > -1 Or Combo4(6).ListIndex > -1 Then
               mError2 = True
               MsgBox "Datos de Vehículos SIN CARGAR", vbExclamation, sMessage
            End If
            If Not mError2 Then
               Label3(40).Caption = "Datos del PEATON"
               Command1(0).Caption = "Cargar Vehículo..."
               Label5(8).Visible = False
               Label5(9).Visible = False
               Label5(9).Caption = ""
               sMoveIconos 1
               sFlexEnabled True
            End If
         End If
         
      Case 1
         If fValidaDatos Then
            MsgBox "Faltan agregar datos o existen errores..."
            Image1_Click 1
         Else
            mTextos = "Existió un Error en la carga de datos, IGUALMENTE SE CARGARON DATOS. " & Chr(13) & "Verifique la FICHA desde la opción Buscar. Enviar esto a sistemas: "
            If Command1(1).Caption = "Grabar" Then 'Grabar datos Nuevos
               sMsgEspere Me, "Grabando... espere un momento", True
               If mObj.bExistFicha(Text1(5).Text) Then
                  Text1(5).Text = mObj.sUltNroOrden
                  MsgBox "Se cambia el N° de Ficha por el " & Text1(5).Text & " por existir en la base.", vbInformation, sMessage
               End If
               If fGrabarDatos(Text1(5).Text) Then
                  MsgBox mTextos & mTablasError, vbCritical, sMessage
               Else
'                  'Agrego lo de legales .NET
'                  Dim mObjLeg As New cLegales
'                  mDaniosLeg = 0
'                  For mI = 1 To 10
'                     If Trim(Text3(mI + 2).Text) <> "" Then
'                        mDaniosLeg = 1
'                     End If
'                  Next
'                  On Error Resume Next
'                  mObjLeg.xInLegales Text1(5).Text, Text1(0).Text, mDaniosLeg
'                  Set mObjLeg = Nothing
'
'                  'Fin .NET
                  Label5(9).Tag = 65
                  sBorrarTodo
               End If
               sMsgEspere Me, "", False
            Else
               sMsgEspere Me, "Actualizando la Ficha...", True
               sRollBack (Left(Combo5.Text, 5))
               If fGrabarDatos(Left(Combo5.Text, 5)) Then
                  MsgBox mTextos & mTablasError, vbCritical, sMessage
               Else
                  sBorrarTodo
               End If
            End If
            sMsgEspere Me, "", False
         End If
   End Select
   Set mObj = Nothing
End Sub

Private Sub Flex1_DblClick(Index As Integer)
   If Flex1(Index).Row > 0 And Flex1(Index).TextMatrix(1, 0) <> "" Then
      sFlexEnabled False
      Command1(1).Enabled = False
      Flex1(Index).Tag = Flex1(Index).Row
      Select Case Index
         Case 0 'Modificar Datos de Vehículos
            'sVerIconos 0, True
            Frame3.Visible = True
            Label5(8).Visible = True
            Label5(9).Visible = True
            Label5(9).Caption = Flex1(0).TextMatrix(Flex1(0).Row, 0)
            Check3.Value = Flex1(0).TextMatrix(Flex1(0).Row, 1)
            For mi = 0 To Combo4(0).ListCount - 1 'tipo vehiculo
               If Trim(Left(Combo4(0).List(mi), 30)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 2), 30)) Then
                  Combo4(0).ListIndex = mi
                  mi = Combo4(0).ListCount
               End If
            Next
            Text4(5).Text = Flex1(0).TextMatrix(Flex1(0).Row, 3) 'dominio
            For mi = 0 To Combo4(1).ListCount - 1 'color
               If Trim(Left(Combo4(1).List(mi), 25)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 4), 25)) Then
                  Combo4(1).ListIndex = mi
                  mi = Combo4(1).ListCount
               End If
            Next
            For mi = 0 To Combo4(2).ListCount - 1 ' marca
               If Trim(Left(Combo4(2).List(mi), 30)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 5), 30)) Then
                  Combo4(2).ListIndex = mi
               End If
            Next
            Text5.Text = Left(Flex1(0).TextMatrix(Flex1(0).Row, 6), 30)
            For mi = 0 To Combo4(4).ListCount - 1 'cia seguro
               If Trim(Left(Combo4(4).List(mi), 30)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 7), 30)) Then
                  Combo4(4).ListIndex = mi
                  mi = Combo4(4).ListCount
               End If
            Next
            Text4(6).Text = Trim(Flex1(0).TextMatrix(Flex1(0).Row, 8)) 'N° Póliza
            For mi = 0 To Combo4(5).ListCount - 1 'Estado Gral
               If Trim(Left(Combo4(5).List(mi), 1)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 9), 1)) Then
                  Combo4(5).ListIndex = mi
               End If
            Next
            For mi = 0 To Combo4(6).ListCount - 1 'Estado Neumático
               If Trim(Left(Combo4(6).List(mi), 1)) = Trim(Left(Flex1(0).TextMatrix(Flex1(0).Row, 9), 1)) Then
                  Combo4(6).ListIndex = mi
               End If
            Next
            For mi = 0 To 3
               Check6(mi).Value = Flex1(0).TextMatrix(Flex1(0).Row, mi + 11)
            Next
            sVerIconos 9, True
                              
         Case 1 'Modificar Datos de Victimas
            Label3(40).Caption = "Modificar " & Flex1(1).TextMatrix(Flex1(1).Row, 0) & " - " & Flex1(1).TextMatrix(Flex1(1).Row, 1)
            With Flex1(1)
               Text4(0).Text = Trim(.TextMatrix(.Tag, 2)) 'Nombre
               Text4(1).Text = Trim(.TextMatrix(.Tag, 3)) 'Domicilio
               For mi = 0 To Combo3(0).ListCount - 1 'Tipo Doc
                  If Trim(Left(Combo3(0).List(mi), 4)) = Trim(Left(.TextMatrix(.Tag, 4), 4)) Then
                     Combo3(0).ListIndex = mi
                  End If
               Next
               Text4(2).Text = Trim(.TextMatrix(.Tag, 5)) 'número
               Text4(3).Text = Trim(.TextMatrix(.Tag, 6)) 'edad
               Text4(4).Text = Trim(.TextMatrix(.Tag, 7)) 'telefono
               For mi = 0 To Combo3(1).ListCount - 1 'traslado
                  If Trim(Left(Combo3(1).List(mi), 25)) = Trim(Left(.TextMatrix(.Tag, 8), 25)) Then
                     Combo3(1).ListIndex = mi
                  End If
               Next
               For mi = 0 To Combo3(2).ListCount - 1 'estado
                  If Trim(Left(Combo3(2).List(mi), 2)) = Trim(Left(.TextMatrix(.Tag, 9), 2)) Then
                     Combo3(2).ListIndex = mi
                  End If
               Next
               If Trim(.TextMatrix(.Tag, 1)) <> "" Then
                  Check4.Value = Trim(.TextMatrix(.Tag, 10)) 'cinturon
               End If
               For mi = 0 To Combo3(3).ListCount - 1 'medio traslado
                  If Trim(Left(Combo3(3).List(mi), 25)) = Trim(Left(.TextMatrix(.Tag, 11), 25)) Then
                     Combo3(3).ListIndex = mi
                  End If
               Next
            End With
            sVerIconos 6, True
                        
         Case 2 'Intervención de personal GCO
            sVerIconos 0, True
            For mi = 0 To Combo8(0).ListCount - 1
               If Trim(Left(Combo8(0).List(mi), 40)) = Flex1(2).TextMatrix(Flex1(2).Row, 1) Then
                  Combo8(0).ListIndex = mi
               End If
            Next
            For mi = 0 To Combo8(1).ListCount - 1
               If Trim(Left(Combo8(1).List(mi), 40)) = Flex1(2).TextMatrix(Flex1(2).Row, 2) Then
                  Combo8(1).ListIndex = mi
               End If
            Next
            Text3(15).Text = Flex1(2).TextMatrix(Flex1(2).Row, 3)
            For mi = 0 To Combo8(2).ListCount - 1
               If Trim(Right(Combo8(2).List(mi), 4)) = Flex1(2).TextMatrix(Flex1(2).Row, 4) Then
                  Combo8(2).ListIndex = mi
               End If
            Next
            
                        
         Case 3 'Intervención de terceros
            sVerIconos 3, True
            For mi = 2 To 4
               Text3(mi + 16).Text = Flex1(3).TextMatrix(Flex1(3).Row, mi)
            Next
            For mi = 0 To Combo2(13).ListCount - 1
               If Right(Combo2(13).List(mi), 2) = Left(Flex1(3).TextMatrix(Flex1(3).Row, 1), 2) Then
                  Combo2(13).ListIndex = mi
               End If
            Next
      End Select
   End If
End Sub

Private Sub Image1_Click(Index As Integer)
Dim mObj As New clRAcc
   If Index = 0 Then
      If Command1(1).Caption = "Grabar" Then
         If Fecha_ok(Text1(0).Text) And Len(Trim(Text1(1).Text)) = 7 Then
            If mObj.bExistFichaCodAlfa(Text1(0).Text, Text1(1).Text) Then
               MsgBox "Existe un ficha para la fecha y código (" & Text1(0).Text & ", " & Text1(1).Text & ")." & Chr(13) _
                  & "Grabar esta ficha provocará datos duplicados con Números de Ficha distintos." & Chr(13) & "Se sugiere CANCELAR la carga de datos."
            End If
         End If
      End If
      For mi = 0 To -15000 Step -100
         Frame1.Left = mi
         Frame2.Left = Frame2.Left - 100
      Next
      Frame2.Left = 60
   Else
      For mi = -15000 To 0 Step 100
         Frame1.Left = mi
         Frame2.Left = Frame2.Left + 100
      Next
      Frame1.Left = 60
   End If
   Set mObj = Nothing
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image1(Index).BorderStyle = 0
End Sub

Private Sub Image2_Click(Index As Integer)
   Dim mColorVeh As String
   Dim mNroVehic As String
   
   Select Case Index
      'CARGA -- INTERVENCION PERSONAL AUTOPISTA
      Case 0 'flex1(2)
         If Image2(0).Tag = "G" Then
            If Combo8(0).ListIndex > -1 And Combo8(2).ListIndex > -1 Then
               Flex1(2).AddItem "R" & vbTab & Trim(Left(Combo8(0).Text, 30)) & vbTab & Trim(Left(Combo8(1).Text, 30)) & vbTab & Trim(Text3(15).Text) & vbTab & Right(Combo8(2).Text, 4)
               sAddFlex 2
               sPintaFlex Flex1(2), Flex1(2).Rows - 1, 5, mColores(Flex1(2).Rows Mod 2)
               sBorrarCampos 2
            End If
         Else 'Modificar
            If Combo8(0).ListIndex > -1 And Combo8(2).ListIndex > -1 Then
               Flex1(2).TextMatrix(Flex1(2).Tag, 1) = Trim(Left(Combo8(0).Text, 30))
               Flex1(2).TextMatrix(Flex1(2).Tag, 2) = Trim(Left(Combo8(1).Text, 30))
               Flex1(2).TextMatrix(Flex1(2).Tag, 3) = Trim(Text3(15).Text)
               Flex1(2).TextMatrix(Flex1(2).Tag, 4) = Trim(Right(Combo8(2).Text, 4))
               Flex1(2).Enabled = True
               sVerIconos 0, False
               sBorrarCampos 2
            End If
         End If
      Case 1 'borrar
         If MsgBox("¿Esta seguro de borrar el dato?", vbYesNo, sMessage) = vbYes Then
            If Flex1(2).Tag = 1 Then
               If Flex1(2).Rows > 2 Then
                  Flex1(2).RemoveItem 1
                  sFlexOrden 2
               Else
                  For mi = 0 To 4
                     Flex1(2).TextMatrix(1, mi) = ""
                  Next
               End If
            Else
               Flex1(2).RemoveItem Flex1(2).Tag
               sFlexOrden 2
            End If
            sBorrarCampos 2
            sVerIconos 0, False
            Command1(1).Enabled = True
         End If
         
      Case 2 'Cancelar Flex1(2)
         sBorrarCampos 2
         sVerIconos 0, False
         
      '--------------------------------------------------------------------------------
      'CARGA -- INTERVENCION TERCEROS
      Case 3 'flex1(3)
         If Image2(3).Tag = "G" Then
            If Combo2(13).ListIndex > -1 And Trim(Text3(19).Text) <> "" Then
               Flex1(3).AddItem "R" & vbTab & Right(Combo2(13).Text, 2) & "-" & Left(Combo2(13).Text, 25) & vbTab & Trim(Text3(18).Text) & vbTab & Trim(Text3(19).Text) & vbTab & Trim(Text3(20).Text)
               sAddFlex 3
               sPintaFlex Flex1(3), Flex1(3).Rows - 1, 5, mColores(Flex1(3).Rows Mod 2)
               sBorrarCampos 3
            End If
         Else 'Modificar
            If Combo2(13).ListIndex > -1 And Trim(Text3(19).Text) <> "" Then
               Flex1(3).TextMatrix(Flex1(3).Tag, 1) = Right(Combo2(13).Text, 2) & "-" & Left(Combo2(13).Text, 25)
               For mi = 2 To 4
                  Flex1(3).TextMatrix(Flex1(3).Tag, mi) = Trim(Text3(mi + 16).Text)
               Next
               Flex1(3).Enabled = True
               sVerIconos 3, False
               sBorrarCampos 3
            End If
         End If
      
      Case 4 'BORRAR
         If MsgBox("¿Esta seguro de borrar el dato?", vbYesNo, sMessage) = vbYes Then
            If Flex1(3).Tag = 1 Then
               If Flex1(3).Rows > 2 Then
                  Flex1(3).RemoveItem 1
                  sFlexOrden 3
               Else
                  For mi = 0 To 4
                     Flex1(3).TextMatrix(1, mi) = ""
                  Next
               End If
            Else
               Flex1(3).RemoveItem Flex1(3).Tag
               sFlexOrden 3
            End If
            sBorrarCampos 3
            sVerIconos 3, False
         End If
      
      Case 5 ' --- Interv.
         sBorrarCampos 3
         sVerIconos 3, False
         Command1(1).Enabled = True
         
      '--------------------------------------------------------------------------------
      'CARGA DE VICTIMAS INVOLUCRADAS
      'Flex1(1)
      Case 6 'GRABAR
         If Trim(Text4(0).Text) <> "" Then
            If Image2(6).Tag = "G" Then
               mNroVehic = Label5(9).Caption
               If Not Command1(0).Visible Then 'Grabo de conductor
                  Command1(0).Visible = True
                  Label3(40).Caption = Mid(Label3(40).Caption, 1, 6) & " de los OCUPANTES."
                  mNroVehic = mNroVehic & Space(15) & "C"
                  Command1(0).Enabled = True
                  Image2(13).Enabled = True
               End If
               Flex1(1).AddItem "d" & vbTab & mNroVehic & vbTab & Trim(Text4(0).Text) & vbTab & Trim(Text4(1).Text) & vbTab & Combo3(0).Text & vbTab & "" _
                     & Trim(Text4(2).Text) & vbTab & Trim(Text4(3).Text) & vbTab & Trim(Text4(4).Text) & vbTab & Combo3(1).Text & vbTab & Combo3(2).Text & vbTab & Check4.Value & vbTab & Combo3(3).Text
               sAddFlex 1
               sPintaFlex Flex1(1), Flex1(1).Rows - 1, 11, mColores(Flex1(1).Rows Mod 2)
            Else 'Grabar Modificación
               With Flex1(1)
                  .TextMatrix(.Tag, 2) = Trim(Text4(0).Text)
                  .TextMatrix(.Tag, 3) = Trim(Text4(1).Text)
                  .TextMatrix(.Tag, 4) = Combo3(0).Text
                  .TextMatrix(.Tag, 5) = Trim(Text4(2).Text)
                  .TextMatrix(.Tag, 6) = Trim(Text4(3).Text)
                  .TextMatrix(.Tag, 7) = Trim(Text4(4).Text)
                  .TextMatrix(.Tag, 8) = Combo3(1).Text
                  .TextMatrix(.Tag, 9) = Combo3(2).Text
                  .TextMatrix(.Tag, 10) = Check4.Value
                  .TextMatrix(.Tag, 11) = Combo3(3).Text
                  .Tag = ""
               End With
               Label3(40).Caption = "Datos del PEATON"
               sFlexEnabled True
               Command1(0).Visible = True
            End If
            sBorrarCampos 1
            sVerIconos 6, False
            Text4(0).BackColor = &HFFFFFF
            If Image2(13).Visible And Not Image2(13).Enabled Then
               Command1_Click 0
               Image2(13).Enabled = True
            End If
         Else
            MsgBox "Faltan cargar textos", vbInformation, sMessage
            Text4(0).BackColor = &H7282F1
         End If
         
      Case 7
         If Right(Flex1(1).TextMatrix(Flex1(1).Tag, 1), 1) <> "C" Then
            If MsgBox("Esta seguro de borrar este OCUPANTE?", vbYesNo, "Atención - " & sMessage) = vbYes Then
               If Flex1(1).Rows > 2 Then
                  Flex1(1).RemoveItem Flex1(1).Tag
               Else
                  For mi = 0 To Flex1(1).Cols - 1
                     Flex1(1).TextMatrix(1, mi) = ""
                  Next
               End If
               sFlexOrden 1
               sBorrarCampos 1
               sVerIconos 6, False
               sFlexEnabled True
            End If
         Else
            MsgBox "Este dato se borra al eliminar el vehículo..", vbInformation, sMessage
         End If
      
      Case 8
         If Flex1(1).Tag <> "" Then
            sVerIconos 6, False
            sFlexEnabled True
            Flex1(1).Tag = ""
         End If
         If Image2(13).Enabled = False And Image2(13).Visible Then
            Command1_Click 0
            Image2(13).Enabled = True
         End If
         sBorrarCampos 1
         Label3(40).Caption = "Datos de PEATÓN"
         'ver tema de caption del label "Datos de Ocupantes"
         
      '--------------------------------------------------------------------------------
      'CARGA -- DATOS DE VEHICULOS
      'FLEX1(0)
      Case 9 'GRABAR
         If Image2(9).Tag = "G" Then
            If Combo4(0).ListIndex > -1 And Combo4(5).ListIndex > -1 Then
                             '                               titular          -         t. vehic               -            dominio          -                color                    -              marca
               Flex1(0).AddItem Label5(9).Caption & vbTab & Check3.Value & vbTab & Combo4(0).Text & vbTab & Trim(Text4(5).Text) & vbTab & Combo4(1).Text & vbTab & Combo4(2).Text & "" _
                  & vbTab & Trim(Text5.Text) & vbTab & Combo4(4).Text & vbTab & Trim(Text4(6).Text) & vbTab & Combo4(5).Text & vbTab & Combo4(6).Text & "" _
                  & vbTab & Check6(0).Value & vbTab & Check6(1).Value & vbTab & Check6(2).Value & vbTab & Check6(3).Value
               
               sAddFlex 0
               sPintaFlex Flex1(0), Flex1(0).Rows - 1, 15, mColores(Flex1(0).Rows Mod 2)
               If Combo6.Visible Then
                  Combo6.AddItem Label5(9).Caption
                  Image2(13).Enabled = False
               End If
               sBorrarCampos 0
               sMoveIconos -1
               Command1(0).Caption = "Fin Carga"
               Command1(0).Visible = False
               Command1(1).Enabled = False
               Image2(8).Visible = False
               Check4.Visible = True
               Frame3.Visible = False
               Label3(40).Caption = "Datos del CONDUCTOR"
               Label5(8).Visible = True
               Label5(9).Caption = Chr(Label5(9).Tag)
               Label5(9).Tag = Label5(9).Tag + 1
               Label5(9).Visible = True
               Combo4(0).BackColor = &HFFFFFF
               Combo4(5).BackColor = &HFFFFFF
               
            Else
               MsgBox "Faltan Cargar Datos...", vbCritical, sMessage
               If Combo4(0).ListIndex = -1 Then
                  Combo4(0).BackColor = &H7282F1
               End If
               If Combo4(5).ListIndex = -1 Then
                  Combo4(5).BackColor = &H7282F1
               End If
            End If
         Else 'Modificar
           If Combo4(0).ListIndex > -1 And Combo4(5).ListIndex > -1 Then
               With Flex1(0)
                  .TextMatrix(.Tag, 1) = Check3.Value
                  .TextMatrix(.Tag, 2) = Combo4(0).Text
                  .TextMatrix(.Tag, 3) = Trim(Text4(5).Text)
                  .TextMatrix(.Tag, 4) = Combo4(1).Text
                  .TextMatrix(.Tag, 5) = Combo4(2).Text
                  '.TextMatrix(.Tag, 6) = Combo4(3).Text
                  .TextMatrix(.Tag, 6) = Text5.Text
                  .TextMatrix(.Tag, 7) = Combo4(4).Text
                  .TextMatrix(.Tag, 8) = Trim(Text4(6).Text)
                  .TextMatrix(.Tag, 9) = Combo4(5).Text
                  .TextMatrix(.Tag, 10) = Combo4(6).Text
                  For mi = 0 To 3
                     .TextMatrix(.Tag, mi + 11) = Check6(mi).Value
                  Next
               End With
               Label5(8).Visible = False
               Label5(9).Visible = False
               Label5(9).Caption = ""
               Frame3.Caption = "Datos de Vehículo"
               Check4.Visible = True
               Frame3.Visible = False
               sBorrarCampos 0
               sFlexEnabled True
               sVerIconos 9, False
            Else
               MsgBox "Faltan Cargar Datos...", vbCritical, sMessage
               If Combo4(0).ListIndex = -1 Then
                  Combo4(0).BackColor = &H7282F1
               End If
               If Combo4(5).ListIndex = -1 Then
                  Combo4(5).BackColor = &H7282F1
               End If
            End If
         End If
         
      Case 10 'BORRAR VEHICULO
         If MsgBox("Borrar este registro implica borrar todo lo cargado " & Chr(13) & "en personas involucradas para este vehículo", vbYesNo, "Atención - " & sMessage) = vbYes Then
            Flex1(0).Tag = Flex1(0).Row
            'borrar todos los ocupantes del vehículo...
            For mi = Flex1(1).Rows - 1 To 2 Step -1
               If Left(Flex1(1).TextMatrix(mi, 1), 1) = Flex1(0).TextMatrix(Flex1(0).Tag, 0) Then
                  Flex1(1).RemoveItem mi
               Else
                  If Trim(Flex1(1).TextMatrix(mi, 1)) <> "" Then
                     If Asc(Flex1(1).TextMatrix(mi, 1)) > Asc(Flex1(0).TextMatrix(Flex1(0).Tag, 0)) Then
                        Flex1(1).TextMatrix(mi, 1) = Chr(Asc(Flex1(1).TextMatrix(mi, 1)) - 1) & Space(15) & Right(Flex1(1).TextMatrix(mi, 1), 1)
                     End If
                  End If
               End If
            Next
            'primer fila??
            If Left(Flex1(1).TextMatrix(1, 1), 1) = Flex1(0).TextMatrix(Flex1(0).Tag, 0) Then
               If Flex1(1).Rows > 2 Then
                  Flex1(1).RemoveItem 1
               Else
                  For mi = 0 To Flex1(1).Cols - 1
                     Flex1(1).TextMatrix(1, mi) = ""
                  Next
               End If
            End If
            'borra vehiculo y calcula letras de nro vehic
            For mi = Flex1(0).Rows - 1 To Flex1(0).Tag Step -1
               Flex1(0).TextMatrix(mi, 0) = Chr(Asc(Flex1(0).TextMatrix(mi, 0)) - 1)
            Next
            If Flex1(0).Rows > 2 Then
               Flex1(0).RemoveItem Flex1(0).Tag
               sFlexOrden 1
            Else
               For mi = 0 To Flex1(0).Cols - 1
                  Flex1(0).TextMatrix(1, mi) = ""
               Next
            End If
            If Label5(9).Tag > 65 Then
               Label5(9).Tag = Label5(9).Tag - 1
            End If
            Label5(8).Visible = False
            Label5(9).Visible = False
            Label5(9).Caption = ""
            Frame3.Visible = False
            sBorrarCampos 0
            sFlexEnabled True
            sVerIconos 9, False
            'esto es para combo5 'letras de vehiculos'
            Combo6.Clear
            If Flex1(0).TextMatrix(1, 0) <> "" Then
               For mi = 1 To Flex1(0).Rows - 1
                  Combo6.AddItem Flex1(0).TextMatrix(mi, 0)
               Next
            End If
            
         End If
      
      Case 11 'Cancelar Carga Vehiculo
         If Image2(10).Visible Then 'Cancelar modif
            Frame3.Caption = "Datos de Vehículo"
         End If
         Frame3.Visible = False
         Label5(8).Visible = False
         Label5(9).Visible = False
         Label5(9).Caption = ""
         Check4.Visible = False
         sBorrarCampos 0
         sFlexEnabled True
         sVerIconos 9, False
         
      Case 12
         If MsgBox("Salir del Sistema?", vbYesNo, sMessage) = vbYes Then
            Unload Me
            ShowMenu 2, True, False
         End If
         
      Case 13  'Agregar Ocupantes a un Vehículo
         If Combo6.ListIndex > -1 And Flex1(0).Enabled And Flex1(1).Enabled Then
            Label5(8).Visible = True
            Label5(9).Visible = True
            Label5(9).Caption = Combo6.Text
            Label3(40).Caption = "Datos de Ocupante"
            sMoveIconos -1
            Command1(0).Caption = "Fin Carga"
            Flex1(0).Enabled = False
            Flex1(1).Enabled = False
            sVerIconos 6, False
            Image2(8).Visible = True
            Image2(13).Enabled = False
         End If
         
   End Select
   If Index <> 12 And Index <> 9 Then
      Command1(1).Enabled = True
   End If
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image2(Index).BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image2(Index).BorderStyle = 0
End Sub

Private Sub Image3_Click(Index As Integer)
   Select Case Index
      Case 0
         RNov2_frm.mTabla = "patrulleros"
         RNov2_frm.Label1 = "Tabla Patrulleros"
         RNov2_frm.mFromAccid = True
         
      Case 1
         RAcc5_frm.mFlagRAccd = True
         RAcc5_frm.Show
         
      Case 2
         RAcc4_frm.mFlagRAccd = True
         RAcc4_frm.mTabla = "CiaSeguros"
         RAcc4_frm.Label1.Caption = "Tabla de Cías de Seguros"
         RAcc4_frm.Show
   End Select
   Me.Enabled = False
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image3(Index).BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image3(Index).BorderStyle = 0
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image4.BorderStyle = 1
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image4.BorderStyle = 0
End Sub

Private Sub Image4_Click()
   RAcc8_frm.Show
End Sub

Private Sub Label3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 21 Then
      Label3(Index).BorderStyle = 1
   End If
End Sub

Private Sub Label3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 21 Then
      Label3(Index).BorderStyle = 0
   End If
End Sub

Private Sub Label3_Click(Index As Integer)
   If Index = 21 Then
      sPatrulleros
   End If
End Sub

Private Sub Label6_Click(Index As Integer)
   Select Case Index
      Case 2
         Combo4_Click 0
'      Case 4
'         Combo4(Index).Clear
'         sLlenoCombo "CiaSeguros", Combo4(4)     'Tipo Vehículo
   End Select
End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 4 Or Index = 2 Then
      Label6(Index).BorderStyle = 1
   End If
End Sub

Private Sub Label6_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Index = 4 Or Index = 2 Then
      Label6(Index).BorderStyle = 0
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
      
      Case 1
         KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
         
      Case 2
         KeyAscii = fKmsKeyPress(Text1(Index), KeyAscii)
         
      Case 3, 4, 6
         KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   End Select
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = fAlfaNumKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 17
         KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
               
      Case 13, 14, 15, 19
         KeyAscii = fAlfaKeyPress(KeyAscii)
         
      Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
         KeyAscii = fNumeroKeyPress(KeyAscii)
      
      Case 18, 20
         KeyAscii = fAlfaNumKeyPress(KeyAscii)
   End Select
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text4_Change(Index As Integer)
   Select Case Index
      Case 7, 8, 9, 10
         Label7(Index).Caption = Len(Text4(Index).Text) & " / 1000"
   End Select
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1, 5, 6
          KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
      Case 7, 8, 9, 10
         If KeyAscii <> 13 Then
            KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
         End If
      Case 0
         KeyAscii = fAlfaKeyPress(KeyAscii)
      Case 2, 3, 4
         KeyAscii = fNumeroKeyPress(KeyAscii)
   End Select
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub sInitForm()
Dim mObj As New clRAcc
Dim mObjRN As New clRNov
Dim mRec As New ADODB.Recordset
Dim mj As Integer
   
   Image4.Visible = False
   Combo9.AddItem "Accidentes" & Space(40) & "01"
   Combo9.AddItem "Incidentes" & Space(40) & "02"
   Combo9.ListIndex = 0
   
   Text1(5).Text = mObj.sUltNroOrden
   Frame1.Left = 60
   Frame2.Left = 15000
   Label3(40).Caption = "Datos de peatón."
   Check4.Visible = False
    
     
   Set mRec = mObjRN.oTabla("ramales", "")
   sLlenoCbo Combo1(1), mRec, 1, 0
     
   Set mRec = mObj.oTablaNotNull("configuracion", "")
   sLlenoCbo Combo2(0), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("EstCalzada", "")
   sLlenoCbo Combo2(1), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("visibilidad", "")
   sLlenoCbo Combo2(2), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("LugarAccid", "") 'Calzada Secundaria
   sLlenoCbo Combo2(3), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("pavimento", "") 'Tipo Pavimento
   sLlenoCbo Combo2(4), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("Clima", "")  'Clima
   sLlenoCbo Combo2(5), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("Iluminacion", "")  'Iluminación
   sLlenoCbo Combo2(6), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("inconvenientes", "") 'Inconvenientes
   sLlenoCbo Combo2(7), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("ConOtroVehic", "") 'Accidente con otro vehiculo
   sLlenoCbo Combo2(8), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("Otros", "") 'Accidente con Otros
   sLlenoCbo Combo2(9), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("ColisionContra", "") 'colision Contra
   sLlenoCbo Combo2(10), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("CausaConductor", "")  'Causas de conductor
   sLlenoCbo Combo2(11), mRec, 1, 0
   Set mRec = mObj.oTablaNotNull("CausaConductor", "")  'Causas de conductor
   sLlenoCbo Combo7(0), mRec, 1, 0
   Set mRec = mObj.oTablaNotNull("CausaConductor", "")  'Causas de conductor
   sLlenoCbo Combo7(1), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("CausaVehic", "")  'Causas Vehículo
   sLlenoCbo Combo2(12), mRec, 1, 0
      
   Set mRec = mObjRN.oTablaNull("patrulleros")
   Do While Not mRec.EOF
      Combo8(0).AddItem mRec!Nombre & Space(50) & mRec!Codigo
      Combo8(1).AddItem mRec!Nombre & Space(50) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObjRN.oMovilesGCO("PAT")
   Do While Not mRec.EOF
      Combo8(2).AddItem mRec!descripcion & Space(20) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTablaNotNull("movilext", "") 'Móviles Externos
   sLlenoCbo Combo2(13), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("TipoDocu", "") 'Móviles Externos
   sLlenoCbo Combo3(0), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("LugarTrasl", "") 'Móviles Externos
   sLlenoCbo Combo3(1), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("estadoocupa", "") 'Estados lesionados
   sLlenoCbo Combo3(2), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("MedioTrasl", "") 'Medios de Traslado
   sLlenoCbo Combo3(3), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("TipoVehiculo", "") 'Tipo Vehículo
   sLlenoCbo Combo4(0), mRec, 1, 0
   
   Set mRec = mObj.oTablaNotNull("CiaSeguros", "") 'Medios de Traslado
   sLlenoCbo Combo4(4), mRec, 1, 0
   
   For mi = 5 To 6
      Combo4(mi).AddItem "B"
      Combo4(mi).AddItem "R"
      Combo4(mi).AddItem "M"
   Next
   
   Set mRec = mObj.oTabla("colores", "") 'Medios de Traslado
   sLlenoCbo Combo4(1), mRec, 1, 0
      
   mColores(0) = &HE0E0E0
   mColores(1) = &H0
   'inicializo Flex
   
   With Flex1(0) 'DATOS DE VEHICULO
      .ColWidth(0) = 450      'Veh
      .ColWidth(1) = 450      'Titular
      .ColWidth(2) = 1400     'T.Vehic
      .ColWidth(3) = 1400     'Dominio
      For mi = 4 To 8
         .ColWidth(mi) = 1400
      Next
      .ColWidth(9) = 800     'Est. Gral.
      .ColWidth(10) = 800     'Est. Neumáticos
      .ColWidth(11) = 700     'Luces
      .ColWidth(12) = 700     'Balizas
      .ColWidth(13) = 700     'VTV
      .ColWidth(14) = 700     'Ruta
      .TextMatrix(0, 0) = "Veh"
      .TextMatrix(0, 1) = "Tit"
      .TextMatrix(0, 2) = "T.Vehic."
      .TextMatrix(0, 3) = "Dominio"
      .TextMatrix(0, 4) = "Color"
      .TextMatrix(0, 5) = "Marca"
      .TextMatrix(0, 6) = "Modelo"
      .TextMatrix(0, 7) = "Cia.Seg."
      .TextMatrix(0, 8) = "N° Pol."
      .TextMatrix(0, 9) = "Est.gral"
      .TextMatrix(0, 10) = "Est.Neum"
      .TextMatrix(0, 11) = "Luces"
      .TextMatrix(0, 12) = "Baliza"
      .TextMatrix(0, 13) = "VTV"
      .TextMatrix(0, 14) = "RUTA"
   End With
   With Flex1(1)   'DATOS DE VICTIMAS
      .ColWidth(0) = 450       'N°
      .ColWidth(1) = 500       'Veh
      .ColWidth(2) = 3500      'Apellido, Nombre
      .ColWidth(3) = 3500      'Domicilio
      .ColWidth(4) = 800       'T.Doc
      .ColWidth(5) = 1500      'Doc. Nro
      .ColWidth(6) = 900       'Edad
      .ColWidth(7) = 1500      'Tel
      .ColWidth(8) = 1800      'Trasl. a
      .ColWidth(9) = 900       'Estado
      .ColWidth(10) = 500      'Cint
      .ColWidth(11) = 1800     'Medio Trasl
      .TextMatrix(0, 0) = "N°"
      .TextMatrix(0, 1) = "Veh"
      .TextMatrix(0, 2) = "Apellido y Nombre"
      .TextMatrix(0, 3) = "Domicilio"
      .TextMatrix(0, 4) = "T.Doc"
      .TextMatrix(0, 5) = "Nro.Doc"
      .TextMatrix(0, 6) = "Edad"
      .TextMatrix(0, 7) = "Tel"
      .TextMatrix(0, 8) = "Trasl.a"
      .TextMatrix(0, 9) = "Est."
      .TextMatrix(0, 10) = "Cint"
      .TextMatrix(0, 11) = "Medio Trasl"
   End With
   With Flex1(2)
      .ColWidth(0) = 450       'N°
      .ColWidth(1) = 3500      'Patrullero 1
      .ColWidth(2) = 3500      'Patrullero 2
      .ColWidth(3) = 3500      'Polad
      .ColWidth(4) = 1800      'Móvil
      .TextMatrix(0, 0) = "N°"
      .TextMatrix(0, 1) = "Patrullero 1"
      .TextMatrix(0, 2) = "Patrullero 2"
      .TextMatrix(0, 3) = "Polad"
      .TextMatrix(0, 4) = "Móvil"
   End With
   With Flex1(3)
      .ColWidth(0) = 450       'N°
      .ColWidth(1) = 2200      'Tipo Vehic
      .ColWidth(2) = 1800      'Móvil
      .ColWidth(3) = 3500      'personal
      .ColWidth(4) = 2600      'Dependencia
      .TextMatrix(0, 0) = "N°"
      .TextMatrix(0, 1) = "Tipo Vehic."
      .TextMatrix(0, 2) = "Móvil"
      .TextMatrix(0, 3) = "Personal"
      .TextMatrix(0, 4) = "Dependencia"
   End With
   For mj = 0 To 3
      Flex1(mj).Row = 0
      For mi = 0 To Flex1(mj).Cols - 1
         Flex1(mj).Col = mi
         Flex1(mj).CellFontBold = True
      Next
   Next
   
   Set mObj = Nothing
   Set mObjRN = Nothing
   Set mRec = Nothing
End Sub

Private Sub sEscape(ByRef pObj As Object, ByVal pKey As Integer)
   If pKey = 27 Then
      pObj.ListIndex = -1
   End If
End Sub

Private Sub sAddFlex(ByVal pIndex As Integer)
   If Flex1(pIndex).Rows > 1 And Flex1(pIndex).TextMatrix(1, 0) = "" Then
      Flex1(pIndex).RemoveItem 1
   End If
   If pIndex <> 0 Then
      For mi = 1 To Flex1(pIndex).Rows - 1
         Flex1(pIndex).TextMatrix(mi, 0) = mi
      Next
   End If
   Flex1(pIndex).ColAlignment(0) = 2
   Flex1(pIndex).ColAlignment(1) = 0
   Flex1(pIndex).ColAlignment(2) = 0
End Sub

Private Sub sBorrarCampos(ByVal pIndex As Integer)   'pIndex = indice del FlexGrid
   Select Case pIndex
      Case 0
         Check3.Value = 0
         Text4(5).Text = ""
         Text4(6).Text = ""
         Text5.Text = ""
         For mi = 0 To 6
            Combo4(mi).ListIndex = -1
         Next
         For mi = 0 To 3
            Check6(mi).Value = 0
         Next
      Case 1
         For mi = 0 To 4
            Text4(mi).Text = ""
         Next
         For mi = 0 To Combo3.UBound
            Combo3(mi).ListIndex = -1
         Next
         Check4.Value = 0
      Case 2
         Combo8(0).ListIndex = -1
         Combo8(1).ListIndex = -1
         Combo8(2).ListIndex = -1
         Text3(15).Text = ""
      Case 3
         Combo2(13).ListIndex = -1
         Text3(18).Text = ""
         Text3(19).Text = ""
         Text3(20).Text = ""
   End Select
   Flex1(pIndex).Tag = ""
End Sub

Private Sub sVerIconos(ByVal pIndex As Integer, ByVal pFlag As Boolean) 'pIndex= nro de indice que cominezan los iconos
   If pFlag Then
      Image2(pIndex).Picture = LoadPicture(App.Path & "\ERP\Imagenes\reload.gif")
      Image2(pIndex).Tag = "M"
      Image2(pIndex).ToolTipText = "Actualizar"
   Else
      Image2(pIndex).Tag = "G"
      Image2(pIndex).ToolTipText = "Agregar"
      Image2(pIndex).Picture = LoadPicture(App.Path & "\ERP\Imagenes\filesaveas.gif")
   End If
   Image2(pIndex + 1).Visible = pFlag
   Image2(pIndex + 2).Visible = pFlag
End Sub

Private Sub sFlexOrden(ByVal pIndex As Integer)
   For mi = 1 To Flex1(pIndex).Rows - 1
      Flex1(pIndex).TextMatrix(mi, 0) = mi
      sPintaFlex Flex1(pIndex), Flex1(pIndex).Rows - 1, 5, mColores(Flex1(pIndex).Rows Mod 2)
   Next
End Sub

Private Sub sFlexEnabled(ByVal pFlag As Boolean)
   Flex1(0).Enabled = pFlag
   Flex1(1).Enabled = pFlag
   Command1(1).Enabled = pFlag
End Sub

Private Sub sMoveIconos(ByVal pValor As Integer)
   For mi = 6 To 8
      Image2(mi).Left = Image2(mi).Left + pValor * 8500
      Image2(mi).Top = Image2(mi).Top - pValor * 240
   Next
   Command1(0).Left = Command1(0).Left - 8000 * pValor
   
End Sub

Private Function fValidaDatos() As Boolean
Dim mObjRN As New clRNov
Dim mError As Boolean
   
   If Fecha_ok(Text1(0).Text) Then
      mError = False
      Text1(0).BackColor = &HFFFFFF
   Else
      mError = True
      Text1(0).BackColor = &H7282F1
   End If
   mError = mError Or fVerError(Text1(1), 1)
   For mi = 0 To Text1.UBound - 1
      mError = mError Or fVerError(Text1(mi), 1)
   Next
   If Not mError Then
      If Not mObjRN.bExistDatoTabla("novedades", "codigo='" & Text1(1).Text & "' and fecha between '" & Format(DateAdd("d", -5, Text1(0).Text), "yyyy-mm-dd") & " 00:00:00' and '" & Format(DateAdd("d", 5, Text1(0).Text), "yyyy-mm-dd") & " 23:59:59'") Then
         MsgBox "El código alfanumérico ingresado no existe en el Sistema de Registro de Novedades.", vbInformation, sMessage
         mError = True
          Text1(1).BackColor = &H7282F1
      Else
          Text1(1).BackColor = &HFFFFFF
      End If
   End If
   'mError = mError And fVerError(Combo1, 2)
   mError = mError Or fVerError(Combo1(1), 2) Or fVerError(Combo1(0), 2)
   If Progresiva_Ok(Text1(2).Text, Trim(Right(Combo1(0).Text, 2))) Then
      If Len(Trim(Text1(2).Text)) <= 5 Then
'         If Leer_Dato(CurrentUser, "sDecimal") = "," Then
'            Text1(2).Text = Replace(Text1(2).Text, ".", ",")
'            Text1(2).Text = Format(Text1(2).Text, "00,00")
'            Text1(2).Text = Replace(Text1(2).Text, ",", ".")
'         Else
'            Text1(2).Text = Format(Text1(2).Text, "00.00")
'         End If
         
      End If
      Text1(2).BackColor = &HFFFFFF
   Else
      mError = mError Or Not mError
      Text1(2).BackColor = &H7282F1
   End If
   If Hora_ok(Text1(3).Text) Then
      Text1(3).BackColor = &HFFFFFF
   Else
      mError = mError Or Not mError
      Text1(3).BackColor = &H7282F1
   End If
   If Hora_ok(Text1(4).Text) Then
      Text1(4).BackColor = &HFFFFFF
   Else
      mError = mError Or Not mError
      Text1(4).BackColor = &H7282F1
   End If
   'mError = mError Or (fVerError(Check1, 3) And fVerError(Combo2(3), 2))
   
   If Trim(Text1(6).Text) <> "" Then
      If Hora_ok(Text1(6).Text) Then
         Text1(6).BackColor = &HFFFFFF
      Else
         mError = mError Or Not mError
         Text1(6).BackColor = &H7282F1
      End If
   End If
   If (fVerError(Check1, 3) And fVerError(Combo2(3), 2)) Then
      mError = True
   Else
      For mi = 0 To Check1.UBound
         Check1(mi).BackColor = &HFFFFFF
      Next
      Combo2(3).BackColor = &HFFFFFF
   End If
   mError = mError Or fVerError(Check2, 3)
   For mi = 0 To 2
      mError = mError Or fVerError(Combo2(mi), 2)
   Next
   For mi = 4 To 7
      mError = mError Or fVerError(Combo2(mi), 2)
   Next
   If (fVerError(Combo2(8), 2) And fVerError(Combo2(9), 2) And fVerError(Combo2(10), 2)) Then
      mError = True
   Else
      'mError = mError And False
      For mi = 8 To 10
         Combo2(mi).BackColor = &HFFFFFF
      Next
   End If
   mError = mError Or fVerError(Flex1(2), 4)
   'datos que falten agregar a Flex1(...)
   If Combo8(0).ListIndex > -1 Or Combo8(1).ListIndex > -1 Or Trim(Text3(15).Text) <> "" Or Combo8(2).ListIndex > -1 Then
      mError = True
      MsgBox "Datos de Intervención Pers. Autopista SIN CARGAR", vbExclamation, sMessage
   End If
   If Combo2(13).ListIndex > -1 Or Trim(Text3(18).Text) <> "" Or Trim(Text3(19).Text) <> "" Or Trim(Text3(20).Text) <> "" Then
      mError = True
      MsgBox "Datos de Intervención de Terceros SIN CARGAR", vbExclamation, sMessage
   End If
   If Trim(Text4(0).Text) <> "" Or Trim(Text4(1).Text) <> "" Or Trim(Text4(2).Text) <> "" Or Trim(Text4(4).Text) <> "" Or Combo3(0).ListIndex > -1 Or Combo3(1).ListIndex > -1 Or Combo3(2).ListIndex > -1 Then
      mError = True
      MsgBox "Datos de Ocupantes SIN CARGAR", vbExclamation, sMessage
   End If
   If Trim(Text4(5).Text) <> "" Or Trim(Text4(6).Text) <> "" Or Combo4(0).ListIndex > -1 Or Combo4(1).ListIndex > -1 Or Combo4(2).ListIndex > -1 Or Combo4(3).ListIndex > -1 Or Combo4(4).ListIndex > -1 Or Combo4(5).ListIndex > -1 Or Combo4(6).ListIndex > -1 Then
      mError = True
      MsgBox "Datos de Vehículos SIN CARGAR", vbExclamation, sMessage
   End If
   
   fValidaDatos = mError
   Set mObjRN = Nothing
End Function

Private Function fVerError(ByRef pObj As Object, ByVal pTipo As Integer) As Boolean 'TIPO: 1-texto, 2-combo
   Select Case pTipo
      Case 1 'Controla TEXTs
         If Trim(pObj.Text) = "" Then
            pObj.BackColor = &H7282F1
            fVerError = True
         Else
            pObj.BackColor = &HFFFFFF
            fVerError = False
         End If
      
      Case 2 'controla COMBOS
         If pObj.ListIndex = -1 Then
            pObj.BackColor = &H7282F1
            fVerError = True
         Else
            pObj.BackColor = &HFFFFFF
            fVerError = False
         End If
         
      Case 3 'controla CHECK
         fVerError = True
         For mi = 0 To pObj.UBound
            If pObj(mi).Value <> 0 Then
               fVerError = False
            End If
         Next
         If fVerError Then
            For mi = 0 To pObj.UBound
               pObj(mi).BackColor = &H7282F1
            Next
         Else
            For mi = 0 To pObj.UBound
               If Combo5.Visible Then
                  pObj(mi).BackColor = &HB3C1CC
               Else
                  pObj(mi).BackColor = &HCECECE
               End If
            Next
         End If
      
      Case 4 'Controla objetos Flexgrid
         fVerError = False
         If pObj.Rows < 3 Then
            If Trim(pObj.TextMatrix(1, 0)) = "" Then
               fVerError = True
               MsgBox "Falta Agregar datos de Intervención de patrullas", vbCritical, sMessage
               Text3(15).BackColor = &H7282F1
               For mi = 0 To 2
                  Combo8(mi).BackColor = &H7282F1
               Next
            Else
               For mi = 0 To 2
                  Combo8(mi).BackColor = &HFFFFFF
               Next
               Text3(15).BackColor = &HFFFFFF
            End If
         End If
   End Select
End Function

Private Sub sBorrarTodo()
Dim mj As Integer
   Text1(5).Tag = Text1(5).Text
   For mi = 0 To Text1.UBound
      Text1(mi).Text = ""
   Next
   Text1(5).Text = Format(Val(Text1(5).Tag) + 1, "00000")
   Text2.Text = ""
   For mi = 0 To Text3.UBound
      If mi <> 13 And mi <> 14 And mi <> 16 Then
         Text3(mi).Text = ""
      End If
   Next
   For mi = 0 To Text4.UBound
      Text4(mi).Text = ""
   Next
   Combo1(1).ListIndex = -1 'Ramal
   Combo1(0).Clear 'Sentido
   
   For mi = 0 To Combo2.UBound
      Combo2(mi).ListIndex = -1
   Next
   For mi = 0 To Combo3.UBound
      Combo3(mi).ListIndex = -1
   Next
   For mi = 0 To Combo4.UBound
      Combo4(mi).ListIndex = -1
   Next
   For mi = 0 To Combo7.UBound
      Combo7(mi).ListIndex = -1
   Next
   For mi = 0 To Check1.UBound
      Check1(mi).Value = 0
   Next
   For mi = 0 To Check2.UBound
      Check2(mi).Value = 0
   Next
   Check3.Value = 0
   Check4.Value = 0
   For mi = 0 To Check5.UBound
      Check5(mi).Value = 0
   Next
   For mi = 0 To Check6.UBound
      Check6(mi).Value = 0
   Next
   For mj = 0 To Flex1.UBound
      Flex1(mj).Tag = ""
      For mi = Flex1(mj).Rows - 1 To 2 Step -1
         Flex1(mj).RemoveItem mi
      Next
      For mi = 0 To Flex1(mj).Cols - 1
         Flex1(mj).TextMatrix(1, mi) = ""
      Next
   Next
   For mi = 0 To 9 Step 3
      sVerIconos mi, False
   Next
   Label3(40).Caption = "Datos de Peatón"
   Frame3.Visible = False
   Image1_Click 1
   For mi = 0 To Flex1.UBound
      Flex1(mi).Enabled = True
   Next
End Sub

Public Sub sInitModif()
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset

   Text1(5).Visible = False
   Image4.Visible = True
   Combo5.Visible = True
   Set mRec = mObj.oBuscar("Ficha", "where fecha >= '2008-03-01' order by 1")
   Do While Not mRec.EOF
      Combo5.AddItem mRec.Fields(0) & " - " & mRec!Fecha
      mRec.MoveNext
   Loop
   mRec.Close
   Command1(1).Caption = "Modificar"
   Me.BackColor = &HB3C1CC
   Frame1.BackColor = &HB3C1CC
   Frame2.BackColor = &HB3C1CC
   Frame3.BackColor = &HB3C1CC
   For mi = 0 To Check1.UBound
      Check1(mi).BackColor = &HB3C1CC
   Next
   For mi = 0 To Check2.UBound
      Check2(mi).BackColor = &HB3C1CC
   Next
   Check3.BackColor = &HB3C1CC
   Check4.BackColor = &HB3C1CC
   For mi = 0 To Check5.UBound
      Check5(mi).BackColor = &HB3C1CC
   Next
   For mi = 0 To Check6.UBound
      Check6(mi).BackColor = &HB3C1CC
   Next
   Check7(0).Visible = True
   Check7(1).Visible = True
   Check7(0).Value = 1
   Check7(1).Value = 1
   Command1(1).BackColor = &HCCC8AC
   Frame1.Enabled = False
   
   Set mObj = Nothing
   Set mRec = Nothing
End Sub

Private Sub sSelectCombo(ByRef pObj As Object, ByVal pDato As String)
   Dim pI As Integer
   pObj.ListIndex = -1
   For pI = 0 To pObj.ListCount - 1 ' Sentido
      If Trim(Right(pObj.List(pI), 2)) = pDato Then
         pObj.ListIndex = pI
      End If
   Next
End Sub

Private Sub sDatosModif(ByVal pNroOrden)
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset
Dim mRec1 As New ADODB.Recordset
Dim mAuxi(10) As String


   
   Set mRec = mObj.oTablaNroOrden("Ficha", pNroOrden, "")
   If Not mRec.EOF Then
      If mRec!codtipoficha = "02" Then
         Combo9.ListIndex = 1
      Else
         Combo9.ListIndex = 0
      End If
      Text1(0).Text = mRec!Fecha
      Text1(1).Text = NVL(mRec!CodAlfa, "")
      Text1(2).Text = Format(mRec!Progresiva, "00.00")
      Text1(3).Text = mRec!hora
      Text1(4).Text = mRec!HoraLlegada
      
      'CON HORA FIN
      'Text1(6).Text = NVL(mRec!HoraFin, "")
      
      
      Text2.Text = NVL(mRec!OBS, "")
      sSelectCombo Combo1(1), NVL(mRec!CodRamal, "")
'      For mI = 0 To Combo1(1).ListCount - 1
'         If CInt(Trim(Right(Combo1(1).List(mI), 3))) = mRec!codramal Then
'            Combo1(1).ListIndex = mI
'         End If
'      Next
      For mi = 0 To Combo1(0).ListCount - 1
         If CInt(Trim(Right(Combo1(0).List(mi), 3))) = mRec!SentidoTrans Then
            Combo1(0).ListIndex = mi
         End If
      Next
      sSelectCombo Combo2(0), NVL(mRec!CODCONFIG, "")
      sSelectCombo Combo2(1), mRec!EstCalzada
      sSelectCombo Combo2(2), NVL(mRec!CODVISIBILIDAD, "")
      sSelectCombo Combo2(3), mRec!lugaraccid
      sSelectCombo Combo2(4), NVL(mRec!CODPAVIM, "")
      sSelectCombo Combo2(5), mRec!Clima1
      sSelectCombo Combo2(6), mRec!Iluminac
      sSelectCombo Combo2(7), NVL(mRec!CODINCONV, "")
      sSelectCombo Combo2(8), mRec!AcciconOtro
      sSelectCombo Combo2(9), mRec!AccidOtro
      sSelectCombo Combo2(10), mRec!CodColisContra1
      sSelectCombo Combo2(11), mRec!CodCausaCond1
      sSelectCombo Combo7(0), mRec!CodCausaCond2
      sSelectCombo Combo7(1), mRec!CodCausaCond3
      sSelectCombo Combo2(12), mRec!causaVehic
      'habría que pasar todos los datos anteriores de foto al nuevo formato de 5 digitos
      If Len(mRec!carril) = 11 Then
         For mi = 0 To Check1.UBound
            Check1(mi).Value = Mid(mRec!carril, mi + 1, 1)
         Next
      End If
      If Len(mRec!DemarcHoriz) = 3 Then
         For mi = 0 To Check2.UBound
            Check2(mi).Value = Mid(mRec!DemarcHoriz, mi + 1, 1)
         Next
      End If
      If Len(mRec!Foto) = 5 Then
         For mi = 0 To Check5.UBound
            Check5(mi).Value = Mid(mRec!Foto, mi + 1, 1)
         Next
      End If
   End If
   mRec.Close
   
   'DAÑOS GCO
   Set mRec = mObj.oTablaNroOrden("daniosgco", pNroOrden, "")
   If Not mRec.EOF Then
      For mi = 1 To 10
         Text3(mi + 2).Text = mRec.Fields(mi)
      Next
   End If
   mRec.Close
   
   For mi = 1 To 10
      mAuxi(mi) = ""
   Next
   
   'Observaciones
   Set mRec = mObj.oTablaNroOrden("fichaobs", pNroOrden, "")
   Do While Not mRec.EOF
      Text3(mRec!Indice).Text = NVL(mRec!descripcion, "")
      mRec.MoveNext
   Loop
   mRec.Close
   
   'VEHICULOS INVOLUCRADOS
   Set mRec = mObj.oTablaNroOrden("VehiculosInvolucr", pNroOrden, "order by letra")
   Image2(13).Visible = False
   Label6(8).Visible = False
   Combo6.Visible = False
   If Not mRec.EOF Then
      Combo6.Clear
      Image2(13).Enabled = True
      Do While Not mRec.EOF
         'Tipo Vehiculo
         mAuxi(10) = mRec!letra
         Combo6.AddItem mRec!letra
         Set mRec1 = mObj.oTabla("TipoVehiculo", " where codtipovehic='" & mRec!CodTipoVehic & "' ")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("Marca", "where codtipovehic='" & mRec!CodTipoVehic & "' and codmarca='" & mRec!CodMarca & "' ")
         If Not mRec1.EOF Then
            mAuxi(2) = mRec1.Fields(2) & Space(40) & mRec1.Fields(1)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("CiaSeguros", "where codciaseguro='" & mRec!CodCiaSeguro & "'")
         If Not mRec1.EOF Then
            mAuxi(3) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         
         Set mRec1 = mObj.oTabla("colores", "where codigo='" & mRec!codcolor & "'")
         If Not mRec1.EOF Then
            mAuxi(4) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         If Len(mRec!ANEXOEST) > 0 Then
            For mi = 1 To 4
               mAuxi(mi + 4) = Mid(mRec!ANEXOEST, mi, 1)
            Next
         End If
         
         Flex1(0).AddItem mRec!letra & vbTab & mRec!TITULAR & vbTab & mAuxi(1) & vbTab & mRec!Dominio & vbTab & mAuxi(4) & vbTab & mAuxi(2) & vbTab & mRec!modelo & vbTab & mAuxi(3) & vbTab & mRec!NroPoliza & vbTab & mRec!ESTGRAL & vbTab & mRec!ESTNEUM & vbTab & mAuxi(5) & vbTab & mAuxi(6) & vbTab & mAuxi(7) & vbTab & mAuxi(8)
         mRec.MoveNext
      Loop
      Flex1(0).RemoveItem 1
      Image2(13).Visible = True
      Label6(8).Visible = True
      Combo6.Visible = True
   End If
   mRec.Close
   If mAuxi(10) <> "" Then
      Label5(9).Tag = Asc(mAuxi(10)) + 1  'PROXIMA LETRA DISPONIBLE PARA UN VEHICULO INVOLUCRADO
   End If
   
   'LLENAR OCUPANTES INVOLUCRADOS
   Set mRec = mObj.oTablaNroOrden("VictimasInvolucr", pNroOrden, "order by NroVictima")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         For mi = 1 To 4
            mAuxi(mi) = ""
         Next
         Set mRec1 = mObj.oTabla("TipoDocu", "where codtipodocu='" & mRec!TipoDocu & "'")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("LugarTrasl", "where codlugartrasl='" & mRec!codlugartrasl & "'")
         If Not mRec1.EOF Then
            mAuxi(2) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("estadoocupa", "where codigo='" & mRec!codestado & "'")
         If Not mRec1.EOF Then
            mAuxi(3) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mAuxi(4) = NVL(mRec!letra, "")
         If mAuxi(4) <> "" And mRec!conductor = "1" Then
            mAuxi(4) = mAuxi(4) & Space(20) & "C"
         End If
         mRec1.Close
         mAuxi(5) = ""
         Set mRec1 = mObj.oTabla("MedioTrasl", "where codmediotrasl='" & mRec!codmediotrasl & "'")
         If Not mRec1.EOF Then
            mAuxi(5) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Flex1(1).AddItem mRec!nrovictima & vbTab & mAuxi(4) & vbTab & mRec!Nombre & vbTab & mRec!domicilio & vbTab & mAuxi(1) & vbTab & mRec!nrodocu & vbTab & mRec!edad & vbTab & mRec!TEL & vbTab & mAuxi(2) & vbTab & mAuxi(3) & vbTab & mRec!cinturon & vbTab & mAuxi(5)
         mRec.MoveNext
      Loop
      Flex1(1).RemoveItem 1
   End If
   mRec.Close
   'LLENAR DESCRIPCION ADICIONAL
   Set mRec = mObj.oTablaNroOrden("fichadescr", pNroOrden, "")
   If Not mRec.EOF Then
      For mi = 1 To 4
         Text4(mi + 6).Text = mRec.Fields(mi)
      Next
   End If
   mRec.Close
   
   Set mRec = mObj.oTablaNroOrden("intergco", pNroOrden, "")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         Flex1(2).AddItem "d" & vbTab & mRec!PATRULLERO1 & vbTab & mRec!PATRULLERO2 & vbTab & mRec!POLAD & vbTab & mRec!CodMovil
         mRec.MoveNext
      Loop
      Flex1(2).RemoveItem 1
      sFlexOrden 2
   End If
   mRec.Close
   
   Set mRec = mObj.oTablaNroOrden("interterceros", pNroOrden, "")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         mAuxi(1) = ""
         Set mRec1 = mObj.oTabla("movilext", " where codigo='" & mRec!codtipo & "' ")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(0) & "-" & mRec1.Fields(1)
         End If
         mRec1.Close
         Flex1(3).AddItem "d" & vbTab & mAuxi(1) & vbTab & mRec!Movil & vbTab & mRec!PERSONAL & vbTab & mRec!dependencia
         mRec.MoveNext
      Loop
      Flex1(3).RemoveItem 1
      sFlexOrden 3
      
   End If
   mRec.Close
   
End Sub

Private Function fGrabarDatos(ByVal pNroOrden As String) As Boolean
Dim mObj As New clRAcc
Dim mObjrnov As New clRNov
Dim mVectDanios(10) As Integer
Dim mParam As String
Dim mCarril As String
Dim mFotoAnexo As String
Dim mSenializ As String
   
   mTablasError = ""
   fGrabarDatos = False
   
'Grabar Ficha
   mCarril = ""
   For mi = 0 To Check1.UBound
     mCarril = mCarril & Check1(mi).Value
   Next
   mSenializ = ""
   For mi = 0 To Check2.UBound
     mSenializ = mSenializ & Check2(mi).Value
   Next
   mFotoAnexo = ""
   For mi = 0 To Check5.UBound
     mFotoAnexo = mFotoAnexo & Check5(mi).Value
   Next
    'sin Hora Fin
   If Command1(1).Caption = "Grabar" Then
      fGrabarDatos = fGrabarDatos Or mObj.xInsertFicha(pNroOrden, Text1(0).Text, Text1(3).Text, Text1(4).Text, Text1(2).Text, mCarril, Right(Combo2(8).Text, 2), Right(Combo2(10).Text, 2), Right(Combo2(9).Text, 2), Trim(Right(Combo1(0).Text, 2)), Right(Combo2(3).Text, 2), Right(Combo2(5).Text, 2), Right(Combo2(1).Text, 2), mSenializ, Right(Combo2(6).Text, 2), Right(Combo2(11).Text, 2), Right(Combo7(0).Text, 2), Right(Combo7(1).Text, 2), Right(Combo2(12).Text, 2), mFotoAnexo, Text1(1).Text, Right(Combo2(0).Text, 2), Right(Combo2(2).Text, 2), Right(Combo2(4).Text, 1), Right(Combo2(7).Text, 2), Trim(Text2.Text), Right(Combo9.Text, 2), Trim(Right(Combo1(1).Text, 2)))
   Else
      fGrabarDatos = fGrabarDatos Or Not mObj.xUpFicha(pNroOrden, Text1(0).Text, Text1(3).Text, Text1(4).Text, Text1(2).Text, mCarril, Right(Combo2(8).Text, 2), Right(Combo2(10).Text, 2), Right(Combo2(9).Text, 2), Trim(Right(Combo1(0).Text, 2)), Right(Combo2(3).Text, 2), Right(Combo2(5).Text, 2), Right(Combo2(1).Text, 2), mSenializ, Right(Combo2(6).Text, 2), Right(Combo2(11).Text, 2), Right(Combo7(0).Text, 2), Right(Combo7(1).Text, 2), Right(Combo2(12).Text, 2), mFotoAnexo, Text1(1).Text, Right(Combo2(0).Text, 2), Right(Combo2(2).Text, 2), Right(Combo2(4).Text, 1), Right(Combo2(7).Text, 2), Trim(Text2.Text), Right(Combo9.Text, 2), Trim(Right(Combo1(1).Text, 2)))
   End If
   
   'begin CU. 07-10-2009 _ Modif. para unir reg. novedades con reg. accidentes.
   If mObjrnov.bExistDatoTabla("enlace", "codigo='" & Trim(Text1(1).Text) & "' and fecha between '" & Format(DateAdd("d", -5, Text1(0).Text), "yyyy-mm-dd") & "' and '" & Format(DateAdd("d", 5, Text1(0).Text), "yyyy-mm-dd") & "'") Then
      mObjrnov.xUpEnlace Trim(Text1(1).Text), Text1(0).Text, pNroOrden
   Else 'codigo, fecha, codoperador, fechaproc
      mObjrnov.xInEnlace Trim(Text1(1).Text), Trim(Text1(5).Text), Trim(Text1(0).Text), Trim(Right(MDI.mUser, 15))
   End If
   'fin CU. 07-10-2009
   If fGrabarDatos Then
      mTablasError = "Fichas "
   End If
    'GRABA OBSERVACIONES
   For mi = 0 To 2
      If Text3(mi).Visible Then
         If Not mObj.xInsertObserv(pNroOrden, mi, Trim(Text3(mi).Text)) Then
            fGrabarDatos = True
            mTablasError = mTablasError & " - Observaciones "
         End If
      End If
   Next
   If Text3(17).Visible Then
      If Not mObj.xInsertObserv(pNroOrden, 17, Trim(Text3(17).Text)) Then
         fGrabarDatos = True
         mTablasError = mTablasError & " - Observaciones "
      End If
   End If
   If Text3(21).Visible Then
      If Not mObj.xInsertObserv(pNroOrden, 21, Trim(Text3(21).Text)) Then
         fGrabarDatos = True
         mTablasError = mTablasError & " - Observaciones "
      End If
   End If
    
   'graba daños GCO
   For mi = 1 To 10
      mVectDanios(mi) = 0
      If Trim(Text3(mi + 2).Text) <> "" Then
         mVectDanios(mi) = Val(Trim(Text3(mi + 2).Text))
      End If
   Next
   If Not mObj.xInsertDaniosGCO(pNroOrden, mVectDanios(1), mVectDanios(2), mVectDanios(3), mVectDanios(4), mVectDanios(5), mVectDanios(6), mVectDanios(7), mVectDanios(8), mVectDanios(9), mVectDanios(10)) Then
      fGrabarDatos = True
      mTablasError = mTablasError & " - Daños GCO "
   End If
   
   'graba intervenciones GCO
   If Flex1(2).TextMatrix(1, 0) <> "" Then
      For mi = 1 To Flex1(2).Rows - 1
         If Not mObj.xInsertIntGCO(pNroOrden, Trim(Flex1(2).TextMatrix(mi, 1)), Trim(Flex1(2).TextMatrix(mi, 2)), Trim(Flex1(2).TextMatrix(mi, 3)), Trim(Flex1(2).TextMatrix(mi, 4))) Then
            fGrabarDatos = True
            mTablasError = mTablasError & " - Interv GCO "
         End If
      Next
   End If
   
   'graba intervenciones terceros
   If Flex1(3).TextMatrix(1, 0) <> "" Then
      For mi = 1 To Flex1(3).Rows - 1
         If Not mObj.xInsertIntTerc(pNroOrden, Trim(Left(Flex1(3).TextMatrix(mi, 1), 2)), Trim(Flex1(3).TextMatrix(mi, 2)), Trim(Flex1(3).TextMatrix(mi, 3)), Trim(Flex1(3).TextMatrix(mi, 4))) Then
            fGrabarDatos = True
            mTablasError = mTablasError & " - Interv Terceros "
         End If
      Next
   End If
   
   'graba descripcion adicional
   If Trim(Text4(7).Text) <> "" Or Trim(Text4(8).Text) <> "" Or Trim(Text4(9).Text) <> "" Or Trim(Text4(10).Text) <> "" Then
      If Not mObj.xInsertDescr(pNroOrden, Left(Trim(Text4(7).Text), 1000), Left(Trim(Text4(8).Text), 1000), Left(Trim(Text4(9).Text), 1000), Left(Trim(Text4(10).Text), 1000)) Then
         fGrabarDatos = True
         mTablasError = mTablasError & " - Descrip "
      End If
   End If
        
   'graba vehiculos
   If Flex1(0).TextMatrix(1, 0) <> "" Then
      With Flex1(0)
         For mi = 1 To Flex1(0).Rows - 1
            mParam = ""
            For mj = 11 To 14
               mParam = mParam & Trim(.TextMatrix(mi, mj))
            Next
            If Not mObj.xInsertVehiculos(pNroOrden, Trim(Right(.TextMatrix(mi, 0), 2)), .TextMatrix(mi, 1), Right(.TextMatrix(mi, 2), 2), Trim(.TextMatrix(mi, 3)), Right(.TextMatrix(mi, 4), 2), Right(.TextMatrix(mi, 5), 2), .TextMatrix(mi, 6), Trim(Right(.TextMatrix(mi, 7), 3)), Trim(.TextMatrix(mi, 8)), Trim(.TextMatrix(mi, 9)), Trim(.TextMatrix(mi, 10)), mParam) Then
               fGrabarDatos = True
               mTablasError = mTablasError & " - Vehiculos"
            End If
         Next
      End With
   End If
   
   'graba ocupantes
   If Flex1(1).TextMatrix(1, 0) <> "" Then
      With Flex1(1)
         For mi = 1 To Flex1(1).Rows - 1
            mj = 0
            If Right(.TextMatrix(mi, 1), 1) = "C" And Len(.TextMatrix(mi, 1)) > 10 Then
               mj = 1
            End If
            If Not mObj.xInsertVictimas(pNroOrden, .TextMatrix(mi, 0), Trim(Left(.TextMatrix(mi, 1), 2)), Trim(.TextMatrix(mi, 2)), Trim(.TextMatrix(mi, 3)), Right(.TextMatrix(mi, 4), 2), Trim(.TextMatrix(mi, 5)), Trim(.TextMatrix(mi, 6)), Trim(.TextMatrix(mi, 7)), Right(.TextMatrix(mi, 8), 2), Trim(Right(.TextMatrix(mi, 11), 2)), Trim(Right(.TextMatrix(mi, 9), 2)), Trim(.TextMatrix(mi, 10)), mj) Then
               fGrabarDatos = True
               mTablasError = mTablasError & " - Victimas"
            End If
         Next
      End With
   End If
   Set mObj = Nothing
   Set mObjrnov = Nothing
End Function

Private Sub sRollBack(ByVal pNroOrden As String)
Dim mObj As New clRAcc
   mObj.xDeleteTable "fichadescr", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "interterceros", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "intergco", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "VehiculosInvolucr", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "VictimasInvolucr", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "daniosgco", " nroorden='" & pNroOrden & "' "
   mObj.xDeleteTable "fichaobs", " nroorden='" & pNroOrden & "' "
   Set mObj = Nothing
End Sub

Private Sub sDatosModifViejos(ByVal pNroOrden As String)
Dim mObj As New clRAcc
Dim mRec As New ADODB.Recordset
Dim mAuxi(10) As String
Dim mRec1 As New ADODB.Recordset

   Set mRec = mObj.oTablaNroOrden("Ficha", pNroOrden, "")
   If Not mRec.EOF Then
      If mRec!codtipoficha = "02" Then
         Combo9.ListIndex = 1
      Else
         Combo9.ListIndex = 0
      End If
      Text1(0).Text = mRec!Fecha
      Text1(1).Text = ""
      Text1(2).Text = mRec!Progresiva
      Text1(3).Text = mRec!hora
      Text1(4).Text = mRec!HoraLlegada
      Text2.Text = NVL(mRec!OBS, "")
      sSelectCombo Combo1, mRec!SentidoTrans
      sSelectCombo Combo2(0), NVL(mRec!CODCONFIG, "")
      sSelectCombo Combo2(1), mRec!EstCalzada
      sSelectCombo Combo2(2), NVL(mRec!CODVISIBILIDAD, "")
      sSelectCombo Combo2(3), mRec!lugaraccid
      sSelectCombo Combo2(4), NVL(mRec!CODPAVIM, "")
      sSelectCombo Combo2(5), mRec!Clima1
      sSelectCombo Combo2(6), mRec!Iluminac
      sSelectCombo Combo2(7), NVL(mRec!CODINCONV, "")
      sSelectCombo Combo2(8), mRec!AcciconOtro
      sSelectCombo Combo2(9), mRec!AccidOtro
      sSelectCombo Combo2(10), mRec!CodColisContra1
      sSelectCombo Combo2(11), mRec!CodCausaCond1
      sSelectCombo Combo2(12), mRec!causaVehic
      'habría que pasar todos los datos anteriores de foto al nuevo formato de 5 digitos
      If Len(mRec!carril) = 6 Then
         For mi = 0 To Check1.UBound
            Check1(mi).Value = Mid(mRec!carril, mi + 1, 1)
         Next
      End If
      If Len(mRec!DemarcHoriz) = 3 Then
         For mi = 0 To Check2.UBound
            Check2(mi).Value = Mid(mRec!DemarcHoriz, mi + 1, 1)
         Next
      End If
      If Len(mRec!Foto) = 5 Then
         For mi = 0 To Check5.UBound
            Check5(mi).Value = Mid(mRec!Foto, mi + 1, 1)
         Next
      End If
   End If
   mRec.Close
   
   'DAÑOS GCO
   Set mRec = mObj.oTablaNroOrden("daniosgco", pNroOrden, "")
   If Not mRec.EOF Then
      For mi = 1 To 10
         Text3(mi + 2).Text = mRec.Fields(mi)
      Next
   End If
   mRec.Close
   
   For mi = 1 To 10
      mAuxi(mi) = ""
   Next
   
   'Observaciones
   Set mRec = mObj.oTablaNroOrden("fichaobs", pNroOrden, "")
   Do While Not mRec.EOF
      Text3(mRec!Indice).Text = NVL(mRec!descripcion, "")
      mRec.MoveNext
   Loop
   mRec.Close
   
   'VEHICULOS INVOLUCRADOS
   Set mRec = mObj.oTablaNroOrden("VehiculosInvolucr", pNroOrden, "order by letra")
   Image2(13).Visible = False
   Label6(8).Visible = False
   Combo6.Visible = False
   If Not mRec.EOF Then
      Combo6.Clear
      Image2(13).Enabled = True
      Do While Not mRec.EOF
         'Tipo Vehiculo
         mAuxi(10) = mRec!letra
         Combo6.AddItem mRec!letra
         Set mRec1 = mObj.oTabla("TipoVehiculo", " where codtipovehic='" & mRec!CodTipoVehic & "' ")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("Marca", "where codtipovehic='" & mRec!CodTipoVehic & "' and codmarca='" & mRec!CodMarca & "' ")
         If Not mRec1.EOF Then
            mAuxi(2) = mRec1.Fields(2) & Space(40) & mRec1.Fields(1)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("CiaSeguros", "where codciaseguro='" & mRec!CodCiaSeguro & "'")
         If Not mRec1.EOF Then
            mAuxi(3) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         
         Set mRec1 = mObj.oTabla("colores", "where codigo='" & mRec!codcolor & "'")
         If Not mRec1.EOF Then
            mAuxi(4) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         If Len(mRec!ANEXOEST) > 0 Then
            For mi = 1 To 4
               mAuxi(mi + 4) = Mid(mRec!ANEXOEST, mi, 1)
            Next
         End If
         
         Flex1(0).AddItem mRec!letra & vbTab & mRec!TITULAR & vbTab & mAuxi(1) & vbTab & mRec!Dominio & vbTab & mAuxi(4) & vbTab & mAuxi(2) & vbTab & mRec!modelo & vbTab & mAuxi(3) & vbTab & mRec!NroPoliza & vbTab & mRec!ESTGRAL & vbTab & mRec!ESTNEUM & vbTab & mAuxi(5) & vbTab & mAuxi(6) & vbTab & mAuxi(7) & vbTab & mAuxi(8)
         mRec.MoveNext
      Loop
      Flex1(0).RemoveItem 1
      Image2(13).Visible = True
      Label6(8).Visible = True
      Combo6.Visible = True
   End If
   mRec.Close
   If mAuxi(10) <> "" Then
      Label5(9).Tag = Asc(mAuxi(10)) + 1  'PROXIMA LETRA DISPONIBLE PARA UN VEHICULO INVOLUCRADO
   End If
   
   'LLENAR OCUPANTES INVOLUCRADOS
   Set mRec = mObj.oTablaNroOrden("VictimasInvolucr", pNroOrden, "order by NroVictima")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         For mi = 1 To 5
            mAuxi(mi) = ""
         Next
         Set mRec1 = mObj.oTabla("TipoDocu", "where codtipodocu='" & mRec!TipoDocu & "'")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("LugarTrasl", "where codlugartrasl='" & mRec!codlugartrasl & "'")
         If Not mRec1.EOF Then
            mAuxi(2) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("MedioTrasl", "where codmediotrasl='" & mRec!codmediotrasl & "'")
         If Not mRec1.EOF Then
            mAuxi(5) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mRec1.Close
         Set mRec1 = mObj.oTabla("estadoocupa", "where codigo='" & mRec!codestado & "'")
         If Not mRec1.EOF Then
            mAuxi(3) = mRec1.Fields(1) & Space(40) & mRec1.Fields(0)
         End If
         mAuxi(4) = NVL(mRec!letra, "")
         If mAuxi(4) <> "" And mRec!conductor = "1" Then
            mAuxi(4) = mAuxi(4) & Space(20) & "C"
         End If
         mRec1.Close
         Flex1(1).AddItem mRec!nrovictima & vbTab & mAuxi(4) & vbTab & mRec!Nombre & vbTab & mRec!domicilio & vbTab & mAuxi(1) & vbTab & mRec!nrodocu & vbTab & mRec!edad & vbTab & mRec!TEL & vbTab & mAuxi(2) & vbTab & mAuxi(3) & vbTab & mRec!cinturon & vbTab & mAuxi(5)
         mRec.MoveNext
      Loop
      Flex1(1).RemoveItem 1
   End If
   mRec.Close
   'LLENAR DESCRIPCION ADICIONAL
   Set mRec = mObj.oTablaNroOrden("fichadescr", pNroOrden, "")
   If Not mRec.EOF Then
      For mi = 1 To 4
         Text4(mi + 6).Text = mRec.Fields(mi)
      Next
   End If
   mRec.Close
   
   Set mRec = mObj.oTablaNroOrden("intergco", pNroOrden, "")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         Flex1(2).AddItem "d" & vbTab & mRec!PATRULLERO1 & vbTab & mRec!PATRULLERO2 & vbTab & mRec!POLAD & vbTab & mRec!CodMovil
         mRec.MoveNext
      Loop
      Flex1(2).RemoveItem 1
      sFlexOrden 2
   End If
   mRec.Close
   
   Set mRec = mObj.oTablaNroOrden("interterceros", pNroOrden, "")
   If Not mRec.EOF Then
      Do While Not mRec.EOF
         mAuxi(1) = ""
         Set mRec1 = mObj.oTabla("movilext", " where codigo='" & mRec!codtipo & "' ")
         If Not mRec1.EOF Then
            mAuxi(1) = mRec1.Fields(0) & "-" & mRec1.Fields(1)
         End If
         mRec1.Close
         Flex1(3).AddItem "d" & vbTab & mAuxi(1) & vbTab & mRec!Movil & vbTab & mRec!PERSONAL & vbTab & mRec!dependencia
         mRec.MoveNext
      Loop
      Flex1(3).RemoveItem 1
      sFlexOrden 3
      
   End If
   mRec.Close
   Set mObj = Nothing
End Sub

Private Sub sChgColorForm(ByVal pTipo As String)
Dim mColor
   Select Case pTipo
      Case "01"
         mColor = &HCECECE
      Case "02"
         mColor = &HC1DBD8
   End Select
   Me.BackColor = mColor
   For mi = 0 To Check1.UBound
      Check1(mi).BackColor = mColor
   Next
   For mi = 0 To Check2.UBound
      Check2(mi).BackColor = mColor
   Next
   Check3.BackColor = mColor
   Check4.BackColor = mColor
   For mi = 0 To Check5.UBound
      Check5(mi).BackColor = mColor
   Next
   Check7(0).BackColor = mColor
   Check7(1).BackColor = mColor
   Frame1.BackColor = mColor
   Frame2.BackColor = mColor
   Frame3.BackColor = mColor
End Sub

Private Sub sPatrulleros()
Dim mObjrnov As New clRNov
Dim mRec As New ADODB.Recordset

   Combo8(0).Clear
   Combo8(1).Clear
   Set mRec = mObjrnov.oTablaNull("patrulleros")
   Do While Not mRec.EOF
      Combo8(0).AddItem mRec!Nombre & Space(50) & mRec!Codigo
      Combo8(1).AddItem mRec!Nombre & Space(50) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close
   
   Set mObjrnov = Nothing
   Set mRec = Nothing
End Sub

Public Sub sLlenoSentido()
Dim mCodRamal As String
Dim mObj As New clRNov
Dim mRec As New ADODB.Recordset
   
   mCodRamal = Right(Combo1(1).Text, 1)
   Combo1(0).Clear
   Set mRec = mObj.oTabla("sentidos", "where codRamal = " & mCodRamal & " order by 2")
   Do While Not mRec.EOF
     Combo1(0).AddItem mRec!descripcion & Space(60) & mRec!Codigo
     mRec.MoveNext
   Loop
   mRec.Close
   Set mObj = Nothing
   Set mRec = Nothing
End Sub
