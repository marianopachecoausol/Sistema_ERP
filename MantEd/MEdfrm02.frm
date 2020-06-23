VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MEdfrm02 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Edilicio"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Elija la opción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   960
      TabIndex        =   53
      Top             =   5160
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "Volver"
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   58
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   57
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Anular Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   55
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Modificar Fecha de Solicitud"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   54
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   480
         Width           =   2010
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Anulación de parte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3120
      TabIndex        =   47
      Top             =   4200
      Width           =   10695
      Begin VB.CommandButton Command3 
         Caption         =   "Volver"
         Height          =   375
         Index           =   1
         Left            =   6480
         TabIndex        =   51
         Top             =   1360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   50
         Top             =   1360
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         MaxLength       =   89
         TabIndex        =   49
         Top             =   800
         Width           =   8655
      End
      Begin VB.Label Label3 
         Caption         =   "Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   2010
      End
      Begin VB.Label Label3 
         Caption         =   "Motivo:"
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
         Left            =   240
         TabIndex        =   48
         Top             =   880
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Corrección de Fechas de Solicitud"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   960
      TabIndex        =   39
      Top             =   2880
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "&Volver"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   45
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   44
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   43
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hora"
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
         Left            =   840
         TabIndex        =   42
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Index           =   1
         Left            =   840
         TabIndex        =   41
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Parte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   1080
         TabIndex        =   40
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   9510
      MaxLength       =   150
      TabIndex        =   17
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   4350
      MaxLength       =   50
      TabIndex        =   16
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   3390
      MaxLength       =   5
      TabIndex        =   15
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   2550
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   12
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   7680
      MaxLength       =   90
      TabIndex        =   6
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   19
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   10080
      TabIndex        =   20
      Top             =   8380
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   595
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   19
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   360
      MaxLength       =   5
      TabIndex        =   10
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   4680
      MaxLength       =   5
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   13590
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   14670
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   15990
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Columns         =   8
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "MEdfrm02.frx":0000
      Left            =   1080
      List            =   "MEdfrm02.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   2640
      Width           =   16965
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4980
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   18705
      _ExtentX        =   32994
      _ExtentY        =   8784
      _Version        =   327680
      Cols            =   21
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   18720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   16
      Left            =   9510
      TabIndex        =   38
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   9270
      X2              =   9270
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Materiales"
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
      Index           =   15
      Left            =   4350
      TabIndex        =   37
      Top             =   1800
      Width           =   885
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   4110
      X2              =   4110
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "$/Tarea"
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
      Left            =   3360
      TabIndex        =   36
      Top             =   1800
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hs/Per"
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
      Left            =   2550
      TabIndex        =   35
      Top             =   1800
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   3270
      X2              =   3270
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "M.Obra"
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
      Left            =   360
      TabIndex        =   34
      Top             =   2760
      Width           =   630
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   240
      X2              =   2400
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   240
      X2              =   18720
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   15870
      X2              =   15870
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Registro de Partes de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7680
      TabIndex        =   33
      Top             =   75
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parte"
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
      Left            =   360
      TabIndex        =   32
      Top             =   600
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha/Hora Solicit."
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
      Left            =   1235
      TabIndex        =   31
      Top             =   600
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Edificio"
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
      Left            =   5400
      TabIndex        =   30
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion de la Solicitud"
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
      Left            =   7680
      TabIndex        =   29
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Prioridad"
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
      Left            =   13620
      TabIndex        =   28
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha / Hora Asistencia"
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
      Left            =   3050
      TabIndex        =   27
      Top             =   600
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tiempos"
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
      Left            =   840
      TabIndex        =   26
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Estim."
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
      Index           =   9
      Left            =   360
      TabIndex        =   25
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Real"
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
      Index           =   10
      Left            =   1140
      TabIndex        =   24
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Admis."
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
      Left            =   1785
      TabIndex        =   23
      Top             =   1800
      Width           =   570
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2880
      X2              =   2880
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   18720
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   480
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1050
      X2              =   1050
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5280
      X2              =   5280
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   7560
      X2              =   7560
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   13470
      X2              =   13470
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   240
      X2              =   18720
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   18720
      X2              =   18720
      Y1              =   480
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   14550
      X2              =   14550
      Y1              =   480
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Mant."
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
      Left            =   14745
      TabIndex        =   22
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rubro"
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
      Left            =   15990
      TabIndex        =   21
      Top             =   600
      Width           =   525
   End
End
Attribute VB_Name = "MEdfrm02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantEd
'Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mRenglon As Integer
Dim mParteElegido As Double
Dim mRenglonElegido As Integer
Dim mColumnaElegida As Integer

Private Sub Command1_Click(Index As Integer)
Dim mEstado As String
Dim mFecPro As String
Dim mFecTer As String
Dim mOkGrb As Boolean
Dim mI As Integer
Dim mJ As Integer
Dim mStrMO As String

Dim mTextoMail As String
Dim mErrMail As Integer
Dim mListaDestinatarios As String

If Index = 0 Then
   If fValida1 Then
      If MsgBox("¿Está Seguro de Grabar esta Orden?", vbYesNo, sMessage) = vbYes Then
         mEstado = "P"
         mFecPro = Now
         mFecTer = ""
         mOkGrb = True
         If MsgBox("¿Está terminado el trabajo?", vbYesNo, sMessage) = vbYes Then
            mOkGrb = False
            mEstado = "T"
            mFecTer = IIf(MSFlexGrid1.TextMatrix(mRenglon, 19) = "G", mFecPro, Now)
            If fValida2 Then
               mOkGrb = True
            End If
            If mOkGrb Then
               If DateDiff("n", CDate(Text1(2).Text & " " & Text1(3).Text & ":00"), CDate(mFecPro)) <= 0 Then
                  mOkGrb = False
                  MsgBox "Verificar la fecha de Asistencia", vbCritical, "Atención"
               End If
            End If
            If mOkGrb Then
               If DateDiff("n", CDate(Text1(1).Text), CDate(Text1(2).Text & " " & Text1(3).Text & ":00")) <= 0 Then
                  mOkGrb = False
                  MsgBox "Verificar la fecha de Asistencia", vbCritical, "Atención"
               End If
            End If
         End If

         If mOkGrb Then
            'Veo el string para ManoObra
            mStrMO = ""
            For mI = 0 To List1.ListCount - 1
               If List1.Selected(mI) Then
                  mStrMO = mStrMO & Left(List1.List(mI), InStr(1, List1.List(mI), "-"))
               End If
            Next
            'Completo el FlexGrid
            MSFlexGrid1.TextMatrix(mRenglon, 4) = Text1(2).Text
            MSFlexGrid1.TextMatrix(mRenglon, 5) = Text1(3).Text
            MSFlexGrid1.TextMatrix(mRenglon, 6) = Text1(4).Text
            MSFlexGrid1.TextMatrix(mRenglon, 7) = Combo1(0).Text
            MSFlexGrid1.TextMatrix(mRenglon, 8) = Text1(5).Text
            MSFlexGrid1.TextMatrix(mRenglon, 10) = Combo1(2).Text
            MSFlexGrid1.TextMatrix(mRenglon, 11) = Combo1(3).Text
            MSFlexGrid1.TextMatrix(mRenglon, 12) = Text1(6).Text
            MSFlexGrid1.TextMatrix(mRenglon, 13) = Text1(7).Text
            MSFlexGrid1.TextMatrix(mRenglon, 14) = Text1(8).Text
            MSFlexGrid1.TextMatrix(mRenglon, 15) = mStrMO
            MSFlexGrid1.TextMatrix(mRenglon, 16) = Text1(9).Text
            MSFlexGrid1.TextMatrix(mRenglon, 17) = Text1(10).Text
            MSFlexGrid1.TextMatrix(mRenglon, 18) = Text1(11).Text
            MSFlexGrid1.TextMatrix(mRenglon, 19) = Text1(12).Text
            'MSFlexGrid1.TextMatrix(mRenglon, 19) = mEstado
             MSFlexGrid1.TextMatrix(mRenglon, 20) = mEstado

            'Actualizo en Registros
            mObj.UpdRegistros Combo1(0).Text, Text1(5).Text, Text1(2).Text, Text1(3).Text, Text1(4).Text, Combo1(2).Text, Left(Combo1(3).Text, 3), Text1(6).Text, Text1(7).Text, Text1(8).Text, mStrMO, Text1(9).Text, Text1(10).Text, Text1(11).Text, Text1(12).Text, mFecPro, mFecTer, mEstado, Text1(0).Text


            If mEstado = "T" Then
               
              
               
               
               '==============================================
               'Limpieza
               'borrado de grilla
               'posicion en 1
               '==============================================
               
               
               mErrMail = 0

               mTextoMail = vbCrLf & "Se ha resuelto el Parte  " & Text1(0).Text & " de Mantenimiento Edilicio: " & vbCrLf & vbCrLf & "     Descripción de la solicitud:  " & Text1(5).Text & vbCrLf & vbCrLf & "Verifique el servicio realizado. Gracias"
               Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email  FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxEdilicio WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' AND FechaBaja IS NULL ")
               'Set mRec = mObj.oEjecutarSelect("SELECT * FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And Email <> '" & mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6) & "' And FechaBaja IS NULL")

               If Not mRec.EOF Then
                  mListaDestinatarios = ""
                  Do While Not mRec.EOF
                     mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                     mRec.MoveNext
                  Loop
                  If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", "MANT. EDILICIO - Repuesta a Solicitud de Servicios", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                     mErrMail = mErrMail + 1
                  End If
               End If
               If mErrMail = 0 Then
                  MsgBox "Se ha grabado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
               Else
                  MsgBox "Se ha grabado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
               End If
            End If
            
            limpiarCampos
               
            If MSFlexGrid1.Rows > 2 Then
               For mI = MSFlexGrid1.Row To MSFlexGrid1.Rows - 2
                  For mJ = 1 To MSFlexGrid1.Cols - 1
                     MSFlexGrid1.TextMatrix(mI, mJ) = MSFlexGrid1.TextMatrix(mI + 1, mJ)
                  Next
               Next
                  MSFlexGrid1.RemoveItem (MSFlexGrid1.Rows - 1)
               Else
                  MSFlexGrid1.AddItem ""
                  MSFlexGrid1.RemoveItem 1
               End If
               
               MSFlexGrid1.Row = 1
            
            

         End If
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub Command2_Click(Index As Integer)
MSFlexGrid1.Enabled = True
If Index = 0 Then
   If fValida3 Then
      mObj.UpdFechaSolic Mid(Label4(0).Caption, 7), Text2(0).Text & " " & Text2(1).Text & ":00"
      MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Text2(0).Text & " " & Text2(1).Text & ":00"
      Text1(1) = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
   End If
End If
Frame1.Visible = False
End Sub

Private Sub Command3_Click(Index As Integer)
'MSFlexGrid1.Enabled = True
Dim mParte As Double
Dim mI  As Integer
Dim mJ As Integer
Dim mListaDestinatarios As String
Dim mTextoMail  As String
Dim mErrMail As Integer
Dim mRow As Integer


If Index = 0 Then

      If MsgBox("¿Está seguro de anular esta tarea?", vbYesNo, sMessage) = vbYes Then
         Frame2.Visible = False
         
         mParte = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
         mRow = MSFlexGrid1.Row
         
         
        mTextoMail = "Se ha anulado una solicitud de servicio de personal de Mant. Edilicio, según detalle:" & vbCrLf & vbCrLf
        'mTextoMail = mTextoMail & vbCrLf & "Parte " & mParte & ":        "
        mTextoMail = mTextoMail & "Parte: " & mParte & vbCrLf & "Fecha Solicitud: " & MSFlexGrid1.TextMatrix(mRow, 3) & vbCrLf & "Lugar: " & MSFlexGrid1.TextMatrix(mRow, 7) & vbCrLf & "Descripción de la solicitud: " & MSFlexGrid1.TextMatrix(mRow, 8) & vbCrLf & vbCrLf & "Motivo de anulación: " & Text3.Text
            
        
         
         If MSFlexGrid1.Rows > 2 Then
            For mI = MSFlexGrid1.Row To MSFlexGrid1.Rows - 2
               For mJ = 1 To MSFlexGrid1.Cols - 1
                  MSFlexGrid1.TextMatrix(mI, mJ) = MSFlexGrid1.TextMatrix(mI + 1, mJ)
               Next
            Next
            MSFlexGrid1.RemoveItem (MSFlexGrid1.Rows - 1)
         Else
            MSFlexGrid1.AddItem ""
            MSFlexGrid1.RemoveItem 1
         End If
         'mObj.DelRegistros mParte
         mObj.AnularParte mParte, Trim(Right(MDI.mUser, 20)), Format(Now, "yyyy-mm-dd hh:mm:ss"), Text3.Text
        
        MSFlexGrid1.Row = 1
        
        
       
         mListaDestinatarios = ""
         If mTextoMail <> "" Then
            Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(mParte) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxEdilicio WHERE CodSuperv = '" & mObj.ObtCodSuperv(mParte) & "' And FechaBaja IS NULL ")
            If Not mRec.EOF Then
               Do While Not mRec.EOF
                  mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                  mRec.MoveNext
               Loop
               

            
               If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", " MANT. EDILICIO - Anulación de Solicitud de Servicio", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
                  mErrMail = mErrMail + 1
               End If
            End If
         End If
      
         If mErrMail = 0 Then
            MsgBox "Se ha anulado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
         Else
            MsgBox "Se ha anulado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
         End If
      End If
   
      

End If
Frame2.Visible = False
limpiarCampos

End Sub


Private Sub Command4_Click(Index As Integer)

If Index = 0 Then
   If Not (Option1.Value) And Not (Option2.Value) Then
      MsgBox "Debe seleccionar alguna de las dos opciones", vbExclamation, "Opciones"
      Exit Sub
   Else
      Frame3.Visible = False
      If Option1.Value Then
         Frame1.Visible = True
         Label4(0) = "Parte " & mParteElegido
         Text2(0).Text = Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3), 10)
         Text2(1).Text = Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3), 12, 5)
         Frame1.Top = MSFlexGrid1.CellTop + 705
      Else
          Text3.Text = ""
          Frame2.Visible = True
          Label3(1) = "Parte " & mParteElegido
          Frame2.Top = MSFlexGrid1.CellTop + 1305
      End If
   End If
Else
   Frame3.Visible = False
End If

limpiarOpciones

End Sub

Private Sub Form_Load()
Dim mI As Integer
MEdfrm02.Top = 100
MEdfrm02.Left = (MDI.Width - MEdfrm02.Width) / 2

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL ")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      'Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      Combo1(0).AddItem mRec!ZonaMantEdil & " - " & mRec!descripcion 'Agrego la ZonaMantEdil
      mRec.MoveNext
   Loop
End If
mRec.Close

Combo1(1).AddItem "Alta"
Combo1(1).AddItem "Media"
Combo1(1).AddItem "Baja"

Combo1(2).AddItem "Preventivo"
Combo1(2).AddItem "Predictivo"
Combo1(2).AddItem "Correctivo"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Rubros WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(3).AddItem mRec!Codigo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 1
MSFlexGrid1.ColWidth(2) = 600
MSFlexGrid1.ColWidth(3) = 1650
MSFlexGrid1.ColWidth(4) = 1000
MSFlexGrid1.ColWidth(5) = 600
MSFlexGrid1.ColWidth(6) = 600
MSFlexGrid1.ColWidth(7) = 1800
MSFlexGrid1.ColWidth(8) = 5000
MSFlexGrid1.ColWidth(9) = 700
MSFlexGrid1.ColWidth(10) = 900
MSFlexGrid1.ColWidth(11) = 1800
MSFlexGrid1.ColWidth(12) = 500
MSFlexGrid1.ColWidth(13) = 500
MSFlexGrid1.ColWidth(14) = 500
MSFlexGrid1.ColWidth(15) = 1300
MSFlexGrid1.ColWidth(16) = 600
MSFlexGrid1.ColWidth(17) = 600
MSFlexGrid1.ColWidth(18) = 1800
MSFlexGrid1.ColWidth(19) = 2300
MSFlexGrid1.ColWidth(20) = 350

For mI = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mI) = 2
Next

MSFlexGrid1.TextMatrix(0, 1) = ""
MSFlexGrid1.TextMatrix(0, 2) = "Parte"
MSFlexGrid1.TextMatrix(0, 3) = "Fecha Solicitud"
MSFlexGrid1.TextMatrix(0, 4) = "Asistencia"
MSFlexGrid1.TextMatrix(0, 5) = "H. Ini."
MSFlexGrid1.TextMatrix(0, 6) = "H. Fin"
MSFlexGrid1.TextMatrix(0, 7) = "Edificio"
MSFlexGrid1.TextMatrix(0, 8) = "Descripcion de la Solicitud"
MSFlexGrid1.TextMatrix(0, 9) = "Prioridad"
MSFlexGrid1.TextMatrix(0, 10) = "Tipo Mant."
MSFlexGrid1.TextMatrix(0, 11) = "Rubro"
MSFlexGrid1.TextMatrix(0, 12) = "Estim."
MSFlexGrid1.TextMatrix(0, 13) = "Real"
MSFlexGrid1.TextMatrix(0, 14) = "Admis."
MSFlexGrid1.TextMatrix(0, 15) = "Mano de Obra"
MSFlexGrid1.TextMatrix(0, 16) = "Hs/Per"
MSFlexGrid1.TextMatrix(0, 17) = "$/Tarea"
MSFlexGrid1.TextMatrix(0, 18) = "Materiales"
MSFlexGrid1.TextMatrix(0, 19) = "Observaciones"
MSFlexGrid1.TextMatrix(0, 20) = "Estado"

Set mRec = mObj.oEjecutarSelect(" SELECT R.* FROM " & _
" Registros R " & _
"  Left Join " & _
" AnulacionesParte A ON R.Parte = A.ParteAnu " & _
" WHERE (Estado <> 'T' or FechaSolic BETWEEN ADDDATE(CURDATE(), INTERVAL -420 DAY) and adddate(CURDATE(),interval 1 day)) " & _
" AND A.ParteAnu IS NULL order by 1 desc; ")

 
If Not mRec.EOF Then
   mI = 1
   Do While Not mRec.EOF
      mI = mI + 1
      MSFlexGrid1.AddItem ""
      MSFlexGrid1.TextMatrix(mI, 1) = ""
      MSFlexGrid1.TextMatrix(mI, 2) = mRec!Parte
      MSFlexGrid1.TextMatrix(mI, 3) = NVL(mRec!FechaSolic, "")
      MSFlexGrid1.TextMatrix(mI, 4) = NVL(mRec!FechaAsist, "")
      MSFlexGrid1.TextMatrix(mI, 5) = NVL(mRec!HoraIniAsist, "")
      MSFlexGrid1.TextMatrix(mI, 6) = NVL(mRec!HoraFinAsist, "")
      
      
      MSFlexGrid1.TextMatrix(mI, 7) = NVL(mRec!CodEdificio, "")
      'MSFlexGrid1.TextMatrix(mi, 7) = Right(NVL(mRec!CodEdificio, ""), Len(NVL(mRec!CodEdificio, "")) - 5) '--Saco el codigo de zona.
      MSFlexGrid1.TextMatrix(mI, 8) = NVL(mRec!DescripSolic, "")
      MSFlexGrid1.TextMatrix(mI, 9) = NVL(mRec!Prioridad, "")
      MSFlexGrid1.TextMatrix(mI, 10) = NVL(mRec!TipoMant, "")
      MSFlexGrid1.TextMatrix(mI, 11) = mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!CodRubro & "'", 0) & " - " & mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!CodRubro & "'", 1)
      MSFlexGrid1.TextMatrix(mI, 12) = NVL(mRec!TiempoEstim, 0)
      MSFlexGrid1.TextMatrix(mI, 13) = NVL(mRec!TiempoReal, 0)
      MSFlexGrid1.TextMatrix(mI, 14) = NVL(mRec!TiempoAdmis, 0)
      MSFlexGrid1.TextMatrix(mI, 15) = NVL(mRec!ManoObra, "")
      MSFlexGrid1.TextMatrix(mI, 16) = NVL(mRec!Horas, "")
      MSFlexGrid1.TextMatrix(mI, 17) = NVL(mRec!Pesos, "")
      MSFlexGrid1.TextMatrix(mI, 18) = NVL(mRec!Materiales, "")
      MSFlexGrid1.TextMatrix(mI, 19) = NVL(mRec!Observaciones, "")
      MSFlexGrid1.TextMatrix(mI, 20) = NVL(mRec!estado, "")
      mRec.MoveNext
   Loop
   MSFlexGrid1.RemoveItem 1
End If
mRec.Close

Set mRec = mObj.oEjecutarSelect("SELECT * FROM ManoObra WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      List1.AddItem mRec!Codigo & "-" & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Text1(0).Enabled = False
Text1(1).Enabled = False
'Combo1(0).Enabled = False
'Text1(5).Enabled = False
Combo1(1).Enabled = False
Text1(7).Enabled = False
Text1(8).Enabled = False
Text1(9).Enabled = False
Text1(10).Enabled = False
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 20, True, False
End Sub

Private Function fValida1() As Boolean
Dim mRet As Boolean
Dim mI As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = (Combo1(2).Text <> "")
   If mRet Then
      mRet = (Combo1(3).Text <> "")
   End If
   If mRet Then
      mRet = (Text1(6).Text <> "")
   End If
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida1 = mRet
End Function

Private Sub List1_Click()
Dim mI As Integer
Dim mCant As Integer
If Text1(2).Text <> "" Then
   mCant = 0
   For mI = 0 To List1.ListCount - 1
      If List1.Selected(mI) Then
         mCant = mCant + 1
      End If
   Next
   Text1(10).Text = mCant * mObj.ObtCostoMO(Right(Text1(2).Text, 4) & Mid(Text1(2).Text, 4, 2))
End If
End Sub

Private Sub MSFlexGrid1_Click()
Dim mI As Integer
Dim mJ As Integer
Dim mPos As Integer

Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False


'If (MSFlexGrid1.MouseCol = 0 Or MSFlexGrid1.MouseCol = 1) And MSFlexGrid1.MouseRow > 0 Then
If (MSFlexGrid1.MouseCol = 0) And MSFlexGrid1.MouseRow > 0 And MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2) <> "" Then

   mRenglonElegido = MSFlexGrid1.MouseRow
   mColumnaElegida = MSFlexGrid1.MouseCol
   mParteElegido = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)



   mRenglon = MSFlexGrid1.MouseRow
   Text1(0).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   Text1(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3)
   Text1(2).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
   Text1(3).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5)
   Text1(4).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6)
   Text1(5).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8)
   For mI = 0 To Combo1(0).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7) = Combo1(0).List(mI) Then
         Combo1(0).ListIndex = mI
      End If
   Next
   For mI = 0 To Combo1(1).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = Combo1(1).List(mI) Then
         Combo1(1).ListIndex = mI
      End If
   Next
   For mI = 0 To Combo1(2).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = Combo1(2).List(mI) Then
         Combo1(2).ListIndex = mI
      End If
   Next
   For mI = 0 To Combo1(3).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11) = Combo1(3).List(mI) Then
         Combo1(3).ListIndex = mI
      End If
   Next
   Text1(6).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12)
   Text1(7).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13)
   Text1(8).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 14)

   'Blanqueo la Lista de mano de obra
   For mJ = 0 To List1.ListCount - 1
      List1.Selected(mJ) = False
   Next
   'Verifico a quien tengo que tildar
   mPos = 1
   For mI = 1 To ContarChar(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15), "-")
      For mJ = 0 To List1.ListCount - 1
         If Mid(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15), mPos, InStr(mPos, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15), "-") - mPos) = Left(List1.List(mJ), InStr(1, List1.List(mJ), "-") - 1) Then
            List1.Selected(mJ) = True
         End If
      Next
      mPos = InStr(mPos, MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 15), "-") + 1
   Next
   List1.ListIndex = -1

   Text1(9).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 16)
   Text1(10).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 17)
   Text1(11).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 18)
   Text1(12).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 19)
Else
   mRenglon = 0
End If


End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 2
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 3, 4
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 6, 7, 8, 9, 10
      KeyAscii = fNumDoubleKeyPress(KeyAscii)
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Index = 2 Then
      If Text1(1).Text <> "" And Text1(2).Text <> "" Then
         If IsDate(Text1(2).Text) Then
            Text1(7).Text = DateDiff("d", CDate(Text1(1).Text), CDate(Text1(2).Text))
         Else
            MsgBox "La fecha de asistencia es inválida", vbCritical, "Atención"
            Text1(Index).SetFocus
            Exit Sub
         End If
      End If
End If
'If Index = 3 Then
'   If Not (Text1(3).Text <> "" And IsDate(Text1(3).Text)) Then
'      MsgBox "La hora inicial de asistencia es inválida", vbCritical, "Atención"
'      Text1(Index).SetFocus
'      Exit Sub
'   End If
'End If
If Index = 4 Then
   If Text1(4).Text <> "" And IsDate(Text1(4).Text) And Text1(3).Text <> "" And IsDate(Text1(3).Text) Then
      Text1(9).Text = DateDiff("n", CDate(Text1(3).Text), CDate(Text1(4).Text)) / 60
   Else
      If Not (IsDate(Text1(3).Text)) Then
         MsgBox "La hora inicial de asistencia es inválida", vbCritical, "Atención"
         Text1(3).SetFocus
         Exit Sub
      ElseIf Not (IsDate(Text1(4).Text)) Then
         MsgBox "La hora final de asistencia es inválida", vbCritical, "Atención"
         Text1(4).SetFocus
         Exit Sub
      End If
   End If
End If
If Index = 6 Then
   Text1(8).Text = Val(Text1(6).Text) * 1.5
End If
End Sub

Private Function fValida2() As Boolean
Dim mRet As Boolean
Dim mI As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = Fecha_ok(Text1(2).Text)
   If mRet Then
      mRet = Hora_ok(Text1(3).Text)
   End If
   If mRet Then
      mRet = Hora_ok(Text1(4).Text)
   End If
   For mI = 2 To Combo1.UBound
      If mRet Then
         mRet = (Combo1(mI).Text <> "")
      End If
   Next
   For mI = 6 To Text1.UBound - 3
      If mRet Then
         mRet = (Text1(mI).Text <> "")
      End If
   Next
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida2 = mRet
End Function



Private Sub MSFlexGrid1_DblClick()
Dim mI As Integer
Dim mJ As Integer
Dim mParte As Double
Dim mMotivo As String
Dim mRowAnterior As Long
Dim iRow As Long
Dim iCol As Integer


If MSFlexGrid1.Row > 0 And MSFlexGrid1.TextMatrix(1, 2) <> "" Then
   If MSFlexGrid1.Col = 1 Then
      Frame3.Visible = True
      Label5.Caption = "Parte " & mParteElegido 'MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
      Frame3.Top = MSFlexGrid1.CellTop + 505
   End If
End If
End Sub


Private Function fValida3() As Boolean
Dim mRet As Boolean
mRet = Fecha_ok(Text2(0).Text)
If mRet Then
   mRet = Hora_ok(Text2(1).Text)
End If
fValida3 = mRet
End Function

Sub limpiarOpciones()
Option1 = False
Option2 = False
End Sub

Sub limpiarCampos()
   Dim mI As Integer

   '--Blanqueo de textsBox
   For mI = 0 To Text1.Count - 1
      Text1(mI).Text = ""
   Next
   
   '--Blanqueo de Combos
   For mI = 0 To Combo1.Count - 1
      Combo1(mI).ListIndex = -1
   Next
   
   '--Blanqueo de la lista de mano de obra
   For mI = 0 To List1.ListCount - 1
      List1.Selected(mI) = False
   Next
End Sub

