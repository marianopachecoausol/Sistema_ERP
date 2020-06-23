VERSION 5.00
Begin VB.Form RAcc1_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulo de Fichas de Accidentes de Tránsito."
   ClientHeight    =   13875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   16143.1
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Probables Causas"
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
      Index           =   5
      Left            =   120
      TabIndex        =   108
      Top             =   11280
      Width           =   11415
      Begin VB.Frame Frame3 
         Height          =   1815
         Left            =   9600
         TabIndex        =   111
         Top             =   360
         Width           =   1575
         Begin VB.CommandButton Command2 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   240
            TabIndex        =   89
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Siguiente"
            Default         =   -1  'True
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Del  Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   18
         Left            =   5640
         TabIndex        =   110
         Top             =   360
         Width           =   3735
         Begin VB.OptionButton Option14 
            Caption         =   "Sin Freno"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   86
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Otras"
            Height          =   255
            Index           =   4
            Left            =   2520
            TabIndex        =   87
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Reventón de Neumático"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   85
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Fallas Mecánica"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   84
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Sin Luces"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   83
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Del Conductor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   17
         Left            =   360
         TabIndex        =   109
         Top             =   360
         Width           =   4935
         Begin VB.ComboBox Combo3 
            Height          =   315
            Index           =   2
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   1320
            Width           =   3855
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Index           =   1
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   840
            Width           =   3855
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Index           =   0
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   360
            Width           =   3855
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Width           =   11415
      Begin VB.Frame Frame1 
         Caption         =   "Demarcación Vertical"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   11
         Left            =   7680
         TabIndex        =   0
         Top             =   240
         Width           =   3615
         Begin VB.OptionButton Option12 
            Caption         =   "Inexistente"
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   55
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Existente"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   54
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Demarcación Horizontal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   12
         Left            =   7680
         TabIndex        =   107
         Top             =   1560
         Width           =   3615
         Begin VB.OptionButton Option11 
            Caption         =   "Inexistente"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   57
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Existente"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   56
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
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
         Height          =   2655
         Index           =   14
         Left            =   3960
         TabIndex        =   104
         Top             =   2880
         Width           =   2295
         Begin VB.OptionButton Option13 
            Caption         =   "De Noche sin Iluminación"
            Height          =   615
            Index           =   2
            Left            =   600
            TabIndex        =   71
            Top             =   1680
            Width           =   1575
         End
         Begin VB.OptionButton Option13 
            Caption         =   "De Noche con Iluminación"
            Height          =   435
            Index           =   1
            Left            =   600
            TabIndex        =   70
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton Option13 
            Caption         =   "De Día"
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   69
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado de Banquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   16
         Left            =   9360
         TabIndex        =   106
         Top             =   2880
         Width           =   1935
         Begin VB.OptionButton Option10 
            Caption         =   "Malo"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   79
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Regular"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   78
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Bueno"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   77
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado de Calzada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   15
         Left            =   6720
         TabIndex        =   105
         Top             =   2880
         Width           =   2175
         Begin VB.OptionButton Option9 
            Caption         =   "Otro"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   76
            Top             =   2040
            Width           =   855
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Seco"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   75
            Top             =   1680
            Width           =   855
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Mojado"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   74
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Húmedo"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   73
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option9 
            Caption         =   "En reparación"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   72
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Condiciones Climáticas"
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
         Index           =   10
         Left            =   3000
         TabIndex        =   103
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton Option7 
            Caption         =   "Viento"
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   53
            Top             =   1920
            Width           =   1095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Polvo"
            Height          =   255
            Index           =   6
            Left            =   2760
            TabIndex        =   52
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Humo"
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   51
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Escarcha o Nieve"
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   49
            Top             =   1920
            Width           =   1695
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Granizo"
            Height          =   255
            Index           =   4
            Left            =   2760
            TabIndex        =   50
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Lluvia"
            Height          =   255
            Index           =   2
            Left            =   600
            TabIndex        =   48
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Neblina"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   47
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Buen Tiempo"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   46
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lugar del Accidente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Index           =   13
         Left            =   120
         TabIndex        =   102
         Top             =   2880
         Width           =   3375
         Begin VB.OptionButton Option6 
            Caption         =   "Otro"
            Height          =   255
            Index           =   10
            Left            =   2040
            TabIndex        =   68
            Top             =   1920
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Peaje"
            Height          =   255
            Index           =   9
            Left            =   2040
            TabIndex        =   67
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Colectora"
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   66
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Desvío"
            Height          =   255
            Index           =   7
            Left            =   2040
            TabIndex        =   65
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Salida"
            Height          =   255
            Index           =   6
            Left            =   2040
            TabIndex        =   64
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Acceso"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   63
            Top             =   2160
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Recta"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   62
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Curva"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   61
            Top             =   1440
            Width           =   975
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Túnel"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   60
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Puente"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   59
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Intersección"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sentido del tránsito"
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
         Index           =   9
         Left            =   120
         TabIndex        =   101
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton Option5 
            Caption         =   "Se Ignora"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   45
            Top             =   1920
            Width           =   1095
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Progr. Descendente"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   44
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Progr. Ascendente"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   43
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Ambos"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   42
            Top             =   480
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
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
      Height          =   3135
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   11415
      Begin VB.Frame Frame1 
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
         Height          =   2775
         Index           =   8
         Left            =   8280
         TabIndex        =   3
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton Option4 
            Caption         =   "Otro"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   41
            Top             =   2280
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Caída de ocupante"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   40
            Top             =   1920
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Salida de la Vía"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   39
            Top             =   1560
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Atropello de Peatón"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   38
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Atropello de Ciclista"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   37
            Top             =   840
            Width           =   1935
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Vuelco"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Colisión Contra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   7
         Left            =   2640
         TabIndex        =   7
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   2
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1440
            Width           =   3735
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   1
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   960
            Width           =   3735
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Index           =   0
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Con Otro Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   6
         Left            =   120
         TabIndex        =   100
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton Option3 
            Caption         =   "En Cadena"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   32
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Lateral"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   31
            Top             =   1440
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "En Ángulo"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   30
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "De Cola"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Frontal"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   28
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   11415
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   8880
         TabIndex        =   27
         Top             =   600
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         Caption         =   "Carril"
         Height          =   735
         Index           =   0
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   4095
         Begin VB.CheckBox Check2 
            Caption         =   "E.V."
            Height          =   255
            Index           =   6
            Left            =   3240
            TabIndex        =   26
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "CC"
            Height          =   255
            Index           =   5
            Left            =   2640
            TabIndex        =   25
            Top             =   360
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "B"
            Height          =   255
            Index           =   4
            Left            =   2160
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
         Begin VB.CheckBox Check2 
            Caption         =   "4"
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Width           =   375
         End
         Begin VB.CheckBox Check2 
            Caption         =   "3"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
         Begin VB.CheckBox Check2 
            Caption         =   "2"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   21
            Top             =   360
            Width           =   375
         End
         Begin VB.CheckBox Check2 
            Caption         =   "1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tramo OE"
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tramo R5"
         Height          =   195
         Index           =   1
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Autoridades de G.C.O."
         Height          =   195
         Left            =   9360
         TabIndex        =   4
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Progresiva"
         Height          =   195
         Left            =   240
         TabIndex        =   99
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Intersección"
         Height          =   195
         Left            =   240
         TabIndex        =   98
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   94
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   4680
         MaxLength       =   5
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Autopista del Oeste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   600
         TabIndex        =   113
         Top             =   360
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Hora Intervención"
         Height          =   435
         Left            =   3600
         TabIndex        =   97
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hora Aviso"
         Height          =   195
         Left            =   3600
         TabIndex        =   96
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   95
         Top             =   1005
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   6000
      TabIndex        =   8
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   "Buscar"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   115
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Todas"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   114
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H80000018&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Index           =   2
         Left            =   3960
         TabIndex        =   15
         Text            =   "88888"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         Height          =   195
         Left            =   3960
         TabIndex        =   93
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Móvil"
         Height          =   255
         Left            =   240
         TabIndex        =   92
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Policía"
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Patrullero"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "RAcc1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mObj As New clRAcc
Dim mRec As New ADODB.Recordset
Public mNroOrden As String
Public Cont As Integer
Public mBusca As Boolean
Dim mI As Integer
Public mConn As New ADODB.Connection
Dim mRec1 As New ADODB.Recordset

Private Sub Form_Load()
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click()
Dim mFlag As Boolean
   mFlag = True
   If mBusca Then
      If Combo4.Text = "" Then
         mFlag = False
         MsgBox "Seleccione un Número de Orden.", vbInformation, sMessage
      Else
         RAcc2_frm.Text1(0).Text = mNroOrden
      End If
   End If
   If mFlag Then
      RAcc1_frm.Visible = False
      RAcc2_frm.Show
      RAcc2_frm.Top = 0
      RAcc2_frm.Left = 0
   End If
End Sub

Private Sub Command2_Click()
   Unload RAcc2_frm
   Unload RAcc3_frm
   Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
   If Index = 0 Then
      Combo4.Clear
      Set mRec = mObj.oTabla("Ficha", "where fecha < '2008-03-01' order by 1")
      Do While Not mRec.EOF
        Combo4.AddItem "" & mRec!nroorden & " - " & mRec!Fecha & ""
        mRec.MoveNext
      Loop
      mRec.Close
      
   Else
      MsgBox "Opción anulada, buscar desde el módulo nuevo", vbInformation, sMessage
   End If
End Sub

Private Sub Option3_DblClick(Index As Integer) 'CON OTRO VEHICULO
   Option3(Index).Value = Not Option3(Index).Value
End Sub

Private Sub Option4_DblClick(Index As Integer) 'Tipo de Accidentes, OTROS
   Option4(Index).Value = Not Option4(Index).Value
End Sub

Private Sub Option14_DblClick(Index As Integer) 'Causas del Vehículo
   Option14(Index).Value = Not Option14(Index).Value
End Sub

Private Sub Combo2_Click(Index As Integer)
   If Index = 0 Then
      Combo2(1).Enabled = True
   Else
      Combo2(2).Enabled = True
   End If
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
     Combo2(Index).ListIndex = -1
   End If
End Sub

Private Sub Combo3_Click(Index As Integer)
   If Index = 0 Then
      Combo3(1).Enabled = True
   Else
      Combo3(2).Enabled = True
   End If
End Sub

Private Sub Combo3_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
     Combo3(Index).ListIndex = -1
   End If
End Sub

Private Sub Combo4_Click() 'BUSCAR
Dim mI As Integer
Dim mJ As String

   If Combo4.Text <> "" Then
      '-BEGIN****** LIMPIA TODO *********
      For mI = 0 To Check2.UBound
         Check2(mI).Value = 0
      Next
      Combo2(0).ListIndex = -1
      Combo2(1).ListIndex = -1
      Combo2(2).ListIndex = -1
      Combo1.ListIndex = -1
      Combo3(0).ListIndex = -1
      Combo3(1).ListIndex = -1
      Combo3(2).ListIndex = -1
      sOptionOff Option1
      sOptionOff Option3
      sOptionOff Option4
      sOptionOff Option5
      sOptionOff Option6
      sOptionOff Option7
      sOptionOff Option9
      sOptionOff Option10
      sOptionOff Option11
      sOptionOff Option12
      sOptionOff Option13
      sOptionOff Option14
      '-END******LIMPIA TODO ************
      mNroOrden = Left(Combo4.Text, 5) 'Nro de Ficha "Publica"
      Text2(2).Text = mNroOrden
      Set mRec = mObj.oTabla("Ficha", "where NroOrden = '" & mNroOrden & "'")
      If Not mRec.EOF Then
         Text1(0).Text = mRec!Fecha
         Text1(1).Text = NVL(mRec!Hora, 0)
         Text1(2).Text = mRec!HoraLlegada
         Text2(0).Text = mRec!POLICIA
         Text2(1).Text = mRec!MovilNro
         Text3(0).Text = mRec!PROGRESIVA
         Text3(1).Text = mRec!Interseccion
         If mRec!Tramo <> "" Then
            mI = mRec!Tramo
            Option1(mI - 1).Value = True
         End If
         If mRec!carril <> "" Then
            For mI = 1 To Len(mRec!carril) Step 2
               mJ = mId(mRec!carril, mI, 2)
               Check2(mJ - 1).Value = 1
            Next
         End If
         Text4.Text = mRec!AutoridGCO
         If mRec!AcciconOtro <> "" Then
            mI = mRec!AcciconOtro
            Option3(mI - 1).Value = True
         End If
         If mRec!CodColisContra1 <> "" Then
            mI = mRec!CodColisContra1
            Combo2(0).ListIndex = (mI - 1)
         End If
         If mRec!CodColisContra2 <> "" Then
            mI = mRec!CodColisContra2
            Combo2(1).ListIndex = (mI - 1)
         End If
         If mRec!CodColisContra3 <> "" Then
            mI = mRec!CodColisContra3
            Combo2(2).ListIndex = (mI - 1)
         End If
         For mI = 0 To Combo1.ListCount - 1
            If Left(Combo1.List(mI), 3) = mRec!CodPatrullero Then
               Combo1.ListIndex = mI
            End If
         Next
         If mRec!AccidOtro <> "" Then
            mI = mRec!AccidOtro
            Option4(mI - 1).Value = True
         End If
         mI = mRec!SentidoTrans
         Option5(mI - 1).Value = True
         mI = mRec!lugaraccid
         Option6(mI - 1).Value = True
         mI = mRec!Clima1
         Option7(mI - 1).Value = True
         mI = mRec!EstCalzada
         Option9(mI - 1).Value = True
         mI = mRec!EstBanquina
         Option10(mI - 1).Value = True
         mI = mRec!DemarcHoriz
         Option12(mI - 1).Value = True
         mI = mRec!DemarcVert
         Option11(mI - 1).Value = True
         mI = mRec!Iluminac
         Option13(mI - 1).Value = True
         If mRec!CodCausaCond1 <> "" Then
            mI = mRec!CodCausaCond1
            Combo3(0).ListIndex = (mI - 1)
         End If
         If mRec!CodCausaCond2 <> "" Then
            mI = mRec!CodCausaCond2
            Combo3(1).ListIndex = (mI - 1)
         End If
         If mRec!CodCausaCond3 <> "" Then
            mI = mRec!CodCausaCond3
            Combo3(2).ListIndex = (mI - 1)
         End If
         If mRec!causaVehic <> "" Then
            mI = mRec!causaVehic
            Option14(mI - 1).Value = True
         End If
         Command1.Enabled = True
      End If
      mRec.Close
      Unload RAcc2_frm
      Unload RAcc3_frm
   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Else
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
 If Index = 1 Then
    KeyAscii = fNumeroKeyPress(KeyAscii)
 Else
    If KeyAscii >= 97 And KeyAscii <= 122 Then
       KeyAscii = KeyAscii - 32
    Else
       If KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 44 Then
          KeyAscii = 0
       End If
    End If
 End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = fAlfaNumKeyPress(KeyAscii)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
   Else
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   End If
End Sub

Private Sub sInitForm()
   Me.Height = 14250
   Me.Top = 0
   Me.Left = 0
   Combo2(1).Enabled = False
   Combo2(2).Enabled = False
   Combo3(1).Enabled = False
   Combo3(2).Enabled = False
   Cont = 0
   Set mRec = mObj.oTabla("Patrullero", "order by nombre")
   sLlenoCbo Me.Combo1, mRec, 1, 0
   Set mRec = mObj.oTabla("ColisionContra", "order by CodColision")
   Do While Not mRec.EOF
      For mI = 0 To Combo2.UBound
         Combo2(mI).AddItem mRec!Descripcion & Space(40) & mRec!CodColision
      Next
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTabla("CausaConductor", "order by CodCausaCond")
   Do While Not mRec.EOF
        For mI = 0 To Combo3.UBound
           Combo3(mI).AddItem mRec!Descripcion & Space(40) & mRec!CodCausacond
        Next
        mRec.MoveNext
   Loop
   mRec.Close
   If Not mBusca Then
      Text2(2).Text = mObj.sTablaDescr("Auxiliar", "1=1", 0)
   Else
       Text2(2).Visible = False
       Combo4.Enabled = True
       Combo4.Visible = True
       Label4.FontSize = 14
       Label4.Top = 260
       Command3(0).Visible = True
       Command3(1).Visible = True
       Me.Caption = Me.Caption & "  ---MODO VISTA Y MODIFICACIÓN--"
   End If
End Sub

Private Sub sOptionOff(ByRef pObjOpt As Object)
Dim mI As Integer
   For mI = 0 To pObjOpt.UBound
      pObjOpt(mI).Value = False
   Next
End Sub

Private Function fValid() As Boolean
   fValid = Fecha_ok(Text1(0).Text)
   fValid = fValid And Hora_ok(Text1(1).Text)
   fValid = fValid And Hora_ok(Text1(2).Text)
   fValid = fValid And Progresiva_Ok(Text3(0).Text)
   fValid = fValid And fValidTxt(Combo1, "Patrullero")
   fValid = fValid And fValidTxt(Text2(1), "Móvil")
   fValid = fValid And fValidOption(Option5, "Sentido de Tránsito")
   fValid = fValid And (fValidOption(Option3, "al menos un Item en Tipo de Accidentes") Or fValidOption(Option4, "Tipo de Accidente") Or fValidTxt(Combo2(0), "Tipo de Accidente"))
   fValid = fValid And fValidOption(Option6, "Lugar del Accidente")
   fValid = fValid And fValidOption(Option7, "Condiciones Climáticas")
   fValid = fValid And fValidOption(Option9, "Estado de Calzada")
   fValid = fValid And fValidOption(Option10, "Estado de Banquina")
   fValid = fValid And fValidOption(Option11, "Demarcación Horizontal")
   fValid = fValid And fValidOption(Option12, "Demarcación Vertical")
   fValid = fValid And fValidOption(Option13, "Iluminación")
End Function

Private Function fValidOption(ByRef pObj As Object, ByVal pTexto As String) As Boolean
   fValidOption = False
   For mI = 0 To pObj.UBound
       fValidOption = fValidOption Or pObj(mI).Value
   Next
   If Not fValidOption Then
      MsgBox "Falta seleccionar " & pTexto, vbExclamation, sMessage
   End If
End Function

Private Function fValidTxt(ByRef pObj As Object, ByVal pTexto As String) As Boolean
   fValidTxt = True
   If Trim(pObj.Text) = "" Then
      MsgBox "Falta seleccionar " & pTexto & ".", vbExclamation, sMessage
      fValidTxt = False
   End If
End Function
