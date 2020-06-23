VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form RAcc7_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Parametrizada"
   ClientHeight    =   11925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11925
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame14 
      Height          =   4815
      Left            =   120
      TabIndex        =   90
      Top             =   7080
      Width           =   11295
      Begin VB.Frame Frame34 
         Caption         =   "Causas del Conductor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   360
         TabIndex        =   131
         Top             =   120
         Width           =   7815
         Begin VB.CheckBox Check11 
            Caption         =   "Todos"
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
            Left            =   6720
            TabIndex        =   144
            Top             =   1320
            Width           =   855
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Otro"
            Height          =   255
            Index           =   11
            Left            =   4920
            TabIndex        =   143
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Escaza Veloc.Carril Izq."
            Height          =   195
            Index           =   10
            Left            =   4920
            TabIndex        =   142
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Falta de Visibilidad"
            Height          =   195
            Index           =   9
            Left            =   4920
            TabIndex        =   141
            Top             =   720
            Width           =   1815
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Maniobra Equivoc."
            Height          =   195
            Index           =   8
            Left            =   4920
            TabIndex        =   140
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Adelantam. Indebido"
            Height          =   195
            Index           =   7
            Left            =   2640
            TabIndex        =   139
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Violación de Señal"
            Height          =   195
            Index           =   6
            Left            =   2640
            TabIndex        =   138
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Competencia"
            Height          =   195
            Index           =   5
            Left            =   2640
            TabIndex        =   137
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Encandilamiento"
            Height          =   195
            Index           =   4
            Left            =   2640
            TabIndex        =   136
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Distracción"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   135
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Exceso de Velocidad"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   134
            Top             =   960
            Width           =   2055
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Se Durmió o Desvaneció"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   133
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Ebriedad o Drogas"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   132
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame Frame23 
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
         Height          =   1695
         Left            =   8640
         TabIndex        =   125
         Top             =   120
         Width           =   2415
         Begin VB.CheckBox Check12 
            Caption         =   "Otra"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   130
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Sin Freno"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   129
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Reventón Neumát."
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   128
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Fallas Mecánicas"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   127
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Sin Luces"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   126
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Tipo Vehículo"
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
         Left            =   360
         TabIndex        =   109
         Top             =   1920
         Width           =   5655
         Begin VB.CheckBox Check13 
            Caption         =   "Todos"
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
            Left            =   3720
            TabIndex        =   124
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Se Ignora"
            Height          =   195
            Index           =   13
            Left            =   3720
            TabIndex        =   123
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Automóvil"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   122
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Tren"
            Height          =   195
            Index           =   12
            Left            =   3720
            TabIndex        =   121
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Tracción Animal"
            Height          =   195
            Index           =   11
            Left            =   3720
            TabIndex        =   120
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Tractor/Maquinan"
            Height          =   195
            Index           =   10
            Left            =   3720
            TabIndex        =   119
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Semi Remolque"
            Height          =   195
            Index           =   9
            Left            =   1920
            TabIndex        =   118
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Camión c/Acopl"
            Height          =   195
            Index           =   8
            Left            =   1920
            TabIndex        =   117
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Camión Simple"
            Height          =   195
            Index           =   7
            Left            =   1920
            TabIndex        =   116
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Micro Omnib"
            Height          =   195
            Index           =   6
            Left            =   1920
            TabIndex        =   115
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Omnibus"
            Height          =   195
            Index           =   5
            Left            =   1920
            TabIndex        =   114
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Camioneta/Jeep"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   113
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Motocicleta"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   112
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Ciclomotor"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   111
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Bicicleta"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   110
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Fallecidos"
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
         Left            =   6360
         TabIndex        =   105
         Top             =   1920
         Width           =   1815
         Begin VB.CheckBox Check14 
            Caption         =   "Después"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   108
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox Check14 
            Caption         =   "Durante el Trasl"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   107
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox Check14 
            Caption         =   "En el Lugar"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   106
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Heridos"
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
         Left            =   8520
         TabIndex        =   102
         Top             =   1920
         Width           =   1335
         Begin VB.CheckBox Check15 
            Caption         =   "Graves"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   104
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check15 
            Caption         =   "Leves"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   103
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Cinturón"
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
         Left            =   9960
         TabIndex        =   99
         Top             =   1920
         Width           =   1095
         Begin VB.CheckBox Check16 
            Caption         =   "No"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   101
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox Check16 
            Caption         =   "Si"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   100
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8520
         TabIndex        =   96
         Top             =   3120
         Width           =   2535
         Begin VB.CheckBox Check17 
            Caption         =   "Femenino"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   98
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Masculino"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   97
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Comp. de Seguros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   94
         Top             =   3720
         Width           =   3255
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame15 
         Height          =   855
         Left            =   5160
         TabIndex        =   91
         Top             =   3840
         Width           =   4335
         Begin VB.CommandButton Command2 
            Caption         =   "&Cancelar"
            Height          =   495
            Left            =   2520
            TabIndex        =   93
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   495
            Left            =   840
            TabIndex        =   92
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   240
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame26 
         Caption         =   "Horario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9000
         TabIndex        =   152
         Top             =   240
         Width           =   2055
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   156
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   5
            TabIndex        =   155
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   1200
            TabIndex        =   154
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   240
            TabIndex        =   153
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5880
         TabIndex        =   147
         Top             =   240
         Width           =   2775
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   151
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   10
            TabIndex        =   149
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   1560
            TabIndex        =   150
            Top             =   240
            Width           =   330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   240
            TabIndex        =   148
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame24 
         Height          =   855
         Left            =   360
         TabIndex        =   145
         Top             =   240
         Width           =   5055
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Consulta Parametrizada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            TabIndex        =   146
            Top             =   240
            Width           =   2970
         End
      End
      Begin VB.Frame Frame16 
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
         Height          =   615
         Left            =   360
         TabIndex        =   88
         Top             =   1320
         Width           =   2655
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Est. de Banquina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5640
         TabIndex        =   84
         Top             =   6120
         Width           =   2775
         Begin VB.CheckBox Check9 
            Caption         =   "Malo"
            Height          =   195
            Index           =   2
            Left            =   1920
            TabIndex        =   87
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Regular"
            Height          =   195
            Index           =   1
            Left            =   960
            TabIndex        =   86
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Bueno"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   85
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Est. de Calzada"
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
         Left            =   5640
         TabIndex        =   78
         Top             =   4920
         Width           =   2775
         Begin VB.CheckBox Check8 
            Caption         =   "Otro"
            Height          =   195
            Index           =   4
            Left            =   1680
            TabIndex        =   83
            Top             =   600
            Width           =   615
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Bueno"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   82
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Resbaladizos"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   81
            Top             =   840
            Width           =   1335
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Baches"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   80
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox Check8 
            Caption         =   "En Reparación"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   79
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
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
         Height          =   2055
         Left            =   8640
         TabIndex        =   74
         Top             =   5040
         Width           =   2415
         Begin VB.CheckBox Check10 
            Caption         =   "De Noche s/Ilumin"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   77
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox Check10 
            Caption         =   "De Noche c/Ilumin"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   76
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox Check10 
            Caption         =   "De Día"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   75
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Lugar Accidente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   62
         Top             =   5640
         Width           =   5055
         Begin VB.CheckBox Check18 
            Caption         =   "Otro"
            Height          =   195
            Index           =   10
            Left            =   3600
            TabIndex        =   73
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Peaje"
            Height          =   195
            Index           =   9
            Left            =   3600
            TabIndex        =   72
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Colectora"
            Height          =   195
            Index           =   8
            Left            =   3600
            TabIndex        =   71
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Desvío"
            Height          =   195
            Index           =   7
            Left            =   1800
            TabIndex        =   70
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Salida"
            Height          =   195
            Index           =   6
            Left            =   1800
            TabIndex        =   69
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Acceso"
            Height          =   195
            Index           =   5
            Left            =   1800
            TabIndex        =   68
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Recta"
            Height          =   195
            Index           =   4
            Left            =   1800
            TabIndex        =   67
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Curva"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   66
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Túnel"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   65
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Puente"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Check18 
            Caption         =   "Intersección"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Demarcación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5640
         TabIndex        =   55
         Top             =   4080
         Width           =   5415
         Begin VB.CheckBox Check19 
            Caption         =   "Inexistente"
            Height          =   195
            Index           =   1
            Left            =   3720
            TabIndex        =   61
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox Check19 
            Caption         =   "Existente"
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   60
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Inexistente"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   59
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Existente"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Vertical"
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
            Left            =   2640
            TabIndex        =   57
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Horizontal"
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
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   1575
         Left            =   2400
         TabIndex        =   46
         Top             =   4080
         Width           =   3015
         Begin VB.CheckBox Check6 
            Caption         =   "Viento"
            Height          =   195
            Index           =   7
            Left            =   1920
            TabIndex        =   54
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Polvo"
            Height          =   195
            Index           =   6
            Left            =   1920
            TabIndex        =   53
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Humo"
            Height          =   195
            Index           =   5
            Left            =   1920
            TabIndex        =   52
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Granizo"
            Height          =   195
            Index           =   4
            Left            =   1920
            TabIndex        =   51
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Escarcha o Nieve"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   50
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Lluvia"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   49
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Neblina"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   48
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Buen  Tiempo"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Sent. Tránsito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   360
         TabIndex        =   41
         Top             =   4080
         Width           =   1575
         Begin VB.CheckBox Check5 
            Caption         =   "Se Ignora"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   45
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Progr. Desc."
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   44
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Progr. Asc."
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Ambos"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Otros - Tipo de Accid."
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
         Left            =   8760
         TabIndex        =   34
         Top             =   2040
         Width           =   2295
         Begin VB.CheckBox Check4 
            Caption         =   "Otro"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   40
            Top             =   1560
            Width           =   975
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Caída de Ocup"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   39
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Salida de Vía"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   38
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Atropello de Peat"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   37
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Atropello de Cicl"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Vuelco"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1935
         Left            =   2880
         TabIndex        =   20
         Top             =   2040
         Width           =   5415
         Begin VB.CheckBox Check3 
            Caption         =   "Todos"
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
            Left            =   4320
            TabIndex        =   33
            Top             =   1560
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Otros"
            Height          =   195
            Index           =   11
            Left            =   2880
            TabIndex        =   32
            Top             =   1560
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Árbol"
            Height          =   195
            Index           =   10
            Left            =   2880
            TabIndex        =   31
            Top             =   1320
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Peaje"
            Height          =   195
            Index           =   9
            Left            =   2880
            TabIndex        =   30
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Material S/Calzada"
            Height          =   195
            Index           =   8
            Left            =   2880
            TabIndex        =   29
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Columna o Poste"
            Height          =   195
            Index           =   7
            Left            =   2880
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Animal"
            Height          =   195
            Index           =   6
            Left            =   2880
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Señal Vial"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   26
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Puente o Alcantarilla"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   25
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Separador Central"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   24
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Vehic. Deten. S/Banquina"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   23
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Vehic. Deten. S/Calzada"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   600
            Width           =   2175
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Baranda Lateral"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Con Otro Vehic"
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
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   2055
         Begin VB.CheckBox Check2 
            Caption         =   "En Cadena"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   19
            Top             =   1440
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Lateral"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   18
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox Check2 
            Caption         =   "En Ángulo"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   17
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Frontal"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Progresiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   720
            MaxLength       =   6
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   1560
            TabIndex        =   11
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Carril"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   1
         Top             =   1320
         Width           =   4695
         Begin VB.CheckBox Check1 
            Caption         =   "Todo"
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
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "EV"
            Height          =   195
            Index           =   6
            Left            =   3960
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "CC"
            Height          =   195
            Index           =   5
            Left            =   3360
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox Check1 
            Caption         =   "B"
            Height          =   195
            Index           =   4
            Left            =   2880
            TabIndex        =   6
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "4"
            Height          =   195
            Index           =   3
            Left            =   2400
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "3"
            Height          =   195
            Index           =   2
            Left            =   1920
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "2"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
         Begin VB.CheckBox Check1 
            Caption         =   "1"
            Height          =   195
            Index           =   0
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "RAcc7_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObj As New clRAcc
Dim mObjAcc As New clAccess
Dim mRec As New ADODB.Recordset
Dim mI As Integer
Public mDatos As String

Private Sub Form_Load()
   Me.Height = 12400
   Me.Width = 11700
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   Combo1.AddItem "TODOS"
   Set mRec = mObj.oTabla("Patrullero", "order by nombre")
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!CodPatrullero & " -  " & mRec!nombre & ""
      mRec.MoveNext
   Loop
   mRec.Close
   Combo1.ListIndex = 0
   Combo2.AddItem "TODOS"
   Set mRec = mObj.oTabla("CiaSeguros", "order by descripcion")
   sLlenoCbo Me.Combo2, mRec, 1, 0
   Combo2.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mData.Close
   Set mData = Nothing
   Set mObj = Nothing
   Set mObjAcc = Nothing
   Set mRec = Nothing
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click()
Dim mCond As String
Dim mCond2 As String
Dim mCond3 As String
Dim mWhere As String
Dim mDesde As String
Dim mHasta As String
Dim mNroOrden1 As String
Dim mNroOrden2 As String
Dim mNroOrden3 As String
Dim mNroOrden4 As String
Dim Flag As Boolean
Dim Vector1(24) As Integer
Dim Vector2(24) As String
Dim mJ, Total As Integer
Dim mAuxi As Variant

   Flag = True
   If sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text) Then
      mDesde = Text1(0).Text 'Format(Text1(0).Text, "yyyy/mm/dd")
      mHasta = Text1(1).Text 'Format(Text1(1).Text, "yyyy/mm/dd")
      mWhere = " Ficha"
      Vector1(0) = 1
      Vector2(0) = "Período de la Consulta desde el " & mDesde & " hasta el " & mHasta & ""
      If Combo1.Text <> "TODOS" Then
         mCond = " AND Ficha.CodPatrullero = '" & Left(Combo1.Text, 3) & "'"
         Vector1(1) = 1
         Vector2(1) = "Patrullero: " & Combo1.Text
      End If
      If Progr_Ok(Text2(0), Text2(1)) Then
         If Hora_ok2(Text3(0), Text3(1)) Then 'Modificación Hora
            If Text2(0).Text <> "" And Text2(1).Text <> "" Then
               mCond = mCond & " AND Ficha.Progresiva >= " & Text2(0).Text & " AND Ficha.Progresiva <= " & Text2(1).Text & ""
               Vector1(2) = 1
               Vector2(2) = "Progresiva desde " & Text2(0).Text & " a " & Text2(1).Text & ""
            End If
            If Not (Check_Vacio(Check1) Or Check_Todo(Check1)) Then
                 mCond = mCond & Check_SQL(Check1, "Ficha.Carril")
                 Vector1(3) = 1
                 Vector2(3) = "Carriles: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check2) Or Check_Todo(Check2)) Then
                 mCond = mCond & Check_SQL(Check2, "Ficha.AcciConOtro")
                 Vector1(4) = 1
                 Vector2(4) = "Cod. Accidentes con Otro: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check3) Or Check_Todo(Check3)) Then
                 mCond = mCond & Check_SQL(Check3, "Ficha.CodColisContra1")
                 Vector1(5) = 1
                 Vector2(5) = "Cod Colisión Contra: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check4) Or Check_Todo(Check4)) Then
                 mCond = mCond & Check_SQL(Check4, "Ficha.AccidOtro")
                 Vector1(6) = 1
                 Vector2(6) = "Cod. Accidentes-Otros: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check5) Or Check_Todo(Check5)) Then
                 mCond = mCond & Check_SQL(Check5, "Ficha.SentidoTrans")
                 Vector1(7) = 1
                 Vector2(7) = "Cod. Sentido de Tránsito: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check6) Or Check_Todo(Check6)) Then
                 mCond = mCond & Check_SQL(Check6, "Ficha.Clima1")
                 Vector1(8) = 1
                 Vector2(8) = "Cod. Clima: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check7) Or Check_Todo(Check7)) Then
                 mCond = mCond & Check_SQL(Check7, "Ficha.DemarcHoriz")
                 Vector1(9) = 1
                 Vector2(9) = "Cod. Demarcación Horizontal: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check19) Or Check_Todo(Check19)) Then
                 mCond = mCond & Check_SQL(Check19, "Ficha.DemarcVert")
                 Vector1(10) = 1
                 Vector2(10) = "Cod. Demarcación Vertical: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check8) Or Check_Todo(Check8)) Then
                 mCond = mCond & Check_SQL(Check8, "Ficha.EstCalzada")
                 Vector1(11) = 1
                 Vector2(11) = "Cod. Estado de Calzada: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check9) Or Check_Todo(Check9)) Then
                 mCond = mCond & Check_SQL(Check9, "Ficha.EstBanquina")
                 Vector1(12) = 1
                 Vector2(12) = "Estado de Banquina: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check10) Or Check_Todo(Check10)) Then
                 mCond = mCond & Check_SQL(Check10, "Ficha.Iluminac")
                 Vector1(13) = 1
                 Vector2(13) = "Cod. Iluminación: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check18) Or Check_Todo(Check18)) Then
                 mCond = mCond & Check_SQL(Check18, "Ficha.LugarAccid")
                 Vector1(14) = 1
                 Vector2(14) = "Cod. Lugar del Accidente: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check11) Or Check_Todo(Check11)) Then
                 mCond = mCond & Check_SQL(Check11, "Ficha.CodCausaCond1")
                 Vector1(15) = 1
                 Vector2(15) = "Cod. Causas Conductor: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check12) Or Check_Todo(Check12)) Then
                 mCond = mCond & Check_SQL(Check12, "Ficha.CausaVehic")
                 Vector1(16) = 1
                 Vector2(16) = "Cod. de Causas del Vehículo: " & mDatos & ""
                 mDatos = ""
            End If
            
            If Not (Check_Vacio(Check13) Or Check_Todo(Check13)) Then
                 mCond2 = mCond2 & Check_SQL(Check13, "VehiculosInvolucr.CodTipoVehic")
                 Vector1(17) = 1
                 Vector2(17) = "Cod. Tipo de Vehículo: " & mDatos & ""
                 mDatos = ""
            End If
            If Combo2.Text <> "TODOS" Then
               mCond2 = mCond2 & " AND VehiculosInvolucr.CodCiaSeguro = '" & Right(Combo2.Text, 2) & "'"
               Vector1(18) = 1
               Vector2(18) = "Cía de Seguro: " & Right(Combo2.Text, 2) & " - " & Trim(Left(Combo2.Text, 30))
            End If
            If mCond2 <> "" Then
               mCond = mCond & " AND Ficha.NroOrden = VehiculosInvolucr.NroOrden" & mCond2
               mWhere = mWhere & ", VehiculosInvolucr"
            End If
            
            If Not (Check_Vacio(Check14) Or Check_Todo(Check14)) Then
                 mCond3 = mCond3 & Check_SQL(Check14, "VictimasInvolucr.Fallecio")
                 Vector1(19) = 1
                 Vector2(19) = "Cod. Fallecido: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check15) Or Check_Todo(Check15)) Then
                 mCond3 = mCond3 & Check_SQL(Check15, "VictimasInvolucr.Herido")
                 Vector1(20) = 1
                 Vector2(20) = "Cod. Herido: " & mDatos & ""
                 mDatos = ""
            End If
            If Not (Check_Vacio(Check16) Or Check_Todo(Check16)) Then
                mCond3 = mCond3 & Check_SQL(Check16, "VictimasInvolucr.Cinturon")
                If Right(mCond3, 3) = "1')" Then
                   mCond3 = Mid(mCond3, 1, (Len(mCond3) - 3))
                   mCond3 = mCond3 & "SI')"
                   mDatos = "SI"
                Else
                   mCond3 = Mid(mCond3, 1, (Len(mCond3) - 3))
                   mCond3 = mCond3 & "NO')"
                   mDatos = "NO"
                End If
                Vector1(21) = 1
                Vector2(21) = "Cinturón: " & mDatos & ""
                mDatos = ""
            End If
            If Not (Check_Vacio(Check17) Or Check_Todo(Check17)) Then
                 mCond3 = mCond3 & Check_SQL(Check17, "VictimasInvolucr.Sexo")
                 Vector1(22) = 1
                 Vector2(22) = "Cod. Sexo: " & mDatos & ""
                 mDatos = ""
            End If
            'Modificación Hora
            If Text3(0).Text <> "" And Text3(1).Text <> "" Then
               mCond = mCond & " AND Ficha.Hora >= '" & Text3(0).Text & "' AND Ficha.Hora <= '" & Text3(1).Text & "'"
               Vector1(23) = 1
               Vector2(23) = "Intervalo de Hora: " & Text3(0).Text & " a " & Text3(1).Text & ""
            End If
            'Fin Modificación
            If mCond3 <> "" Then
               mCond = mCond & " AND Ficha.NroOrden = VictimasInvolucr.NroOrden" & mCond3
               mWhere = mWhere & ", VictimasInvolucr"
            End If
            
            'Modificación Nro Ordenes
            Set mRec = mObj.oParametrizada(mDesde, mHasta, mWhere, mCond)
             If Not mRec.EOF Then
                mNroOrden1 = "Nro de Ordenes: "
                mJ = 0
                Do While Not mRec.EOF
                   Select Case mJ
                   Case Is < 19     'Antes 13, modificado por Diego. Sería bueno usar un vector...
                       mNroOrden1 = mNroOrden1 & mRec!nroorden & "-"
                       mJ = mJ + 1
                   Case Is < 37     'Antes 25, modificado por Diego
                        mNroOrden2 = mNroOrden2 & mRec!nroorden & "-"
                        mJ = mJ + 1
                   Case Is < 55     'Antes 37, modificado por Diego
                        mNroOrden3 = mNroOrden3 & mRec!nroorden & "-"
                        mJ = mJ + 1
                   Case Else
                        mNroOrden4 = mNroOrden4 & mRec!nroorden & "-"
                        mJ = mJ + 1
                   End Select
                   mRec.MoveNext
                Loop
                mRec.Close
             End If
             Total = mJ
             Flag = True
             mCond = ""
             mCond2 = ""
             mCond3 = ""
             mJ = 0
             For mI = 0 To 6
                If Vector1(mI) = 1 Then
                   If Flag Then
                      mCond = "'" & Vector2(mI) & "'"
                      mJ = mJ + 1
                      Flag = False
                   Else
                      mCond = mCond & ",'" & Vector2(mI) & "'"
                      mJ = mJ + 1
                   End If
                End If
             Next
             For mI = 7 To 15
                If Vector1(mI) = 1 Then
                      mCond2 = mCond2 & ",'" & Vector2(mI) & "'"
                      mJ = mJ + 1
                End If
             Next
             For mI = 16 To 23
                If Vector1(mI) = 1 Then
                      mCond3 = mCond3 & ",'" & Vector2(mI) & "'"
                      mJ = mJ + 1
                End If
             Next
             For mI = 1 To (24 - mJ)
               mCond3 = mCond3 & ",''"
             Next
             mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
             mData.Execute ("CREATE TABLE Auxi (C1 String,C2 String,C3 String,C4 String,C5 String,C6 String,C7 String,C8 String,C9 String,C10 String,C11 String,C12 String,C13 String,C14 String,C15 String,C16 String,C17 String,C18 String,C19 String,C20 String,C21 String,C22 String,C23 String,C24 String,C25 Integer,C26 String,C27 String,C28 String, C29 String)")
             mData.Execute ("INSERT INTO Auxi (C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13,C14,C15,C16,C17,C18,C19,C20,C21,C22,C23,C24,C25,C26,C27,C28,C29) VALUES (" & mCond & mCond2 & mCond3 & "," & Total & ",'" & mNroOrden1 & "','" & mNroOrden2 & "','" & mNroOrden3 & "','" & mNroOrden4 & "')")
             Set mAuxi = mData.OpenRecordset("SELECT * FROM Auxi")
             mAuxi.Close
             For mI = 0 To mData.TableDefs.Count - 1
                If mData.TableDefs(mI).Name = "Auxi" Then
                   CrystalReport1.WindowTitle = "Reporte Consulta Parametrizada "
                   CrystalReport1.DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
                   CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep19.rpt"
                   CrystalReport1.WindowState = crptMaximized
                   CrystalReport1.Action = 1
                   mI = 300
                End If
             Next
             'Limpia todo
             Text1(0).Text = ""
             Text1(1).Text = ""
             Text2(0).Text = ""
             Text2(1).Text = ""
             Text3(0).Text = ""
             Text3(1).Text = ""
             Combo1.ListIndex = 0
             Combo2.ListIndex = 0
             Check_Clear Check1
             Check_Clear Check2
             Check_Clear Check3
             Check_Clear Check4
             Check_Clear Check5
             Check_Clear Check6
             Check_Clear Check7
             Check_Clear Check8
             Check_Clear Check9
             Check_Clear Check10
             Check_Clear Check11
             Check_Clear Check12
             Check_Clear Check13
             Check_Clear Check14
             Check_Clear Check15
             Check_Clear Check16
             Check_Clear Check17
             Check_Clear Check18
             Check_Clear Check19
         End If
      End If
   End If
End Sub

Private Sub Check1_Click(Index As Integer)
   If Index = 7 Then
      For mI = 0 To Check1.UBound - 1
         Check1(mI).Value = Check1(Index).Value
      Next
   End If
End Sub

Private Sub Check3_Click(Index As Integer)
   If Index = 12 Then
      For mI = 0 To Check3.UBound - 1
         Check3(mI).Value = Check3(Index).Value
      Next
   End If
End Sub

Private Sub Check11_Click(Index As Integer)
   If Index = 12 Then
      For mI = 0 To Check11.UBound - 1
         Check11(mI).Value = Check11(Index).Value
      Next
   End If
End Sub

Private Sub Check13_Click(Index As Integer)
   If Index = 14 Then
      For mI = 0 To Check13.UBound - 1
         Check13(mI).Value = Check13(Index).Value
      Next
   End If
End Sub

Private Sub Command2_Click()
   Unload RAcc7_frm
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fHoraKeyPress(Text3(Index), KeyAscii)
End Sub

Public Function Check_SQL(mObj As Object, mCodigo As String) As String
Dim Flag As Boolean
Dim Cond As String
Dim mI, mJ As Integer
   If mCodigo = "CodColisContra1" Or mCodigo = "Carril" Or mCodigo = "CodTipoVehic" Then 'xxx
     mJ = 1
   Else
     mJ = 0
   End If
   Flag = True
   For mI = 0 To mObj.UBound - mJ
      If mObj(mI).Value = 1 Then
         If Left(mCodigo, 8) = "Victimas" Then
            If Flag Then
               Cond = " AND (" & mCodigo & " = '" & (mI + 1) & "'"
               mDatos = "" & (mI + 1) & ""
               Flag = False
            Else
               Cond = Cond & " OR " & mCodigo & " = '" & (mI + 1) & "'"
               mDatos = mDatos & " - " & (mI + 1) & ""
            End If
         Else
            If Flag Then
               Cond = " AND (" & mCodigo & " = '" & Format((mI + 1), "00") & "'"
               Flag = False
               mDatos = "" & Format((mI + 1), "00") & ""
            Else
               Cond = Cond & " OR " & mCodigo & " = '" & Format((mI + 1), "00") & "'"
               mDatos = mDatos & "-" & Format((mI + 1), "00") & ""
            End If
         End If
      End If
   Next
   Cond = Cond & ")"
   Check_SQL = Cond
End Function

Public Function Progr_Ok(mObj1 As Object, mObj2 As Object) As Boolean
Dim Flag As Boolean
Dim Num1, Num2 As Double
If mObj1.Text <> "" And mObj2.Text <> "" Then
   Num1 = mObj1.Text
   Num2 = mObj2.Text
   If (Num1 >= 12.95 And Num1 <= 65.14) And (Num2 >= 12.95 And Num2 <= 65.14) And Num1 <= Num2 Then
      Flag = True
   Else
      MsgBox "Error en Progresiva.", vbCritical, "Atención!"
      Flag = False
   End If
Else
   If mObj1.Text <> "" Or mObj2.Text <> "" Then
      MsgBox "Falta Completar una Progresiva", vbCritical, "Atención!"
      Flag = False
   Else
      Flag = True
   End If
End If
Progr_Ok = Flag
End Function

Public Function Check_Vacio(mObj As Object) As Boolean
   Check_Vacio = True
   For mI = 0 To mObj.UBound
      If mObj(mI).Value <> 0 Then
         Check_Vacio = False
      End If
   Next
End Function

Public Function Check_Todo(mObj As Object) As Boolean
   Check_Todo = True
   For mI = 0 To mObj.UBound
      If mObj(mI).Value <> 1 Then
          Check_Todo = False
      End If
   Next
End Function

Public Sub Check_Clear(mObj As Object)
   For mI = 0 To mObj.UBound
     mObj(mI).Value = 0
   Next
End Sub
