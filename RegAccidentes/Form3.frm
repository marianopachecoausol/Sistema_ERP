VERSION 5.00
Begin VB.Form RAcc3_frm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulo de Víctimas Involucradas."
   ClientHeight    =   12720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12720
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame12 
      Caption         =   "Intervinieron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   50
      Top             =   9120
      Width           =   11535
      Begin VB.CheckBox Check1 
         Caption         =   "Foto Común"
         Height          =   375
         Index           =   1
         Left            =   7320
         TabIndex        =   82
         Top             =   2640
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Foto Digital"
         Height          =   375
         Index           =   0
         Left            =   6360
         TabIndex        =   81
         Top             =   2640
         Width           =   735
      End
      Begin VB.Frame Frame13 
         Height          =   1695
         Left            =   9120
         TabIndex        =   77
         Top             =   1680
         Width           =   2055
         Begin VB.CommandButton Command2 
            Caption         =   "Siguiente"
            Height          =   615
            Index           =   1
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   -960
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Volver"
            Height          =   615
            Index           =   0
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   13
         Left            =   5880
         MaxLength       =   25
         TabIndex        =   76
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   12
         Left            =   8400
         MaxLength       =   25
         TabIndex        =   75
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   11
         Left            =   5880
         MaxLength       =   25
         TabIndex        =   74
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   10
         Left            =   8400
         MaxLength       =   25
         TabIndex        =   73
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   5880
         MaxLength       =   25
         TabIndex        =   72
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   66
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   360
         MaxLength       =   25
         TabIndex        =   64
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   62
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   360
         MaxLength       =   25
         TabIndex        =   61
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   58
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   360
         MaxLength       =   25
         TabIndex        =   55
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   2640
         MaxLength       =   25
         TabIndex        =   52
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   360
         MaxLength       =   25
         TabIndex        =   51
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Otra Autoridad"
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
         Left            =   5880
         TabIndex        =   71
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Dependencia"
         Height          =   195
         Left            =   8400
         TabIndex        =   70
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Policía Científica"
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
         Left            =   5880
         TabIndex        =   69
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   8400
         TabIndex        =   68
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Chofer Grúa"
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
         Left            =   5880
         TabIndex        =   67
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Dependencia"
         Height          =   195
         Left            =   2640
         TabIndex        =   65
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Gendarme"
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
         Left            =   360
         TabIndex        =   63
         Top             =   2520
         Width           =   870
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Dependencia"
         Height          =   195
         Left            =   2640
         TabIndex        =   60
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Policía"
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
         Left            =   360
         TabIndex        =   59
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Dependencia"
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Ambulanciero"
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
         Left            =   360
         TabIndex        =   56
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   2640
         TabIndex        =   54
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bombero"
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
         Left            =   360
         TabIndex        =   53
         Top             =   360
         Width           =   750
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Daños a Autopista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   48
      Top             =   7800
      Width           =   11535
      Begin VB.TextBox Text3 
         Height          =   360
         Index           =   0
         Left            =   360
         MaxLength       =   255
         TabIndex        =   49
         Top             =   600
         Width           =   10695
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         Height          =   195
         Left            =   360
         TabIndex        =   78
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   41
      Top             =   4200
      Width           =   11535
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   480
         Width           =   11055
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   45
         Top             =   2160
         Width           =   11055
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Nro- Nacionali- Lugar Trasl - Medio Trasl. - Herid-Fallec- Vehic- Edad- Sexo - Est.Civil- Cintu"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         Width           =   9975
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc - Nro Doc "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   9240
         TabIndex        =   44
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nro - Nombre                   - Domicilio                                           -"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   9030
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.Frame Frame10 
         Height          =   975
         Left            =   9360
         TabIndex        =   42
         Top             =   3120
         Width           =   1815
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar"
            Height          =   615
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   5
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2400
         Width           =   855
      End
      Begin VB.Frame Frame8 
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
         Height          =   495
         Left            =   8040
         TabIndex        =   40
         Top             =   3120
         Width           =   975
         Begin VB.OptionButton Option4 
            Height          =   195
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Fallecido"
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
         Left            =   4080
         TabIndex        =   39
         Top             =   2040
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "Después"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Durante el Traslado"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton Option2 
            Caption         =   "En elLugar"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   1095
         Left            =   9480
         TabIndex        =   20
         Top             =   2040
         Width           =   1575
         Begin VB.OptionButton Option3 
            Caption         =   "Femenino"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Masculino"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Traslado"
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
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   2415
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1200
            Width           =   1695
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Medio"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Lugar"
            Height          =   195
            Left            =   360
            TabIndex        =   37
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   14
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   6
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         Caption         =   "Herido"
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
         Left            =   2640
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
         Begin VB.OptionButton Option1 
            Caption         =   "Grave"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Leve"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   9600
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   8520
         MaxLength       =   9
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   3240
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   3375
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Víctimas Involucradas"
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
            Left            =   360
            TabIndex        =   26
            Top             =   240
            Width           =   2715
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Height          =   525
         Index           =   1
         Left            =   10440
         MaxLength       =   2
         TabIndex        =   22
         Text            =   "88"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Height          =   525
         Index           =   0
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   21
         Text            =   "888888"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Vehículo"
         Height          =   195
         Left            =   6480
         TabIndex        =   35
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Edad"
         Height          =   195
         Left            =   8040
         TabIndex        =   34
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Estado Civil"
         Height          =   195
         Left            =   6480
         TabIndex        =   33
         Top             =   3000
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nacionalidad"
         Height          =   195
         Left            =   9600
         TabIndex        =   31
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nro."
         Height          =   195
         Left            =   8520
         TabIndex        =   30
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Doc."
         Height          =   195
         Left            =   7440
         TabIndex        =   29
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   3240
         TabIndex        =   28
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nombre/Apellido"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N° Víctima"
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
         Left            =   9000
         TabIndex        =   24
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Orden"
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
         Left            =   6120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "RAcc3_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mRec As New ADODB.Recordset
Dim mNum As Integer
Dim mI As Integer

Private Sub Form_Load()
sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mRec = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim mJ As Integer
   If Index = 0 Then
      mI = Combo1(0).ListIndex
      If mI > -1 Then
         RAcc2_frm!List1.ListIndex = mI
         RAcc2_frm!List2.ListIndex = mI
         Text2(0).Text = Trim(mId(RAcc2_frm!List2.Text, 5, 50))
         Combo1(5).ListIndex = mI
         If Trim(mId(RAcc2_frm!List1.Text, 83, 2)) <> "" Then
            mJ = Trim(mId(RAcc2_frm!List1.Text, 83, 2))
            For mI = 0 To Combo1(1).ListCount - 1
               If mJ = Right(Combo1(1).List(mI), 2) Then
                  Combo1(1).ListIndex = mI            'Llena Tipo Docu
               End If
            Next
          End If
         Text2(1).Text = Trim(mId(RAcc2_frm!List1.Text, 93, 9))
      End If
   End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
   If Index = 0 Then
      If KeyAscii >= 97 And KeyAscii <= 122 Then
         KeyAscii = KeyAscii - 32
      Else
         If KeyAscii = 241 Then
            KeyAscii = 209
         Else
            If Not (KeyAscii >= 65 And KeyAscii <= 90) And KeyAscii <> 8 And KeyAscii <> 32 And KeyAscii <> 44 And KeyAscii <> 209 And KeyAscii <> 46 Then
               KeyAscii = 0
            End If
         End If
      End If
   End If
End Sub

Private Sub Command1_Click()
   If Command1.Caption = "Actualizar" Then
      Command1.Caption = "Agregar"
      Command1.Picture = LoadPicture("checkmrk.ico")
   End If
   If Combo1(0).Text <> "" Then
      If Option3(0).Value Or Option3(1).Value Then
          sLlenaLista
          Text1(1).Text = Format((Val(Text1(1).Text) + 1), "00")
          sLimpiarForm3
      Else
         MsgBox "Debe Seleccionar el Sexo", vbCritical, sMessage
      End If
   Else
      MsgBox "Debe Ingresar El Nombre, si no se Conoce Ingrese: N.N.", vbCritical, sMessage
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   If Index = 0 Then  'Volver
      Me.Visible = False
      RAcc2_frm.Show
      RAcc2_frm.Top = 0
      RAcc2_frm.Left = 0
   Else 'Siguiente
      RAcc1_frm.Visible = False
      RAcc2_frm.Visible = False
      RAcc3_frm.Visible = False
      RAcc9_frm.Visible = True
      RAcc9_frm.Top = 0
      RAcc9_frm.Left = 0
   End If
End Sub

Private Sub Command3_Click()
Dim OptionArray(14) As String
Dim mArray(6) As String
Dim mJ As String
Dim mI As Integer
Dim mFlag As Boolean
  
   If MsgBox("¿Desea Grabar los Datos?", vbYesNo, sMessage) = vbYes Then
      If Not RAcc1_frm.mBusca Then
         RAcc1_frm.mObj.xUpAuxiliar (Format((Val(RAcc1_frm!Text2(2).Text) + 1), "00000"))
      End If
      OptionArray(0) = SelectOption(RAcc1_frm!Option1) 'TRAMO
      mJ = ""
      For mI = 0 To 6
         If RAcc1_frm!Check2(mI).Value = 1 Then
            mJ = mJ & Format((mI + 1), "00")
         End If
      Next
      OptionArray(1) = mJ
      OptionArray(2) = SelectOption(RAcc1_frm!Option3) 'ACCIDE CON OTRO'
      OptionArray(3) = SelectOption(RAcc1_frm!Option4) 'OTRO
      OptionArray(4) = SelectOption(RAcc1_frm!Option5) 'SENTIDO TRANS
      OptionArray(5) = SelectOption(RAcc1_frm!Option6) 'LUGAR
      OptionArray(6) = SelectOption(RAcc1_frm!Option7) 'CLIMA
      OptionArray(7) = SelectOption(RAcc1_frm!Option9) 'CALZADA
      OptionArray(8) = SelectOption(RAcc1_frm!Option10) 'BANQUINA
      OptionArray(9) = SelectOption(RAcc1_frm!Option11) 'DEM. HOR
      OptionArray(10) = SelectOption(RAcc1_frm!Option12) 'DEM. VERT
      OptionArray(11) = SelectOption(RAcc1_frm!Option13) 'ILUM
      OptionArray(12) = SelectOption(RAcc1_frm!Option14) 'CAUSA VEHIC
      If Check1(0).Value = 1 And Check1(1).Value = 1 Then
         OptionArray(13) = "03"
      Else
         If Check1(0).Value = 1 And Check1(1).Value = 0 Then
            OptionArray(13) = "01"
         Else
            If Check1(0).Value = 0 And Check1(1).Value = 1 Then
               OptionArray(13) = "02"
            End If
         End If
      End If
      For mI = 0 To RAcc1_frm!Combo2.UBound
         If RAcc1_frm!Combo2(mI).Text <> "" Then
            mArray(mI) = Trim(Right(RAcc1_frm!Combo2(mI).Text, 2))
         End If
      Next
      mJ = 3 'SE USA COMO CONTADOR
      For mI = 0 To RAcc1_frm!Combo3.UBound
         If RAcc1_frm!Combo3(mI).Text <> "" Then
            mArray(mI) = Trim(Right(RAcc1_frm!Combo3(mI).Text, 3))
         End If
         mJ = mJ + 1
      Next
      If Not RAcc1_frm.mBusca Then
         mFlag = RAcc1_frm.mObj.xInsFichaOlder(Trim(RAcc1_frm!Text2(2).Text), Trim(Left(RAcc1_frm!Combo1.Text, 3)), Trim(RAcc1_frm!Text2(0).Text), Trim(RAcc1_frm!Text2(1).Text), Trim(RAcc1_frm!Text1(0).Text), Trim(RAcc1_frm!Text1(1).Text), Trim(RAcc1_frm!Text1(2).Text), Trim(RAcc1_frm!Text3(0).Text), OptionArray(0), RAcc1_frm!Text3(1).Text, OptionArray(1), _
                                               RAcc1_frm!Text4.Text, OptionArray(2), mArray(0), mArray(1), mArray(2), OptionArray(3), OptionArray(4), OptionArray(5), OptionArray(6), OptionArray(7), OptionArray(8), OptionArray(9), OptionArray(10), OptionArray(11), mArray(3), mArray(4), mArray(5), OptionArray(12), _
                                               Trim(Text3(0).Text), Trim(Text3(1).Text), Trim(Text3(2).Text), Trim(Text3(3).Text), Trim(Text3(4).Text), Trim(Text3(5).Text), Trim(Text3(6).Text), Trim(Text3(7).Text), Trim(Text3(8).Text), Trim(Text3(9).Text), Trim(Text3(10).Text), Trim(Text3(11).Text), Trim(Text3(12).Text), Trim(Text3(13).Text), OptionArray(13))
      Else
         mFlag = RAcc1_frm.mObj.xUpFichaOlder(RAcc1_frm!Text2(2).Text, Trim(Left(RAcc1_frm!Combo1.Text, 3)), Trim(RAcc1_frm!Text2(0).Text), Trim(RAcc1_frm!Text2(1).Text), Trim(RAcc1_frm!Text1(0).Text), Trim(RAcc1_frm!Text1(1).Text), _
                Trim(RAcc1_frm!Text1(2).Text), Trim(RAcc1_frm!Text3(0).Text), OptionArray(0), RAcc1_frm!Text3(1).Text, OptionArray(1), RAcc1_frm!Text4.Text, OptionArray(2), mArray(0), mArray(1), _
                mArray(2), OptionArray(3), OptionArray(4), OptionArray(5), OptionArray(6), OptionArray(7), OptionArray(8), OptionArray(9), OptionArray(10), OptionArray(11), mArray(3), _
                mArray(4), mArray(5), OptionArray(12), Trim(Text3(0).Text), Trim(Text3(1).Text), Trim(Text3(2).Text), Trim(Text3(3).Text), Trim(Text3(4).Text), Trim(Text3(5).Text), Trim(Text3(6).Text), _
                Trim(Text3(7).Text), Trim(Text3(8).Text), Trim(Text3(9).Text), Trim(Text3(10).Text), Trim(Text3(11).Text), Trim(Text3(12).Text), Trim(Text3(13).Text), OptionArray(13))
         If mFlag Then
            mFlag = RAcc1_frm.mObj.xDeleteTable("VehiculosInvolucr", "nroorden='" & Trim(RAcc1_frm!Text2(2).Text) & "'")
            mFlag = RAcc1_frm.mObj.xDeleteTable("VictimasInvolucr", "nroorden='" & Trim(RAcc1_frm!Text2(2).Text) & "'")
         End If
      End If
      
      If RAcc2_frm!List1.ListCount > 0 Then
         For mI = 0 To RAcc2_frm!List1.ListCount - 1
            RAcc2_frm!List1.ListIndex = mI
            RAcc2_frm!List2.ListIndex = mI
            mFlag = RAcc1_frm.mObj.xInsVehicOlder(Trim(RAcc1_frm!Text2(2).Text), Trim(mId(RAcc2_frm!List1.Text, 1, 2)), Trim(mId(RAcc2_frm!List1.Text, 6, 2)), Trim(mId(RAcc2_frm!List1.Text, 16, 2)), _
                  Trim(mId(RAcc2_frm!List1.Text, 26, 15)), Trim(mId(RAcc2_frm!List1.Text, 44, 8)), Trim(mId(RAcc2_frm!List1.Text, 55, 25)), Trim(mId(RAcc2_frm!List1.Text, 83, 2)), Trim(mId(RAcc2_frm!List1.Text, 93, 9)), _
                  Trim(mId(RAcc2_frm!List2.Text, 5, 50)), Trim(mId(RAcc2_frm!List2.Text, 58, 15)), Trim(mId(RAcc2_frm!List2.Text, 76, 2)), Trim(mId(RAcc2_frm!List2.Text, 86, 20)))
         Next
      End If
      If List1.ListCount > 0 Then
         For mI = 0 To List1.ListCount - 1
            List1.ListIndex = mI
            List2.ListIndex = mI
            mFlag = RAcc1_frm.mObj.xInsVictimasOlder(Text1(0).Text, Trim(Left(List1.Text, 2)), Trim(mId(List1.Text, 6, 25)), Trim(mId(List1.Text, 34, 50)), Trim(mId(List1.Text, 87, 2)), Trim(mId(List1.Text, 97, 9)), Trim(mId(List2.Text, 6, 2)), Trim(mId(List2.Text, 16, 2)), Trim(mId(List2.Text, 31, 2)), Trim(mId(List2.Text, 48, 1)), Trim(mId(List2.Text, 55, 1)), Trim(mId(List2.Text, 60, 1)), Trim(mId(List2.Text, 75, 1)), Trim(mId(List2.Text, 68, 3)), Trim(mId(List2.Text, 80, 1)), Trim(mId(List2.Text, 91, 2)))
         Next
      End If
      Unload RAcc1_frm
      Unload RAcc2_frm
      Unload Me
      RAcc1_frm.Show
      If Not RAcc1_frm.mBusca Then
         RAcc1_frm.Combo4.Visible = False
      End If
  End If
End Sub

Private Sub Option1_DblClick(Index As Integer)
   Option1(Index).Value = Not Option1(Index).Value
End Sub

Private Sub Option2_DblClick(Index As Integer)
   Option2(Index).Value = Not Option2(Index).Value
End Sub

Private Sub Option4_DblClick()
   Option4.Value = Not Option4.Value
End Sub

Private Sub List1_Click()
   List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
   List1.ListIndex = List2.ListIndex
End Sub

Private Sub List1_DblClick()
Dim mnro As String
Dim mResp, mI, mJ As Integer
Dim mFlag As Boolean
   mnro = Text1(1).Text
   mFlag = True
   List2.ListIndex = List1.ListIndex
   If (List1.ListCount - 1) = List1.ListIndex Then
      If MsgBox("¿Desea Eliminar el Registro?", vbYesNo, sMessage) = vbYes Then
         mJ = List1.ListIndex
         Text1(1).Text = Trim(Left(List1.Text, 3))
         List1.RemoveItem (mJ)
         List2.RemoveItem (mJ)
         mFlag = False
      End If
   End If
   If mFlag Then
      If MsgBox("¿Desea Modificar el Registro?", vbYesNo, sMessage) = vbYes Then
         Command1.Caption = "Actualizar"
         Command1.Picture = LoadPicture("erase02.ico")
         List2.ListIndex = List1.ListIndex
         If Not RAcc1_frm.mBusca Then
            mNum = Text1(1).Text
         End If
         Text1(1).Text = Trim(Left(List1.Text, 3))
         Combo1(0).Text = Trim(mId(List1.Text, 6, 25)) 'nombre
         Text2(0).Text = Trim(mId(List1.Text, 34, 50)) 'domicilio
         If Trim(mId(List1.Text, 87, 2)) <> "" Then
            For mI = 0 To Combo1(1).ListCount - 1
               If Trim(mId(List1.Text, 87, 2)) = Trim(Right(Combo1(1).List(mI), 2)) Then 'Llena Tipo Docu
                  Combo1(1).ListIndex = mI
               End If
            Next
         End If
         Text2(1).Text = Trim(mId(List1.Text, 97, 8)) 'nro docu

     '////////////******  LIST 2 **************////////
         If Trim(mId(List2.Text, 6, 2)) <> "" Then
            mJ = Trim(mId(List2.Text, 6, 2))
            For mI = 0 To Combo1(2).ListCount - 1   'Llena nacion
              If mJ = Trim(Right(Combo1(2).List(mI), 2)) Then
                 Combo1(2).ListIndex = mI
               End If
            Next
         End If
 
         If Trim(mId(List2.Text, 16, 2)) <> "" Then
            mJ = Trim(mId(List2.Text, 16, 2))
            For mI = 0 To Combo1(3).ListCount - 1
               If mJ = Trim(Right(Combo1(3).List(mI), 2)) Then
                  Combo1(3).ListIndex = mI            'Llena Lugar Traslado
               End If
            Next
         End If
         If Trim(mId(List2.Text, 31, 2)) <> "" Then
            mJ = Trim(mId(List2.Text, 31, 2))
            For mI = 0 To Combo1(1).ListCount - 1
               If mJ = Trim(Right(Combo1(4).List(mI), 2)) Then
                  Combo1(4).ListIndex = mI              'Llena Medio Traslado
               End If
            Next
         End If
         If Trim(mId(List2.Text, 48, 1)) <> "" Then
            mI = Trim(mId(List2.Text, 48, 1))
            Option1(mI - 1).Value = True                   ' HERIDO
         End If
         If Trim(mId(List2.Text, 55, 1)) <> "" Then
            mI = Trim(mId(List2.Text, 55, 1))            'FALLECIDO
            Option2(mI - 1).Value = True
         End If
         If Trim(mId(List2.Text, 60, 1)) <> "" Then
            For mI = 0 To Combo1(5).ListCount - 1
              If Combo1(5).List(mI) = Trim(mId(List2.Text, 60, 1)) Then
                   Combo1(5).ListIndex = mI                           'Llena Vehiculo
              End If
            Next
         End If
         Text2(2).Text = Trim(mId(List2.Text, 68, 3))         'LLENA EDAD
         mI = Trim(mId(List2.Text, 75, 1))
         Option3(mI - 1).Value = True               'SEXO
         If Trim(mId(List2.Text, 80, 1)) <> "" Then
            mJ = Trim(mId(List2.Text, 80, 1))
            For mI = 0 To Combo1(6).ListCount - 1
               If mJ = Trim(Left(Combo1(6).List(mI), 1)) Then
                  Combo1(6).ListIndex = mI                     'Llena Estado Civil
               End If
            Next
         End If
         If Trim(mId(List2.Text, 91, 2)) = "SI" Then
            Option4.Value = True                            'LLENA CINTURON
         End If
         mJ = List1.ListIndex
         List1.RemoveItem (mJ)
         List2.RemoveItem (mJ)
      End If
   End If
End Sub

Private Sub List2_DblClick()
   List1.ListIndex = List2.ListIndex
   List1_DblClick
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Else
      KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
   End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
End Sub

Private Function sLlenaLista() As Boolean
Dim mOpt1 As String
Dim mOpt2 As String
Dim mOpt3 As String
Dim mOpt4 As String
  
   mOpt1 = SelectOption(Option1)
   mOpt2 = SelectOption(Option2)
   mOpt3 = SelectOption(Option3)
   mOpt4 = "NO"
   If Option4.Value Then
      mOpt4 = "SI"
   End If
                   ' ************ Nro *******************   /    *********** NOMBRE   *****************    /  *********   DIRECCION  **************       /    *******************************     TIPO DOCUMENTO ***************************    /   *************** NRO DOCUMENTO  *******
  List1.AddItem "" & Right(Space(2) & Text1(1).Text, 2) & " - " & Left(Combo1(0).Text & Space(25), 25) & " - " & Left(Text2(0).Text & Space(50), 50) & " - " & Right(Space(7) & Right(Combo1(1).Text, 2) & " " & Left(Combo1(1).Text, 4), 7) & " - " & Right(Space(9) & Text2(1).Text, 9) & ""
                   '  ***********  NRO  *****************   /   ************************  NACIONALIDAD ******************************************   / ***********************************  LUGAR DE TRASLADO  ****************************   /  *************************************  MEDIO DE TRASLADO  ************************   /   ******* HERIDO  ************   /   ******* FALLECIDO  **********    /   ********  VEHICULO   ******************   /   **********  EDAD   ******************   /  **************  SEXO  *********   /  ************** ESTADO CIVIL  *********   / ********* CINTURON  *************
  List2.AddItem "" & Right(Space(2) & Text1(1).Text, 2) & " - " & Right(Space(7) & Right(Combo1(2).Text, 2) & " " & Left(Combo1(2).Text, 4), 7) & " - " & Right(Space(12) & Right(Combo1(3).Text, 2) & " " & Left(Combo1(3).Text, 9), 12) & " - " & Right(Space(12) & Right(Combo1(4).Text, 2) & " " & Left(Combo1(4).Text, 9), 12) & " -  " & Left(mOpt1 & Space(3), 3) & " -  " & Left(mOpt2 & Space(3), 3) & " - " & Left(Combo1(5).Text & Space(4), 4) & " -  " & Left(Text2(2).Text & Space(3), 3) & "-  " & Left(mOpt3 & Space(3), 3) & " - " & Left(Combo1(6).Text & Space(8), 8) & " - " & Left(mOpt4 & Space(3), 3) & " "
End Function

Private Function SelectOption(mObj As Object) As String
   For mI = 0 To mObj.UBound
      If mObj(mI).Value Then
         SelectOption = Format((mI + 1), "00")
         mI = 60
      Else
         SelectOption = ""
      End If
   Next
End Function

Private Sub sLimpiarForm3()
   Combo1(0).Text = ""
   For mI = 0 To Combo1.UBound
      Combo1(mI).ListIndex = -1
   Next
   For mI = 0 To Text2.UBound
      Text2(mI).Text = ""
   Next
   Option1(0).Value = False
   Option1(1).Value = False
   Option2(0).Value = False
   Option2(1).Value = False
   Option2(2).Value = False
   Option3(0).Value = False
   Option3(1).Value = False
   Option4.Value = False
End Sub
   
Private Sub sInitForm()
Dim mTipoDocu As String
Dim mNacion As String
Dim mLugar As String
Dim mMedio As String
Dim mEstCivil As String

   Me.Top = 0
   Me.Left = 0
   Me.Height = 13100
   Me.Width = 11900
   If RAcc1_frm.mBusca Then
     Command3.Caption = "&Actualizar"
     Command3.BackColor = &H5256FE
     Me.Caption = Me.Caption & "  ---MODO VISTA Y MODIFICACIÓN--"
   End If
   Text1(0).Text = RAcc1_frm!Text2(2).Text
   Text1(1).Text = "01"
   For mI = 0 To Text3.UBound
      Text3(mI).Text = ""
   Next
   Set mRec = RAcc1_frm.mObj.oTabla("LugarTrasl", "order by descripcion")
   sLlenoCbo Me.Combo1(3), mRec, 1, 0
   Set mRec = RAcc1_frm.mObj.oTabla("MedioTrasl", "order by descripcion")
   sLlenoCbo Me.Combo1(4), mRec, 1, 0
   Set mRec = RAcc1_frm.mObj.oTabla("TipoDocu", "")
   sLlenoCbo Me.Combo1(1), mRec, 1, 0
   Set mRec = RAcc1_frm.mObj.oTabla("Nacionalidad", "order by descripcion")
   sLlenoCbo Me.Combo1(2), mRec, 1, 0
   Set mRec = RAcc1_frm.mObj.oTabla("EstadoCivil", "order by descripcion")
   sLlenoCbo Me.Combo1(6), mRec, 1, 0
   If RAcc1_frm.mBusca Then
      mI = 0
      Set mRec = RAcc1_frm.mObj.oTabla("VictimasInvolucr", " where nroorden='" & RAcc1_frm.mNroOrden & "'")
      Do While Not mRec.EOF
         mTipoDocu = RAcc1_frm.mObj.sTablaDescr("TipoDocu", "CodTipoDocu = '" & NVL(mRec!TipoDocu, "") & "'", 1)
         mNacion = RAcc1_frm.mObj.sTablaDescr("Nacionalidad", "CodNacion = '" & NVL(mRec!CodNacion, "") & "'", 1)
         mLugar = RAcc1_frm.mObj.sTablaDescr("LugarTrasl", "CodLugarTrasl = '" & NVL(mRec!CodLugarTrasl, "") & "'", 1)
         mMedio = RAcc1_frm.mObj.sTablaDescr("MedioTrasl", "CodMedioTrasl = '" & NVL(mRec!CodMedioTrasl, "") & "'", 1)
         mEstCivil = RAcc1_frm.mObj.sTablaDescr("EstadoCivil", "CodEstCivil = '" & mRec!CodEstCivil & "'", 1)
         List1.AddItem "" & Right(Space(2) & mRec!NroVictima, 2) & " - " & Left(mRec!Nombre & Space(25), 25) & " - " & Left(mRec!domicilio & Space(50), 50) & " - " & Left(Left(mRec!TipoDocu, 2) & " " & Left(mTipoDocu, 4) & Space(7), 7) & " - " & Left(mRec!NroDocu & Space(9), 9) & ""
         List2.AddItem "" & Right(Space(2) & mRec!NroVictima, 2) & " - " & Right(Space(7) & Right(mRec!CodNacion, 2) & " " & Left(mNacion, 4), 7) & " - " & Right(Space(12) & Right(mRec!CodLugarTrasl, 2) & " " & Left(mLugar, 9), 12) & " - " & Right(Space(12) & Right(mRec!CodMedioTrasl, 2) & " " & Left(mMedio, 9), 12) & " -  " & Left(Format(mRec!Herido, "00") & Space(3), 3) & " -  " & Left(Format(mRec!Fallecio, "00") & Space(3), 3) & " - " & Left(mRec!Letra & Space(4), 4) & " -  " & Left(mRec!Edad & Space(3), 3) & "-  " & Left(Format(mRec!Sexo, "00") & Space(3), 3) & " - " & Left(mRec!CodEstCivil & " " & Left(mEstCivil, 5) & Space(8), 8) & " - " & Left(mRec!Cinturon & Space(3), 3) & " "
         mI = mRec!NroVictima
         mRec.MoveNext
      Loop
      mRec.Close
      Text1(1).Text = Format(mI + 1, "00")
      Set mRec = RAcc1_frm.mObj.oTabla("Ficha", "where nroorden='" & RAcc1_frm.mNroOrden & "'")
      If Not mRec.EOF Then
         Text3(0).Text = NVL(mRec!DanosGCO, "")
         Text3(1).Text = NVL(mRec!BombName, "")
         Text3(2).Text = NVL(mRec!BombDepto, "")
         Text3(3).Text = NVL(mRec!AmbulName, "")
         Text3(4).Text = NVL(mRec!AmbulDepend, "")
         Text3(5).Text = NVL(mRec!PoliciaName, "")
         Text3(6).Text = NVL(mRec!PoliciaDepend, "")
         Text3(7).Text = NVL(mRec!GendarName, "")
         Text3(8).Text = NVL(mRec!GendarDepend, "")
         Text3(9).Text = NVL(mRec!GruaName, "")
         Text3(10).Text = NVL(mRec!GruaEmpr, "")
         Text3(11).Text = NVL(mRec!PoliCientName, "")
         Text3(12).Text = NVL(mRec!PoliCientDepend, "")
         Text3(13).Text = NVL(mRec!OtraAutoridad, "")
         Select Case mRec!Foto
            Case "01"
               Check1(0).Value = 1
            Case "02"
               Check1(1).Value = 1
            Case "03"
               Check1(0).Value = 1
               Check1(1).Value = 1
         End Select
      End If
      mRec.Close
   End If
End Sub
