VERSION 5.00
Begin VB.Form RAcc8_frm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Módulo Parámetros de Búsqueda"
   ClientHeight    =   5385
   ClientLeft      =   5040
   ClientTop       =   1335
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Para tener en cuenta "
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
      Height          =   5055
      Left            =   15000
      TabIndex        =   48
      Top             =   240
      Width           =   3555
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Presionando la tecla ""ESC"" sobre un combo de selección, este quederá en blanco si estuviera seleccionado."
         Height          =   675
         Index           =   4
         Left            =   120
         TabIndex        =   54
         Top             =   3780
         Width           =   3255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   195
         Left            =   2880
         MouseIcon       =   "Form8.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   4740
         Width           =   525
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form8.frx":0152
         Height          =   855
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AB%6"
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
         Index           =   2
         Left            =   360
         TabIndex        =   51
         Top             =   2340
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form8.frx":0206
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   1140
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Al ingresar un texto a buscar se podrá utilizar un caracter comodín, este es el símbolo porcentaje:   %"
         Height          =   675
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   7
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   4380
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E8E8E3&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4860
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   13
      Left            =   5280
      MaxLength       =   20
      TabIndex        =   38
      Top             =   4260
      Width           =   2385
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   12
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   37
      Top             =   3900
      Width           =   2385
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   11
      Left            =   5280
      MaxLength       =   18
      TabIndex        =   36
      Top             =   3540
      Width           =   2385
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   10
      Left            =   5280
      MaxLength       =   30
      TabIndex        =   35
      Top             =   3180
      Width           =   2385
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   6
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   3960
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   5
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3540
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   9
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   32
      Top             =   3180
      Width           =   1845
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   5280
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1980
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   5280
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1560
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2400
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1980
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   1260
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   8
      Left            =   6360
      MaxLength       =   7
      TabIndex        =   14
      Top             =   420
      Width           =   1305
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   7
      Left            =   2400
      MaxLength       =   5
      TabIndex        =   12
      Top             =   420
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   6
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   11
      Top             =   420
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C1DBD8&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   5220
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4860
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   5
      Left            =   7080
      MaxLength       =   5
      TabIndex        =   7
      Top             =   780
      Width           =   550
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   4
      Left            =   6360
      MaxLength       =   5
      TabIndex        =   6
      Top             =   780
      Width           =   550
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   3
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   5
      Top             =   780
      Width           =   550
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   2
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   4
      Top             =   780
      Width           =   550
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   1
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   3
      Top             =   780
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Index           =   0
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   2
      Top             =   780
      Width           =   1000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cómo buscar ?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   180
      MouseIcon       =   "Form8.frx":02D1
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   5100
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Vehic:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   17
      Left            =   180
      TabIndex        =   46
      Top             =   3660
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos día y lugar"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos víctimas"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   4080
      TabIndex        =   43
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos vehículo"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   120
      TabIndex        =   42
      Top             =   2880
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos personal GCO"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Left            =   4080
      TabIndex        =   41
      Top             =   1260
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos accidente"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   40
      Top             =   1260
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   16
      Left            =   4260
      TabIndex        =   31
      Top             =   4320
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   15
      Left            =   4260
      TabIndex        =   30
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Doc:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   14
      Left            =   4260
      TabIndex        =   29
      Top             =   3600
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apell, Nombre:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   13
      Left            =   3840
      TabIndex        =   28
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   12
      Left            =   540
      TabIndex        =   27
      Top             =   4080
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   11
      Left            =   600
      TabIndex        =   26
      Top             =   4500
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patente:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   10
      Left            =   420
      TabIndex        =   25
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Móvil GCO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   9
      Left            =   4140
      TabIndex        =   23
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patrullero:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   8
      Left            =   4200
      TabIndex        =   20
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Causa Veh:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   2460
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Causa Prob:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   17
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inconv:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Alfa:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   13
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Ficha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Km:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   5820
      TabIndex        =   8
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   660
   End
End
Attribute VB_Name = "RAcc8_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clRAcc
Dim mRec As ADODB.Recordset

Private Sub Form_Load()
   sAlinearForm Me
   sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   RAcc1beta.Enabled = True
End Sub

Private Sub Combo1_Click(Index As Integer)
   If Index = 5 Then
      Combo1(6).Clear
      If Combo1(5).ListIndex > -1 Then
         Set mRec = mObj.oMarcasVehic(Right(Combo1(5).Text, 2))
         sLlenoCbo Combo1(6), mRec, 2, 1
      End If
   End If
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 27 Then
      Combo1(Index).ListIndex = -1
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mSQL As String
Dim mTablas As String
Dim mWhere As String

   If Index = 0 Then
      mTablas = "Ficha A"
      mWhere = ""
      mSQL = ""
      If fValid Then
         If Combo1(3).ListIndex > -1 Or Combo1(4).ListIndex > -1 Then
            mTablas = mTablas & ", intergco B"
            mWhere = " A.nroorden=B.nroorden AND "
         End If
         If Text1(9).Text <> "" Or Combo1(5).ListIndex > -1 Or Combo1(6).ListIndex > -1 Or Combo1(7).ListIndex > -1 Then
            mTablas = mTablas & ", VehiculosInvolucr C"
            mWhere = " A.nroorden=C.nroorden AND "
         End If
         If Text1(10).Text <> "" Or Text1(11).Text <> "" Or Text1(12).Text <> "" Or Text1(13).Text <> "" Then
            mTablas = mTablas & ", VictimasInvolucr D"
            mWhere = " A.nroorden=D.nroorden AND "
         End If
         mSQL = fArmarSQL
         If Trim(mSQL) = "where" Then
            mSQL = ""
         End If
         If Trim(mWhere) <> "" Then
            mWhere = Mid(mWhere, 1, Len(mWhere) - 4)
            mSQL = mSQL & " and (" & mWhere & ")"
         End If
         Set mRec = mObj.oBuscar(mTablas, mSQL & " order by nroorden")
         If Not mRec.EOF Then
            RAcc1beta.Combo5.Clear
            Do While Not mRec.EOF
               RAcc1beta.Combo5.AddItem mRec.Fields(0) & " - " & mRec!Fecha
               mRec.MoveNext
            Loop
         End If
         mRec.Close
         Unload Me
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub sInitForm()
Dim mObjRN As New clRNov

   RAcc1beta.Enabled = False
   Set mRec = mObj.oTabla("inconvenientes", " where fechabaja is null")
   sLlenoCbo Combo1(0), mRec, 1, 0
   Set mRec = mObj.oTabla("CausaConductor", " where fechabaja is null")
   sLlenoCbo Combo1(1), mRec, 1, 0
   Set mRec = mObj.oTabla("CausaVehic", " where fechabaja is null")
   sLlenoCbo Combo1(2), mRec, 1, 0
   Set mRec = mObjRN.oTablaNull("patrulleros")
   sLlenoCbo Combo1(3), mRec, 1, 0
   Set mRec = mObjRN.oMovilesGCO("PAT")
   sLlenoCbo Combo1(4), mRec, 1, 0
   Set mObjRN = Nothing
   Set mRec = mObj.oTabla("TipoVehiculo", "")
   sLlenoCbo Combo1(5), mRec, 1, 0
   Set mRec = mObj.oTabla("colores", "") 'traer colores
   sLlenoCbo Combo1(7), mRec, 1, 0
End Sub

Private Function fValid() As Boolean
   fValid = True
   'Fecha
   If Trim(Text1(0).Text) <> "" Then
      fValid = Fecha_ok(Text1(0).Text)
      If Trim(Text1(1).Text) <> "" Then
          fValid = sValidFechaDesdeHasta(Text1(0).Text, Text1(1).Text)
      Else
         Text1(1).Text = Text1(0).Text
      End If
   End If
   'Horas
   If Trim(Text1(2).Text) <> "" Then
      fValid = fValid And Hora_ok(Text1(2).Text)
      If Trim(Text1(3).Text) <> "" Then
          fValid = fValid And Hora_ok(Text1(3).Text)
      Else
         Text1(3).Text = Text1(2).Text
      End If
   End If
   'Kms
   If Trim(Text1(4).Text) <> "" Then
      If Trim(Text1(5).Text) <> "" Then
         If Val(Text1(5).Text) < Val(Text1(4).Text) Then
            Text1(4).Tag = Text1(5).Text
            Text1(5).Text = Text1(4).Text
            Text1(4).Text = Text1(4).Tag
         End If
      Else
         Text1(5).Text = Text1(4).Text
      End If
   End If
End Function

Private Sub Label3_Click()
Dim mI As Integer
   For mI = Frame1.Left To 4140 Step -100
         Frame1.Left = mI
   Next
   Frame1.Left = 4140
End Sub

Private Sub Label5_Click()
   Frame1.Left = 15000
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1 'fechas
         KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
         
      Case 2, 3 'horas
         KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
      
      Case 4, 5 'kms
         If KeyAscii <> 46 Then
            KeyAscii = fNumeroKeyPress(KeyAscii)
         End If
         
      Case 6, 7 'nro fichas
         If KeyAscii <> 37 Then
            KeyAscii = fNumeroKeyPress(KeyAscii)
         End If
      
      Case 8 To 12 'patente, nombre, nro doc., domicilio
         If KeyAscii <> 37 Then
            KeyAscii = fAlfaNumKeyPress(KeyAscii)
         End If
      
      Case 13 'teléfono
         If KeyAscii <> 37 Then
            KeyAscii = fNumeroKeyPress(KeyAscii)
         End If
   End Select
End Sub

Private Function fArmarSQL() As String
Dim mTexto As String
Dim mConector As String

   mTexto = ""
   mConector = " AND "
   'fecha
   If Trim(Text1(0).Text) <> "" Then
      mTexto = " A.fecha between '" & Format(Text1(0).Text, "yyyy-mm-dd") & "' and '" & Format(Text1(1).Text, "yyyy-mm-dd") & "' " & mConector
   End If
   'fecha
   If Trim(Text1(2).Text) <> "" Then
      mTexto = mTexto & " A.hora between '" & Format(Text1(2).Text, "hh:mm") & "' and '" & Format(Text1(3).Text, "hh:mm") & "' " & mConector
   End If
   'km
   If Trim(Text1(4).Text) <> "" Then
      mTexto = mTexto & " A.progesiva between " & Text1(4).Text & " and " & Text1(5).Text & " " & mConector
   End If
   'nro ficha
   If Trim(Text1(6).Text) <> "" Then
      mTexto = mTexto & " A.nroorden between '" & Text1(6).Text & "' and '" & Text1(7).Text & "' " & mConector
   End If
   'codigo alfa
   If Trim(Text1(8).Text) <> "" Then
      mTexto = mTexto & " A.codalfa like '" & Text1(8).Text & "' " & mConector
   End If
   'inconvenientes
   If Combo1(0).ListIndex > -1 Then
      mTexto = mTexto & " A.codinconv = '" & Right(Combo1(0).Text, 2) & "' " & mConector
   End If
   'causa probable
   If Combo1(1).ListIndex > -1 Then
      mTexto = mTexto & " A.codcausacond1 = '" & Right(Combo1(1).Text, 2) & "' " & mConector
   End If
   'causa vehic
   If Combo1(2).ListIndex > -1 Then
      mTexto = mTexto & " A.causavehic = '" & Right(Combo1(2).Text, 2) & "' " & mConector
   End If
   
   'patrullero
   If Combo1(3).ListIndex > -1 Then
      mTexto = mTexto & " B.patrullero1 = '" & Trim(Left(Combo1(3).Text, 25)) & "' " & mConector
   End If
   'móvil GCO
   If Combo1(4).ListIndex > -1 Then
      mTexto = mTexto & " B.codmovil = '" & Right(Combo1(4).Text, 4) & "' " & mConector
   End If
   
   'patente
   If Trim(Text1(9).Text) <> "" Then
      mTexto = mTexto & " C.dominio LIKE '" & Trim(Text1(9).Text) & "' " & mConector
   End If
   'tipovehic
   If Combo1(5).ListIndex > -1 Then
      mTexto = mTexto & " C.codtipovehic = '" & Right(Combo1(5).Text, 2) & "' " & mConector
      If Combo1(6).ListIndex > -1 Then
         mTexto = mTexto & " C.codmarca = '" & Right(Combo1(6).Text, 2) & "' " & mConector
      End If
   End If
   'color
   If Combo1(7).ListIndex > -1 Then
      mTexto = mTexto & " C.codcolor = '" & Right(Combo1(5).Text, 2) & "' " & mConector
   End If
   
   'nombre
   If Trim(Text1(10).Text) <> "" Then
      mTexto = mTexto & " D.nombre LIKE '" & Trim(Text1(10).Text) & "' " & mConector
   End If
   'nro doc
   If Trim(Text1(11).Text) <> "" Then
      mTexto = mTexto & " D.nrodocu LIKE '" & Trim(Text1(11).Text) & "' " & mConector
   End If
   'domicilio
   If Trim(Text1(12).Text) <> "" Then
      mTexto = mTexto & " D.domicilio LIKE '" & Trim(Text1(12).Text) & "' " & mConector
   End If
   'nombre
   If Trim(Text1(13).Text) <> "" Then
      mTexto = mTexto & " D.tel LIKE '" & Trim(Text1(13).Text) & "' " & mConector
   End If
   
   If mTexto <> "" Then
      mTexto = Mid(mTexto, 1, Len(mTexto) - 4)
   End If
   fArmarSQL = " where " & mTexto
End Function
