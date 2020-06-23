VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MantElect02auxiliar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Ordenes de Trabajo"
   ClientHeight    =   13650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   27105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   13650
   ScaleWidth      =   27105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   11295
      Index           =   3
      Left            =   120
      TabIndex        =   46
      Top             =   1320
      Width           =   26775
      Begin VB.CommandButton Command3ve 
         Caption         =   "Grabar Asigancion"
         Height          =   375
         Left            =   7560
         TabIndex        =   69
         Top             =   5040
         Width           =   3135
      End
      Begin VB.Frame Frame5 
         Caption         =   "Vehículos Asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   240
         TabIndex        =   67
         Top             =   5520
         Width           =   26295
         Begin MSFlexGridLib.MSFlexGrid FlexVehAsignado 
            Height          =   3855
            Left            =   600
            TabIndex        =   68
            Top             =   360
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   6800
            _Version        =   327680
            Cols            =   5
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Vehículos disponibles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   240
         TabIndex        =   65
         Top             =   120
         Width           =   26295
         Begin MSFlexGridLib.MSFlexGrid FlexVehDisponible 
            Height          =   3855
            Left            =   360
            TabIndex        =   66
            Top             =   360
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   6800
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.TextBox Text1ve 
         Height          =   285
         Index           =   1
         Left            =   4200
         TabIndex        =   62
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox Text1ve 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   61
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Command2ve 
         Height          =   495
         Index           =   1
         Left            =   6480
         Picture         =   "MantElect02auxiliar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton Command2ve 
         Height          =   495
         Index           =   0
         Left            =   5760
         Picture         =   "MantElect02auxiliar.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4920
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "Km Final"
         Height          =   255
         Left            =   3360
         TabIndex        =   64
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Km Inicial"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   5040
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   240
      TabIndex        =   44
      Top             =   480
      Width           =   26775
      _ExtentX        =   47228
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tablero de Requerimientos"
            Object.Tag             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asiganción Materiales"
            Object.Tag             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asignación Mano de Obra"
            Object.Tag             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asignación Vehículos"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "FrameManoObra"
      Height          =   11295
      Index           =   2
      Left            =   120
      TabIndex        =   34
      Top             =   1320
      Width           =   26775
      Begin VB.CommandButton Command2mo 
         Height          =   495
         Index           =   1
         Left            =   10560
         Picture         =   "MantElect02auxiliar.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton Command2mo 
         Height          =   495
         Index           =   0
         Left            =   10560
         Picture         =   "MantElect02auxiliar.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4800
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Caption         =   "Mano de Obra Asiganda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   120
         TabIndex        =   52
         Top             =   6360
         Width           =   26535
         Begin MSFlexGridLib.MSFlexGrid FlexMoAsignada 
            Height          =   4095
            Left            =   480
            TabIndex        =   53
            Top             =   360
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   7223
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Mano de Obra Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1295
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   26535
         Begin VB.CommandButton GrabarMO 
            Caption         =   "Grabar MO"
            Height          =   495
            Left            =   11160
            TabIndex        =   58
            Top             =   5160
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid FlexMoDisponible 
            Height          =   3855
            Left            =   240
            TabIndex        =   51
            Top             =   600
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   6800
            _Version        =   327680
            Cols            =   3
         End
         Begin VB.Label Label6 
            Caption         =   "Tecnicos asignados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12600
            TabIndex        =   57
            Top             =   2400
            Width           =   3615
         End
         Begin VB.Label Label5 
            Caption         =   "Tecnicos disponibles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   56
            Top             =   2400
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FrameMateriales"
      Height          =   1295
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   26775
      Begin VB.ComboBox Combo1i 
         Height          =   315
         Left            =   4200
         TabIndex        =   49
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox Text1i 
         Height          =   285
         Left            =   1200
         TabIndex        =   48
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   18840
         Picture         =   "MantElect02auxiliar.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5880
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Index           =   2
         Left            =   18000
         Picture         =   "MantElect02auxiliar.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5760
         Width           =   495
      End
      Begin VB.TextBox Text2i 
         Height          =   285
         Left            =   9960
         MaxLength       =   60
         TabIndex        =   39
         Top             =   5880
         Width           =   7575
      End
      Begin VB.Frame frameEgreso 
         Caption         =   "Egresos por Orden de Trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   37
         Top             =   6360
         Width           =   26295
         Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
            Height          =   2775
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   4895
            _Version        =   327680
            Cols            =   7
         End
      End
      Begin VB.Frame frameProdDispo 
         Caption         =   "Productos Disponibles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5295
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   26295
         Begin MSFlexGridLib.MSFlexGrid FlexProd 
            Height          =   4695
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   25875
            _ExtentX        =   45641
            _ExtentY        =   8281
            _Version        =   327680
            Cols            =   7
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
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
         Left            =   360
         TabIndex        =   43
         Top             =   5880
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Motivo"
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
         Index           =   14
         Left            =   3480
         TabIndex        =   42
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   8400
         TabIndex        =   41
         Top             =   5880
         UseMnemonic     =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FrameGral"
      Height          =   1295
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   26775
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   16200
         MaxLength       =   5
         TabIndex        =   19
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   15480
         MaxLength       =   5
         TabIndex        =   18
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   9360
         MaxLength       =   3
         TabIndex        =   17
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   6360
         MaxLength       =   150
         TabIndex        =   16
         Top             =   1320
         Width           =   8660
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   14
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   9960
         MaxLength       =   90
         TabIndex        =   13
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   16320
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   17400
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   19200
         MaxLength       =   5
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   15240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2280
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   18120
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   7740
         Left            =   240
         TabIndex        =   45
         Top             =   3360
         Width           =   26355
         _ExtentX        =   46487
         _ExtentY        =   13653
         _Version        =   327680
         Cols            =   16
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Horas"
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
         Left            =   16245
         TabIndex        =   33
         Top             =   2040
         Width           =   510
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   16200
         X2              =   19800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   360
         X2              =   19805
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   8
         X1              =   3360
         X2              =   3360
         Y1              =   1920
         Y2              =   2760
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
         Left            =   480
         TabIndex        =   32
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha / Hora Solicit."
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
         Left            =   1155
         TabIndex        =   31
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lugar"
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
         Left            =   3240
         TabIndex        =   30
         Top             =   1080
         Width           =   495
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
         Index           =   3
         Left            =   6360
         TabIndex        =   29
         Top             =   1080
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
         Index           =   4
         Left            =   15270
         TabIndex        =   28
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha / Hora Inicio"
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
         Left            =   16290
         TabIndex        =   27
         Top             =   1320
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Segunda Descripcion"
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
         Left            =   10080
         TabIndex        =   26
         Top             =   2040
         Width           =   1830
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Unid"
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
         Left            =   9360
         TabIndex        =   25
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cant."
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
         Left            =   15510
         TabIndex        =   24
         Top             =   2040
         Width           =   465
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   3000
         X2              =   3000
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   19805
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   360
         X2              =   360
         Y1              =   960
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   1080
         X2              =   1080
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   18000
         X2              =   18000
         Y1              =   1200
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   6240
         X2              =   6240
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   15120
         X2              =   15120
         Y1              =   960
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   12
         X1              =   360
         X2              =   19805
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   9240
         X2              =   9240
         Y1              =   1920
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   19800
         X2              =   19800
         Y1              =   960
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   16200
         X2              =   16200
         Y1              =   960
         Y2              =   1920
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
         Index           =   8
         Left            =   555
         TabIndex        =   23
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Rubro"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha / Hora Fin"
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
         Left            =   18150
         TabIndex        =   21
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Asistencia"
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
         Left            =   17640
         TabIndex        =   20
         Top             =   1005
         Width           =   885
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   9840
         X2              =   9840
         Y1              =   1920
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   15360
         X2              =   15360
         Y1              =   1920
         Y2              =   2760
      End
      Begin VB.Line Line1 
         Index           =   16
         X1              =   16080
         X2              =   16080
         Y1              =   1920
         Y2              =   2760
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   12120
      TabIndex        =   1
      Top             =   12720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   0
      Top             =   12720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Registro de Reparaciones de Trabajo"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   26835
   End
End
Attribute VB_Name = "MantElect02auxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clMantElect
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Dim mRenglon As Integer
Dim mObjLuser As New clLogUser

Dim mObjInv As New clInven
Dim mRenglonProducto As Integer
Private Type Movimiento
    CodProducto As String
    Cantidad As Double
End Type

Dim vMovimientos() As Movimiento
Dim contDim As Integer

Dim contDimMObra As Integer
Dim mRenglonMObraDispo As Integer
Dim mRenglonMOAsignada As Integer

Dim mRenglonVehDispo As Integer
Dim mRenglonVehAsignado As Integer

Private Sub Combo1_Click(Index As Integer)
Select Case Index
   Case 2
      Combo1(3).Clear
      Set mRec = mObj.oEjecutarSelect("SELECT * FROM SubRubros WHERE CodRubro = '" & Left(Combo1(2).Text, 8) & "' AND FechaBaja IS NULL ORDER BY Codigo")
      If Not mRec.EOF Then
         Do While Not mRec.EOF
            Combo1(3).AddItem mRec!Codigo & "-" & mRec!descripcion
            mRec.MoveNext
         Loop
      End If
      mRec.Close
   Case 3
      Set mRec = mObj.oEjecutarSelect("SELECT * FROM SubRubros WHERE CodRubro = '" & Left(Combo1(2).Text, 8) & "' AND FechaBaja IS NULL ORDER BY Codigo")
      Text1(7).Text = mRec!Unidad
      mRec.Close
End Select
End Sub


Private Sub Command1_Click(Index As Integer)
Dim mEstado As String
Dim mFecPro As String
Dim mFecTer As String
Dim mOkGrb As Boolean
Dim mEstadoAnt As String
Dim mTextoMail As String
Dim mErrMail As Integer
Dim mListaDestinatarios As String
Dim mSectorAire As String

If Index = 0 Then
   If fValida1 Then
      If MsgBox("¿Está Seguro de Grabar esta Orden?", vbYesNo, sMessage) = vbYes Then
         mEstadoAnt = MSFlexGrid1.TextMatrix(mRenglon, 14)
         mEstado = "P"
         mFecPro = Now
         mFecTer = ""
         mOkGrb = True
         If MsgBox("¿Está terminado el trabajo?", vbYesNo, sMessage) = vbYes Then
            mOkGrb = False
            mEstado = "T"
            mFecTer = IIf(MSFlexGrid1.TextMatrix(mRenglon, 14) = "G", mFecPro, Now)
            If fValida2 Then
               mOkGrb = True
            End If
            If mOkGrb Then
               If DateDiff("n", CDate(Text1(1).Text), CDate(Text1(3).Text & " " & Text1(4).Text & ":00")) <= 0 Then
                  mOkGrb = False
                  MsgBox "Verificar la fecha de Asistencia", vbCritical, "Atención"
               End If
            End If
         End If

         If mOkGrb Then
            'Completo el FlexGrid
            MSFlexGrid1.TextMatrix(mRenglon, 6) = Text1(3).Text & " " & Text1(4).Text & ":00"
            MSFlexGrid1.TextMatrix(mRenglon, 7) = IIf(Text1(5).Text <> "", Text1(5).Text & " " & Text1(6).Text & ":00", "")
            MSFlexGrid1.TextMatrix(mRenglon, 8) = Text1(8).Text
            MSFlexGrid1.TextMatrix(mRenglon, 9) = Combo1(2).Text
            MSFlexGrid1.TextMatrix(mRenglon, 10) = Combo1(3).Text
            MSFlexGrid1.TextMatrix(mRenglon, 11) = Text1(7).Text
            MSFlexGrid1.TextMatrix(mRenglon, 12) = Text1(9).Text
            MSFlexGrid1.TextMatrix(mRenglon, 13) = Text1(10).Text
            MSFlexGrid1.TextMatrix(mRenglon, 14) = mEstado

            mSectorAire = IIf(MSFlexGrid1.TextMatrix(mRenglon, 15) = "Si", "1", "0")
            
            'Actualizo en Registros
            'mObj.UpdRegistros Text1(3).Text & " " & Text1(4).Text & ":00", IIf(Text1(5).Text <> "", Text1(5).Text & " " & Text1(6).Text & ":00", ""), Text1(8).Text, Left(Combo1(2).Text, 8), Left(Combo1(3).Text, 6), IIf(Text1(9).Text <> "", Text1(9).Text, ""), IIf(Text1(10).Text <> "", Text1(10).Text, ""), mEstado, IIf(mEstadoAnt = "G", Trim(Right(MDI.mUser, 20)), ""), IIf(mEstadoAnt = "G", mFecPro, ""), IIf(mEstado = "T", Trim(Right(MDI.mUser, 20)), ""), IIf(mFecTer <> "", mFecTer, ""), Text1(0).Text
            If mEstado = "T" Then
               mErrMail = 0
               
               mTextoMail = vbCrLf & "Se ha resuelto el Parte  " & Text1(0).Text & " de Mantenimiento Eléctrico: " & vbCrLf & vbCrLf & "     Descripción de la solicitud:  " & Text1(2).Text & vbCrLf & vbCrLf & "Verifique el servicio realizado. Gracias"
               Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT Email  FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And FechaBaja IS NULL UNION SELECT DISTINCT Email FROM MailsxElectrico WHERE SectorAire = " & mSectorAire & " AND FechaBaja IS NULL ")
               'Set mRec = mObj.oEjecutarSelect("SELECT * FROM MailsxSuperv WHERE CodSuperv = '" & mObj.ObtCodSuperv(Text1(0).Text) & "' And Email <> '" & mObjLuser.sCampoDescrip("USUARIOS", "CodUsuario = '" & Trim(Right(MDI.mUser, 20)) & "'", 6) & "' And FechaBaja IS NULL")
 
               If Not mRec.EOF Then
                  mListaDestinatarios = ""
                  Do While Not mRec.EOF
                     mListaDestinatarios = mListaDestinatarios & mRec!Email & ";"
                     mRec.MoveNext
                  Loop
'                  If Not fEnviar_Mail_CDO("", mListaDestinatarios, "ausolmail@ausol.com.ar", "Repuesta a Solicitud de Servicios", mTextoMail, "", "587", "ausolmail@ausol.com.ar", "sgedosmildiecisiete1$", True, False) Then
'                     mErrMail = mErrMail + 1
'                  End If
               End If
               If mErrMail = 0 Then
                  MsgBox "Se ha grabado la solicitud y enviado el mensaje correctamente a los responsables.", vbInformation, "Atención"
               Else
                  MsgBox "Se ha grabado la solicitud correctamente, pero se NO ha enviado el correo correctamente", vbExclamation, "Atención"
               End If
            End If
         End If
      End If
   End If
Else
   Unload Me
End If
End Sub

Private Sub FlexMoAsignada_Click()
   Dim mI As Integer
   
   If FlexMoAsignada.MouseRow > 0 Then
   
      If mRenglonMOAsignada <> 0 Then
         FlexMoAsignada.Row = mRenglonMOAsignada
         For mI = 1 To FlexMoAsignada.Cols - 1
            FlexMoAsignada.Col = mI
            FlexMoAsignada.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonMOAsignada = FlexMoAsignada.MouseRow
   
      FlexMoAsignada.Row = mRenglonMOAsignada
      For mI = 1 To FlexMoAsignada.Cols - 1
         FlexMoAsignada.Col = mI
         FlexMoAsignada.CellBackColor = &H8000000D
      Next
   Else
      
      FlexMoAsignada.Row = mRenglonMOAsignada
      For mI = 1 To FlexMoAsignada.Cols - 1
         FlexMoAsignada.Col = mI
         FlexMoAsignada.CellBackColor = vbWhite
      Next
      
      mRenglonMOAsignada = 0
     
   End If

End Sub

Private Sub FlexProd_Click()

   Dim mI As Integer
   If FlexProd.MouseRow > 0 Then
   
      If mRenglonProducto <> 0 Then
         FlexProd.Row = mRenglonProducto
         For mI = 1 To FlexProd.Cols - 1
            FlexProd.Col = mI
            FlexProd.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonProducto = FlexProd.MouseRow
   
      FlexProd.Row = mRenglonProducto
      For mI = 1 To FlexProd.Cols - 1
         FlexProd.Col = mI
         FlexProd.CellBackColor = &H8000000D
      Next
   Else
     
      FlexProd.Row = mRenglonProducto
      For mI = 1 To FlexProd.Cols - 1
         FlexProd.Col = mI
         FlexProd.CellBackColor = vbWhite
      Next
      
      mRenglonProducto = 0
      
   End If
   
End Sub


Private Sub Form_Load()
Dim mI As Integer
MantElect02auxiliar.Top = 100
MantElect02auxiliar.Left = (MDI.Width - MantElect02auxiliar.Width) / 2


 Frame1(0).Visible = True
 Frame1(1).Visible = False
 Frame1(2).Visible = False
 Frame1(3).Visible = False


Set mRec = mObj.oEjecutarSelect("SELECT * FROM Edificios WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(0).AddItem mRec!Tramo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

Combo1(1).AddItem "Alta"
Combo1(1).AddItem "Media"
Combo1(1).AddItem "Baja"

Set mRec = mObj.oEjecutarSelect("SELECT * FROM Rubros WHERE FechaBaja IS NULL")
If Not mRec.EOF Then
   Do While Not mRec.EOF
      Combo1(2).AddItem mRec!Codigo & " - " & mRec!descripcion
      mRec.MoveNext
   Loop
End If
mRec.Close

MSFlexGrid1.ColWidth(0) = 200
MSFlexGrid1.ColWidth(1) = 500
MSFlexGrid1.ColWidth(2) = 1700
MSFlexGrid1.ColWidth(3) = 3000
MSFlexGrid1.ColWidth(4) = 4000
MSFlexGrid1.ColWidth(5) = 750
MSFlexGrid1.ColWidth(6) = 1700
MSFlexGrid1.ColWidth(7) = 1700
MSFlexGrid1.ColWidth(8) = 4000
MSFlexGrid1.ColWidth(9) = 2200
MSFlexGrid1.ColWidth(10) = 4000
MSFlexGrid1.ColWidth(11) = 500
MSFlexGrid1.ColWidth(12) = 500
MSFlexGrid1.ColWidth(13) = 600
MSFlexGrid1.ColWidth(14) = 400
MSFlexGrid1.ColWidth(15) = 0

For mI = 0 To MSFlexGrid1.Cols - 1
   MSFlexGrid1.ColAlignment(mI) = 2
Next

MSFlexGrid1.TextMatrix(0, 1) = "Parte"
MSFlexGrid1.TextMatrix(0, 2) = "Fecha Solicitud"
MSFlexGrid1.TextMatrix(0, 3) = "Lugar"
MSFlexGrid1.TextMatrix(0, 4) = "Descripcion de la Solicitud"
MSFlexGrid1.TextMatrix(0, 5) = "Prioridad"
MSFlexGrid1.TextMatrix(0, 6) = "Fecha Ini. Asist."
MSFlexGrid1.TextMatrix(0, 7) = "Fecha Fin Asist."
MSFlexGrid1.TextMatrix(0, 8) = "Segunda Descripcion"
MSFlexGrid1.TextMatrix(0, 9) = "Rubro"
MSFlexGrid1.TextMatrix(0, 10) = "Sub Rubro"
MSFlexGrid1.TextMatrix(0, 11) = "Unid."
MSFlexGrid1.TextMatrix(0, 12) = "Cant."
MSFlexGrid1.TextMatrix(0, 13) = "Horas"
MSFlexGrid1.TextMatrix(0, 14) = "Est."
MSFlexGrid1.TextMatrix(0, 15) = "Sector Aire"

Set mRec = mObj.oEjecutarSelect("SELECT R.* " & _
                                    "FROM Registros R " & _
                                        "Inner Join " & _
                                    "MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                "WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "';")
If Not mRec.EOF Then
   mI = 1
   Do While Not mRec.EOF
      mI = mI + 1
      MSFlexGrid1.AddItem ""
      MSFlexGrid1.TextMatrix(mI, 1) = mRec!Parte
      MSFlexGrid1.TextMatrix(mI, 2) = NVL(mRec!FechaSolic, "")
      MSFlexGrid1.TextMatrix(mI, 3) = NVL(mRec!CodEdificio, "")
      MSFlexGrid1.TextMatrix(mI, 4) = NVL(mRec!descripcion, "")
      MSFlexGrid1.TextMatrix(mI, 5) = NVL(mRec!Prioridad, "")
      MSFlexGrid1.TextMatrix(mI, 6) = NVL(mRec!FechaIniAsist, "")
      MSFlexGrid1.TextMatrix(mI, 7) = NVL(mRec!FechaFinAsist, "")
      MSFlexGrid1.TextMatrix(mI, 8) = NVL(mRec!SegundaDesc, "")
      MSFlexGrid1.TextMatrix(mI, 9) = NVL(mRec!Rubro, "") & " - " & mObj.sCampoDescrip("Rubros", "Codigo = '" & mRec!Rubro & "'", 1)
      MSFlexGrid1.TextMatrix(mI, 10) = NVL(mRec!SubRubro, "") & "-" & mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 2)
      MSFlexGrid1.TextMatrix(mI, 11) = mObj.sCampoDescrip("SubRubros", "Codigo = '" & mRec!SubRubro & "'", 3)
      MSFlexGrid1.TextMatrix(mI, 12) = NVL(mRec!Cantidad, "")
      MSFlexGrid1.TextMatrix(mI, 13) = NVL(mRec!Horas, "")
      MSFlexGrid1.TextMatrix(mI, 14) = NVL(mRec!estado, "")
      MSFlexGrid1.TextMatrix(mI, 15) = IIf(mRec!SectorAire = 1, "Si", "No")
      
      mRec.MoveNext
   Loop
   MSFlexGrid1.RemoveItem 1
End If
mRec.Close

Text1(0).Enabled = False
Text1(1).Enabled = False

Text1(2).Enabled = True
Combo1(0).Enabled = False
Combo1(1).Enabled = False
Text1(7).Enabled = False
Text1(10).Enabled = False

InicioInventario
InicioManoObra
InicioVehiculos

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mRec = Nothing
ShowMenu 47, True, False
End Sub

Private Function fValida1() As Boolean
Dim mRet As Boolean
Dim mI As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = Fecha_ok(Text1(3).Text)
   If mRet Then
      mRet = Hora_ok(Text1(4).Text)
   End If
   If mRet Then
      mRet = DateDiff("s", CDate(Text1(1).Text), CDate(Text1(3).Text & " " & Text1(4).Text & ":00")) > 0
   End If
   If mRet Then
      mRet = (Combo1(2).Text <> "")
   End If
   If mRet Then
      mRet = (Combo1(3).Text <> "")
   End If
   If mRet Then
      mRet = (Text1(8).Text <> "")
   End If
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida1 = mRet
End Function



Private Sub MSFlexGrid1_Click()
Dim mI As Integer
Dim mJ As Integer
Dim mFound As Boolean
Dim mHoraIniAsist As String
Dim mHoraFinAsist As String

If MSFlexGrid1.MouseCol = 0 And MSFlexGrid1.MouseRow > 0 Then
   mRenglon = MSFlexGrid1.MouseRow
   Text1(0).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
   Text1(1).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 2)
   For mI = 0 To Combo1(0).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 3) = Combo1(0).List(mI) Then
         Combo1(0).ListIndex = mI
      End If
   Next
   Text1(2).Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 4)
   For mI = 0 To Combo1(1).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 5) = Combo1(1).List(mI) Then
         Combo1(1).ListIndex = mI
      End If
   Next
   Text1(3).Text = NVL(Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 10), "")
   
   
   mHoraIniAsist = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 24), "")
   If mHoraIniAsist <> "" Then
      mHoraIniAsist = Format(mHoraIniAsist, "hh:mm")
   End If
   Text1(4).Text = mHoraIniAsist
   'Text1(4).Text = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 6), 8), "")
   
   
   
   
   Text1(5).Text = NVL(Left(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 10), "")
   
   
   mHoraFinAsist = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 24), "")
   If mHoraFinAsist <> "" Then
      mHoraFinAsist = Format(mHoraFinAsist, "hh:mm")
   End If
   Text1(6).Text = mHoraFinAsist
   'Text1(6).Text = NVL(Right(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 7), 8), "")
   
   mFound = False
   For mI = 0 To Combo1(2).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 9) = Combo1(2).List(mI) Then
         mFound = True
         Combo1(2).ListIndex = mI
      End If
   Next
   If Not mFound Then
      Combo1(2).ListIndex = -1
   End If
   
   mFound = False
   For mI = 0 To Combo1(3).ListCount - 1
      If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 10) = Combo1(3).List(mI) Then
         mFound = True
         Combo1(3).ListIndex = mI
      End If
   Next
   If Not mFound Then
      Combo1(3).ListIndex = -1
   End If
   
   Text1(7).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 11), "")
   Text1(8).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 8), "")
   Text1(9).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 12), "")
   Text1(10).Text = NVL(MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 13), "")
   
   ActualizaGridMoSegunParte (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))
   ActualizaGridVeSegunParte (MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1))

Else
   mRenglon = 0
End If
End Sub

Private Sub TabStrip1_Click()
   Dim i As Integer
   Dim j As Integer
    
    i = TabStrip1.SelectedItem.Index
         
   
   For j = 1 To TabStrip1.Tabs.Count
      If j = i Then
         Frame1(j - 1).Visible = True
         
      Else
         Frame1(j - 1).Visible = False
      End If
   Next
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
   Case 3, 5
      KeyAscii = fDateKeyPress(Text1(Index), KeyAscii)
   Case 4, 6
      KeyAscii = fHoraKeyPress(Text1(Index), KeyAscii)
   Case 8
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   Case 9, 10
      KeyAscii = fNumDoubleKeyPress(KeyAscii)
End Select
End Sub

Private Function fValida2() As Boolean
Dim mRet As Boolean
Dim mI As Integer
mRet = mRenglon <> 0
If mRet Then
   mRet = Fecha_ok(Text1(5).Text)
   If mRet Then
      mRet = Hora_ok(Text1(6).Text)
   End If
   If mRet Then
      mRet = DateDiff("s", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) >= 0
   End If
   If mRet Then
      mRet = (Text1(9).Text <> "")
   End If
   If mRet Then
      mRet = (Text1(10).Text <> "")
   End If
   If Not mRet Then
      MsgBox "Verifique que los datos estén ingresados correctamente", vbCritical, "Atención"
   End If
Else
   MsgBox "Seleccione un pedido de la grilla", vbCritical, "Atención"
End If
fValida2 = mRet
End Function

Private Sub Text1_LostFocus(Index As Integer)
Dim mRet As Boolean
Select Case Index
   Case 3, 4, 5, 6
      mRet = (Text1(3).Text <> "" And Text1(4).Text <> "" And Text1(5).Text <> "" And Text1(6).Text <> "")
      If mRet Then
         If DateDiff("s", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) >= 0 Then
            'Text1(10).Text = Redondeo(DateDiff("n", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) / 60, 2)
            Text1(10).Text = Replace(Redondeo(DateDiff("n", CDate(Text1(3).Text & " " & Text1(4).Text & ":00"), CDate(Text1(5).Text & " " & Text1(6).Text & ":00")) / 60, 2), ",", ".")
         Else
            MsgBox "Verifique las fechas de Asistencia", vbCritical, "Atención"
            Text1(Index).Text = ""
            Text1(Index).SetFocus
         End If
      End If
End Select
End Sub


'*************************************************************************************************************************************************************

'---------------------------------------------------------------------FRAME PRODCUTOS(INVENTARIO)-------------------------------------------------------------

'*************************************************************************************************************************************************************

Private Sub InicioInventario()
  Dim mI As Integer
  
  contDim = 0
  
  '--------------------------------------------------GRILLA PRODUCTOS DISPONIBLES---------------------------------------------------------------
   FlexProd.ColWidth(0) = 200
   FlexProd.ColWidth(1) = 1000
   FlexProd.ColWidth(2) = 16460
   FlexProd.ColWidth(3) = 1200
   FlexProd.ColWidth(4) = 1200
   FlexProd.ColWidth(5) = 2000
   FlexProd.ColWidth(6) = 3000
   
   FlexProd.TextMatrix(0, 1) = "Código"
   FlexProd.TextMatrix(0, 2) = "Descripcion"
   FlexProd.TextMatrix(0, 3) = "Stock"
   FlexProd.TextMatrix(0, 4) = "Stock Mínimo"
   FlexProd.TextMatrix(0, 5) = "Unidad de Medida"
   FlexProd.TextMatrix(0, 6) = "Sector"
   
   Set mRec = mObjInv.oEjecutarSelect("SELECT P.Codigo, P.Descripcion, P.Stock, P.Stock_Min, U.Descripcion AS UnidadMedida, S.Descripcion AS Sector " & _
     " From " & _
     " Producto P  " & _
     " Inner Join  " & _
     " UnidadMedida U ON P.CodUnidadMedida = U.Codigo  " & _
     " Inner Join  " & _
     " Sector S ON P.CodSector = S.Codigo " & _
     " where P.Fecha_Baja is null " & _
     " ORDER BY Codigo; ")

   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         mI = mI + 1

         FlexProd.AddItem ""
         FlexProd.TextMatrix(mI, 1) = mRec!Codigo
         FlexProd.TextMatrix(mI, 2) = mRec!descripcion
         FlexProd.TextMatrix(mI, 3) = mRec!Stock
         FlexProd.TextMatrix(mI, 4) = mRec!Stock_Min
         FlexProd.TextMatrix(mI, 5) = mRec!UnidadMedida
         FlexProd.TextMatrix(mI, 6) = mRec!Sector

         mRec.MoveNext
      Loop
      FlexProd.RemoveItem 1
   End If
   mRec.Close
   
   '-FIN: GRILLA PRODUCTOS DISPONIBLES-----------------------------------------------------------------------------------------------------------------

   '-------------------------------------------------GRILLA PRODUCTOS CONSUMIDOS----------------------------------------------------------------------
   FlexEgreso.ColWidth(0) = 200
   FlexEgreso.ColWidth(1) = 1000
   FlexEgreso.ColWidth(2) = 12400
   FlexEgreso.ColWidth(3) = 1200
   FlexEgreso.ColWidth(4) = 3240
   FlexEgreso.ColWidth(5) = 7000
   FlexEgreso.ColWidth(6) = 0
   
   FlexEgreso.TextMatrix(0, 1) = "Código"
   FlexEgreso.TextMatrix(0, 2) = "Descripcion"
   FlexEgreso.TextMatrix(0, 3) = "Cantidad"
   FlexEgreso.TextMatrix(0, 4) = "Motivo"
   FlexEgreso.TextMatrix(0, 5) = "Observaciones"
   FlexEgreso.TextMatrix(0, 6) = "CodMotivo"
   
   FlexEgreso.ColAlignment(4) = 2
  '--------------------------------------------------FIN GRILLA PRODUCTOS CONSUMIDOS---------------------------------------------------------------
  
  '--CARGO COMBO MOTIVOS EGRESO
   
   Set mRec = mObjInv.oTabla("MotivosEgreso", "")
   Do While Not mRec.EOF
      Combo1i.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   
   '--FIN: CARGO COMBO MOTIVOS
  
  
End Sub

Private Sub Command2_Click()
   Dim i As Integer

   If fValidaEgreso() Then
      ReDim Preserve vMovimientos(0 To contDim) As Movimiento
      
      vMovimientos(contDim).CodProducto = FlexProd.TextMatrix(mRenglonProducto, 1)
      vMovimientos(contDim).Cantidad = CDbl(Replace(Trim(Text1i.Text), ".", ","))
      
      FlexEgreso.AddItem vbTab & FlexProd.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 2) & vbTab & Text1i.Text & vbTab & Combo1i.Text & vbTab & Text2i.Text & vbTab & Left(Combo1i.Text, 2)
      If FlexEgreso.TextMatrix(1, 1) = "" Then
         FlexEgreso.RemoveItem 1
      End If
      contDim = contDim + 1
      Text1i.Text = ""
   End If
End Sub

Private Function fValidaEgreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim j As Integer
   Dim mCantidadMovida As Double
   Dim mCantidaStock As Double
   Dim mCodProducto As String
   
   mRet = True
      
   If mRenglonProducto = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un producto de la grilla superior"
   End If
      
   If mRet Then
      If mRenglonProducto <> 0 And FlexProd.TextMatrix(mRenglonProducto, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un producto de la grilla superior"
      End If
   End If
      
      
   If mRet Then
      If Trim(Text1i.Text) = "" Or Trim(Combo1i.Text) = "" Or Trim(Text2i.Text) = "" Then
         mRet = False
         mMensajeError = "Debe completar todos los datos"
      End If
   End If
      
   If mRet Then
      If Not IsNumeric(Replace(Text1i.Text, ".", ",")) Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no es un valor numérico"
      End If
   End If
      
   'Vvalido si el saldo del stock es insuficiente
   If mRet Then
   
      mCantidadMovida = 0
      mCodProducto = FlexProd.TextMatrix(mRenglonProducto, 1)
   
      mCantidaStock = mObjInv.sTablaDescr("Producto", "Codigo = '" & mCodProducto & "'", 4)
      
      If contDim <> 0 Then
         For j = 0 To contDim - 1
           If vMovimientos(j).CodProducto = mCodProducto Then
              mCantidadMovida = mCantidadMovida + vMovimientos(j).Cantidad
           End If
         Next
      End If
      
      mCantidadMovida = mCantidadMovida + CDbl(Replace(Trim(Text1i.Text), ".", ","))
      
      If mCantidaStock < mCantidadMovida Then
         mRet = False
         mMensajeError = "Stock disponible insuficiente"
      End If
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaEgreso = mRet
End Function



'*************************************************************************************************************************************************************

'---------------------------------------------------------------------FRAME MANO DE OBRA----------------------------------------------------------------------

'*************************************************************************************************************************************************************


Private Sub InicioManoObra()
  
  Dim mI As Integer
  contDimMObra = 0
  
  '--------------------------------------------------GRILLA MANO OBRA DISPONIBLE---------------------------------------------------------------
   FlexMoDisponible.ColWidth(0) = 200
   FlexMoDisponible.ColWidth(1) = 24860
   FlexMoDisponible.ColWidth(2) = 0
   
   FlexMoDisponible.TextMatrix(0, 1) = "Técnico Disponible"
   FlexMoDisponible.TextMatrix(0, 2) = "Codigo"
   
   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja is null ORDER BY Descripcion;")

   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         mI = mI + 1

         FlexMoDisponible.AddItem ""
         FlexMoDisponible.TextMatrix(mI, 1) = mRec!descripcion
         FlexMoDisponible.TextMatrix(mI, 2) = mRec!Codigo

         mRec.MoveNext
         
         
      Loop
      FlexMoDisponible.RemoveItem 1
      
   End If
   mRec.Close
   '----------------------------------------------FIN: GRILLA MANO OBRA DISPONIBLE--------------------------------------------------------------



   '-------------------------------------------------GRILLA MANO OBRA ASIGNADA-------------------------------------------------------------------------
   FlexMoAsignada.ColWidth(0) = 200
   FlexMoAsignada.ColWidth(1) = 24860
   FlexMoAsignada.ColWidth(2) = 0

   FlexMoAsignada.TextMatrix(0, 1) = "Técnico Asignado"
   FlexMoAsignada.TextMatrix(0, 2) = "Codigo"
  
  '--------------------------------------------------FIN GRILLA MANO OBRA ASIGNADA-----------------------------------------------------------------------

End Sub

Private Sub FlexMoDisponible_Click()
   Dim mI As Integer
   
   If FlexMoDisponible.MouseRow > 0 Then
   
      If mRenglonMObraDispo <> 0 Then
         FlexMoDisponible.Row = mRenglonMObraDispo
         For mI = 1 To FlexMoDisponible.Cols - 1
            FlexMoDisponible.Col = mI
            FlexMoDisponible.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonMObraDispo = FlexMoDisponible.MouseRow
   
      FlexMoDisponible.Row = mRenglonMObraDispo
      For mI = 1 To FlexMoDisponible.Cols - 1
         FlexMoDisponible.Col = mI
         FlexMoDisponible.CellBackColor = &H8000000D
      Next
   Else
      FlexMoDisponible.Row = mRenglonMObraDispo
      For mI = 1 To FlexMoDisponible.Cols - 1
         FlexMoDisponible.Col = mI
         FlexMoDisponible.CellBackColor = vbWhite
      Next
      
      mRenglonMObraDispo = 0
   End If
End Sub


Private Sub Command2mo_Click(Index As Integer)
   If Index = 0 Then
      If mRenglonMObraDispo > 0 Then
         If Trim(FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 1)) <> "" Then
            FlexMoAsignada.AddItem vbTab & FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 1) & vbTab & FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 2)
            
            If FlexMoAsignada.TextMatrix(1, 2) = "" Then
               FlexMoAsignada.RemoveItem 1
            End If
         End If
         
         If FlexMoDisponible.Rows > 2 Then
            FlexMoDisponible.RemoveItem mRenglonMObraDispo
         
            mRenglonMObraDispo = 0
         Else
            If Trim(FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 1)) <> "" Then
               FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 1) = ""
               FlexMoDisponible.TextMatrix(mRenglonMObraDispo, 2) = ""
         
               mRenglonMObraDispo = 0
            End If
         End If
      End If
   Else
      If mRenglonMOAsignada > 0 Then
         
         If Trim(FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 1)) <> "" Then
            FlexMoDisponible.AddItem vbTab & FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 1) & vbTab & FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 2)
            
            If FlexMoDisponible.TextMatrix(1, 2) = "" Then
               FlexMoDisponible.RemoveItem 1
            End If
         End If
         
         If FlexMoAsignada.Rows > 2 Then
            FlexMoAsignada.RemoveItem mRenglonMOAsignada

            mRenglonMOAsignada = 0
         Else
            If Trim(FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 1)) <> "" Then
               FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 1) = ""
               FlexMoAsignada.TextMatrix(mRenglonMOAsignada, 2) = ""
         
               mRenglonMOAsignada = 0
            End If
         End If
      End If
   End If

  
End Sub

Private Sub ActualizaGridMoSegunParte(pParte As Double)
   Dim mI As Integer
   Dim mJ As Integer
   
      'Limpio la Grilla modispo
      For mI = 1 To FlexMoDisponible.Rows
         
         If FlexMoDisponible.Rows > 2 Then
            FlexMoDisponible.RemoveItem 1
         Else
            FlexMoDisponible.TextMatrix(1, 1) = ""
            FlexMoDisponible.TextMatrix(1, 2) = ""
         End If
      Next
   
      'Limpio la Grilla MoAsignada
      For mI = 1 To FlexMoAsignada.Rows
         
         If FlexMoAsignada.Rows > 2 Then
            FlexMoAsignada.RemoveItem 1
         Else
            FlexMoAsignada.TextMatrix(1, 1) = ""
            FlexMoAsignada.TextMatrix(1, 2) = ""
         End If
      Next
      
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja is null ORDER BY Descripcion;")

      If Not mRec.EOF Then
         mI = 1
         Do While Not mRec.EOF
            mI = mI + 1

            FlexMoDisponible.AddItem ""
            FlexMoDisponible.TextMatrix(mI, 1) = mRec!descripcion
            FlexMoDisponible.TextMatrix(mI, 2) = mRec!Codigo

            mRec.MoveNext
         Loop
         FlexMoDisponible.RemoveItem 1
      End If
      mRec.Close


      Set mRec = mObj.oEjecutarSelect(" SELECT M.Codigo, M.Descripcion " & _
                                      " From " & _
                                      " MO_Tecnicos M " & _
                                          " Inner Join" & _
                                      " Partes_MoTecnicos P " & _
                                          " ON P.CodMoTecnico = M.Codigo " & _
                                       " Where P.Parte = " & pParte & " " & _
                                       " ORDER BY M.Descripcion ;")
   
      mJ = 1
      Do While Not mRec.EOF

         mJ = mJ + 1
         FlexMoAsignada.AddItem ""
         FlexMoAsignada.TextMatrix(mJ, 1) = mRec!descripcion
         FlexMoAsignada.TextMatrix(mJ, 2) = mRec!Codigo
   
      
   
         For mI = 1 To FlexMoDisponible.Rows
            If Trim(FlexMoDisponible.TextMatrix(mI, 2)) = Trim(mRec!Codigo) Then
               
               If FlexMoDisponible.Rows > 2 Then
                  FlexMoDisponible.RemoveItem mI
               Else
                  'blanquear tantas veces como columnas tenga
                  FlexMoDisponible.TextMatrix(1, 1) = ""
                  FlexMoDisponible.TextMatrix(1, 2) = ""
               End If
               
               mI = 99
            End If
         Next
         
         mRec.MoveNext
      Loop
       
       
      If FlexMoAsignada.Rows > 2 Then
         FlexMoAsignada.RemoveItem 1
      Else
         'MsgBox "menorigual2"
      End If
      
      mRec.Close

  
End Sub

Private Sub GrabarMO_Click()

   GrabarMoTecnicos MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)

End Sub

Private Sub GrabarMoTecnicos(ByVal pParte As Double)
   Dim mI As Integer

   If Not mObj.xDeleteMoTecnicos(pParte) Then
      MsgBox "ERROR al borrar borrar Mano Obra(técnicos)...", vbCritical, sMessage
   End If
   
   For mI = 1 To FlexMoAsignada.Rows - 1
      If Trim(FlexMoAsignada.TextMatrix(mI, 2)) <> "" Then
            mObj.xInsertMoTecnico pParte, Trim(FlexMoAsignada.TextMatrix(mI, 2))
      End If
   Next
End Sub


Private Sub ActualizaGridMoSegunParte2(pParte As Double)
   Dim mI As Integer
   Dim mJ As Integer
   
     
      'Limpio la Grilla MoDispo
      For mI = 1 To FlexMoDisponible.Rows
         
         If FlexMoDisponible.Rows > 2 Then
            FlexMoDisponible.RemoveItem 1
         Else
            FlexMoDisponible.TextMatrix(1, 1) = ""
            FlexMoDisponible.TextMatrix(1, 2) = ""
         End If
      Next
   
      'Limpio la Grilla MoAsignado
      For mI = 1 To FlexMoAsignada.Rows
         
         If FlexMoAsignada.Rows > 2 Then
            FlexMoAsignada.RemoveItem 1
         Else
            FlexMoAsignada.TextMatrix(1, 1) = ""
            FlexMoAsignada.TextMatrix(1, 2) = ""
         End If
      Next
      
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja is null ORDER BY Descripcion;")

      If Not mRec.EOF Then
         mI = 1
         Do While Not mRec.EOF
            mI = mI + 1

            FlexMoDisponible.AddItem ""
            FlexMoDisponible.TextMatrix(mI, 1) = mRec!descripcion
            FlexMoDisponible.TextMatrix(mI, 2) = mRec!Codigo

            mRec.MoveNext
         Loop
         FlexMoDisponible.RemoveItem 1
      End If
      mRec.Close


      Set mRec = mObj.oEjecutarSelect(" SELECT M.Codigo, M.Descripcion " & _
                                      " From " & _
                                      " MO_Tecnicos M " & _
                                          " Inner Join" & _
                                      " Partes_MoTecnicos P " & _
                                          " ON P.CodMoTecnico = M.Codigo " & _
                                       " Where P.Parte = " & pParte & " " & _
                                       " ORDER BY M.Descripcion ;")
   
      mJ = 1
      Do While Not mRec.EOF
   
         mJ = mJ + 1
         FlexMoAsignada.AddItem ""
         FlexMoAsignada.TextMatrix(mJ, 1) = mRec!descripcion
         FlexMoAsignada.TextMatrix(mJ, 2) = mRec!Codigo
         
         For mI = 1 To FlexMoDisponible.Rows
            If Trim(FlexMoDisponible.TextMatrix(mI, 2)) = Trim(mRec!Codigo) Then
               
               If FlexMoDisponible.Rows > 2 Then
                  FlexMoDisponible.RemoveItem mI
               Else
                  'blanquear tantas veces como columnas tenga
                  FlexMoDisponible.TextMatrix(1, 1) = ""
                  FlexMoDisponible.TextMatrix(1, 2) = ""
               End If
               
               mI = 99
            End If
         Next
         
         mRec.MoveNext
      Loop
       
       
      If FlexMoAsignada.Rows > 2 Then
         FlexMoAsignada.RemoveItem 1
      Else
         'MsgBox "menorigual2"
      End If
      
      mRec.Close
End Sub







'*************************************************************************************************************************************************************

'---------------------------------------------------------------------FRAME VEHICULOS-------------------------------------------------------------------------

'*************************************************************************************************************************************************************



Private Sub InicioVehiculos()
  
  Dim mI As Integer
  
  '--------------------------------------------------GRILLA VEHICULO DISPONIBLE---------------------------------------------------------------
   FlexVehDisponible.ColWidth(0) = 200
   FlexVehDisponible.ColWidth(1) = 24860
   FlexVehDisponible.ColWidth(2) = 0
   
   FlexVehDisponible.TextMatrix(0, 1) = "Vehículo Disponible"
   FlexVehDisponible.TextMatrix(0, 2) = "Codigo"
   
   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Vehiculos Where Fecha_Baja is null ORDER BY Descripcion;")

   If Not mRec.EOF Then
      mI = 1
      Do While Not mRec.EOF
         mI = mI + 1

         FlexVehDisponible.AddItem ""
         FlexVehDisponible.TextMatrix(mI, 1) = mRec!descripcion
         FlexVehDisponible.TextMatrix(mI, 2) = mRec!Codigo
         
         mRec.MoveNext
      Loop
      FlexVehDisponible.RemoveItem 1
   End If
   mRec.Close
   
   '----------------------------------------------FIN: GRILLA VEHICULO DISPONIBLE--------------------------------------------------------------


   '-------------------------------------------------GRILLA VEHICULO ASIGNADO-------------------------------------------------------------------------
   FlexVehAsignado.ColWidth(0) = 200
   FlexVehAsignado.ColWidth(1) = 12430
   FlexVehAsignado.ColWidth(2) = 6215
   FlexVehAsignado.ColWidth(3) = 6215
   FlexVehAsignado.ColWidth(4) = 0

   FlexVehAsignado.TextMatrix(0, 1) = "Vehículo Asignado"
   FlexVehAsignado.TextMatrix(0, 2) = "Km Inicial"
   FlexVehAsignado.TextMatrix(0, 3) = "Km Final"
   FlexVehAsignado.TextMatrix(0, 4) = "Codigo"
  
  '--------------------------------------------------FIN GRILLA VEHICULO ASIGNADO-----------------------------------------------------------------------

End Sub

Private Sub FlexVehDisponible_Click()
   Dim mI As Integer
   
   If FlexVehDisponible.MouseRow > 0 Then
   
      If mRenglonVehDispo <> 0 Then
         FlexVehDisponible.Row = mRenglonVehDispo
         For mI = 1 To FlexVehDisponible.Cols - 1
            FlexVehDisponible.Col = mI
            FlexVehDisponible.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonVehDispo = FlexVehDisponible.MouseRow
   
      FlexVehDisponible.Row = mRenglonVehDispo
      For mI = 1 To FlexVehDisponible.Cols - 1
         FlexVehDisponible.Col = mI
         FlexVehDisponible.CellBackColor = &H8000000D
      Next
   Else
      FlexVehDisponible.Row = mRenglonVehDispo
      For mI = 1 To FlexVehDisponible.Cols - 1
         FlexVehDisponible.Col = mI
         FlexVehDisponible.CellBackColor = vbWhite
      Next
      
      mRenglonVehDispo = 0
   End If
End Sub

Private Sub FlexVehAsignado_Click()
   Dim mI As Integer
   
   If FlexVehAsignado.MouseRow > 0 Then
   
      If mRenglonVehAsignado <> 0 Then
         FlexVehAsignado.Row = mRenglonVehAsignado
         For mI = 1 To FlexVehAsignado.Cols - 1
            FlexVehAsignado.Col = mI
            FlexVehAsignado.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonVehAsignado = FlexVehAsignado.MouseRow
   
      FlexVehAsignado.Row = mRenglonVehAsignado
      For mI = 1 To FlexVehAsignado.Cols - 1
         FlexVehAsignado.Col = mI
         FlexVehAsignado.CellBackColor = &H8000000D
      Next
   Else
      
      FlexVehAsignado.Row = mRenglonVehAsignado
      For mI = 1 To FlexVehAsignado.Cols - 1
         FlexVehAsignado.Col = mI
         FlexVehAsignado.CellBackColor = vbWhite
      Next
      
      mRenglonVehAsignado = 0
     
   End If
End Sub

Private Sub Command2ve_Click(Index As Integer)
   
   If Index = 0 Then
      If fValidaVehiculo() Then
         If mRenglonVehDispo > 0 Then
            If Trim(FlexVehDisponible.TextMatrix(mRenglonVehDispo, 1)) <> "" Then
               
               
               
               FlexVehAsignado.AddItem vbTab & FlexVehDisponible.TextMatrix(mRenglonVehDispo, 1) & vbTab & Text1ve(0).Text & vbTab & Text1ve(1).Text & vbTab & FlexVehDisponible.TextMatrix(mRenglonVehDispo, 2)
               
               Text1ve(0).Text = ""
               Text1ve(1).Text = ""
               
               If FlexVehAsignado.TextMatrix(1, 2) = "" Then
                  FlexVehAsignado.RemoveItem 1
               End If
            End If
            
            If FlexVehDisponible.Rows > 2 Then
               FlexVehDisponible.RemoveItem mRenglonVehDispo
            
               mRenglonVehDispo = 0
            Else
               If Trim(FlexVehDisponible.TextMatrix(mRenglonVehDispo, 1)) <> "" Then
                  FlexVehDisponible.TextMatrix(mRenglonVehDispo, 1) = ""
                  FlexVehDisponible.TextMatrix(mRenglonVehDispo, 2) = ""
            
                  mRenglonVehDispo = 0
               End If
            End If
         End If
      End If
   Else
      If mRenglonVehAsignado > 0 Then
         
         If Trim(FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 1)) <> "" Then
            FlexVehDisponible.AddItem vbTab & FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 1) & vbTab & FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 2)
            
            If FlexVehDisponible.TextMatrix(1, 2) = "" Then
               FlexVehDisponible.RemoveItem 1
            End If
         End If
         
         If FlexVehAsignado.Rows > 2 Then
            FlexVehAsignado.RemoveItem mRenglonVehAsignado

            mRenglonVehAsignado = 0
         Else
            If Trim(FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 1)) <> "" Then
               FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 1) = ""
               FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 2) = ""
               FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 3) = ""
               FlexVehAsignado.TextMatrix(mRenglonVehAsignado, 4) = ""
               mRenglonVehAsignado = 0
            End If
         End If
      End If
   End If

End Sub

Private Function fValidaVehiculo() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String

   mRet = True
      
   If mRenglonVehDispo = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un vehículo de la grilla 'Vehículos Disponibles'"
   End If
      
   If mRet Then
      If mRenglonVehDispo <> 0 And FlexProd.TextMatrix(mRenglonProducto, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un vehículo de la grilla 'Vehículos Disponibles'"
      End If
   End If
      
   If mRet Then
      If Trim(Text1ve(0).Text) = "" Or Trim(Text1ve(1).Text) = "" Then
         mRet = False
         mMensajeError = "Debe completar los kilometrajes de Incio y Fin"
      End If
   End If
      
   If mRet Then
      If Not IsNumeric(Replace(Text1ve(0).Text, ".", ",")) Then
         mRet = False
         mMensajeError = "El 'Km Incial' no es un valor numérico"
      End If
   End If
   
   If mRet Then
      If Not IsNumeric(Replace(Text1ve(1).Text, ".", ",")) Then
         mRet = False
         mMensajeError = "El 'Km Final' no es un valor numérico"
      End If
   End If
   
   If mRet Then
      If CDbl(Replace(Trim(Text1ve(1).Text), ".", ",")) < CDbl(Replace(Trim(Text1ve(0).Text), ".", ",")) Then
         mRet = False
         mMensajeError = "El 'Km Final' no puede ser menor que el 'Km Inicial'"
      End If
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaVehiculo = mRet
End Function

Private Sub Text1ve_KeyPress(Index As Integer, KeyAscii As Integer)
      
   Select Case Index
      Case 0, 1
         If KeyAscii <> 46 Then
            KeyAscii = fNumeroKeyPress(KeyAscii)
         End If
   End Select

End Sub

Private Sub GrabarVehAsignados(ByVal pParte As Double)
   Dim mI As Integer

   If Not mObj.xDeleteVehAsignados(pParte) Then
      MsgBox "ERROR al actualizar Vehiculos Asignados...", vbCritical, sMessage
   End If
   For mI = 1 To FlexVehAsignado.Rows - 1
      If Trim(FlexVehAsignado.TextMatrix(mI, 4)) <> "" Then
            mObj.xInsertVehAsignados pParte, Trim(FlexVehAsignado.TextMatrix(mI, 4)), Trim(FlexVehAsignado.TextMatrix(mI, 2)), Trim(FlexVehAsignado.TextMatrix(mI, 3))
      End If
   Next
   
End Sub

Private Sub Command3ve_Click()
   GrabarVehAsignados MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
End Sub


Private Sub ActualizaGridVeSegunParte(pParte As Double)
   Dim mI As Integer
   Dim mJ As Integer
   
      'Limpio la Grilla VehDispo
      For mI = 1 To FlexVehDisponible.Rows
         
         If FlexVehDisponible.Rows > 2 Then
            FlexVehDisponible.RemoveItem 1
         Else
            FlexVehDisponible.TextMatrix(1, 1) = ""
            FlexVehDisponible.TextMatrix(1, 2) = ""
         End If
      Next
   
      'Limpio la Grilla VehAsignado
      For mI = 1 To FlexVehAsignado.Rows
         
         If FlexVehAsignado.Rows > 2 Then
            FlexVehAsignado.RemoveItem 1
         Else
            FlexVehAsignado.TextMatrix(1, 1) = ""
            FlexVehAsignado.TextMatrix(1, 2) = ""
            FlexVehAsignado.TextMatrix(1, 3) = ""
            FlexVehAsignado.TextMatrix(1, 4) = ""
         End If
      Next
      
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Vehiculos Where Fecha_Baja is null ORDER BY Descripcion;")

      If Not mRec.EOF Then
         mI = 1
         Do While Not mRec.EOF
            mI = mI + 1

            FlexVehDisponible.AddItem ""
            FlexVehDisponible.TextMatrix(mI, 1) = mRec!descripcion
            FlexVehDisponible.TextMatrix(mI, 2) = mRec!Codigo

            mRec.MoveNext
         Loop
         FlexVehDisponible.RemoveItem 1
      End If
      mRec.Close

      Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.Descripcion, P.KmInicio, P.KmFin " & _
                                      " From " & _
                                      " Vehiculos V " & _
                                          " Inner Join" & _
                                      " Partes_Vehiculos P " & _
                                          " ON P.CodVehiculo = V.Codigo " & _
                                       " Where P.Parte = " & pParte & " " & _
                                       " ORDER BY V.Descripcion ;")
   
      mJ = 1
      Do While Not mRec.EOF
   
         mJ = mJ + 1
         FlexVehAsignado.AddItem ""
         FlexVehAsignado.TextMatrix(mJ, 1) = mRec!descripcion
         FlexVehAsignado.TextMatrix(mJ, 2) = mRec!KmInicio
         FlexVehAsignado.TextMatrix(mJ, 3) = mRec!KmFin
         FlexVehAsignado.TextMatrix(mJ, 4) = mRec!Codigo

   
         For mI = 1 To FlexVehDisponible.Rows
            If Trim(FlexVehDisponible.TextMatrix(mI, 2)) = Trim(mRec!Codigo) Then
               
               If FlexVehDisponible.Rows > 2 Then
                  FlexVehDisponible.RemoveItem mI
               Else
                  'blanquear tantas veces como columnas tenga
                  FlexVehDisponible.TextMatrix(1, 1) = ""
                  FlexVehDisponible.TextMatrix(1, 2) = ""
               End If
               
               mI = 99
            End If
         Next
         
         mRec.MoveNext
      Loop
       
       
      If FlexVehAsignado.Rows > 2 Then
         FlexVehAsignado.RemoveItem 1
      Else
         'MsgBox "menorigual2"
      End If
      
      mRec.Close
End Sub









