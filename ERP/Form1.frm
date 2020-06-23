VERSION 5.00
Begin VB.Form ERP1_frm 
   BorderStyle     =   0  'None
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   495
   ClientWidth     =   1845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   1845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   8
      Left            =   80
      TabIndex        =   35
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Usuarios"
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
         Index           =   8
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "ABM Usuario"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   25
         Left            =   120
         TabIndex        =   39
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Permisos a Usuarios"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   38
         Top             =   1995
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   11
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":0000
         Tag             =   "8.ABM Usuarios"
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   10
         Left            =   600
         MouseIcon       =   "Form1.frx":0376
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":04C8
         Tag             =   "8.Permisos"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Cambio de Clave"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   23
         Left            =   90
         TabIndex        =   37
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   9
         Left            =   600
         MouseIcon       =   "Form1.frx":07D2
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0924
         Tag             =   "8.Cambio Clave"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6700
         Index           =   8
         Left            =   0
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   7
      Left            =   80
      TabIndex        =   32
      Top             =   2220
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Varios"
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
         Index           =   7
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Inventario"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   21
         Left            =   465
         TabIndex        =   59
         Top             =   5520
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   12
         Left            =   600
         Picture         =   "Form1.frx":0C2E
         Tag             =   "7. Inventario"
         Top             =   5070
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Mant. Electrico"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   15
         Left            =   280
         TabIndex        =   58
         Top             =   4380
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   47
         Left            =   600
         Picture         =   "Form1.frx":0F38
         Tag             =   "7. Mant. Elect."
         Top             =   3840
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Mant. Edilicio"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   29
         Left            =   375
         TabIndex        =   56
         Top             =   3240
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   20
         Left            =   600
         Picture         =   "Form1.frx":1418
         Tag             =   "7. Mant. Ed."
         Top             =   2610
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Ctrl. Obras"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   27
         Left            =   465
         TabIndex        =   54
         Top             =   2100
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   42
         Left            =   600
         MouseIcon       =   "Form1.frx":205A
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":21AC
         Tag             =   "7. Ctrl Proveed"
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Sist. Legales"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   18
         Left            =   360
         TabIndex        =   52
         Top             =   1050
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   40
         Left            =   600
         MouseIcon       =   "Form1.frx":583D
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":598F
         Tag             =   "7. Legales"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6400
         Index           =   7
         Left            =   0
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   6
      Left            =   80
      TabIndex        =   27
      Top             =   1920
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "R.R.H.H."
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
         Index           =   6
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   28
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Sistema de Sanciones"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   19
         Left            =   720
         TabIndex        =   53
         Top             =   4260
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   41
         Left            =   120
         MouseIcon       =   "Form1.frx":8E54
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":915E
         Stretch         =   -1  'True
         Tag             =   "6.Sist. Sanciones"
         Top             =   4140
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Administrador de Ausencias"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   38
         Left            =   660
         TabIndex        =   48
         Top             =   3480
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   38
         Left            =   120
         MouseIcon       =   "Form1.frx":9481
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":978B
         Stretch         =   -1  'True
         Tag             =   "6.Adm.Ausencias"
         Top             =   3420
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Sistema de Formularios"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   32
         Left            =   720
         TabIndex        =   47
         Top             =   2700
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   32
         Left            =   120
         MouseIcon       =   "Form1.frx":AA8F
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":AD99
         Stretch         =   -1  'True
         Tag             =   "6.Sist.Formularios"
         Top             =   2640
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Gestión CV"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   20
         Left            =   720
         TabIndex        =   45
         Top             =   1980
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   28
         Left            =   120
         MouseIcon       =   "Form1.frx":B3D3
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":B6DD
         Stretch         =   -1  'True
         Tag             =   "6.Gestion CV"
         Top             =   1860
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Archivos txt"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   11
         Left            =   720
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   15
         Left            =   120
         MouseIcon       =   "Form1.frx":CB6A
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":CE74
         Tag             =   "6.Gestion Ausencias"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000C&
         Caption         =   "Gestión de Cursos"
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   10
         Left            =   720
         TabIndex        =   29
         Top             =   540
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   14
         Left            =   180
         MouseIcon       =   "Form1.frx":D17E
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":D2D0
         Tag             =   "6.Cursos"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   6
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   5
      Left            =   80
      TabIndex        =   22
      Top             =   1620
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Sistemas"
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
         Index           =   5
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   23
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Mantenimiento"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   14
         Left            =   315
         TabIndex        =   50
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   8
         Left            =   600
         MouseIcon       =   "Form1.frx":D5DA
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":F64C
         Tag             =   "5.Mantenimiento"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Telecarga Peas"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   44
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   25
         Left            =   600
         MouseIcon       =   "Form1.frx":FC80
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":FDD2
         Tag             =   "5.Telecarga"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   5
         Left            =   0
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   4
      Left            =   80
      TabIndex        =   18
      Top             =   1320
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Ventas"
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
         Index           =   4
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   19
         Top             =   0
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   45
         Left            =   570
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":10445
         Tag             =   "4.Exc. Peaje"
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Exenc. Peaje"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   30
         Left            =   90
         TabIndex        =   57
         Top             =   3075
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Seg. Empresas"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   1995
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   22
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":1398A
         Tag             =   "4.Vacio"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   7
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":13FF1
         Tag             =   "4.Pasadas"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Pasadas"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   20
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   4
         Left            =   0
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   3
      Left            =   80
      TabIndex        =   14
      Top             =   1020
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Validación"
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
         Index           =   3
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   705
         Index           =   31
         Left            =   550
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":142FB
         Tag             =   "31.Compensaciones"
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Compensaciones"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   26
         Left            =   240
         TabIndex        =   46
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   650
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":15DAD
         Tag             =   "3.Valid"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Valid"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   16
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   3
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   2
      Left            =   80
      TabIndex        =   9
      Top             =   720
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Supervisión"
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
         Index           =   2
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   10
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "POLAD"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   17
         Left            =   80
         TabIndex        =   51
         Top             =   5010
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   615
         Index           =   19
         Left            =   490
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":16347
         Tag             =   "2. POLAD"
         Top             =   4320
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Arqueo Supervisores"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   8
         Left            =   80
         TabIndex        =   49
         Top             =   4005
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":17BE1
         Tag             =   "2.Arqueo Sup"
         Top             =   3360
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Consulta de TAGs"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   17
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":17EEB
         Tag             =   "2.ConsulTAG"
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":181F5
         Tag             =   "20. Aux. Vía"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Gestión Sup."
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   12
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":184FF
         Tag             =   "2.Violaciones"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Violaciones"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1995
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   1
      Left            =   80
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   420
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Seguridad Vial"
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
         Index           =   1
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Mensajes en Vias"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   28
         Left            =   120
         TabIndex        =   55
         Top             =   4200
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   555
         Index           =   43
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":18809
         Tag             =   "43.TeleCargas"
         Top             =   3600
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Balanza Móvil"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   43
         Top             =   3120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Reg. Accidentes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   18
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":19753
         Tag             =   "1.Registro Accidentes"
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   600
         MouseIcon       =   "Form1.frx":19A5D
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":19BAF
         Tag             =   "1.RegNov"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Reg. Novedades"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   7
         Tag             =   "Registro Novedades"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":19EB9
         Tag             =   "1.Registro Accidentes"
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Reg. Accidentes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   1
         Left            =   0
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   300
      Index           =   0
      Left            =   80
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton Command1 
         Caption         =   "Operaciones"
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
         Left            =   0
         MousePointer    =   12  'No Drop
         TabIndex        =   1
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "Cont. Peek"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   13
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":1A2FB
         Tag             =   "0.Peek"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "ConsPea"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   1050
         Width           =   1515
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   600
         MousePointer    =   12  'No Drop
         Picture         =   "Form1.frx":1A605
         Tag             =   "0.ConsPea"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         Height          =   6100
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "ERP1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mCtaFor As Integer

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mSuma As Integer
Dim mFlag As Boolean
mFlag = (Frame1(Index).Height = 300)
If Command1(Index).MousePointer = 0 And Frame1(Index).Height <> 6700 Then 'Con esto logramos que se vean los programas del menu
   mSuma = 120
   For mCtaFor = 0 To Frame1.UBound
      Frame1(mCtaFor).Height = 300
      Frame1(mCtaFor).Top = mSuma
      mSuma = mSuma + 300
   Next
   If mFlag Then
      Frame1(Index).Height = 5700
      mCtaFor = Index + 1
      For mCtaFor = mCtaFor To Frame1.UBound
         Frame1(mCtaFor).Top = Frame1(mCtaFor).Top + 5400
      Next
   End If
End If
End Sub

Private Sub Command1_KeyPress(Index As Integer, KeyAscii As Integer)
If Frame1(Index).Height > 300 Then
   Select Case Index
      Case 0   'OPERACIONES
         Select Case KeyAscii
            Case 49
               Image1_Click 0
            Case 50
               Image1_Click 13
         End Select
      Case 1   'SEGURIDAD VIAL
         Select Case KeyAscii
            Case 49
               Image1_Click 1   'REG. NOVEDADES
            Case 50
               Image1_Click 2   'Registro de Accidentes
            Case 51
               Image1_Click 18  'Sistema de Balanzas
            Case 52
               Image1_Click 43  'Sistema de Tele Cargas de Display de Vía
         End Select
      Case 2   'SUPERVISION
         Select Case KeyAscii
            Case 49
               Image1_Click 3
            Case 50
               Image1_Click 5
            Case 51
               Image1_Click 17
            Case 52
               Image1_Click 4
            Case 53
               Image1_Click 19
         End Select
      Case 3   'VALIDACIONES
         Select Case KeyAscii
            Case 49
               Image1_Click 6
            Case 50
               Image1_Click 31
         End Select
      Case 4   'VENTAS
         Select Case KeyAscii
            Case 49
               Image1_Click 7
            Case 50
               Image1_Click 22
            Case 51
               Image1_Click 45
         End Select
      Case 5   'SISTEMAS
         Select Case KeyAscii
            Case 49
               Image1_Click 25 'TELECARGAS PEAS
            Case 50
               Image1_Click 8 'MANTENIMIENTO DE SISTEMAS
         End Select
      Case 6   'RRHH
         Select Case KeyAscii
            Case 49
               Image1_Click 14 'GESTION DE CURSOS
            Case 50
               Image1_Click 15 'GESTION DE AUSENCIAS
            Case 51
               Image1_Click 28 'GESTION DE CVs
            Case 52
               Image1_Click 32 'FORMULARIOS
            Case 53
               Image1_Click 38 'ADM AUSENCIAS
            Case 54
               Image1_Click 41 'SANCIONES
         End Select
      Case 7   'VARIOS
         Select Case KeyAscii
'            Case 49
'               Image1_Click 12 'PROVEEDORES
            Case 49
               Image1_Click 40 'LEGALES
            Case 50
               Image1_Click 42 'CTRL. OBRAS
            Case 51
               Image1_Click 20 'MANT. EDILICIO
            Case 52
               Image1_Click 47 'MANT. ELECTRICO
            Case 53
               Image1_Click 12 'INVENTARIO
         End Select
      Case 8   'USUARIOS
         Select Case KeyAscii
            Case 49
               Image1_Click 9 'CAMBIO DE CLAVE
            Case 50
               Image1_Click 10 'GESTION DE PERMISOS
            Case 51
               Image1_Click 11 'ABM de USUARIOS
         End Select
   End Select
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Dim mObj As New clLogUser
If MDI.ERP_VenWind(Index).Visible = False Then
   If MDI.mMenuActivo <> Index And Image1(Index).MousePointer = 99 Then
      ERP1_frm.Visible = False
      ShowMenu Index, True, False
      mObj.sActivarMenu Trim(Right(MDI.mUser, 15)), Index
      
   Else
      MsgBox "Acceso Denegado!", vbCritical, sMessage
   End If
Else
   MDI.ERP_VenWind_Click (Index)
End If
Set mObj = Nothing
End Sub
