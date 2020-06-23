VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MantElect08old 
   Caption         =   "Cierre de Orden de Trabajo"
   ClientHeight    =   9405
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   11160
      TabIndex        =   33
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   30
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9450
      TabIndex        =   27
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar O.T."
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   26
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(4)"
      Height          =   15
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1125
      Width           =   16920
      Begin VB.Frame Frame4 
         Caption         =   "Egresos"
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
         Height          =   2475
         Left            =   0
         TabIndex        =   59
         Top             =   4740
         Width           =   16900
         Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
            Height          =   2050
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   3625
            _Version        =   327680
            Cols            =   8
         End
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   1320
         TabIndex        =   57
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   0
         Left            =   2880
         Picture         =   "MantElect08.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   1
         Left            =   3375
         Picture         =   "MantElect08.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   4320
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Caption         =   "Selección del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2895
         Left            =   0
         TabIndex        =   50
         Top             =   1320
         Width           =   16900
         Begin VB.TextBox Text9 
            Height          =   315
            Left            =   2640
            TabIndex        =   53
            Top             =   380
            Width           =   10455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Buscar"
            Height          =   315
            Left            =   13320
            TabIndex        =   52
            Top             =   380
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid FlexProd 
            Height          =   1935
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   3413
            _Version        =   327680
            Cols            =   8
         End
         Begin VB.Label Label10 
            Caption         =   "Contiene texto:"
            Height          =   375
            Left            =   1080
            TabIndex        =   54
            Top             =   440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Información del Consumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   0
         TabIndex        =   39
         Top             =   120
         Width           =   16900
         Begin VB.TextBox Text8 
            Height          =   315
            Left            =   12520
            MaxLength       =   9
            TabIndex        =   49
            Top             =   510
            Width           =   1215
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            ItemData        =   "MantElect08.frx":0614
            Left            =   2240
            List            =   "MantElect08.frx":0616
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   510
            Width           =   3235
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Vale a cargo/recambio"
            Height          =   375
            Left            =   14300
            TabIndex        =   43
            Top             =   280
            Width           =   2555
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Vale de retiro múltiple"
            Height          =   375
            Left            =   14300
            TabIndex        =   42
            Top             =   640
            Width           =   2555
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   280
            Width           =   3235
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   740
            Width           =   3235
         End
         Begin VB.Label Label9 
            Caption         =   "Retirar de Bodega:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   575
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Retirado por:"
            Height          =   255
            Left            =   6000
            TabIndex        =   47
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Vale número:"
            Height          =   255
            Left            =   11080
            TabIndex        =   46
            Top             =   575
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Autorizado por:"
            Height          =   255
            Left            =   6000
            TabIndex        =   45
            Top             =   740
            Width           =   1455
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Cantidad:"
         Height          =   360
         Left            =   120
         TabIndex        =   58
         Top             =   4380
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7650
      Index           =   3
      Left            =   15
      TabIndex        =   4
      Top             =   1125
      Width           =   16920
      Begin VB.Frame Frame10 
         Caption         =   "Detalle de consumo"
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
         Height          =   7440
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   16440
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   12000
            TabIndex        =   34
            Top             =   6834
            Width           =   1575
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   22
            Top             =   540
            Width           =   3735
         End
         Begin MSFlexGridLib.MSFlexGrid FlexProduct 
            Height          =   5640
            Left            =   240
            TabIndex        =   21
            Top             =   1200
            Width           =   15975
            _ExtentX        =   28178
            _ExtentY        =   9948
            _Version        =   327680
            Cols            =   9
         End
         Begin VB.Label Label2 
            Caption         =   "Retirar de:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7300
      Index           =   2
      Left            =   15
      TabIndex        =   3
      Top             =   1125
      Width           =   16920
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   1
         Left            =   8400
         Picture         =   "MantElect08.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3840
         Width           =   495
      End
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   0
         Left            =   7440
         Picture         =   "MantElect08.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3840
         Width           =   495
      End
      Begin VB.Frame Frame7 
         Caption         =   "Rubros/Subrubros asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2820
         Left            =   240
         TabIndex        =   12
         Top             =   4320
         Width           =   16455
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubrosAsign 
            Height          =   2250
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   15975
            _ExtentX        =   28178
            _ExtentY        =   3969
            _Version        =   327680
            Cols            =   5
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Rubros/Subrubros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3615
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   16455
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   360
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubros 
            Height          =   2655
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   15975
            _ExtentX        =   28178
            _ExtentY        =   4683
            _Version        =   327680
            Cols            =   5
         End
         Begin VB.Label Label1 
            Caption         =   "Rubro:"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   420
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   7300
      Index           =   1
      Left            =   15
      TabIndex        =   2
      Top             =   1125
      Width           =   16920
      Begin VB.Frame Frame9 
         Caption         =   "Vehículos Especiales asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3015
         Left            =   9120
         TabIndex        =   18
         Top             =   240
         Width           =   7335
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   5160
            TabIndex        =   38
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   3120
            TabIndex        =   37
            Top             =   2640
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid FlexVehEspAsign 
            Height          =   2205
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3889
            _Version        =   327680
            Cols            =   5
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Mano de Obra asignada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3210
         Left            =   4440
         TabIndex        =   8
         Top             =   3720
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexMoAsig 
            Height          =   2535
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   4471
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Vehículos asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3015
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   7335
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   5160
            TabIndex        =   36
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   2880
            TabIndex        =   35
            Top             =   2640
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid FlexVehAsign 
            Height          =   2205
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3889
            _Version        =   327680
            Cols            =   6
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7300
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   1125
      Width           =   16920
      Begin VB.Frame Frame12 
         Caption         =   "Partes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   6975
         Left            =   360
         TabIndex        =   24
         Top             =   120
         Width           =   16275
         Begin MSFlexGridLib.MSFlexGrid FlexPartes 
            Height          =   6135
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   10821
            _Version        =   327680
            Cols            =   7
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   720
      Width           =   16920
      _ExtentX        =   29845
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asignación de partes"
            Object.Tag             =   "SupervisorMantElectrico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asignación  Mano de Obra / Vehículos "
            Object.Tag             =   "SupervisorMantElectrico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Fallas"
            Object.Tag             =   "SupervisorMantElectrico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Asignacion Materiales"
            Object.Tag             =   "SupervisorMantElectrico"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   " Fecha Fin:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9840
      TabIndex        =   32
      Top             =   165
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   " Fecha Inicio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5280
      TabIndex        =   31
      Top             =   165
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "O.T.  -  Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   165
      Width           =   1575
   End
End
Attribute VB_Name = "MantElect08old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset
Dim mRenglonPartes As Integer
Dim mRenglonVehAsign As Integer
Dim mRenglonVehEspAsign As Integer
Dim mRenglonMoAsign As Integer
Dim mRenglonSubRubroDispo As Integer
Dim mRenglonSubRubroAsign As Integer
Dim mRenglonProdDispo As Integer

Dim mRenglonProducto As Integer
Dim mRenglonEgreso As Integer
               

Dim XLS As EXCEL.Application

Dim filaAnt As Integer
Dim columnAnt As Integer
Dim filaAntVehAsign As Integer
Dim columnAntVehAsign As Integer
Dim filaAntVehAsignKmFinal As Integer
Dim columnAntVehAsignKmFinal As Integer

Dim filaAntVehEspAsignKmInicio As Integer
Dim columnAntVehEspAsignKmInicio As Integer
Dim filaAntVehEspAsignKmFinal As Integer
Dim columnAntVehEspAsignKmFinal As Integer


'TODO: Ver si es necesario utilizar las siguientes variables:
Dim mCodParte As Integer
Dim mCodMO As String
Dim mCodSubrubro As String
Dim mCodVeh As String
Dim mCodVehEsp As String
Dim mCodProducto As String

Dim cboListIndex As Integer

Dim mEsOTcerrada As Boolean



Private Sub Combo1_Click()
 Dim mi As Integer
 Dim sListaSubrubrosSeleccionados As String
   
   'Elimino los registros  de la grilla superior
   For mi = FlexSubRubros.Rows To 3 Step -1
      FlexSubRubros.RemoveItem mi
   Next
   
   If FlexSubRubrosAsign.Rows > 2 Then
      For mi = 2 To FlexSubRubrosAsign.Rows - 1
         sListaSubrubrosSeleccionados = sListaSubrubrosSeleccionados & "'" & FlexSubRubrosAsign.TextMatrix(mi, 4) & "',"
      Next
      sListaSubrubrosSeleccionados = Left(sListaSubrubrosSeleccionados, Len(sListaSubrubrosSeleccionados) - 1)
   End If
    
   mRenglonSubRubroDispo = 0
   
   If FlexSubRubrosAsign.Rows > 2 Then
      
         Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
         "  From " & _
         " Rubros R " & _
         "  Inner Join " & _
         " SubRubros S ON S.CodRubro = R.Codigo " & _
         " WHERE S.Codigo NOT IN (" & sListaSubrubrosSeleccionados & ")" & _
         " AND R.Codigo ='" & Right(Combo1.Text, 8) & "' ORDER BY RubroDesc, SubRubroDesc ;")
   Else
      Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
         "  From " & _
         " Rubros R " & _
         "  Inner Join " & _
         " SubRubros S ON S.CodRubro = R.Codigo" & _
         " WHERE R.Codigo ='" & Right(Combo1.Text, 8) & "' ORDER BY RubroDesc, SubRubroDesc;")
      
   End If
                                
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
      
         With FlexSubRubros
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!RubroDesc
            .TextMatrix(mi, 2) = mRec!SubRubroDesc
            .TextMatrix(mi, 3) = mRec!CodRubro
            .TextMatrix(mi, 4) = mRec!CodSubrubro
         End With
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub Combo3_Click()
   Dim mi As Integer
   Dim mIdOT As Integer
   Dim mCodUbicacion As String


'MsgBox "Form ancho: " & Me.Width
'MsgBox "Form mantelec08 ancho: " & MantElect08.Width
'MsgBox "Frame1 ancho: " & Frame1(0).Width
'MsgBox "FrameInterior ancho: " & Frame12.Width
'MsgBox "Flex ancho: " & FlexPartes.Width
'MsgBox "------ altos -------: " & FlexPartes.Width
'MsgBox "Form alto: " & Me.Height
'MsgBox "Form mantelec08 alto: " & MantElect08.Height
   mIdOT = Left(Combo3.Text, 10)

   '---FECHA INICIO Y FECHA FIN OT-------------------------------------------------------------------------------------
   
   'Habilito Textboxs Fecha (Inicio y Fin)
   Text3(0).Enabled = True
   Text3(1).Enabled = True
   Text3(0).Text = ""
   Text3(1).Text = ""
   
   
   Set mRec = mObj.oEjecutarSelect("SELECT IdOT, FechaInicio, FechaFin FROM OT_H WHERE IdOT = " & mIdOT & " and FechaFin <> '0000-00-00 00:00:00'; ")
   
   If Not mRec.EOF Then
      Text3(0).Text = mRec!FechaInicio
      Text3(1).Text = mRec!FechaFin
      InhabilitarControlesOTCerrada
      
   Else
      mEsOTcerrada = False
      'Habilito Textboxs Fecha (Inicio y Fin)
      Text3(0).Enabled = True
      Text3(1).Enabled = True
      Command2(0).Enabled = True
      'Habilito Botones 'Subrurbro'
      CommandSubRubro(0).Enabled = True
      CommandSubRubro(1).Enabled = True
   End If
   mRec.Close
   
   '--PARTES-------------------------------------------------------------------------------------------------------------------
   'Elimino los registros (de la consulta anterior) de la grilla superior
   
   mRenglonPartes = 0
   
   For mi = FlexPartes.Rows To 3 Step -1
      FlexPartes.RemoveItem mi
   Next

   'TODO: VER SI EN ESTA QUERY NO TENGO QUE CONTEMPLAR DIFERENCIAR OT de aire y d electricos puro
   Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire " & _
                                       "FROM " & _
                                          "OT_Partes OH " & _
                                       "Inner Join " & _
                                          "Registros R ON OH.Parte = R.Parte " & _
                                        "where OH.IdOT = " & mIdOT & " " & _
                                        "AND Cancelado = 0; ")
                                        '"AND Finalizado = 'NO'; ")
                                        
                                        
                                        

   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
      
         With FlexPartes
            .AddItem ""
            .TextMatrix(mi, 0) = "X"
            .TextMatrix(mi, 1) = mRec!Parte
            .TextMatrix(mi, 2) = NVL(mRec!FechaSolic, "")
            .TextMatrix(mi, 3) = NVL(mRec!CodEdificio, "")
            .TextMatrix(mi, 4) = NVL(mRec!descripcion, "")
            .TextMatrix(mi, 5) = NVL(mRec!Prioridad, "")
            
            .TextMatrix(mi, 6) = IIf(mRec!SectorAire = 1, "Si", "No")
            
         End With
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
'----------------------------------------------------------------------------------------------------------------------------------




'--VEHICULOS------------------------------------------------------------------------------------------------------------------------
   
   mRenglonVehAsign = 0
   Text4.Text = ""
   Text4.Visible = False
   Text5.Text = ""
   Text5.Visible = False
   FlexVehAsign.ScrollBars = flexScrollBarVertical
   
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexVehAsign.Rows To 3 Step -1
      FlexVehAsign.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion, OV.KmInicial, OV.KmFinal " & _
                                   " From " & _
                                   "   Vehiculos V " & _
                                   " Inner Join " & _
                                   "   OT_Vehiculos OV ON V.Codigo = OV.CodVehiculo " & _
                                   " Left Join " & _
                                   "   Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where OV.IdOT = " & mIdOT & _
                                   " AND CodUbicacion Is Null; ")
   
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         With FlexVehAsign
            .AddItem ""
            .TextMatrix(mi, 0) = "X"
            .TextMatrix(mi, 2) = mRec!descripcion
            .TextMatrix(mi, 3) = mRec!KmInicial
            .TextMatrix(mi, 4) = mRec!KmFinal
            .TextMatrix(mi, 5) = NVL(mRec!Codigo, "")
         End With

         mRec.MoveNext
      Loop
   End If
   mRec.Close
'--------------------------------------------------------------------------------------------------------------------------------------

'--VEHICULOS ESPECIAL (tambien completo grilla materiales)------------------------------------------------------------------------------------------------------------------
   
   mRenglonVehEspAsign = 0
   Text6.Text = ""
   Text6.Visible = False
   Text7.Text = ""
   Text7.Visible = False
   FlexVehEspAsign.ScrollBars = flexScrollBarVertical
      
      
   Text1.Text = ""
      
   'Elimino los registros de la grilla
   For mi = FlexVehEspAsign.Rows To 3 Step -1
      FlexVehEspAsign.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion, V.CodUbicacion, OV.KmInicial, OV.KmFinal  " & _
                                   " From " & _
                                   "   Vehiculos V " & _
                                   " Inner Join " & _
                                   "   OT_Vehiculos OV ON V.Codigo = OV.CodVehiculo " & _
                                   " Left Join " & _
                                   "   Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where OV.IdOT = " & mIdOT & _
                                   " AND CodUbicacion Is NOT Null; ")
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         With FlexVehEspAsign
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!descripcion
            .TextMatrix(mi, 2) = mRec!KmInicial
            .TextMatrix(mi, 3) = mRec!KmFinal
            .TextMatrix(mi, 4) = NVL(mRec!Codigo, "")
         End With
         
         mCodUbicacion = NVL(mRec!CodUbicacion, "")
         
         Text1.Enabled = False
         Text1 = mRec!descripcion & Space(100) & mRec!Codigo

         mRec.MoveNext
      Loop
   End If
   mRec.Close
   
   mRenglonProdDispo = 0
   Text2.Text = ""
   Text2.Visible = False
   FlexProduct.ScrollBars = flexScrollBarVertical
   
   
   
   
'''  'Elimino los registros  de la grilla
'''  For mi = FlexProduct.Rows To 3 Step -1
'''      FlexProduct.RemoveItem mi
'''   Next
'''
'''
''''  Set mRec = mObjInven.oEjecutarSelect(" SELECT IdMov, Fecha, CodTipoMovimiento, CodProducto, P.Descripcion AS Producto, CodUbicacion, " & _
''''   " U.Descripcion AS Ubicacion, 0 AS Cantidad , 0 AS CantidadBD, Stock, Med.Descripcion AS UnidadMedida, P.CodigoSap, CodUsuario,   Observaciones " & _
''''   " From " & _
''''   "  Movimientos2 M " & _
''''   " Inner Join " & _
''''   "  Producto P ON M.CodProducto = P.Codigo " & _
''''   " Inner Join " & _
''''   "  Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
''''   " Inner Join " & _
''''   "  UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
''''   " WHERE Fecha = (SELECT MAX(Fecha) " & _
''''   "                From Movimientos2 " & _
''''   "               WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
''''   " AND U.Codigo ='" & mCodUbicacion & "' " & _
''''   " ORDER BY  P.Descripcion; ")
'''
'''   Set mRec = mObjInven.oEjecutarSelect(" " & _
'''   " SELECT " & _
'''   "   P.CodigoSap, AUX.CodProducto, AUX.CodUbicacion, P.Descripcion AS Producto, " & _
'''   "   Med.Descripcion AS UnidadMedida,  SUM(AUX.Cantidad) AS Cantidad, " & _
'''   "   SUM(AUX.CantidadBD) AS CantidadBD, SUM(AUX.Stock) AS Stock " & _
'''   " FROM  " & _
'''   " (   SELECT CodProducto, CodUbicacion, 0 AS Cantidad, 0 AS CantidadBD, Stock " & _
'''   "     FROM  " & _
'''   "     Inventario.Movimientos2 M " & _
'''   "     WHERE M.Fecha = (  SELECT MAX(Fecha) From Inventario.Movimientos2 " & _
'''   "                        WHERE CodProducto = M.CodProducto " & _
'''   "                        AND CodUbicacion = M.CodUbicacion) " & _
'''   "     AND M.CodUbicacion ='" & mCodUbicacion & "' " & _
'''   "     UNION ALL " & _
'''   "     SELECT  MV.CodProducto, MV.CodUbicacion,   MV.Cantidad, MV.Cantidad AS CantidadBD, 0 AS Stock " & _
'''   "     FROM " & _
'''   "     MantElect.OT_Materiales OM " & _
'''   "     INNER JOIN Inventario.Movimientos2 MV ON OM.IdMov = MV.IdMov " & _
'''   "     WHERE idOT = " & mIdOT & " AND esMovimientoFinal = 1 " & _
'''   " ) AS AUX " & _
'''   " INNER JOIN Inventario.Producto P ON AUX.CodProducto = P.Codigo " & _
'''   " INNER JOIN Inventario.Ubicaciones U ON  AUX.CodUbicacion = U.Codigo " & _
'''   " INNER JOIN Inventario.UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
'''   " GROUP BY P.CodigoSap, AUX.CodProducto, AUX.CodUbicacion, P.Descripcion, Med.Descripcion " & _
'''   " ORDER BY P.Descripcion; ")
'''
'''
'''
'''   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
'''   If Not mRec.EOF Then
'''      mi = 1
'''      Do While Not mRec.EOF
'''         mi = mi + 1
'''
'''         FlexProduct.AddItem ""
'''         FlexProduct.TextMatrix(mi, 1) = mRec!CodigoSap
'''         FlexProduct.TextMatrix(mi, 2) = mRec!Producto
'''         FlexProduct.TextMatrix(mi, 3) = mRec!Cantidad
'''         FlexProduct.TextMatrix(mi, 4) = mRec!Stock
'''         FlexProduct.TextMatrix(mi, 5) = mRec!UnidadMedida
'''         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
'''         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
'''         FlexProduct.TextMatrix(mi, 8) = mRec!CantidadBD
'''
'''         mRec.MoveNext
'''      Loop
'''   End If
'''   mRec.Close

   llenoGrillaConsumo mIdOT, mCodUbicacion

   
   
   
'--------------------------------------------------------------------------------------------------------------------------------------


'--TECNICOS----------------------------------------------------------------------------------------------------------------------------
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexMoAsig.Rows To 3 Step -1
      FlexMoAsig.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect(" SELECT Codigo,Descripcion " & _
                                   " From " & _
                                   "   MO_Tecnicos M " & _
                                   " Inner Join " & _
                                   "  OT_MO_Tecnicos OM ON OM.CodMo_Tecnico = M.Codigo " & _
                                   " WHERE OM.IdOT = " & mIdOT & ";")
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         With FlexMoAsig
            .AddItem ""
            .TextMatrix(mi, 0) = "X"
            .TextMatrix(mi, 1) = mRec!descripcion
            .TextMatrix(mi, 2) = NVL(mRec!Codigo, "")
         End With

         mRec.MoveNext
      Loop
   End If
   mRec.Close


'--SUBRUBROS--------------------------------------------------------------------------------------------------------------------------
   Combo1.Clear
   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Rubros Where FechaBaja IS NULL;")
   
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!descripcion & Space(50) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close

   'Elimino los registros de la grilla
   For mi = FlexSubRubros.Rows To 3 Step -1
      FlexSubRubros.RemoveItem mi
   Next

   'Elimino los registros de la grilla
   For mi = FlexSubRubrosAsign.Rows To 3 Step -1
      FlexSubRubrosAsign.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect(" SELECT R.Codigo As CodRubro, R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc FROM " & _
                                   " OT_Subrubros OS " & _
                                   "   Inner Join " & _
                                   " SubRubros S ON OS.CodSubrubro = S.Codigo " & _
                                   "   Inner Join " & _
                                   " Rubros R ON S.CodRubro = R.Codigo " & _
                                   " WHERE OS.IdOT = " & mIdOT & " ORDER BY RubroDesc, SubRubroDesc;")
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         With FlexSubRubrosAsign
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!RubroDesc
            .TextMatrix(mi, 2) = mRec!SubRubroDesc
            .TextMatrix(mi, 3) = mRec!CodRubro
            .TextMatrix(mi, 4) = mRec!CodSubrubro
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub



Private Sub llenoGrillaConsumo(pIdOT As Integer, pCodUbicacion As String)
 Dim mi As Integer
 
 'Elimino los registros  de la grilla
  For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   
   Set mRec = mObj.getConsumoMatXidOTyUbicacion(pIdOT, pCodUbicacion)

   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 2) = mRec!Producto
         FlexProduct.TextMatrix(mi, 3) = mRec!Cantidad
         FlexProduct.TextMatrix(mi, 4) = mRec!Stock
         FlexProduct.TextMatrix(mi, 5) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
         FlexProduct.TextMatrix(mi, 8) = mRec!CantidadBD
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub






Private Sub Command1_Click()
   Dim mi As Integer
   Dim mj As Integer

   mRenglonProducto = 0

   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProd.Rows To 3 Step -1
      FlexProd.RemoveItem mi
   Next

   Set mRec = mObjInven.getStockXBodegaConFiltroProducto(Left(Combo5.Text, 4), Text9.Text)

   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1

         FlexProd.AddItem ""
         FlexProd.TextMatrix(mi, 1) = mRec!Producto
         FlexProd.TextMatrix(mi, 2) = mRec!Ubicacion
         FlexProd.TextMatrix(mi, 3) = mRec!Stock
         FlexProd.TextMatrix(mi, 4) = mRec!UnidadMedida
         FlexProd.TextMatrix(mi, 5) = mRec!CodigoSap
         FlexProd.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProd.TextMatrix(mi, 7) = mRec!CodUbicacion

         mRec.MoveNext
      Loop
   End If
   mRec.Close

   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" descontando el consumo de la grilla inferior
   For mi = 2 To FlexProd.Rows - 1
      For mj = 2 To FlexEgreso.Rows - 1
         If FlexProd.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mj, 6) And FlexProd.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mj, 7) Then
            FlexProd.TextMatrix(mi, 3) = CDbl(Replace(Trim(FlexProd.TextMatrix(mi, 3)), ".", ",")) - CDbl(Replace(Trim(FlexEgreso.TextMatrix(mj, 3)), ".", ","))
            mj = 999
         End If
      Next
   Next
End Sub
Private Sub Command2_Click(Index As Integer)
   If Index = 0 Then
      Dim IdOT As Integer
      'Dim OTgenerada As Integer
      Dim vPartes() As Double
      
      Dim vVeh_Codigo() As String
      Dim vVeh_KmIni() As Integer
      Dim vVeh_KmFin() As Integer
      
      Dim vVehEsp_Codigo() As String
      Dim vVehEsp_KmIni() As Integer
      Dim vVehEsp_KmFin() As Integer
      
      Dim vMO_OT() As String
      Dim vSubrub_OT() As String
      
      Dim vMat_CodProd() As String
      Dim vMat_CodUbic() As String
      Dim vMat_Cantidad() As Double
      Dim vMat_CantidadBD() As Double
      
      Dim mi As Integer
      Dim fecIni As Date
      Dim FecFin As Date
      
      If fValidaOT Then
     
         preparaArrayPartes vPartes()
         preparaArrayVehiculos vVeh_Codigo(), vVeh_KmIni(), vVeh_KmFin()
         preparaArrayVehiculosEsp vVehEsp_Codigo(), vVehEsp_KmIni(), vVehEsp_KmFin()
         preparaArrayMO_Tecnicos vMO_OT()
         preparaArraySubrubros vSubrub_OT()
         preparaArrayMateriales vMat_CodProd(), vMat_CodUbic(), vMat_Cantidad(), vMat_CantidadBD()
         
         IdOT = CInt(Left(Combo3.Text, 10))
         fecIni = CDate(Text3(0).Text)
         FecFin = CDate(Text3(1).Text)
         
         
         mObj.xCerrarOT IdOT, fecIni, FecFin, Trim(Right(MDI.mUser, 15)), vPartes(), _
                        vVeh_Codigo(), vVeh_KmIni(), vVeh_KmFin(), vVehEsp_Codigo(), _
                        vVehEsp_KmIni(), vVehEsp_KmFin(), vMO_OT(), vSubrub_OT(), _
                        vMat_CodProd(), vMat_CodUbic(), vMat_Cantidad(), vMat_CantidadBD()

         
         InhabilitarControlesOTCerrada

         MsgBox "Se ha cerrado correctamente la O.T.: " & IdOT
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub preparaArrayPartes(ByRef pvPartes_OT() As Double)
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartes.Rows - 2
   If cantPartes > 0 Then
      ReDim pvPartes_OT(0 To cantPartes - 1) As Double
         
      For mj = 2 To FlexPartes.Rows - 1
         pvPartes_OT(mj - 2) = FlexPartes.TextMatrix(mj, 1)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      'Igualmente no deberia pasar debido a que no antes de confirmar valido si hay algun parte elegido.
      ReDim pvPartes_OT(0)
      pvPartes_OT(0) = 0
   End If
End Sub

Private Sub preparaArrayVehiculos(ByRef pvVeh_Codigo() As String, ByRef pvVeh_KmIni() As Integer, ByRef pvVeh_KmFin() As Integer)
   Dim mj As Integer
   Dim cantVehiculos As Integer

   cantVehiculos = FlexVehAsign.Rows - 2
   If cantVehiculos > 0 Then
      ReDim pvVeh_Codigo(0 To cantVehiculos - 1) As String
      ReDim pvVeh_KmIni(0 To cantVehiculos - 1) As Integer
      ReDim pvVeh_KmFin(0 To cantVehiculos - 1) As Integer
         
      For mj = 2 To FlexVehAsign.Rows - 1
         pvVeh_Codigo(mj - 2) = FlexVehAsign.TextMatrix(mj, 5)
         pvVeh_KmIni(mj - 2) = FlexVehAsign.TextMatrix(mj, 3)
         pvVeh_KmFin(mj - 2) = FlexVehAsign.TextMatrix(mj, 4)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      'Podria ocurrir que no tenga vehiculo asociado a la ot
      ReDim pvVeh_Codigo(0)
      pvVeh_Codigo(0) = "00"
   End If
End Sub

Private Sub preparaArrayVehiculosEsp(ByRef pvVehEsp_Codigo() As String, ByRef pvVehEsp_KmIni() As Integer, ByRef pvVehEsp_KmFin() As Integer)
   Dim mj As Integer
   Dim cantVehiculos As Integer

   cantVehiculos = FlexVehEspAsign.Rows - 2
   If cantVehiculos > 0 Then
      ReDim pvVehEsp_Codigo(0 To cantVehiculos - 1) As String
      ReDim pvVehEsp_KmIni(0 To cantVehiculos - 1) As Integer
      ReDim pvVehEsp_KmFin(0 To cantVehiculos - 1) As Integer
      For mj = 2 To FlexVehEspAsign.Rows - 1
         pvVehEsp_Codigo(mj - 2) = FlexVehEspAsign.TextMatrix(mj, 4)
         pvVehEsp_KmIni(mj - 2) = FlexVehEspAsign.TextMatrix(mj, 2)
         pvVehEsp_KmFin(mj - 2) = FlexVehEspAsign.TextMatrix(mj, 3)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      'Podria ocurrir que no tenga vehiculo asociado a la ot
      ReDim pvVehEsp_Codigo(0)
      pvVehEsp_Codigo(0) = "00"
   End If
End Sub

Private Function fValidaOT() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mCodTipoVale As String
   Dim mRec1 As New ADODB.Recordset
   Dim mi As Integer
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   Dim iStock As Double

   mRet = True
   
   'Valido si existe Orden de Trabajo
   If Trim(Combo3.Text) = "" Then
      mRet = False
      mMensajeError = "Debe seleccionar una Orden de trabajo"
   End If
   
   'Valido Fecha Inicio Valida
   If mRet Then
      If Not IsDate(Text3(0).Text) Then
         mRet = False
         mMensajeError = "La Fecha de Inicio no es válida"
      End If
   End If
   
   'Valido Fecha Fin Valida
   If mRet Then
      If Not IsDate(Text3(1).Text) Then
         mRet = False
         mMensajeError = "La Fecha de Fin no es válida"
      End If
   End If
  
  'Valido Fecha OT <= Fecha Inicio
  If mRet Then
      If CDate(Mid(Combo3.Text, 14, 10)) > CDate(Text3(0).Text) Then
         mRet = False
         mMensajeError = "La Fecha de la OT no puede ser mayor a la 'Fecha Inicio'"
      End If
   End If
   
   'Valido Fecha Inicio <= Fecha Actual
   If mRet Then
      If CDate(Text3(0).Text) > CDate(Format(Now, "dd/mm/yyyy")) Then
         mRet = False
         mMensajeError = "La 'Fecha Inicio' no puede ser mayor a la 'Fecha Actual'"
      End If
   End If
   
   'Valido Fecha Fin <= Fecha Actual
   If mRet Then
      If CDate(Text3(1).Text) > CDate(Format(Now, "dd/mm/yyyy")) Then
         mRet = False
         mMensajeError = "La 'Fecha Fin' no puede ser mayor a la 'Fecha Actual'"
      End If
   End If
   
   'Valido Fecha Inicio <= Fecha Fin
   If mRet Then
      If CDate(Text3(0).Text) > CDate(Text3(1).Text) Then
         mRet = False
         mMensajeError = "La 'Fecha Inicio' no puede ser mayor a la 'Fecha Fin'"
      End If
   End If
   
   'Valido que al menos exista un tecnico
   If mRet Then
      If FlexMoAsig.Rows <= 2 Then
         mRet = False
         Me.TabStrip1.Tabs(2).Selected = True
         mMensajeError = "Al menos se debe seleccionar un técnico"
      End If
   End If

   'Valido que al menos exista un Subrubro.
   If mRet Then
      If FlexMoAsig.Rows <= 2 Then
         mRet = False
         Me.TabStrip1.Tabs(3).Selected = True
         mMensajeError = "Al menos se debe seleccionar un técnico"
      End If
   End If
   
   'Valido Km iniciales de los vehiculo comunes
   If mRet Then
      For mi = 2 To FlexVehAsign.Rows - 1
   
         posInstr = InStr(1, Replace(FlexVehAsign.TextMatrix(mi, 3), ".", ","), ",")
      
         qtyDecimales = 0
         If posInstr <> 0 Then
            qtyDecimales = Len(Right(Trim(FlexVehAsign.TextMatrix(mi, 3)), Len(Trim(FlexVehAsign.TextMatrix(mi, 3))) - posInstr))
         End If
   
         'Valido Km Inicial que sea númerico.
         If Not IsNumeric(Replace(FlexVehAsign.TextMatrix(mi, 3), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "Se ha cargado un 'Km Inicial' incorrecto para el Vehículo: '" & FlexVehAsign.TextMatrix(mi, 2) & "'"
            mi = 9999
         End If
            
         'Valido Km Inicial que no sea decimal.
         If mRet Then
            If qtyDecimales > 0 Then
               mRet = False
               Me.TabStrip1.Tabs(2).Selected = True
               mMensajeError = "El 'Km Inicial' ingresado para '" & FlexVehAsign.TextMatrix(mi, 2) & "' no puede tener decimales"
               mi = 9999
            End If
         End If
         
         'Valido Km Inicial que sea > a cero.
         If Replace(FlexVehAsign.TextMatrix(mi, 3), ".", ",") <= 0 Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Inicial' para el Vehículo: '" & FlexVehAsign.TextMatrix(mi, 2) & "' no puede ser menor o igual a cero."
            mi = 9999
         End If
        Next
   End If
   
   'Valido Km finales de los vehiculo comunes
   If mRet Then
      For mi = 2 To FlexVehAsign.Rows - 1
         
         posInstr = InStr(1, Replace(FlexVehAsign.TextMatrix(mi, 4), ".", ","), ",")
      
         qtyDecimales = 0
         If posInstr <> 0 Then
            qtyDecimales = Len(Right(Trim(FlexVehAsign.TextMatrix(mi, 4)), Len(Trim(FlexVehAsign.TextMatrix(mi, 4))) - posInstr))
         End If
   
         'Valido Km Inicial que sea númerico.
         If Not IsNumeric(Replace(FlexVehAsign.TextMatrix(mi, 4), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "Se ha cargado un 'Km Final' incorrecto para el Vehículo: '" & FlexVehAsign.TextMatrix(mi, 2) & "'"
            mi = 9999
         End If
            
         'Valido Km Inicial que no sea decimal.
         If mRet Then
            If qtyDecimales > 0 Then
               mRet = False
               Me.TabStrip1.Tabs(2).Selected = True
               mMensajeError = "El 'Km Final' ingresado para '" & FlexVehAsign.TextMatrix(mi, 2) & "' no puede tener decimales"
               mi = 9999
            End If
         End If
         
         'Valido Km Final que sea > a cero.
         If Replace(FlexVehAsign.TextMatrix(mi, 4), ".", ",") <= 0 Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Final' para el Vehículo: '" & FlexVehAsign.TextMatrix(mi, 2) & "' no puede ser menor o igual a cero."
            mi = 9999
         End If
      Next
   End If
   
   'Valido Km Inicial < Km Final (vehiculos comunes)
   If mRet Then
      For mi = 2 To FlexVehAsign.Rows - 1
         If CDbl(Replace(FlexVehAsign.TextMatrix(mi, 3), ".", ",")) > CDbl(Replace(FlexVehAsign.TextMatrix(mi, 4), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Incial' no puede ser mayor que el 'Km Final' para '" & FlexVehAsign.TextMatrix(mi, 2) & "' "
            mi = 9999
         End If
      Next
   End If
   
   '--------------------------------------------------
   'Valido Km iniciales de los vehiculo espciales
   
    If mRet Then
      For mi = 2 To FlexVehEspAsign.Rows - 1
   
         posInstr = InStr(1, Replace(FlexVehEspAsign.TextMatrix(mi, 2), ".", ","), ",")
      
         qtyDecimales = 0
         If posInstr <> 0 Then
            qtyDecimales = Len(Right(Trim(FlexVehEspAsign.TextMatrix(mi, 2)), Len(Trim(FlexVehEspAsign.TextMatrix(mi, 2))) - posInstr))
         End If
   
         'Valido Km Inicial que sea númerico.
         If Not IsNumeric(Replace(FlexVehEspAsign.TextMatrix(mi, 2), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "Se ha cargado un 'Km Inicial' incorrecto para el Vehículo Especial: '" & FlexVehEspAsign.TextMatrix(mi, 1) & "'"
            mi = 9999
         End If
            
         'Valido Km Inicial que no sea decimal.
         If mRet Then
            If qtyDecimales > 0 Then
               mRet = False
               Me.TabStrip1.Tabs(2).Selected = True
               mMensajeError = "El 'Km Inicial' ingresado para '" & FlexVehEspAsign.TextMatrix(mi, 1) & "' no puede tener decimales"
               mi = 9999
            End If
         End If
         
         'Valido Km Inicial que sea > a cero.
         If Replace(FlexVehEspAsign.TextMatrix(mi, 2), ".", ",") <= 0 Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Inicial' para el Vehículo: '" & FlexVehEspAsign.TextMatrix(mi, 1) & "' no puede ser menor o igual a cero."
            mi = 9999
         End If

         
      Next
   End If

   'Valido Km finales de los vehiculo especiales
   If mRet Then
      For mi = 2 To FlexVehEspAsign.Rows - 1

         posInstr = InStr(1, Replace(FlexVehEspAsign.TextMatrix(mi, 3), ".", ","), ",")

         qtyDecimales = 0
         If posInstr <> 0 Then
            qtyDecimales = Len(Right(Trim(FlexVehEspAsign.TextMatrix(mi, 3)), Len(Trim(FlexVehEspAsign.TextMatrix(mi, 3))) - posInstr))
         End If

         'Valido Km final que sea númerico.
         If Not IsNumeric(Replace(FlexVehEspAsign.TextMatrix(mi, 3), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "Se ha cargado un 'Km Final' incorrecto para el Vehículo Especial: '" & FlexVehEspAsign.TextMatrix(mi, 1) & "'"
            mi = 9999
         End If

         'Valido Km Final que no sea decimal.
         If mRet Then
            If qtyDecimales > 0 Then
               mRet = False
               Me.TabStrip1.Tabs(2).Selected = True
               mMensajeError = "El 'Km Final' ingresado para '" & FlexVehEspAsign.TextMatrix(mi, 1) & "' no puede tener decimales"
               mi = 9999
            End If
         End If
         
         'Valido Km Inicial que sea > a cero.
         If Replace(FlexVehEspAsign.TextMatrix(mi, 3), ".", ",") <= 0 Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Final' para el Vehículo: '" & FlexVehEspAsign.TextMatrix(mi, 1) & "' no puede ser menor o igual a cero."
            mi = 9999
         End If
      Next
   End If
   
   'Valido Km Inicial < Km Final (vehiculos Especiales)
   If mRet Then
      For mi = 2 To FlexVehEspAsign.Rows - 1
         If CDbl(Replace(FlexVehEspAsign.TextMatrix(mi, 2), ".", ",")) > CDbl(Replace(FlexVehEspAsign.TextMatrix(mi, 3), ".", ",")) Then
            mRet = False
            Me.TabStrip1.Tabs(2).Selected = True
            mMensajeError = "El 'Km Incial' no puede ser mayor que el 'Km Final' para '" & FlexVehEspAsign.TextMatrix(mi, 1) & "' "
            mi = 9999
         End If
      Next
   End If
   
   
   'Valido Cantidad valida, cantidad decimales <2 t  saldo del stock insuficiente para ese Producto/Ubicación
   If mRet Then
      If mRet Then
         For mi = 2 To FlexProduct.Rows - 1
            Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                                      " FROM Movimientos2 M " & _
                                                      " WHERE CodProducto  = '" & FlexProduct.TextMatrix(mi, 6) & "' and CodUbicacion = '" & FlexProduct.TextMatrix(mi, 7) & "'" & _
                                                      " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close
            
            posInstr = InStr(1, Replace(FlexProduct.TextMatrix(mi, 3), ".", ","), ",")
      
            qtyDecimales = 0
            If posInstr <> 0 Then
               qtyDecimales = Len(Right(Trim(FlexProduct.TextMatrix(mi, 3)), Len(Trim(FlexProduct.TextMatrix(mi, 3))) - posInstr))
            End If
            
            
            'Valido valor numerico
            If Not IsNumeric(Replace(FlexProduct.TextMatrix(mi, 3), ".", ",")) Then
               mRet = False
               Me.TabStrip1.Tabs(4).Selected = True
               mMensajeError = "Se ha cargado un valor incorrecto para el producto: '" & FlexProduct.TextMatrix(mi, 2) & "'"
               mi = 9999
            End If
            
            'Valido cantidad decimales
            If mRet Then
               If qtyDecimales > 2 Then
                  mRet = False
                  Me.TabStrip1.Tabs(4).Selected = True
                  mMensajeError = "La Cantidad ingresada para  ' " & FlexProduct.TextMatrix(mi, 2) & " ' no puede tener mas de dos decimales"
                  mi = 9999
               End If
            End If
            
            'Valido saldo insuficiente
            If mRet Then
               If CDbl(Replace(Trim(FlexProduct.TextMatrix(mi, 3)), ".", ",")) > iStock Then
                  mRet = False
                  Me.TabStrip1.Tabs(4).Selected = True
                  mMensajeError = "El stock es insuficiente para ' " & FlexProduct.TextMatrix(mi, 2) & " '"
                  mi = 9999
               End If
            End If
         Next
      End If
   End If
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If

   fValidaOT = mRet
End Function

Private Sub imprimirExcelOT(ByVal NroOT As Integer, FechaOT As Date, Supervisor As String)
         sMsgEspere Me, "Procesando datos...", True
         'mFechaEjec = Now()
         Set XLS = CreateObject("Excel.Application")
         sPlanilla1 NroOT, FechaOT, Supervisor
         XLS.Worksheets(1).Select
         sMsgEspere Me, "", False
         XLS.Application.Visible = True
End Sub

Private Sub sPlanilla1(NroOT As Integer, FechaOT As Date, Supervisor As String)
'   mI = 10
   sCabecera1 NroOT, FechaOT, Supervisor
End Sub

Private Sub sCabecera1(NroOT As Integer, FechaOT As Date, Supervisor As String)
   Dim mi As Integer
   Dim primerColumna As Boolean
   
   mi = 10
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Orden de Trabajo"
      .Columns("A:A").ColumnWidth = 1.14 '
      .Columns("B:B").ColumnWidth = 6.86 '
      .Columns("C:C").ColumnWidth = 24.29 '
      .Columns("J:J").ColumnWidth = 1.14 '

      .Range("B1:J500").Select
      .Selection.Font.Size = 7
      .Selection.Font.Bold = True
      .Selection.RowHeight = 10.5

'---------------------------------ENCABEZADO HOJA-------------------------------------------------------
      .Cells(1, 2).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(2, 4).Formula = "PLANILLA DE ORDEN DE TRABAJO"
      
      .Cells(4, 2).Formula = "Fecha: " & FechaOT
      .Cells(5, 2).Formula = "Tipo Tarea"
      .Cells(6, 2).Formula = "Supervisor: " & Supervisor
      
      .Cells(4, 8).Formula = "Nº OT"
      .Cells(5, 8).Formula = "Hora Inicio"
      .Cells(6, 8).Formula = "Hora Fin"
      .Cells(4, 9).Formula = NroOT

      .Range("H4:H6").Select
      .Selection.Interior.ColorIndex = 15

      .Range("H4:I6").Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'---------------------------------ENCABEZADO TECNICOS---------------------------------------------------
       .Cells(9, 4).Formula = "TECNICOS QUE INTERVIENEN"
      
      .Range("B9:H9").Select
      .Selection.Interior.ColorIndex = 15
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TECNICOS----------------------------------------------------
      primerColumna = 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion " & _
                                          "FROM OT_MO_Tecnicos O " & _
                                              "Inner Join " & _
                                          "MO_Tecnicos M ON O.CodMO_Tecnico = M.Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "';")
                                
      Do While Not mRec.EOF
         .Range("B" & mi & ":H" & mi).Select
         With .Selection.Borders(xlBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         .Range("E" & mi & ":E" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         If primerColumna Then
            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mi, 5).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mi = mi + 1
         End If
   
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO VEHICULOS------------------------------------------------

      mi = mi + 2

      .Range("B" & mi & ":H" & (mi + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mi & ":H" & (mi + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("E" & (mi + 1) & ":H" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
       With .Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
      
      .Cells(mi, 4).Formula = "VEHICULOS QUE INTERVIENEN"
      mi = mi + 1
      .Cells(mi, 2).Formula = "Vehículo"
      .Cells(mi, 5).Formula = "Km Inicial"
      .Cells(mi, 7).Formula = "Km Final"

'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE VEHICULOS------------------------------------------------
      mi = mi + 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo,Descripcion FROM " & _
                                          "OT_Vehiculos O " & _
                                              "Inner Join " & _
                                          "Vehiculos V ON O.CodVehiculo = Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "'; ")
                                
      Do While Not mRec.EOF
         .Range("B" & mi & ":H" & mi).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("E" & mi & ":E" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("G" & mi & ":G" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
         
         With XLS
            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mi = mi + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO TAREAS--------------------------------------------------
      mi = mi + 2
      
      .Cells(mi, 5).Formula = "TAREAS"
      .Cells(mi + 1, 2).Formula = "Parte"
      .Cells(mi + 1, 3).Formula = "Lugar"
      .Cells(mi + 1, 4).Formula = "Descripcion"
      .Cells(mi + 1, 9).Formula = "¿Finalizado?"

      .Range("B" & mi & ":I" & (mi + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mi & ":I" & (mi + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("C" & (mi + 1) & ":C" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("D" & (mi + 1) & ":D" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("I" & (mi + 1) & ":I" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TAREAS------------------------------------------------------
      mi = mi + 2
      Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,R.CodEdificio, R.Descripcion FROM " & _
                                          "OT_Partes O " & _
                                              "Inner Join " & _
                                          "Registros R ON O.Parte = R.Parte " & _
                                          "WHERE IDOT = '" & NroOT & "' " & _
                                          "ORDER BY R.parte; ")
                                
      Do While Not mRec.EOF
         .Range("B" & mi & ":I" & mi).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         .Range("C" & mi & ":C" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("D" & mi & ":D" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         .Range("I" & mi & ":I" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With XLS
            .Cells(mi, 2).Formula = NVL(mRec!Parte, "")
            .Cells(mi, 3).Formula = NVL(mRec!CodEdificio, "")
            .Cells(mi, 4).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mi = mi + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------
'---------------------------------ENCABEZADO SUBRUBROS------------------------------------------------
      mi = mi + 2
      
      .Cells(mi, 5).Formula = "FALLAS"
      .Cells(mi + 1, 2).Formula = "Subrubro"
      .Cells(mi + 1, 6).Formula = "Subrubro"

      .Range("B" & mi & ":I" & (mi + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mi & ":I" & (mi + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("E" & mi + 1 & ":E" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
      End With

      .Range("I" & mi + 1 & ":I" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE SUBRUBROS-----------------------------------------------
      mi = mi + 2
      Set mRec = mObj.oEjecutarSelect("SELECT S.Codigo,S.Descripcion FROM " & _
                                       "SubRubros S " & _
                                          "Inner Join " & _
                                       "OT_Subrubros O ON O.CodSubrubro = S.Codigo " & _
                                       "WHERE IDOT = '" & NroOT & "' " & _
                                       "ORDER BY S.Descripcion; ")
                                
      primerColumna = True
      Do While Not mRec.EOF
   
         If primerColumna Then
            .Range("B" & mi & ":I" & mi).Select
            With .Selection.Borders(xlBottom)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlTop)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
      
            With .Selection.Borders(xlEdgeRight)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
         
            .Range("E" & mi & ":E" & mi).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
         
            .Range("F" & mi & ":F" & mi).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlMedium
              .ColorIndex = xlAutomatic
            End With
      
            .Range("I" & mi & ":I" & mi).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
            
            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mi, 6).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mi = mi + 1
         End If
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------

'---------------------------------ENCABEZADO Materiales-----------------------------------------------
      mi = mi + 2
      .Cells(mi, 4).Formula = "                 MATERIALES"
      .Cells(mi + 1, 2).Formula = "Cód.Sap"
      .Cells(mi + 1, 3).Formula = "Descripción"
      .Cells(mi + 1, 6).Formula = "Consumido"
      .Cells(mi + 1, 7).Formula = "Unid. Media"
      
      .Range("B" & mi & ":H" & (mi + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mi & ":H" & (mi + 1)).Select
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      .Range("C" & (mi + 1) & ":C" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("F" & (mi + 1) & ":F" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      .Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE Materiales--------------------------------------------------
      mi = mi + 2
      Set mRec = mObj.oEjecutarSelect("SELECT  idMov,  M.Fecha,  P.CodigoSap,  P.Descripcion,  Stock,  UM.Descripcion AS UnidadMedidad FROM " & _
                                          "Inventario.Movimientos2 M " & _
                                              "Inner Join " & _
                                          "Inventario.Producto P ON M.CodProducto = P.Codigo " & _
                                              "Inner Join " & _
                                          "Inventario.UnidadMedida UM ON P.CodUnidadMedida = UM.Codigo " & _
                                              "Inner Join " & _
                                          "Inventario.Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
                                             "Inner Join " & _
                                          "Vehiculos V ON U.Codigo = V.CodUbicacion " & _
                                             "Inner Join " & _
                                          "OT_Vehiculos OV ON OV.CodVehiculo = V.Codigo " & _
                                          "WHERE M.Fecha = (SELECT MAX(Fecha) " & _
                                          "                From Inventario.Movimientos2 " & _
                                          "                WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
                                          "AND U.Codigo = '0006'" & _
                                          "and OV.IDOT = '" & NroOT & "' and stock > 0; ")
                                
      Do While Not mRec.EOF
         .Range("B" & mi & ":H" & mi).Select
         With .Selection.Borders(xlBottom)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlTop)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         With .Selection.Borders(xlEdgeRight)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("C" & mi & ":C" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("F" & mi & ":F" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         .Range("G" & mi & ":G" & mi).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         With XLS
            .Cells(mi, 2).Formula = NVL(mRec!CodigoSap, "")
            .Cells(mi, 3).Formula = NVL(mRec!descripcion, "")
            .Cells(mi, 7).Formula = NVL(mRec!UnidadMedidad, "")
         End With
         mRec.MoveNext
         mi = mi + 1
      Loop
      mRec.Close
   
'-----------------------------------------------------------------------------------------------------
 
 
'----------------------------------------------OBSERVACIONES------------------------------------------
      mi = mi + 2
      .Cells(mi, 2).Formula = "OBSERVACIONES"
      mi = mi + 1
      .Range("B" & mi & ":I" & (mi + 4)).Select
      .Selection.RowHeight = 16.5
      With .Selection.Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      With .Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

'-----------------------------------------------------------------------------------------------------

 
'----------------------------------------------FIRMAS-------------------------------------------------
      mi = mi + 8
      .Cells(mi, 3).Formula = "              SUPERVISOR"
      .Cells(mi, 6).Formula = "     ENCARGADO BODEGA"
      
      .Range("C" & mi & ":C" & mi).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("F" & mi & ":G" & mi).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
End With
'-----------------------------------------------------------------------------------------------------
 
'  Configuracion de margenes.
   With ActiveSheet.PageSetup
      .LeftMargin = Application.CentimetersToPoints(0)
      .RightMargin = Application.CentimetersToPoints(0)
      .TopMargin = Application.CentimetersToPoints(0)
      .BottomMargin = Application.CentimetersToPoints(0)
   End With
   
End Sub

Private Sub preparaArrayMO_Tecnicos(ByRef pvMO_Tecnicos_OT() As String)
   Dim mj As Integer
   Dim cantMO_Tecnicos As Integer

   cantMO_Tecnicos = FlexMoAsig.Rows - 2
   If cantMO_Tecnicos > 0 Then
      ReDim pvMO_Tecnicos_OT(0 To cantMO_Tecnicos - 1) As String
         
      For mj = 2 To FlexMoAsig.Rows - 1
         
         pvMO_Tecnicos_OT(mj - 2) = FlexMoAsig.TextMatrix(mj, 2)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvMO_Tecnicos_OT(0)
      pvMO_Tecnicos_OT(0) = "00"
   End If
End Sub

Private Sub preparaArraySubrubros(ByRef pvSubrubros_OT() As String)
   Dim mj As Integer
   Dim cantSubrubros As Integer

   cantSubrubros = FlexSubRubrosAsign.Rows - 2
   If cantSubrubros > 0 Then
      ReDim pvSubrubros_OT(0 To cantSubrubros - 1) As String
         
      For mj = 2 To FlexSubRubrosAsign.Rows - 1
        pvSubrubros_OT(mj - 2) = FlexSubRubrosAsign.TextMatrix(mj, 4)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvSubrubros_OT(0)
      pvSubrubros_OT(0) = "000000"
   End If
End Sub


Private Sub preparaArrayMateriales(ByRef pvMat_CodProd() As String, ByRef pvMat_CodUbic() As String, ByRef pvMat_Cantidad() As Double, ByRef pvMat_CantidadBD() As Double)
   Dim mj As Integer
   Dim cantMateriales As Integer

   cantMateriales = FlexProduct.Rows - 2
   If cantMateriales > 0 Then
      
      ReDim pvMat_CodProd(0 To cantMateriales - 1) As String
      ReDim pvMat_CodUbic(0 To cantMateriales - 1) As String
      ReDim pvMat_Cantidad(0 To cantMateriales - 1) As Double
      ReDim pvMat_CantidadBD(0 To cantMateriales - 1) As Double
      
      For mj = 2 To FlexProduct.Rows - 1
        pvMat_CodProd(mj - 2) = FlexProduct.TextMatrix(mj, 6)
        pvMat_CodUbic(mj - 2) = FlexProduct.TextMatrix(mj, 7)
        pvMat_Cantidad(mj - 2) = CDbl(Replace(FlexProduct.TextMatrix(mj, 3), ".", ","))
        pvMat_CantidadBD(mj - 2) = CDbl(Replace(FlexProduct.TextMatrix(mj, 8), ".", ","))
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvMat_CodProd(0)
      pvMat_CodProd(0) = "000000"
   End If
End Sub
Private Sub Command3_Click(Index As Integer)
   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset

   If Index = 0 Then
      If fValidaEgreso() Then
            FlexEgreso.AddItem vbTab & FlexProd.TextMatrix(mRenglonProducto, 1) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 2) & vbTab & Text10.Text & vbTab & FlexProd.TextMatrix(mRenglonProducto, 4) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 5) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 6) & vbTab & FlexProd.TextMatrix(mRenglonProducto, 7)
            FlexProd.TextMatrix(mRenglonProducto, 3) = CDbl(Replace(Trim(FlexProd.TextMatrix(mRenglonProducto, 3)), ".", ",")) - CDbl(Replace(Trim(Text10.Text), ".", ","))
            Text10.Text = ""
            Text10.SetFocus
      End If
   Else
      For mi = 2 To FlexProd.Rows - 1

         If FlexProd.TextMatrix(mi, 6) = FlexEgreso.TextMatrix(mRenglonEgreso, 6) And FlexProd.TextMatrix(mi, 7) = FlexEgreso.TextMatrix(mRenglonEgreso, 7) Then
            Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                       " FROM Movimientos2 M " & _
                                       " WHERE CodProducto  = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
                                       " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")

            If Not mRec1.EOF Then
               iStock = mRec1!Stock
            Else
               iStock = 0
            End If
            mRec1.Close

            FlexProd.TextMatrix(mi, 3) = iStock

            mi = 9999
         End If
      Next

      If FlexEgreso.Rows > 2 And mRenglonEgreso > 1 Then
         FlexEgreso.RemoveItem (mRenglonEgreso)
      End If

      mRenglonEgreso = 0
   End If
End Sub

Private Function fValidaEgreso() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mj As Integer
   Dim mCantidaStock As Double
   Dim sStock As String
   Dim iStock As Double
   Dim mRec1 As New ADODB.Recordset
   Dim posInstr As Integer
   Dim qtyDecimales As Integer
   Dim mCodTipoVale As String

   mRet = True

   If Trim(Text8.Text) = "" Then
      mRet = False
      mMensajeError = "Debe completar el Número de Vale"
   End If

   If mRet Then
      If Not IsNumeric(Trim(Text8.Text)) Then
         mRet = False
         mMensajeError = "El Nro. Vale debe ser numérico !!"
      End If
   End If


   If mRet Then
      If Len(Trim(Text8.Text)) <> 9 Then
         mRet = False
         mMensajeError = "El Nro. Vale debe tener 9 caracteres numéricos !!"
      End If
   End If

   If mRet Then
      If ((Not Option1.Value) And (Not Option2.Value)) Then
         mRet = False
         mMensajeError = "Debe completar el Tipo de Vale"
      End If
   End If


   If mRet Then
      If Option1.Value Then
         mCodTipoVale = "C"
      Else
         mCodTipoVale = "M"
      End If

      Set mRec1 = mObjInven.oEjecutarSelect("SELECT NroVale,CodTipoVale FROM Consumos_H WHERE NroVale = " & Trim(Text8.Text) & " and CodTipoVale = '" & mCodTipoVale & "'; ")
      If Not mRec1.EOF Then
         mRet = False
         mMensajeError = "Ya se han realizado consumos para ese Número y Tipo de Vale !!!"
      End If
      mRec1.Close
   End If

   If mRet Then
      If mRenglonProducto = 0 Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If

   If mRet Then
      If mRenglonProducto <> 0 And FlexProd.TextMatrix(mRenglonProducto, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If

   If mRet Then
      If Trim(Text10.Text) = "" Then
         mRet = False
         mMensajeError = "Debe completar el campo: 'Cantidad'. "
      End If
   End If

   If mRet Then
      If Not IsNumeric(Replace(Text10.Text, ".", ",")) Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no es un valor numérico"
      End If
   End If

   If mRet Then
      If CDbl(Replace(Trim(Text10.Text), ".", ",")) <= 0 Then
         mRet = False
         mMensajeError = "La Cantidad ingresada no puede ser menor o igual a cero."
      End If
   End If

   'Valido que no supere los 2 digitos decimales
   If mRet Then
      posInstr = InStr(1, Replace(Trim(Text10.Text), ".", ","), ",")

      If posInstr <> 0 Then
         qtyDecimales = Len(Right(Trim(Text10.Text), Len(Trim(Text10.Text)) - posInstr))
      End If

      If qtyDecimales > 2 Then
         mRet = False
         mMensajeError = "El campo 'Cantidad' solo admite hasta dos dígitos decimales."
      End If
   End If

   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
   If mRet Then
      For mj = 2 To FlexEgreso.Rows - 1
         If FlexEgreso.TextMatrix(mj, 6) = FlexProd.TextMatrix(mRenglonProducto, 6) And FlexEgreso.TextMatrix(mj, 7) = FlexProd.TextMatrix(mRenglonProducto, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
            mj = 999
         End If
      Next
   End If

   'Valido si el saldo del stock es insuficiente para ese Producto/Ubicación
   If mRet Then

      Set mRec1 = mObjInven.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
                                                " FROM Movimientos2 M " & _
                                                " WHERE CodProducto  = '" & FlexProd.TextMatrix(mRenglonProducto, 6) & "' and CodUbicacion = '" & FlexProd.TextMatrix(mRenglonProducto, 7) & "'" & _
                                                " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
      If Not mRec1.EOF Then
         iStock = mRec1!Stock
      Else
         iStock = 0
      End If
      mRec1.Close

      If CDbl(Replace(Trim(Text10.Text), ".", ",")) > iStock Then
         mRet = False
         mMensajeError = "El stock es insuficiente para ese Producto en esa Ubicación"
      End If
   End If

   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaEgreso = mRet
End Function

Private Sub CommandSubRubro_Click(Index As Integer)
   Dim sListaSubrubrosSeleccionados
   Dim mj As Integer
   sListaSubrubrosSeleccionados = ""
   
   If Index = 0 Then
      If mRenglonSubRubroDispo > 0 Then
         If Trim(FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1)) <> "" Then
            FlexSubRubrosAsign.AddItem vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 2) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 3) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 4)
         End If
         
         If FlexSubRubros.Rows > 2 Then
            FlexSubRubros.RemoveItem mRenglonSubRubroDispo
         
            mRenglonSubRubroDispo = 0
         Else
            If Trim(FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1)) <> "" Then
               FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1) = ""
               FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 2) = ""
         
               mRenglonSubRubroDispo = 0
            End If
         End If
      End If
   Else
      If FlexSubRubrosAsign.Rows > 2 And mRenglonSubRubroAsign > 1 Then
         
         FlexSubRubrosAsign.RemoveItem (mRenglonSubRubroAsign)
         
         If FlexSubRubrosAsign.Rows > 2 Then
            For mj = 2 To FlexSubRubrosAsign.Rows - 1
               sListaSubrubrosSeleccionados = sListaSubrubrosSeleccionados & "'" & FlexSubRubrosAsign.TextMatrix(mj, 4) & "',"
            Next
            sListaSubrubrosSeleccionados = Left(sListaSubrubrosSeleccionados, Len(sListaSubrubrosSeleccionados) - 1)
        End If
            
         mRenglonSubRubroDispo = 0
         
         FlexSubRubros.Clear
         'Elimino los registros  de la grilla superior
         For mj = FlexSubRubros.Rows To 3 Step -1
            FlexSubRubros.RemoveItem mj
         Next
         
         With FlexSubRubros
            .TextMatrix(0, 1) = "Rubro"
            .TextMatrix(0, 2) = "SubRubro"
         
            .RowHeight(1) = 0
         End With
         
         If FlexSubRubrosAsign.Rows > 2 Then
            Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
               "  From " & _
               " Rubros R " & _
               "  Inner Join " & _
               " SubRubros S ON S.CodRubro = R.Codigo " & _
               " WHERE S.Codigo NOT IN (" & sListaSubrubrosSeleccionados & ")" & _
               " AND R.Codigo ='" & Right(Combo1.Text, 8) & "' ORDER BY RubroDesc, SubRubroDesc;")
         Else
            Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
               "  From " & _
               " Rubros R " & _
               "  Inner Join " & _
               " SubRubros S ON S.CodRubro = R.Codigo" & _
               " WHERE R.Codigo ='" & Right(Combo1.Text, 8) & "' ORDER BY RubroDesc, SubRubroDesc;")
         End If
         
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
            
               With FlexSubRubros
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!RubroDesc
                  .TextMatrix(mj, 2) = mRec!SubRubroDesc
                  .TextMatrix(mj, 3) = mRec!CodRubro
                  .TextMatrix(mj, 4) = mRec!CodSubrubro
               End With
               
               mRec.MoveNext
            Loop
         End If
         mRec.Close

      End If
      mRenglonSubRubroAsign = 0
   End If
End Sub

Private Sub FlexEgreso_Click()
   Dim mi As Integer
   If FlexEgreso.MouseRow > 0 Then
      If mRenglonEgreso <> 0 Then
         If FlexEgreso.Rows > mRenglonEgreso Then
            FlexEgreso.Row = mRenglonEgreso
            For mi = 1 To FlexEgreso.Cols - 1
               FlexEgreso.Col = mi
               FlexEgreso.CellBackColor = vbWhite
            Next
         End If
      End If
      mRenglonEgreso = FlexEgreso.MouseRow
      FlexEgreso.Row = mRenglonEgreso
      For mi = 1 To FlexEgreso.Cols - 1
         FlexEgreso.Col = mi
         FlexEgreso.CellBackColor = &H80000003
      Next
      If mRenglonEgreso > 1 Then
          mCodProducto = FlexEgreso.TextMatrix(mRenglonEgreso, 4)
      End If
   Else
      FlexEgreso.Row = mRenglonEgreso
      If FlexEgreso.Row > 0 Then
         For mi = 1 To FlexProd.Cols - 1
            FlexEgreso.Col = mi
            FlexEgreso.CellBackColor = vbWhite
         Next
      End If
      mRenglonEgreso = 0
   End If
End Sub

Private Sub FlexMoAsig_Click()
   Dim mi As Integer
   Dim resultado As String
   
   If FlexMoAsig.MouseCol = 0 And FlexMoAsig.MouseRow > 0 Then
      If mRenglonMoAsign <> 0 Then
         FlexMoAsig.Row = mRenglonMoAsign
         For mi = 1 To FlexMoAsig.Cols - 1
            FlexMoAsig.Col = mi
            FlexMoAsig.CellBackColor = vbWhite
         Next
      End If
      mRenglonMoAsign = FlexMoAsig.MouseRow
      FlexMoAsig.Row = mRenglonMoAsign
      For mi = 1 To FlexMoAsig.Cols - 1
         FlexMoAsig.Col = mi
         FlexMoAsig.CellBackColor = &H80000003
      Next
      If mRenglonMoAsign > 1 Then
         If Not mEsOTcerrada Then
            If FlexMoAsig.Rows > 3 Then
               mCodMO = FlexMoAsig.TextMatrix(mRenglonMoAsign, 1)
               resultado = MsgBox(" ¿ Desea quitar al Técnico  " & mCodMO & " de esta Orden de Trabajo ?", vbOKCancel, "Quitar Técnico de OT")
               If resultado = vbOK Then
                  If FlexMoAsig.Rows > 2 Then
                     FlexMoAsig.RemoveItem mRenglonMoAsign
                     mRenglonMoAsign = 0
                  Else
                     If Trim(FlexMoAsig.TextMatrix(mRenglonMoAsign, 1)) <> "" Then
                        FlexMoAsig.TextMatrix(mRenglonMoAsign, 1) = ""
                        FlexMoAsig.TextMatrix(mRenglonMoAsign, 2) = ""
                        mRenglonMoAsign = 0
                     End If
                  End If
               End If
            Else
               MsgBox "No es posible quitar todos los Técnicos de una OT"
            End If
         Else
            MsgBox "No es posible realizar esta operación cuando la O.T. está cerrada", vbExclamation
         End If
      End If
   Else
      FlexMoAsig.Row = mRenglonMoAsign
      If FlexMoAsig.Row > 0 Then
         For mi = 1 To FlexMoAsig.Cols - 1
            FlexMoAsig.Col = mi
            FlexMoAsig.CellBackColor = vbWhite
         Next
      End If
      mRenglonMoAsign = 0
   End If
End Sub

Private Sub FlexPartes_Click()
   Dim mi As Integer
   Dim resultado As String
   
If FlexPartes.MouseCol = 0 Then
   If FlexPartes.MouseCol = 0 And FlexPartes.MouseRow > 0 Then
      If mRenglonPartes <> 0 Then
         FlexPartes.Row = mRenglonPartes
         For mi = 1 To FlexPartes.Cols - 1
            FlexPartes.Col = mi
            FlexPartes.CellBackColor = vbWhite
         Next
      End If
      mRenglonPartes = FlexPartes.MouseRow
      FlexPartes.Row = mRenglonPartes
      For mi = 1 To FlexPartes.Cols - 1
         FlexPartes.Col = mi
         FlexPartes.CellBackColor = &H80000003
      Next
      If mRenglonPartes > 1 Then
         '-------
         If Not mEsOTcerrada Then
         
            If FlexPartes.Rows > 3 Then
               mCodParte = FlexPartes.TextMatrix(mRenglonPartes, 1)
               resultado = MsgBox(" ¿ Desea eliminar el Parte número " & mCodParte & " de esta Orden de Trabajo ?", vbOKCancel, "Eliminar Parte de OT")
               If resultado = vbOK Then
                  If FlexPartes.Rows > 2 Then
                     FlexPartes.RemoveItem mRenglonPartes
                     mRenglonPartes = 0
                  Else
                     If Trim(FlexPartes.TextMatrix(mRenglonPartes, 1)) <> "" Then
                        FlexPartes.TextMatrix(mRenglonPartes, 1) = ""
                        FlexPartes.TextMatrix(mRenglonPartes, 2) = ""
                        mRenglonPartes = 0
                     End If
                  End If
               End If
               '--   End If
            Else
               MsgBox "No es posible eliminar todos los partes de una O.T.", vbExclamation
            End If
         Else
            MsgBox "No es posible realizar esta operación cuando la O.T. está cerrada", vbExclamation
         End If
         
         
      End If
   Else
      FlexPartes.Row = mRenglonPartes
      If FlexPartes.Row <> 0 Then
         For mi = 1 To FlexPartes.Cols - 1
            FlexPartes.Col = mi
            FlexPartes.CellBackColor = vbWhite
         Next
      End If
         mRenglonPartes = 0
      End If
'
'
'Else
'   MsgBox "No es posible realizar esta operación cuando la O.T. está cerrada", vbExclamation
'   mRenglonPartes = 0
End If
End Sub

Private Sub FlexProd_Click()
   Dim mi As Integer
   If FlexProd.MouseRow > 0 Then
      If mRenglonProducto <> 0 Then
         FlexProd.Row = mRenglonProducto
         For mi = 1 To FlexProd.Cols - 1
            FlexProd.Col = mi
            FlexProd.CellBackColor = vbWhite
         Next
      End If
      mRenglonProducto = FlexProd.MouseRow
      FlexProd.Row = mRenglonProducto
      For mi = 1 To FlexProd.Cols - 1
         FlexProd.Col = mi
         FlexProd.CellBackColor = &H80000003
      Next
      If mRenglonProducto > 1 Then
          mCodProducto = FlexProd.TextMatrix(mRenglonProducto, 4)
      End If
   Else
      FlexProd.Row = mRenglonProducto
      If FlexProd.Row > 0 Then
         For mi = 1 To FlexProd.Cols - 1
            FlexProd.Col = mi
            FlexProd.CellBackColor = vbWhite
         Next
      End If
      mRenglonProducto = 0
   End If
End Sub

Private Sub FlexProduct_Click()
   Dim mi As Integer
   
   If FlexProduct.MouseRow > 0 Then
      If Not mEsOTcerrada Then
         'En este caso 3 es la columna que seria editable
         If FlexProduct.Col = 3 And FlexProduct.Row <> 1 Then
            Text2.Text = FlexProduct.Text
            Text2.Width = FlexProduct.ColWidth(FlexProduct.Col)
            Text2.Left = FlexProduct.ColPos(FlexProduct.Col) + FlexProduct.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text2.Top = FlexProduct.Top + FlexProduct.RowPos(FlexProduct.Row)
            Text2.Visible = True
            Text2.SetFocus
            FlexProduct.ScrollBars = flexScrollBarNone
         Else
            Text2.Visible = False
            FlexProduct.ScrollBars = flexScrollBarVertical
         End If
      
         filaAnt = FlexProduct.Row
         columnAnt = FlexProduct.Col
      End If
      If mRenglonProdDispo <> 0 Then
         FlexProduct.Row = mRenglonProdDispo
         For mi = 1 To FlexProduct.Cols - 1
            FlexProduct.Col = mi
            FlexProduct.CellBackColor = vbWhite
         Next
      End If
      mRenglonProdDispo = FlexProduct.MouseRow
      FlexProduct.Row = mRenglonProdDispo
      For mi = 1 To FlexProduct.Cols - 1
         FlexProduct.Col = mi
         FlexProduct.CellBackColor = &H80000003
      Next
      If mRenglonProdDispo > 1 Then
          mCodProducto = FlexProduct.TextMatrix(mRenglonProdDispo, 4)
      End If
   Else
      FlexProduct.Row = mRenglonProdDispo
      If FlexProduct.Row > 0 Then
         For mi = 1 To FlexProduct.Cols - 1
            FlexProduct.Col = mi
            FlexProduct.CellBackColor = vbWhite
         Next
      End If
      mRenglonProdDispo = 0
   End If
End Sub

Private Sub FlexSubRubros_Click()
   Dim mi As Integer
   If FlexSubRubros.MouseRow > 0 Then
      If mRenglonSubRubroDispo <> 0 Then
         FlexSubRubros.Row = mRenglonSubRubroDispo
         For mi = 1 To FlexSubRubros.Cols - 1
            FlexSubRubros.Col = mi
            FlexSubRubros.CellBackColor = vbWhite
         Next
      End If
      mRenglonSubRubroDispo = FlexSubRubros.MouseRow
      FlexSubRubros.Row = mRenglonSubRubroDispo
      For mi = 1 To FlexSubRubros.Cols - 1
         FlexSubRubros.Col = mi
         FlexSubRubros.CellBackColor = &H80000003
      Next
      If mRenglonSubRubroDispo > 1 Then
          mCodSubrubro = FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 4)
      End If
   Else
      FlexSubRubros.Row = mRenglonSubRubroDispo
      If FlexSubRubros.Row > 0 Then
         For mi = 1 To FlexSubRubros.Cols - 1
            FlexSubRubros.Col = mi
            FlexSubRubros.CellBackColor = vbWhite
         Next
      End If
      mRenglonSubRubroDispo = 0
   End If
End Sub

Private Sub FlexSubRubrosAsign_Click()
   Dim mi As Integer
   If FlexSubRubrosAsign.MouseRow > 0 Then
      If mRenglonSubRubroAsign <> 0 Then
         FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
         For mi = 1 To FlexSubRubrosAsign.Cols - 1
            FlexSubRubrosAsign.Col = mi
            FlexSubRubrosAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonSubRubroAsign = FlexSubRubrosAsign.MouseRow
      FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
      For mi = 1 To FlexSubRubrosAsign.Cols - 1
         FlexSubRubrosAsign.Col = mi
         FlexSubRubrosAsign.CellBackColor = &H80000003
      Next
      If mRenglonSubRubroAsign > 1 Then
          mCodSubrubro = FlexSubRubrosAsign.TextMatrix(mRenglonSubRubroAsign, 4)
      End If
   Else
      FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
      If FlexSubRubrosAsign.Row > 0 Then
         For mi = 1 To FlexSubRubrosAsign.Cols - 1
            FlexSubRubrosAsign.Col = mi
            FlexSubRubrosAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonSubRubroAsign = 0
   End If
End Sub

Private Sub FlexVehAsign_Click()
   Dim mi As Integer
   Dim mColVehAsign As Integer
   Dim resultado As String
   
   mColVehAsign = FlexVehAsign.Col
   If FlexVehAsign.MouseRow > 0 Then
       If Not mEsOTcerrada Then
         'En este caso 3 es la columna que seria editable
         If FlexVehAsign.Col = 3 And FlexVehAsign.Row <> 1 Then
            Text4.Text = FlexVehAsign.Text
            Text4.Width = FlexVehAsign.ColWidth(FlexVehAsign.Col)
            Text4.Left = FlexVehAsign.ColPos(FlexVehAsign.Col) + FlexVehAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text4.Top = FlexVehAsign.Top + FlexVehAsign.RowPos(FlexVehAsign.Row)
            Text4.Visible = True
            Text4.SetFocus
            FlexVehAsign.ScrollBars = flexScrollBarNone
         Else
            Text4.Visible = False
            If FlexVehAsign.Col <> 4 Then
               FlexVehAsign.ScrollBars = flexScrollBarVertical
            End If
         End If
         filaAntVehAsign = FlexVehAsign.Row
         columnAntVehAsign = FlexVehAsign.Col
          'En este caso 4 es la columna que seria editable
         If FlexVehAsign.Col = 4 And FlexVehAsign.Row <> 1 Then
            Text5.Text = FlexVehAsign.Text
            Text5.Width = FlexVehAsign.ColWidth(FlexVehAsign.Col)
            Text5.Left = FlexVehAsign.ColPos(FlexVehAsign.Col) + FlexVehAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text5.Top = FlexVehAsign.Top + FlexVehAsign.RowPos(FlexVehAsign.Row)
            Text5.Visible = True
            Text5.SetFocus
            FlexVehAsign.ScrollBars = flexScrollBarNone
         Else
            Text5.Visible = False
            If FlexVehAsign.Col <> 3 Then
               FlexVehAsign.ScrollBars = flexScrollBarVertical
            End If
         End If
         filaAntVehAsignKmFinal = FlexVehAsign.Row
         columnAntVehAsignKmFinal = FlexVehAsign.Col
      End If
      If mRenglonVehAsign <> 0 Then
         FlexVehAsign.Row = mRenglonVehAsign
         For mi = 1 To FlexVehAsign.Cols - 1
            FlexVehAsign.Col = mi
            FlexVehAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonVehAsign = FlexVehAsign.MouseRow
      FlexVehAsign.Row = mRenglonVehAsign
      For mi = 1 To FlexVehAsign.Cols - 1
         FlexVehAsign.Col = mi
         FlexVehAsign.CellBackColor = &H80000003
      Next
      If mRenglonVehAsign > 1 Then
         If mColVehAsign = 1 Then
            If Not mEsOTcerrada Then
               mCodVeh = FlexVehAsign.TextMatrix(mRenglonVehAsign, 2)
               resultado = MsgBox(" ¿ Desea eliminar el Vehiculo  " & mCodVeh & " de esta Orden de Trabajo ?", vbOKCancel, "Eliminar Vehículo de OT")
               If resultado = vbOK Then
                  If FlexVehAsign.Rows > 2 Then
                     FlexVehAsign.RemoveItem mRenglonVehAsign
                     mRenglonVehAsign = 0
                  Else
                     If Trim(FlexVehAsign.TextMatrix(mRenglonVehAsign, 1)) <> "" Then
                        FlexVehAsign.TextMatrix(mRenglonVehAsign, 1) = ""
                        FlexVehAsign.TextMatrix(mRenglonVehAsign, 2) = ""
                        mRenglonVehAsign = 0
                     End If
                  End If
               End If
            Else
               MsgBox "No es posible realizar esta operación cuando la O.T. está cerrada", vbExclamation
            End If
         End If
      End If
   Else
      FlexVehAsign.Row = mRenglonVehAsign
      If FlexVehAsign.Row > 0 Then
         For mi = 1 To FlexVehAsign.Cols - 1
            FlexVehAsign.Col = mi
            FlexVehAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonVehAsign = 0
   End If
End Sub



Private Sub Text10_KeyPress(KeyAscii As Integer)
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
End Sub

Private Sub text8_KeyPress(KeyAscii As Integer)
         KeyAscii = fNumeroKeyPress(KeyAscii)
End Sub



Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = fNumeroKeyPress(KeyAscii)
   
   If KeyAscii = 13 Then
      FlexVehAsign.TextMatrix(filaAntVehAsign, columnAntVehAsign) = Text4.Text
      Text4.Visible = False
      FlexVehAsign.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text4_LostFocus()
   If FlexVehAsign.Col <> columnAntVehAsign Or FlexVehAsign.Row <> filaAntVehAsign Then
      'En este caso 3 es la columna que seria editable
      If columnAntVehAsign = 3 Then
         FlexVehAsign.TextMatrix(filaAntVehAsign, columnAntVehAsign) = Text4.Text
      End If
   End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = fNumeroKeyPress(KeyAscii)

   If KeyAscii = 13 Then
      FlexVehAsign.TextMatrix(filaAntVehAsignKmFinal, columnAntVehAsignKmFinal) = Text5.Text
      Text5.Visible = False
      FlexVehAsign.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text5_LostFocus()
   If FlexVehAsign.Col <> columnAntVehAsignKmFinal Or FlexVehAsign.Row <> filaAntVehAsignKmFinal Then
      'En este caso 4 es la columna que seria editable
      If columnAntVehAsignKmFinal = 4 Then
         FlexVehAsign.TextMatrix(filaAntVehAsignKmFinal, columnAntVehAsignKmFinal) = Text5.Text
      End If
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = fNumeroKeyPress(KeyAscii)

   If KeyAscii = 13 Then
      FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmInicio, columnAntVehEspAsignKmInicio) = Text6.Text
      Text6.Visible = False
      FlexVehEspAsign.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text6_LostFocus()
   If FlexVehEspAsign.Col <> columnAntVehEspAsignKmInicio Or FlexVehEspAsign.Row <> filaAntVehEspAsignKmInicio Then
      'En este caso 2 es la columna que seria editable
      If columnAntVehEspAsignKmInicio = 2 Then
         FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmInicio, columnAntVehEspAsignKmInicio) = Text6.Text
      End If
   End If
End Sub


Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = fNumeroKeyPress(KeyAscii)
   
   If KeyAscii = 13 Then
      FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmFinal, columnAntVehEspAsignKmFinal) = Text7.Text
      Text7.Visible = False
      FlexVehEspAsign.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text7_LostFocus()
   If FlexVehEspAsign.Col <> columnAntVehEspAsignKmFinal Or FlexVehEspAsign.Row <> filaAntVehEspAsignKmFinal Then
      'En este caso 3 es la columna que seria editable
      If columnAntVehEspAsignKmFinal = 3 Then
         FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmFinal, columnAntVehEspAsignKmFinal) = Text7.Text
      End If
   End If
End Sub

Private Sub FlexVehEspAsign_Click()
   Dim mi As Integer
   
   If FlexVehEspAsign.MouseRow > 0 Then
      If Not mEsOTcerrada Then
          'En este caso 2 es la columna que seria editable
         If FlexVehEspAsign.Col = 2 And FlexVehEspAsign.Row <> 1 Then
            Text6.Text = FlexVehEspAsign.Text
            Text6.Width = FlexVehEspAsign.ColWidth(FlexVehEspAsign.Col)
            Text6.Left = FlexVehEspAsign.ColPos(FlexVehEspAsign.Col) + FlexVehEspAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text6.Top = FlexVehEspAsign.Top + FlexVehEspAsign.RowPos(FlexVehEspAsign.Row)
            Text6.Visible = True
            Text6.SetFocus
            FlexVehEspAsign.ScrollBars = flexScrollBarNone
         Else
            Text6.Visible = False
            If FlexVehEspAsign.Col <> 3 Then
               FlexVehEspAsign.ScrollBars = flexScrollBarVertical
            End If
         End If
         filaAntVehEspAsignKmInicio = FlexVehEspAsign.Row
         columnAntVehEspAsignKmInicio = FlexVehEspAsign.Col
          'En este caso 3 es la columna que seria editable
         If FlexVehEspAsign.Col = 3 And FlexVehEspAsign.Row <> 1 Then
            Text7.Text = FlexVehEspAsign.Text
            Text7.Width = FlexVehEspAsign.ColWidth(FlexVehEspAsign.Col)
            Text7.Left = FlexVehEspAsign.ColPos(FlexVehEspAsign.Col) + FlexVehEspAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
            Text7.Top = FlexVehEspAsign.Top + FlexVehEspAsign.RowPos(FlexVehEspAsign.Row)
            Text7.Visible = True
            Text7.SetFocus
            FlexVehEspAsign.ScrollBars = flexScrollBarNone
         Else
            Text7.Visible = False
            If FlexVehEspAsign.Col <> 2 Then
               FlexVehEspAsign.ScrollBars = flexScrollBarVertical
            End If
         End If
         filaAntVehEspAsignKmFinal = FlexVehEspAsign.Row
         columnAntVehEspAsignKmFinal = FlexVehEspAsign.Col
      End If
      
      If mRenglonVehEspAsign <> 0 Then
         FlexVehEspAsign.Row = mRenglonVehEspAsign
         For mi = 1 To FlexVehEspAsign.Cols - 1
            FlexVehEspAsign.Col = mi
            FlexVehEspAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonVehEspAsign = FlexVehEspAsign.MouseRow
      FlexVehEspAsign.Row = mRenglonVehEspAsign
      For mi = 1 To FlexVehEspAsign.Cols - 1
         FlexVehEspAsign.Col = mi
         FlexVehEspAsign.CellBackColor = &H80000003
      Next
      If mRenglonVehEspAsign > 1 Then
          mCodVeh = FlexVehEspAsign.TextMatrix(mRenglonVehEspAsign, 2)
      End If
   Else
      FlexVehEspAsign.Row = mRenglonVehEspAsign
      If FlexVehEspAsign.Row > 0 Then
         For mi = 1 To FlexVehEspAsign.Cols - 1
            FlexVehEspAsign.Col = mi
            FlexVehEspAsign.CellBackColor = vbWhite
         Next
      End If
      mRenglonVehEspAsign = 0
   End If
End Sub

Private Sub Form_Load()
   
   mEsOTcerrada = False
   'Inhabilito Boton 'Cerrar OT'
   Command2(0).Enabled = False
   'Inhabilito Botones 'Subrurbro'
   CommandSubRubro(0).Enabled = False
   CommandSubRubro(1).Enabled = False
   
   If mObj.esSupervisorElectrico(Trim(Right(MDI.mUser, 20))) Then
      Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CONVERT( CONCAT(LPAD(O.IdOT,10,'0'),' - ',Date_Format(O.Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
         " FROM MantElect.OT_H O " & _
         " Inner Join " & _
         " OT_Partes OP ON O.IdOT = OP.IdOT " & _
         " Inner Join " & _
         " Registros R ON OP.Parte = R.Parte " & _
         " Where SectorAire = 0 " & _
         " ORDER BY O.IdOT DESC; ")
   Else
      Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CONVERT( CONCAT(LPAD(O.IdOT,10,'0'),' - ',Date_Format(O.Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
         " FROM MantElect.OT_H O " & _
         " Inner Join " & _
         " OT_Partes OP ON O.IdOT = OP.IdOT " & _
         " Inner Join " & _
         " Registros R ON OP.Parte = R.Parte " & _
         " Where SectorAire = 1 " & _
         " ORDER BY O.IdOT DESC; ")
   End If
   
   
'   Set mRec = mObj.oEjecutarSelect("SELECT CONVERT( CONCAT(LPAD(IdOT,10,'0'),' - ',Date_Format(Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
'                           " FROM MantElect.OT_H O " & _
'                           " ORDER BY IdOT DESC; ")





'SELECT CONVERT( CONCAT(LPAD(O.IdOT,10,'0'),' - ',Date_Format(O.Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha
'FROM MantElect.OT_H O
'Inner Join
'OT_Partes OP ON O.IdOT = OP.IdOT
'Inner Join
'Registros R ON OP.Parte = R.Parte
'Where SectorAire = 1




   Do While Not mRec.EOF
      Combo3.AddItem mRec!OT_Fecha
      mRec.MoveNext
   Loop
   mRec.Close
   
   Me.Width = 17090
   Me.Height = 9920
   sAlinearForm Me
   
   Frame1(0).Visible = True
   Frame1(1).Visible = False
   Frame1(2).Visible = False
   Frame1(3).Visible = False
   Frame1(4).Visible = False
   
   initPartes
   initManoObra
   initVehiculos
   initVehiculosEspecial
   initRubros_SubRubros
   initMateriales
   initAbastecimiento
   
End Sub

Private Sub initPartes()
   mRenglonPartes = 0

   With FlexPartes
      .ColWidth(0) = 200
      .ColWidth(1) = 500
      .ColWidth(2) = 2000
      .ColWidth(3) = 3000
      .ColWidth(4) = 8800
      .ColWidth(5) = 750
      .ColWidth(6) = 0
      
      .TextMatrix(0, 1) = "Parte"
      .TextMatrix(0, 2) = "Fecha Solicitud"
      .TextMatrix(0, 3) = "Lugar"
      .TextMatrix(0, 4) = "Descripcion de la Solicitud"
      .TextMatrix(0, 5) = "Prioridad"
      .TextMatrix(0, 6) = "Sector Aire"
      
      .RowHeight(1) = 0
   End With
End Sub

Private Sub initManoObra()
   Dim mi As Integer

   mRenglonMoAsign = 0
   With FlexMoAsig
      .ColWidth(0) = 200
      .ColWidth(1) = 6150
      .ColWidth(2) = 0
      
      .TextMatrix(0, 1) = "Técnico"
      .TextMatrix(0, 2) = "Codigo"
 
      .RowHeight(1) = 0
   End With
End Sub

Private Sub initVehiculos()
   mRenglonVehAsign = 0
   
   filaAntVehAsign = 0
   columnAntVehAsign = 0
   Text4.Visible = False

   filaAntVehAsignKmFinal = 0
   columnAntVehAsignKmFinal = 0
   Text5.Visible = False

   With FlexVehAsign
      .ColWidth(0) = 200
      .ColWidth(1) = 0
      .ColWidth(2) = 3150
      .ColWidth(3) = 1500
      .ColWidth(4) = 1500
      .ColWidth(5) = 0
      
      .TextMatrix(0, 2) = "Vehículo"
      .TextMatrix(0, 3) = "Km Inicial"
      .TextMatrix(0, 4) = "Km Final"
      .TextMatrix(0, 5) = "Codigo"
 
      .RowHeight(1) = 0
   End With

End Sub

Private Sub initVehiculosEspecial()
   mRenglonVehEspAsign = 0

   filaAntVehEspAsignKmInicio = 0
   columnAntVehEspAsignKmInicio = 0
   Text6.Visible = False

   filaAntVehEspAsignKmFinal = 0
   columnAntVehEspAsignKmFinal = 0
   Text7.Visible = False

   With FlexVehEspAsign
      .ColWidth(0) = 200
      .ColWidth(1) = 3150
      .ColWidth(2) = 1500
      .ColWidth(3) = 1500
      .ColWidth(4) = 0
      
      .TextMatrix(0, 1) = "Vehículo especial"
      .TextMatrix(0, 2) = "Km. Inicial"
      .TextMatrix(0, 3) = "Km Final"
      .TextMatrix(0, 4) = "Codigo"
      
      .RowHeight(1) = 0
   End With
End Sub

Private Sub initRubros_SubRubros()
   Dim mi As Integer
   
   For mi = FlexSubRubros.Rows To 3 Step -1
      FlexSubRubros.RemoveItem mi
   Next

   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Rubros Where FechaBaja IS NULL;")
   
   Do While Not mRec.EOF
       Combo1.AddItem "" & mRec!descripcion & Space(50) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close

   mRenglonSubRubroDispo = 0

   With FlexSubRubros
      .ColWidth(0) = 200
      .ColWidth(1) = 5000
      .ColWidth(2) = 10250
      .ColWidth(3) = 0
      .ColWidth(4) = 0
      
      .TextMatrix(0, 1) = "Rubro"
      .TextMatrix(0, 2) = "SubRubro"
      .TextMatrix(0, 3) = "CodRubro"
      .TextMatrix(0, 4) = "CodSubRubro"

      .RowHeight(1) = 0
   End With
   
   With FlexSubRubrosAsign
      .ColWidth(0) = 200
      .ColWidth(1) = 5000
      .ColWidth(2) = 10250
      .ColWidth(3) = 0
      .ColWidth(4) = 0

      .TextMatrix(0, 1) = "Rubro"
      .TextMatrix(0, 2) = "SubRubro"
      .TextMatrix(0, 3) = "CodRubro"
      .TextMatrix(0, 4) = "CodSubRubro"

      .RowHeight(1) = 0
   End With
End Sub

Private Sub initMateriales()
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   With FlexProduct
      .ColWidth(0) = 200
      .ColWidth(1) = 1250
      .ColWidth(2) = 9700
      .ColWidth(3) = 1250
      .ColWidth(4) = 1250
      .ColWidth(5) = 1900
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      
      .TextMatrix(0, 1) = "Cód.Sap"
      .TextMatrix(0, 2) = "Producto"
      .TextMatrix(0, 3) = "Cantidad"
      .TextMatrix(0, 4) = "Stock"
      .TextMatrix(0, 5) = "Unid.Medida"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      .TextMatrix(0, 8) = "CantidadBD"
      
      .RowHeight(1) = 0
   End With
End Sub

Private Sub initAbastecimiento()

   Combo4.Enabled = False
   Combo2.Enabled = False
   'TODO(Realizado): Debe traer solo las bodegas que puede administrar el usuario. Tabla Futura Tabla: Usuarios-Bodegas (o sera mejor hacerlo por Almacen)
   Set mRec = mObjInven.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
   
   Do While Not mRec.EOF
      Combo5.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close

   FlexProd.ColWidth(0) = 200
   FlexProd.ColWidth(1) = 9700
   FlexProd.ColWidth(2) = 2500
   FlexProd.ColWidth(3) = 1000
   FlexProd.ColWidth(4) = 1500
   FlexProd.ColWidth(5) = 1250
   FlexProd.ColWidth(6) = 0
   FlexProd.ColWidth(7) = 0
   
   FlexProd.TextMatrix(0, 1) = "Producto"
   FlexProd.TextMatrix(0, 2) = "Ubicación"
   FlexProd.TextMatrix(0, 3) = "Stock"
   FlexProd.TextMatrix(0, 4) = "Unid.Medida"
   FlexProd.TextMatrix(0, 5) = "Cód.Sap"
   FlexProd.TextMatrix(0, 6) = "Cód. Producto"
   FlexProd.TextMatrix(0, 7) = "Cód. Ubicacion"
   
   FlexProd.RowHeight(1) = 0

   FlexEgreso.ColWidth(0) = 200
   FlexEgreso.ColWidth(1) = 9700
   FlexEgreso.ColWidth(2) = 2500
   FlexEgreso.ColWidth(3) = 1000
   FlexEgreso.ColWidth(4) = 1500
   FlexEgreso.ColWidth(5) = 1250
   FlexEgreso.ColWidth(6) = 0
   FlexEgreso.ColWidth(7) = 0

   FlexEgreso.TextMatrix(0, 1) = "Producto"
   FlexEgreso.TextMatrix(0, 2) = "Ubicación"
   FlexEgreso.TextMatrix(0, 3) = "Cantidad"
   FlexEgreso.TextMatrix(0, 4) = "Unid.Medida"
   FlexEgreso.TextMatrix(0, 5) = "Cód.Sap"
   FlexEgreso.TextMatrix(0, 6) = "Cód. Producto"
   FlexEgreso.TextMatrix(0, 7) = "Cód. Ubicacion"

   FlexEgreso.RowHeight(1) = 0

   cboListIndex = Combo5.ListIndex

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


Private Function fValidaAsignaMateriales() As Boolean
'   Dim mRet As Boolean
'   Dim mMensajeError As String
'   Dim mJ As Integer
'
'   mRet = True
'
'   If mRenglonProdDispo = 0 Then
'      mRet = False
'      mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
'   End If
'
'   If mRet Then
'      If mRenglonProdDispo <> 0 And FlexProduct.TextMatrix(mRenglonProdDispo, 1) = "" Then
'         mRet = False
'         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
'      End If
'   End If
'
'   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
'   If mRet Then
'      For mJ = 2 To FlexEgreso.Rows - 1
'         If FlexEgreso.TextMatrix(mJ, 6) = FlexProduct.TextMatrix(mRenglonProdDispo, 6) And FlexEgreso.TextMatrix(mJ, 7) = FlexProduct.TextMatrix(mRenglonProdDispo, 7) Then
'            mRet = False
'            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
'            mJ = 999
'         End If
'      Next
'   End If
'
'   If Not mRet Then
'         MsgBox mMensajeError, vbCritical, "Atención"
'   End If
'   fValidaAsignaMateriales = mRet
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
   
   If KeyAscii = 13 Then
      FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
      Text2.Visible = False
      FlexProduct.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text2_LostFocus()
   If FlexProduct.Col <> columnAnt Or FlexProduct.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fDateKeyPress(Text3(Index), KeyAscii)
End Sub
  

Private Sub Combo5_Click()

   Dim mi As Integer

   Combo4.Enabled = True
   Combo2.Enabled = True

   If cboListIndex <> Combo5.ListIndex Then
      sLlenoUsuariosRet
      sLlenoUsuariosAut
      If (cboListIndex <> -1) Then
         'Si tengo algun registro en la grilla inferior(Egresos)
         If FlexEgreso.Rows > 2 Then
            If MsgBox("Si selecciona otra Bodega se perderán los consumos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Bodega") = vbYes Then
               Text9.Text = ""
               Text10.Text = ""

               'Elimino los registros de la grilla superior (productos)
               For mi = FlexProd.Rows To 3 Step -1
                  FlexProd.RemoveItem mi
               Next

               'Elimino los registros de la grilla inferior (consumos)
               For mi = FlexEgreso.Rows To 3 Step -1
                  FlexEgreso.RemoveItem mi
               Next

               mRenglonProducto = 0
               mRenglonEgreso = 0
            Else
               Combo5.ListIndex = cboListIndex
               sLlenoUsuariosRet
               sLlenoUsuariosAut
            End If
         Else
            Text9.Text = ""
            Text10.Text = ""

            'Elimino los registros de la grilla superior (productos)
            For mi = FlexProd.Rows To 3 Step -1
               FlexProd.RemoveItem mi
            Next

         End If

         cboListIndex = Combo5.ListIndex

      Else
         cboListIndex = Combo5.ListIndex
      End If

   End If







End Sub



Private Sub sLlenoUsuariosRet()
Dim mCodBodega As String
Dim mObjInven2 As New clInven
Dim mRec2 As New ADODB.Recordset

   mCodBodega = Trim(Left(Combo5.Text, 4))
   Combo4.Clear

   Set mRec2 = mObjInven.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
   " UsuariosRet_Bodegas UB " & _
   " Inner Join " & _
   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
   " ORDER BY P.Apellido;")


   Do While Not mRec2.EOF
      Combo4.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
      mRec2.MoveNext
   Loop
   mRec2.Close
   Set mObjInven2 = Nothing
   Set mRec2 = Nothing
End Sub

Private Sub sLlenoUsuariosAut()
Dim mCodBodega As String
Dim mObjInven2 As New clInven
Dim mRec2 As New ADODB.Recordset

   mCodBodega = Trim(Left(Combo5.Text, 4))
   Combo2.Clear

   Set mRec2 = mObjInven.oEjecutarSelect("SELECT CONCAT (P.Apellido,',', P.Nombres) AS Descripcion,P.CodUsuario AS CodUsuario FROM " & _
   " UsuariosAut_Bodegas UB " & _
   " Inner Join " & _
   " Personal P ON UB.CodUsuario = P.CodUsuario " & _
   " WHERE UB.CodBodega = '" & mCodBodega & "' AND P.Fecha_Baja IS NULL " & _
   " ORDER BY P.Apellido;")


   Do While Not mRec2.EOF
      Combo2.AddItem mRec2!descripcion & Space(60) & mRec2!CodUsuario
      mRec2.MoveNext
   Loop
   mRec2.Close
   Set mObjInven2 = Nothing
   Set mRec2 = Nothing
End Sub

Private Sub InhabilitarControlesOTCerrada()
      mEsOTcerrada = True
'      Text3(0).Text = mRec!FechaInicio
'      Text3(1).Text = mRec!FechaFin
      'Inhabilito Textboxs Fecha (Inicio y Fin)
      Text3(0).Enabled = False
      Text3(1).Enabled = False
      Command2(0).Enabled = False
      'Inhabilito Botones 'Subrurbro'
      CommandSubRubro(0).Enabled = False
      CommandSubRubro(1).Enabled = False
End Sub




