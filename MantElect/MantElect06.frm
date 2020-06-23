VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MantElect06 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nueva Orden de Trabajo"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16965
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9450
      TabIndex        =   51
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crear OT"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   50
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1(4)"
      Height          =   15
      Index           =   4
      Left            =   15
      TabIndex        =   5
      Top             =   480
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   3
      Left            =   15
      TabIndex        =   4
      Top             =   480
      Width           =   21100
      Begin VB.CommandButton CommandProd 
         Height          =   495
         Index           =   1
         Left            =   10797
         Picture         =   "MantElect06.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton CommandProd 
         Height          =   495
         Index           =   0
         Left            =   9807
         Picture         =   "MantElect06.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5400
         Width           =   495
      End
      Begin VB.Frame Frame11 
         Caption         =   "Egresos"
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
         Height          =   4455
         Left            =   120
         TabIndex        =   41
         Top             =   6120
         Width           =   20895
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   278
            Left            =   18840
            TabIndex        =   45
            Top             =   240
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid FlexEgreso 
            Height          =   3615
            Left            =   240
            TabIndex        =   42
            Top             =   600
            Width           =   20415
            _ExtentX        =   36010
            _ExtentY        =   6376
            _Version        =   327680
            Cols            =   11
         End
      End
      Begin VB.Frame Frame10 
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
         Height          =   4935
         Left            =   70
         TabIndex        =   34
         Top             =   120
         Width           =   20895
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   420
            Width           =   3735
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   2640
            TabIndex        =   37
            Top             =   960
            Width           =   10455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Buscar"
            Height          =   315
            Left            =   13320
            TabIndex        =   36
            Top             =   960
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid FlexProduct 
            Height          =   2775
            Left            =   240
            TabIndex        =   35
            Top             =   1440
            Width           =   20415
            _ExtentX        =   36010
            _ExtentY        =   4895
            _Version        =   327680
            Cols            =   8
         End
         Begin VB.Label Label2 
            Caption         =   "Retirar de:"
            Height          =   255
            Left            =   1440
            TabIndex        =   40
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Contiene texto:"
            Height          =   375
            Left            =   1080
            TabIndex        =   38
            Top             =   1020
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8025
      Index           =   2
      Left            =   15
      TabIndex        =   3
      Top             =   480
      Width           =   16920
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   1
         Left            =   9000
         Picture         =   "MantElect06.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   0
         Left            =   8010
         Picture         =   "MantElect06.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4440
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
         Left            =   120
         TabIndex        =   22
         Top             =   5040
         Width           =   16635
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubrosAsign 
            Height          =   2370
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   16215
            _ExtentX        =   28601
            _ExtentY        =   4180
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
         Height          =   4335
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   16635
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   420
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubros 
            Height          =   3255
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   16215
            _ExtentX        =   28601
            _ExtentY        =   5741
            _Version        =   327680
            Cols            =   5
         End
         Begin VB.Label Label1 
            Caption         =   "Rubro:"
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   520
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   8025
      HelpContextID   =   7300
      Index           =   1
      Left            =   15
      TabIndex        =   2
      Top             =   480
      Width           =   16920
      Begin VB.CommandButton CommandVehEsp 
         Height          =   495
         Index           =   1
         Left            =   8280
         Picture         =   "MantElect06.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4335
         Width           =   495
      End
      Begin VB.CommandButton CommandVehEsp 
         Height          =   495
         Index           =   0
         Left            =   8280
         Picture         =   "MantElect06.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3345
         Width           =   495
      End
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
         Height          =   1935
         Left            =   9120
         TabIndex        =   30
         Top             =   3000
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexVehEspAsign 
            Height          =   1335
            Left            =   240
            TabIndex        =   31
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   2355
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Vehículos Especiales"
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
         Height          =   1935
         Left            =   645
         TabIndex        =   28
         Top             =   3000
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexVehEspDispo 
            Height          =   1335
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   2355
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.CommandButton CommandVeh 
         Height          =   495
         Index           =   1
         Left            =   8280
         Picture         =   "MantElect06.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton CommandVeh 
         Height          =   495
         Index           =   0
         Left            =   8280
         Picture         =   "MantElect06.frx":1546
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   885
         Width           =   495
      End
      Begin VB.CommandButton CommandMO 
         Height          =   495
         Index           =   1
         Left            =   8280
         Picture         =   "MantElect06.frx":1850
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6810
         Width           =   495
      End
      Begin VB.CommandButton CommandMO 
         Height          =   495
         Index           =   0
         Left            =   8280
         Picture         =   "MantElect06.frx":1B5A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5820
         Width           =   495
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
         Height          =   2755
         Left            =   9121
         TabIndex        =   14
         Top             =   5160
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexMoAsig 
            Height          =   2055
            Left            =   240
            TabIndex        =   15
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3625
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mano de Obra"
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
         Height          =   2775
         Left            =   643
         TabIndex        =   12
         Top             =   5160
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexMoDispo 
            Height          =   2055
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3625
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
         Height          =   2775
         Left            =   9121
         TabIndex        =   10
         Top             =   120
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexVehAsign 
            Height          =   2175
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3836
            _Version        =   327680
            Cols            =   3
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vehículos"
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
         Height          =   2775
         Left            =   643
         TabIndex        =   8
         Top             =   120
         Width           =   7335
         Begin MSFlexGridLib.MSFlexGrid FlexVehDispo 
            Height          =   2175
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   6900
            _ExtentX        =   12171
            _ExtentY        =   3836
            _Version        =   327680
            Cols            =   3
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   8025
      Index           =   0
      Left            =   15
      TabIndex        =   1
      Top             =   480
      Width           =   16920
      Begin VB.Frame Frame13 
         Caption         =   "Partes asignados"
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
         Height          =   3855
         Left            =   240
         TabIndex        =   48
         Top             =   4080
         Width           =   16500
         Begin MSFlexGridLib.MSFlexGrid FlexPartAsignados 
            Height          =   3375
            Left            =   360
            TabIndex        =   49
            Top             =   360
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   5953
            _Version        =   327680
            Cols            =   7
         End
      End
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
         Height          =   3495
         Left            =   240
         TabIndex        =   46
         Top             =   120
         Width           =   16500
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   300
            Width           =   2415
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   300
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid FlexPartes 
            Height          =   2535
            Left            =   360
            TabIndex        =   47
            Top             =   840
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   4471
            _Version        =   327680
            Cols            =   7
         End
         Begin VB.Label Label5 
            Caption         =   "Detalle:"
            Height          =   255
            Left            =   6000
            TabIndex        =   55
            Top             =   405
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Origen:"
            Height          =   255
            Left            =   720
            TabIndex        =   53
            Top             =   405
            Width           =   975
         End
      End
      Begin VB.CommandButton CommandPartes 
         Height          =   375
         Index           =   1
         Left            =   8707
         Picture         =   "MantElect06.frx":1E64
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3720
         Width           =   330
      End
      Begin VB.CommandButton CommandPartes 
         Height          =   375
         Index           =   0
         Left            =   7717
         Picture         =   "MantElect06.frx":216E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3720
         Width           =   330
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   15
      TabIndex        =   0
      Top             =   120
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
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
End
Attribute VB_Name = "MantElect06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mObjInven As New clInven
Dim mRec As New ADODB.Recordset
Dim mRenglonPartes As Integer
Dim mRenglonPartAsignados As Integer
Dim mRenglonVehDispo As Integer
Dim mRenglonVehAsign As Integer
Dim mRenglonVehEspDispo As Integer
Dim mRenglonVehEspAsign As Integer
Dim mRenglonMoDispo As Integer
Dim mRenglonMoAsign As Integer
Dim mRenglonSubRubroDispo As Integer
Dim mRenglonSubRubroAsign As Integer
Dim mRenglonProdDispo As Integer
Dim mRenglonProdAsign As Integer

Dim XLS As EXCEL.Application


Dim filaAnt As Integer
Dim columnAnt As Integer

'TODO: Ver si es necesario utilizar las siguientes variables:
Dim mCodParte As Integer
Dim mCodMO As String
Dim mCodSubrubro As String
Dim mCodVeh As String
Dim mCodVehEsp As String
Dim mCodProducto As String

Dim cboOrigenListIndex As Integer
Dim cboDetalleListIndex As Integer

Dim mLinea As Integer


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
         " AND R.Codigo ='" & Right(Combo1.Text, 8) & "'  ORDER BY RubroDesc, SubRubroDesc;")
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

Private Sub Combo2_Click()

   Dim mi As Integer
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   Text1.Text = ""
   mRenglonProdDispo = 0
End Sub


Private Sub Combo3_Click()
   If cboOrigenListIndex <> Combo3.ListIndex And FlexPartAsignados.Rows > 2 Then
         If MsgBox("Si selecciona otro Origen se perderán los partes cargados hasta el momento en la grilla inferior. ¿ Desea continuar ? ", vbYesNo, "Origen") = vbYes Then
            eliminoGrillaPartes
            eliminoGrillaPartesAsignados
            sLlenoCboDetalle
            cboDetalleListIndex = -99
         Else
            Combo3.ListIndex = cboOrigenListIndex
         End If
         cboOrigenListIndex = Combo3.ListIndex
   Else
      If Combo3.ListIndex <> cboOrigenListIndex Then
         Combo4.Enabled = True
         eliminoGrillaPartes
         eliminoGrillaPartesAsignados
         sLlenoCboDetalle
         cboDetalleListIndex = -99
      End If
      cboOrigenListIndex = Combo3.ListIndex
   End If
End Sub
Private Sub eliminoGrillaPartes()
   Dim mi As Integer
   'Elimino los registros grilla superior
   For mi = FlexPartes.Rows To 3 Step -1
      FlexPartes.RemoveItem mi
   Next
   mRenglonPartes = 0
End Sub

Private Sub eliminoGrillaPartesAsignados()
   Dim mi As Integer
   'Elimino los registros grilla inferior
   For mi = FlexPartAsignados.Rows To 3 Step -1
     FlexPartAsignados.RemoveItem mi
   Next
   mRenglonPartAsignados = 0
End Sub

Private Sub Combo4_Click()
   Dim mi As Integer
   Dim mNroComunicado As String
   Dim mTramo As String
   Dim mRamal As String
   Dim Origen As String
   
   If cboDetalleListIndex <> Combo4.ListIndex And FlexPartAsignados.Rows > 2 Then
         If MsgBox("Si selecciona otra opción se perderán los partes cargados hasta el momento en la grilla inferior. ¿ Desea continuar ? ", vbYesNo, "Detalle") = vbYes Then
            eliminoGrillaPartes
            eliminoGrillaPartesAsignados
            Origen = Trim(Right(Combo3.Text, 4))
            Select Case Origen
               Case "OPE"
                  mTramo = Trim(Left(Combo4.Text, 2))
                  cargarGrillaConPartesOperaciones mTramo, "-1"
               Case "REL"
                  mRamal = Trim(Left(Combo4.Text, 50))
                  cargarGrillaConPartesDeRelevamientos mRamal, "-1"
               Case "COM"
                  mNroComunicado = Trim(Combo4.Text)
                  cargarGrillaConPartesDeComunicado mNroComunicado, "-1"
            End Select
         Else
            Combo4.ListIndex = cboDetalleListIndex
         End If
         cboDetalleListIndex = Combo4.ListIndex
   Else
      If Combo4.ListIndex <> cboDetalleListIndex Then
         eliminoGrillaPartes
         eliminoGrillaPartesAsignados
         Origen = Trim(Right(Combo3.Text, 4))
         Select Case Origen
            Case "OPE"
               mTramo = Trim(Left(Combo4.Text, 2))
               cargarGrillaConPartesOperaciones mTramo, "-1"
            Case "REL"
               mRamal = Trim(Left(Combo4.Text, 50))
               cargarGrillaConPartesDeRelevamientos mRamal, "-1"
            Case "COM"
               mNroComunicado = Trim(Combo4.Text)
               cargarGrillaConPartesDeComunicado mNroComunicado, "-1"
         End Select
      End If
      cboDetalleListIndex = Combo4.ListIndex
   End If
End Sub

Private Sub Command1_Click()
Dim mi As Integer
   Dim mj As Integer
   
   mRenglonProdDispo = 0
   
   'Elimino los registros (de la consulta anterior) de la grilla superior
   For mi = FlexProduct.Rows To 3 Step -1
      FlexProduct.RemoveItem mi
   Next
   'TODO: Dado que estamos consumiendo, seria ideal que el Store siguiente solo muestre los productos con stock > 0 en esa ubicacion.
   Set mRec = mObjInven.getStockXUbicacionConFiltroProducto(Right(Combo2.Text, 4), Text1.Text)
   'Set mRec = mObjInven.getStockXUbicacionConFiltroProducto("0001", Text1.Text)
   
   
   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         
         FlexProduct.AddItem ""
         FlexProduct.TextMatrix(mi, 1) = mRec!Producto
         FlexProduct.TextMatrix(mi, 2) = mRec!Ubicacion
         FlexProduct.TextMatrix(mi, 3) = mRec!Stock
         FlexProduct.TextMatrix(mi, 4) = mRec!UnidadMedida
         FlexProduct.TextMatrix(mi, 5) = mRec!CodigoSap
         FlexProduct.TextMatrix(mi, 6) = mRec!CodProducto
         FlexProduct.TextMatrix(mi, 7) = mRec!CodUbicacion
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
     
'   'Si presiono el boton Buscar y algun "producto/ubicacion" de la grilla de arriba, esta en la grilla inferior
'   'entonces que me actualice en la grilla superior el stock de ese o esos "producto/ubicacion" descontando el consumo de la grilla inferior
'   For mI = 2 To FlexProduct.Rows - 1
'      For mJ = 2 To FlexEgreso.Rows - 1
'         If FlexProduct.TextMatrix(mI, 6) = FlexEgreso.TextMatrix(mJ, 6) And FlexProduct.TextMatrix(mI, 7) = FlexEgreso.TextMatrix(mJ, 7) Then
'            FlexProduct.TextMatrix(mI, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mI, 3)), ".", ",")) - CDbl(Replace(Trim(FlexEgreso.TextMatrix(mJ, 3)), ".", ","))
'            mJ = 999
'         End If
'      Next
'   Next

End Sub

Private Sub Command2_Click(Index As Integer)
   If Index = 0 Then
      Dim OTgenerada As Integer
      Dim vPartes_OT() As Double
      Dim vVehiculos_OT() As String
      Dim vVehiculosEsp_OT() As String
      Dim vMO_Tecnicos_OT() As String
      Dim vSubrubros_OT() As String
      Dim mi As Integer
      Dim fecOt As Date
      
      fecOt = Now()

      'TODO: Validar los datos de la orden de trabajo a generar
      If fValidaOT Then
     
         preparaArrayPartes vPartes_OT()
         preparaArrayVehiculos vVehiculos_OT()
         preparaArrayVehiculosEsp vVehiculosEsp_OT()
         preparaArrayMO_Tecnicos vMO_Tecnicos_OT()
         preparaArraySubrubros vSubrubros_OT()
         
         OTgenerada = mObj.xinsOT(Trim(Right(MDI.mUser, 15)), vPartes_OT(), vVehiculos_OT(), vVehiculosEsp_OT(), vMO_Tecnicos_OT(), vSubrubros_OT(), fecOt)
         
         If OTgenerada <> 0 Then
            MsgBox "Se ha generado la Orden de Trabajo: " & OTgenerada, vbInformation, "Nueva Orden de Trabajo."
            imprimirExcelOT OTgenerada, fecOt, Trim(Left(MDI.mUser, 40))
            
            sLlenoCboOrigen
            cboOrigenListIndex = -99
            cboDetalleListIndex = -99
            InicializoCboDetalle
   
            initPartes False
            initManoObra False
            initVehiculos False
            initVehiculosEspecial False
            initRubros_SubRubros False
            
            Me.TabStrip1.Tabs(1).Selected = True
            
         End If
      End If
   Else
      'TODO: Ver el evento unload
      Unload Me
   End If
End Sub

Private Function fValidaOT() As Boolean


   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mCodTipoVale As String
   Dim mRec1 As New ADODB.Recordset
   
   mRet = True
      
   If FlexPartAsignados.Rows <= 2 Then
      mRet = False
      Me.TabStrip1.Tabs(1).Selected = True
      mMensajeError = "Al menos se debe seleccionar un Parte"
   End If
   
   If mRet Then
      If FlexMoAsig.Rows <= 2 Then
         mRet = False
         Me.TabStrip1.Tabs(2).Selected = True
         mMensajeError = "Al menos se debe seleccionar un técnico"
      End If
   End If
   
   If mRet Then
      If FlexSubRubrosAsign.Rows <= 2 Then
         mRet = False
         Me.TabStrip1.Tabs(3).Selected = True
         mMensajeError = "Al menos se debe seleccionar un Subrubro"
      End If
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   
   fValidaOT = mRet
End Function

Private Sub imprimirExcelOT(ByVal NroOT As Integer, FechaOT As Date, Supervisor As String)
         sMsgEspere Me, "Generando Formulario para OT: " & NroOT, True
         
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
   
'   Set mRec = mObj.oEjecutarSelect(" SELECT  CodProducto,CodigoSAP, P.Descripcion AS Producto, CodBodega, B.Descripcion AS Bodega, SUM(Stock) AS Stock, Med.Descripcion AS UnidadMedida " & _
'   "FROM  " & _
'   " Movimientos2 M " & _
'   "  INNER JOIN " & _
'   " Producto P ON M.CodProducto = P.Codigo " & _
'   "  INNER JOIN " & _
'   " Ubicaciones U ON  M.CodUbicacion = U.Codigo AND U.CodBodega = '" & Left(Combo1.Text, 4) & "' " & _
'   "  INNER JOIN " & _
'   " Bodegas B ON B.Codigo = U.CodBodega  " & _
'   "  INNER JOIN " & _
'   " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
'   " WHERE Fecha = (SELECT MAX(Fecha) " & _
'   "                 From Movimientos2 " & _
'   "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'   " GROUP BY   CodProducto, P.Descripcion,CodBodega, B.Descripcion,Med.Descripcion " & _
'   " ORDER BY   P.Descripcion ;")
'
'   Do While Not mRec.EOF
'      With XLS
'
'      .Cells(mI, 1).Formula = NVL(mRec!CodProducto, "")
'      .Cells(mI, 2).Formula = NVL(mRec!CodigoSap, "")
'      .Cells(mI, 3).Formula = NVL(mRec!Producto, "")
'      .Cells(mI, 4).Formula = NVL(mRec!CodBodega, "")
'      .Cells(mI, 5).Formula = NVL(mRec!Bodega, "")
'      .Cells(mI, 6).Formula = NVL(mRec!Stock, "")
'      .Cells(mI, 7).Formula = NVL(mRec!UnidadMedida, "")
'
'      End With
'      mRec.MoveNext
'      mI = mI + 1
'   Loop
'   mRec.Close
End Sub

'Private Sub sCabecera1(NroOT As Integer, FechaOT As Date, Supervisor As String)
'   Dim mi As Integer
'   Dim primerColumna As Boolean
'
'   mi = 10
'   With XLS
'      .WorkBooks.Add
'      .Worksheets(1).Select
'      .Worksheets(1).Name = "Orden de Trabajo"
'      .Columns("A:A").ColumnWidth = 1.14 '
'      .Columns("B:B").ColumnWidth = 6.86 '
'      .Columns("C:C").ColumnWidth = 24.29 '
'      .Columns("J:J").ColumnWidth = 1.14 '
'
'      .Range("B1:J500").Select
'      .Selection.Font.Size = 7
'      .Selection.Font.Bold = True
'      .Selection.RowHeight = 10.5
'
''---------------------------------ENCABEZADO HOJA-------------------------------------------------------
'      .Cells(1, 2).Formula = "AUTOPISTAS DEL SOL S.A."
'      .Cells(2, 4).Formula = "PLANILLA DE ORDEN DE TRABAJO"
'
'      .Cells(4, 2).Formula = "Fecha: " & FechaOT
'      .Cells(5, 2).Formula = "Tipo Tarea"
'      .Cells(6, 2).Formula = "Supervisor: " & Supervisor
'
'      .Cells(4, 8).Formula = "Nº OT"
'      .Cells(5, 8).Formula = "Hora Inicio"
'      .Cells(6, 8).Formula = "Hora Fin"
'      .Cells(4, 9).Formula = NroOT
'
'      .Range("H4:H6").Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("H4:I6").Select
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''---------------------------------ENCABEZADO TECNICOS---------------------------------------------------
'       .Cells(9, 4).Formula = "TECNICOS QUE INTERVIENEN"
'
'      .Range("B9:H9").Select
'      .Selection.Interior.ColorIndex = 15
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
''---------------------------------DETALLE TECNICOS----------------------------------------------------
'      primerColumna = 1
'      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion " & _
'                                          "FROM OT_MO_Tecnicos O " & _
'                                              "Inner Join " & _
'                                          "MO_Tecnicos M ON O.CodMO_Tecnico = M.Codigo " & _
'                                      "WHERE IdOT = '" & NroOT & "';")
'
'      Do While Not mRec.EOF
'         .Range("B" & mi & ":H" & mi).Select
'         With .Selection.Borders(xlBottom)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlTop)
'            .LineStyle = xlContinuous
'            .Weight = xlThin
'            .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeRight)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("E" & mi & ":E" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         If primerColumna Then
'            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
'            primerColumna = False
'         Else
'            .Cells(mi, 5).Formula = NVL(mRec!descripcion, "")
'            primerColumna = True
'            mi = mi + 1
'         End If
'
'         mRec.MoveNext
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
'
'
'
''---------------------------------ENCABEZADO VEHICULOS------------------------------------------------
'
'      mi = mi + 2
'
'      .Range("B" & mi & ":H" & (mi + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mi & ":H" & (mi + 1)).Select
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("E" & (mi + 1) & ":H" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
'       With .Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'       End With
'
'      .Cells(mi, 4).Formula = "VEHICULOS QUE INTERVIENEN"
'      mi = mi + 1
'      .Cells(mi, 2).Formula = "Vehículo"
'      .Cells(mi, 5).Formula = "Km Inicial"
'      .Cells(mi, 7).Formula = "Km Final"
'
''-----------------------------------------------------------------------------------------------------
'
'
''---------------------------------DETALLE VEHICULOS------------------------------------------------
'      mi = mi + 1
'      Set mRec = mObj.oEjecutarSelect("SELECT Codigo,Descripcion FROM " & _
'                                          "OT_Vehiculos O " & _
'                                              "Inner Join " & _
'                                          "Vehiculos V ON O.CodVehiculo = Codigo " & _
'                                      "WHERE IdOT = '" & NroOT & "'; ")
'
'      Do While Not mRec.EOF
'         .Range("B" & mi & ":H" & mi).Select
'         With .Selection.Borders(xlBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlTop)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeRight)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("E" & mi & ":E" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("G" & mi & ":G" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
'         End With
'         mRec.MoveNext
'         mi = mi + 1
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
'
'
'
''---------------------------------ENCABEZADO TAREAS--------------------------------------------------
'      mi = mi + 2
'
'      .Cells(mi, 5).Formula = "TAREAS"
'      .Cells(mi + 1, 2).Formula = "Parte"
'      .Cells(mi + 1, 3).Formula = "Lugar"
'      .Cells(mi + 1, 4).Formula = "Descripcion"
'      .Cells(mi + 1, 9).Formula = "¿Finalizado?"
'
'      .Range("B" & mi & ":I" & (mi + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mi & ":I" & (mi + 1)).Select
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("C" & (mi + 1) & ":C" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("D" & (mi + 1) & ":D" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("I" & (mi + 1) & ":I" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
''---------------------------------DETALLE TAREAS------------------------------------------------------
'      mi = mi + 2
'      Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,R.CodEdificio, R.Descripcion FROM " & _
'                                          "OT_Partes O " & _
'                                              "Inner Join " & _
'                                          "Registros R ON O.Parte = R.Parte " & _
'                                          "WHERE IDOT = '" & NroOT & "' " & _
'                                          "ORDER BY R.parte; ")
'
'      Do While Not mRec.EOF
'         .Range("B" & mi & ":I" & mi).Select
'         With .Selection.Borders(xlBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlTop)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeRight)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("C" & mi & ":C" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("D" & mi & ":D" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("I" & mi & ":I" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mi, 2).Formula = NVL(mRec!Parte, "")
'            .Cells(mi, 3).Formula = NVL(mRec!CodEdificio, "")
'            .Cells(mi, 4).Formula = NVL(mRec!descripcion, "")
'         End With
'         mRec.MoveNext
'         mi = mi + 1
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
''---------------------------------ENCABEZADO SUBRUBROS------------------------------------------------
'      mi = mi + 2
'
'      .Cells(mi, 5).Formula = "FALLAS"
'      .Cells(mi + 1, 2).Formula = "Subrubro"
'      .Cells(mi + 1, 6).Formula = "Subrubro"
'
'      .Range("B" & mi & ":I" & (mi + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mi & ":I" & (mi + 1)).Select
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("E" & mi + 1 & ":E" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlMedium
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("I" & mi + 1 & ":I" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
''---------------------------------DETALLE SUBRUBROS-----------------------------------------------
'      mi = mi + 2
'      Set mRec = mObj.oEjecutarSelect("SELECT S.Codigo,S.Descripcion FROM " & _
'                                       "SubRubros S " & _
'                                          "Inner Join " & _
'                                       "OT_Subrubros O ON O.CodSubrubro = S.Codigo " & _
'                                       "WHERE IDOT = '" & NroOT & "' " & _
'                                       "ORDER BY S.Descripcion; ")
'
'      primerColumna = True
'      Do While Not mRec.EOF
'
'         If primerColumna Then
'            .Range("B" & mi & ":I" & mi).Select
'            With .Selection.Borders(xlBottom)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            With .Selection.Borders(xlTop)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            With .Selection.Borders(xlEdgeRight)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Range("E" & mi & ":E" & mi).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Range("F" & mi & ":F" & mi).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlMedium
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Range("I" & mi & ":I" & mi).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(mi, 2).Formula = NVL(mRec!descripcion, "")
'            primerColumna = False
'         Else
'            .Cells(mi, 6).Formula = NVL(mRec!descripcion, "")
'            primerColumna = True
'            mi = mi + 1
'         End If
'         mRec.MoveNext
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
'
'
'
'
''---------------------------------ENCABEZADO Materiales-----------------------------------------------
'      mi = mi + 2
'      .Cells(mi, 4).Formula = "                 MATERIALES"
'      .Cells(mi + 1, 2).Formula = "Cód.Sap"
'      .Cells(mi + 1, 3).Formula = "Descripción"
'      '.Cells(mi + 1, 6).Formula = "Consumido"
'      .Cells(mi + 1, 7).Formula = "Consumido"
'      '.Cells(mi + 1, 7).Formula = "U.Medida"
'      .Cells(mi + 1, 8).Formula = "U.Medida"
'
'      .Range("B" & mi & ":H" & (mi + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mi & ":H" & (mi + 1)).Select
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("C" & (mi + 1) & ":C" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      '.Range("F" & (mi + 1) & ":F" & (mi + 1)).Select
'      .Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      '.Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
'      .Range("H" & (mi + 1) & ":H" & (mi + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
'
'
''---------------------------------DETALLE Materiales--------------------------------------------------
'      mi = mi + 2
''      Set mRec = mObj.oEjecutarSelect("SELECT  idMov,  M.Fecha,  P.CodigoSap,  P.Descripcion,  Stock,  UM.Descripcion AS UnidadMedidad FROM " & _
''                                          "Inventario.Movimientos2 M " & _
''                                              "Inner Join " & _
''                                          "Inventario.Producto P ON M.CodProducto = P.Codigo " & _
''                                              "Inner Join " & _
''                                          "Inventario.UnidadMedida UM ON P.CodUnidadMedida = UM.Codigo " & _
''                                              "Inner Join " & _
''                                          "Inventario.Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
''                                             "Inner Join " & _
''                                          "Vehiculos V ON U.Codigo = V.CodUbicacion " & _
''                                             "Inner Join " & _
''                                          "OT_Vehiculos OV ON OV.CodVehiculo = V.Codigo " & _
''                                          "WHERE M.Fecha = (SELECT MAX(Fecha) " & _
''                                          "                From Inventario.Movimientos2 " & _
''                                          "                WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
''                                          "and OV.IDOT = '" & NroOT & "' and stock > 0; ")
'
'
'
'      Set mRec = mObj.oEjecutarSelect(" SELECT P.CodigoSap,P.descripcion,UM.Descripcion AS UnidadMedidad " & _
'                                       " From " & _
'                                       " Matriz_Reposicion_Ubicaciones MR " & _
'                                       " Inner Join " & _
'                                       " Inventario.Ubicaciones U ON U.Codigo = MR.CodUbicacion " & _
'                                       " Inner Join " & _
'                                       " Vehiculos V ON V.CodUbicacion = U.Codigo " & _
'                                       " Inner Join " & _
'                                       " OT_Vehiculos OTV ON OTV.CodVehiculo = V.Codigo " & _
'                                       " Inner Join " & _
'                                       " Inventario.Producto P ON P.Codigo = MR.CodProducto " & _
'                                       " Left Join " & _
'                                       " Inventario.Movimientos2 M ON M.CodProducto = MR.CodProducto AND M.CodUbicacion = MR.CodUbicacion " & _
'                                       " Inner Join " & _
'                                       " Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
'                                       " Where OTV.IdOT = '" & NroOT & "' " & _
'                                       " AND M.Fecha = (SELECT MAX(Fecha) " & _
'                                       " From Inventario.Movimientos2 " & _
'                                       " WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'                                       " AND MR.FechaHasta = '0000-00-00 00:00:00'; ")
'
'
'
'
'
'
'      Do While Not mRec.EOF
'         .Range("B" & mi & ":H" & mi).Select
'         With .Selection.Borders(xlBottom)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlTop)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With .Selection.Borders(xlEdgeRight)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("C" & mi & ":C" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         '.Range("F" & mi & ":F" & mi).Select
'         .Range("G" & mi & ":G" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         '.Range("G" & mi & ":G" & mi).Select
'         .Range("H" & mi & ":H" & mi).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mi, 2).Formula = NVL(mRec!CodigoSap, "")
'            .Cells(mi, 3).Formula = NVL(mRec!descripcion, "")
'            .Cells(mi, 8).Formula = NVL(mRec!UnidadMedidad, "")
'         End With
'         mRec.MoveNext
'         mi = mi + 1
'      Loop
'      mRec.Close
'
''-----------------------------------------------------------------------------------------------------
'
'
''----------------------------------------------OBSERVACIONES------------------------------------------
'      mi = mi + 2
'      .Cells(mi, 2).Formula = "OBSERVACIONES"
'      mi = mi + 1
'      .Range("B" & mi & ":I" & (mi + 4)).Select
'    '  .Selection.RowHeight = 16.5
'      With .Selection.Borders(xlBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      With .Selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
''-----------------------------------------------------------------------------------------------------
'
'
''----------------------------------------------FIRMAS-------------------------------------------------
'      mi = mi + 8
'      .Cells(mi, 3).Formula = "              SUPERVISOR"
'      .Cells(mi, 6).Formula = "     ENCARGADO BODEGA"
'
'      .Range("C" & mi & ":C" & mi).Select
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("F" & mi & ":G" & mi).Select
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'End With
''-----------------------------------------------------------------------------------------------------
'
'
'
'
'
'
'
'
'
'
'
'
'
''  Configuracion de margenes.
'   With ActiveSheet.PageSetup
'      .LeftMargin = Application.CentimetersToPoints(0)
'      .RightMargin = Application.CentimetersToPoints(0)
'      .TopMargin = Application.CentimetersToPoints(0)
'      .BottomMargin = Application.CentimetersToPoints(0)
'   End With
'
'End Sub




Private Sub sCabecera1(NroOT As Integer, FechaOT As Date, Supervisor As String)
   'Dim mi As Integer
   Dim mLineasXpagina As Integer
   Dim primerColumna As Boolean
   Dim mj As Integer
   mLinea = 1
   mLineasXpagina = 81
   'mi = 250
   
   With XLS
      .WorkBooks.Add
      .Worksheets(1).Select
      .Worksheets(1).Name = "Orden de Trabajo"
      .Columns("A:A").ColumnWidth = 1.14 '
      .Columns("B:B").ColumnWidth = 6.86 '
      .Columns("C:C").ColumnWidth = 24.29 '
      .Columns("D:D").ColumnWidth = 9.71 '
      .Columns("F:F").ColumnWidth = 10.29 '
      .Columns("G:G").ColumnWidth = 10.29 '
      .Columns("I:I").ColumnWidth = 9.86 '
      .Columns("J:J").ColumnWidth = 1.14 '

      .Range("B1:J500").Select
      .Selection.Font.Size = 7
      .Selection.Font.Bold = True
      .Selection.RowHeight = 10.5

'---------------------------------ENCABEZADO HOJA-------------------------------------------------------
      .Cells(mLinea, 2).Formula = "AUTOPISTAS DEL SOL S.A."
      .Cells(mLinea + 1, 4).Formula = "PLANILLA DE ORDEN DE TRABAJO"
      
      .Cells(mLinea + 3, 2).Formula = "Fecha: " & FechaOT
      .Cells(mLinea + 4, 2).Formula = "Tipo Tarea"
      .Cells(mLinea + 5, 2).Formula = "Supervisor: " & Supervisor
      
      .Cells(mLinea + 3, 8).Formula = "Nº OT"
      .Cells(mLinea + 4, 8).Formula = "Hora Inicio"
      .Cells(mLinea + 5, 8).Formula = "Hora Fin"
      .Cells(mLinea + 3, 9).Formula = NroOT

      .Range("H4:H6").Select
      .Selection.Interior.ColorIndex = 15

      .Range("H" & (mLinea + 3) & ":I" & (mLinea + 5)).Select
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
      
      mLinea = mLinea + 8
'---------------------------------ENCABEZADO TECNICOS---------------------------------------------------
       If mLinea Mod mLineasXpagina = 0 Then
         MsgBox "FIN: MOD = 0 Encabezado Tecnicos"
         'Repetir lo del else
       Else
          .Cells(mLinea, 4).Formula = "TECNICOS QUE INTERVIENEN"
         
         .Range("B" & mLinea & ":H" & mLinea).Select
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
      End If
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TECNICOS----------------------------------------------------
      primerColumna = 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion " & _
                                          "FROM OT_MO_Tecnicos O " & _
                                              "Inner Join " & _
                                          "MO_Tecnicos M ON O.CodMO_Tecnico = M.Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "';")
                                
      mLinea = mLinea + 1
      Do While Not mRec.EOF
         
         'if linea mod then
            'Imprimir encaabezado
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         
         .Range("B" & mLinea & ":H" & mLinea).Select
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
   
         .Range("E" & mLinea & ":E" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         If primerColumna Then
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mLinea, 5).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mLinea = mLinea + 1
         End If
   
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO VEHICULOS------------------------------------------------

      mLinea = mLinea + 2
         'if (mlinea mod = 0) or (mLinea+1 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
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

      .Range("E" & (mLinea + 1) & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("G" & (mLinea + 1) & ":G" & (mLinea + 1)).Select
       With .Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
       End With
      
      .Cells(mLinea, 4).Formula = "VEHICULOS QUE INTERVIENEN"
      mLinea = mLinea + 1
      .Cells(mLinea, 2).Formula = "Vehículo"
      .Cells(mLinea, 5).Formula = "Km Inicial"
      .Cells(mLinea, 7).Formula = "Km Final"

'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE VEHICULOS------------------------------------------------
      mLinea = mLinea + 1
      Set mRec = mObj.oEjecutarSelect("SELECT Codigo,Descripcion FROM " & _
                                          "OT_Vehiculos O " & _
                                              "Inner Join " & _
                                          "Vehiculos V ON O.CodVehiculo = Codigo " & _
                                      "WHERE IdOT = '" & NroOT & "'; ")
                                
      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         
         .Range("B" & mLinea & ":H" & mLinea).Select
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
      
         .Range("E" & mLinea & ":E" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("G" & mLinea & ":G" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
         
         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------



'---------------------------------ENCABEZADO TAREAS--------------------------------------------------
      
      
      
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      .Cells(mLinea, 5).Formula = "TAREAS"
      .Cells(mLinea + 1, 2).Formula = "Parte"
      .Cells(mLinea + 1, 3).Formula = "Lugar"
      .Cells(mLinea + 1, 4).Formula = "Descripcion"
      .Cells(mLinea + 1, 9).Formula = "¿Finalizado?"

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
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

      .Range("C" & (mLinea + 1) & ":C" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("D" & (mLinea + 1) & ":D" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("I" & (mLinea + 1) & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE TAREAS------------------------------------------------------
      mLinea = mLinea + 2
      Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,R.CodEdificio, R.Descripcion, Length(R.Descripcion) lenDesc " & _
                                          "FROM " & _
                                          "OT_Partes O " & _
                                              "Inner Join " & _
                                          "Registros R ON O.Parte = R.Parte " & _
                                          "WHERE IDOT = '" & NroOT & "' " & _
                                          "ORDER BY R.parte; ")
                                
      Do While Not mRec.EOF
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
         
         .Range("B" & mLinea & ":I" & mLinea).Select
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
   
         .Range("C" & mLinea & ":C" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         .Range("D" & mLinea & ":D" & mLinea).Select
         If mRec!lenDesc > 75 Then
            .Selection.Font.Size = 6
         Else
         .Selection.Font.Size = 7
         End If
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         .Range("I" & mLinea & ":I" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With

         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!Parte, "")
            .Cells(mLinea, 3).Formula = NVL(mRec!CodEdificio, "")
            .Cells(mLinea, 4).Formula = NVL(mRec!descripcion, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------
'---------------------------------ENCABEZADO SUBRUBROS------------------------------------------------
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      
      .Cells(mLinea, 5).Formula = "FALLAS"
      .Cells(mLinea + 1, 2).Formula = "Subrubro"
      .Cells(mLinea + 1, 6).Formula = "Subrubro"

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":I" & (mLinea + 1)).Select
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

      .Range("E" & mLinea + 1 & ":E" & (mLinea + 1)).Select
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

      .Range("I" & mLinea + 1 & ":I" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------
'---------------------------------DETALLE SUBRUBROS-----------------------------------------------
      mLinea = mLinea + 2
      Set mRec = mObj.oEjecutarSelect("SELECT S.Codigo,S.Descripcion FROM " & _
                                       "SubRubros S " & _
                                          "Inner Join " & _
                                       "OT_Subrubros O ON O.CodSubrubro = S.Codigo " & _
                                       "WHERE IDOT = '" & NroOT & "' " & _
                                       "ORDER BY S.Descripcion; ")
                                
      primerColumna = True
      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
   
         If primerColumna Then
            .Range("B" & mLinea & ":I" & mLinea).Select
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
         
            .Range("E" & mLinea & ":E" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
         
            .Range("F" & mLinea & ":F" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlMedium
              .ColorIndex = xlAutomatic
            End With
      
            .Range("I" & mLinea & ":I" & mLinea).Select
            With .Selection.Borders(xlEdgeLeft)
              .LineStyle = xlContinuous
              .Weight = xlThin
              .ColorIndex = xlAutomatic
            End With
            
            .Cells(mLinea, 2).Formula = NVL(mRec!descripcion, "")
            primerColumna = False
         Else
            .Cells(mLinea, 6).Formula = NVL(mRec!descripcion, "")
            primerColumna = True
            mLinea = mLinea + 1
         End If
         mRec.MoveNext
      Loop
      mRec.Close
'-----------------------------------------------------------------------------------------------------




'---------------------------------ENCABEZADO Materiales-----------------------------------------------
      
      'if (mlinea mod = 0) or (mLinea+1 mod = 0) or (mLinea+2 mod = 0) then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda ''mLinea = mLinea + 2
      'End If
      
      mLinea = mLinea + 2 'Borrar cuando descomente lo de arriba.
      
      
      .Cells(mLinea, 4).Formula = "                 MATERIALES"
      .Cells(mLinea + 1, 2).Formula = "Cód.Sap"
      .Cells(mLinea + 1, 3).Formula = "Descripción"
      '.Cells(mi + 1, 6).Formula = "Consumido"
      .Cells(mLinea + 1, 7).Formula = "Consumido"
      '.Cells(mi + 1, 7).Formula = "U.Medida"
      .Cells(mLinea + 1, 8).Formula = "U.Medida"
      
      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
      .Selection.Interior.ColorIndex = 15

      .Range("B" & mLinea & ":H" & (mLinea + 1)).Select
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
      
      .Range("C" & (mLinea + 1) & ":C" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      '.Range("F" & (mi + 1) & ":F" & (mi + 1)).Select
      .Range("G" & (mLinea + 1) & ":G" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
      
      '.Range("G" & (mi + 1) & ":G" & (mi + 1)).Select
      .Range("H" & (mLinea + 1) & ":H" & (mLinea + 1)).Select
      With .Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
'-----------------------------------------------------------------------------------------------------


'---------------------------------DETALLE Materiales--------------------------------------------------
      mLinea = mLinea + 2
'      Set mRec = mObj.oEjecutarSelect("SELECT  idMov,  M.Fecha,  P.CodigoSap,  P.Descripcion,  Stock,  UM.Descripcion AS UnidadMedidad FROM " & _
'                                          "Inventario.Movimientos2 M " & _
'                                              "Inner Join " & _
'                                          "Inventario.Producto P ON M.CodProducto = P.Codigo " & _
'                                              "Inner Join " & _
'                                          "Inventario.UnidadMedida UM ON P.CodUnidadMedida = UM.Codigo " & _
'                                              "Inner Join " & _
'                                          "Inventario.Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
'                                             "Inner Join " & _
'                                          "Vehiculos V ON U.Codigo = V.CodUbicacion " & _
'                                             "Inner Join " & _
'                                          "OT_Vehiculos OV ON OV.CodVehiculo = V.Codigo " & _
'                                          "WHERE M.Fecha = (SELECT MAX(Fecha) " & _
'                                          "                From Inventario.Movimientos2 " & _
'                                          "                WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'                                          "and OV.IDOT = '" & NroOT & "' and stock > 0; ")
                                          
                                          
                                          
      Set mRec = mObj.oEjecutarSelect(" SELECT P.CodigoSap,P.descripcion,UM.Descripcion AS UnidadMedidad " & _
                                       " From " & _
                                       " Matriz_Reposicion_Ubicaciones MR " & _
                                       " Inner Join " & _
                                       " Inventario.Ubicaciones U ON U.Codigo = MR.CodUbicacion " & _
                                       " Inner Join " & _
                                       " Vehiculos V ON V.CodUbicacion = U.Codigo " & _
                                       " Inner Join " & _
                                       " OT_Vehiculos OTV ON OTV.CodVehiculo = V.Codigo " & _
                                       " Inner Join " & _
                                       " Inventario.Producto P ON P.Codigo = MR.CodProducto " & _
                                       " Left Join " & _
                                       " Inventario.Movimientos2 M ON M.CodProducto = MR.CodProducto AND M.CodUbicacion = MR.CodUbicacion " & _
                                       " Inner Join " & _
                                       " Inventario.UnidadMedida UM ON UM.Codigo = P.CodUnidadMedida " & _
                                       " Where OTV.IdOT = '" & NroOT & "' " & _
                                       " AND M.Fecha = (SELECT MAX(Fecha) " & _
                                       " From Inventario.Movimientos2 " & _
                                       " WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
                                       " AND MR.FechaHasta = '0000-00-00 00:00:00'; ")
                                          
                                          
                                          
                                
  

      Do While Not mRec.EOF
      
         'if (mlinea mod = 0)  then
            'Imprimir encaabezado (deberia modificar el valor mLinea)
            'Incrementar mHoja
            'Poscionar mLinea donde corresponda
         'End If
      
         .Range("B" & mLinea & ":H" & mLinea).Select
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
      
         .Range("C" & mLinea & ":C" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         '.Range("F" & mi & ":F" & mi).Select
         .Range("G" & mLinea & ":G" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
   
         '.Range("G" & mi & ":G" & mi).Select
         .Range("H" & mLinea & ":H" & mLinea).Select
         With .Selection.Borders(xlEdgeLeft)
           .LineStyle = xlContinuous
           .Weight = xlThin
           .ColorIndex = xlAutomatic
         End With
      
         With XLS
            .Cells(mLinea, 2).Formula = NVL(mRec!CodigoSap, "")
            .Cells(mLinea, 3).Formula = NVL(mRec!descripcion, "")
            .Cells(mLinea, 8).Formula = NVL(mRec!UnidadMedidad, "")
         End With
         mRec.MoveNext
         mLinea = mLinea + 1
      Loop
      mRec.Close
   
'-----------------------------------------------------------------------------------------------------
 
 
'----------------------------------------------OBSERVACIONES------------------------------------------
      mLinea = mLinea + 2
      
'      For mj = mLinea To mLinea + 10
'         If mLinea Mod 81 = 0 Then
'            mEsCorte = True
'            mj = 9999
'         End If
'      Next
'
'      If mEsCorte Then
'         'imprimirEncabezado
'      Else
'         'mLinea = mj
'      End If
'
      
      
      .Cells(mLinea, 2).Formula = "OBSERVACIONES"
      mLinea = mLinea + 1
      .Range("B" & mLinea & ":I" & (mLinea + 4)).Select
    '  .Selection.RowHeight = 16.5
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
      mLinea = mLinea + 8
      .Cells(mLinea, 3).Formula = "              SUPERVISOR"
      .Cells(mLinea, 6).Formula = "     ENCARGADO BODEGA"
      
      .Range("C" & mLinea & ":C" & mLinea).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With

      .Range("F" & mLinea & ":G" & mLinea).Select
      With .Selection.Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
      End With
End With
'-----------------------------------------------------------------------------------------------------
 
 
''  Configuracion de margenes.
'   With ActiveSheet.PageSetup
'      .LeftMargin = Application.CentimetersToPoints(0)
'      .RightMargin = Application.CentimetersToPoints(0)
'      .TopMargin = Application.CentimetersToPoints(0)
'      .BottomMargin = Application.CentimetersToPoints(0)
'   End With
'
   
   
   '  Configuracion de margenes.
'   ActiveSheet.PageSetup.LeftMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.RightMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.TopMargin = Application.CentimetersToPoints(0)
'   ActiveSheet.PageSetup.BottomMargin = Application.CentimetersToPoints(0)
   
   
End Sub







Private Sub preparaArrayPartes(ByRef pvPartes_OT() As Double)
   
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartAsignados.Rows - 2
   If cantPartes > 0 Then
      ReDim pvPartes_OT(0 To cantPartes - 1) As Double
         
      For mj = 2 To FlexPartAsignados.Rows - 1
         pvPartes_OT(mj - 2) = FlexPartAsignados.TextMatrix(mj, 1)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvPartes_OT(0)
      pvPartes_OT(0) = 0
   End If
End Sub

Private Sub preparaArrayVehiculos(ByRef pvVehiculos_OT() As String)
   Dim mj As Integer
   Dim cantVehiculos As Integer

   cantVehiculos = FlexVehAsign.Rows - 2
   If cantVehiculos > 0 Then
      ReDim pvVehiculos_OT(0 To cantVehiculos - 1) As String
         
      For mj = 2 To FlexVehAsign.Rows - 1
         pvVehiculos_OT(mj - 2) = FlexVehAsign.TextMatrix(mj, 2)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvVehiculos_OT(0)
      pvVehiculos_OT(0) = "00"
   End If
End Sub

Private Sub preparaArrayVehiculosEsp(ByRef pvVehiculosEsp_OT() As String)
   Dim mj As Integer
   Dim cantVehiculosEsp As Integer

   cantVehiculosEsp = FlexVehEspAsign.Rows - 2
   If cantVehiculosEsp > 0 Then
      ReDim pvVehiculosEsp_OT(0 To cantVehiculosEsp - 1) As String
         
      For mj = 2 To FlexVehEspAsign.Rows - 1
         pvVehiculosEsp_OT(mj - 2) = FlexVehEspAsign.TextMatrix(mj, 2)
      Next
   Else
      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
      ReDim pvVehiculosEsp_OT(0)
      pvVehiculosEsp_OT(0) = "00"
   End If
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

Private Sub CommandMO_Click(Index As Integer)
   Dim sListaMOSeleccionados
   Dim mj As Integer
   sListaMOSeleccionados = ""
   
   If Index = 0 Then
      If mRenglonMoDispo > 0 Then
         If Trim(FlexMoDispo.TextMatrix(mRenglonMoDispo, 1)) <> "" Then
            
            FlexMoAsig.AddItem vbTab & FlexMoDispo.TextMatrix(mRenglonMoDispo, 1) & vbTab & FlexMoDispo.TextMatrix(mRenglonMoDispo, 2)

         End If
         
         If FlexMoDispo.Rows > 2 Then
            FlexMoDispo.RemoveItem mRenglonMoDispo
         
            mRenglonMoDispo = 0
         Else
            If Trim(FlexMoDispo.TextMatrix(mRenglonMoDispo, 1)) <> "" Then
               FlexMoDispo.TextMatrix(mRenglonMoDispo, 1) = ""
               FlexMoDispo.TextMatrix(mRenglonMoDispo, 2) = ""
         
               mRenglonMoDispo = 0
            End If
         End If
      End If
   Else
      
      If FlexMoAsig.Rows > 2 And mRenglonMoAsign > 1 Then
         
         FlexMoAsig.RemoveItem (mRenglonMoAsign)
         
         If FlexMoAsig.Rows > 2 Then
            For mj = 2 To FlexMoAsig.Rows - 1
               sListaMOSeleccionados = sListaMOSeleccionados & "'" & FlexMoAsig.TextMatrix(mj, 2) & "',"
            Next
            sListaMOSeleccionados = Left(sListaMOSeleccionados, Len(sListaMOSeleccionados) - 1)
        End If
            
         mRenglonMoDispo = 0


         FlexMoDispo.Clear
         'Elimino los registros  de la grilla superior
         For mj = FlexMoDispo.Rows To 3 Step -1
            FlexMoDispo.RemoveItem mj
         Next
         
         With FlexMoDispo
            .TextMatrix(0, 1) = "Técnico"
            .TextMatrix(0, 2) = "Codigo"
         
            .RowHeight(1) = 0
         End With
         
         If FlexMoAsig.Rows > 2 Then
            Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja IS NULL " & _
            " AND Codigo NOT IN (" & sListaMOSeleccionados & ");")
         Else
         
            Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja IS NULL;")
         End If
         
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
      
               With FlexMoDispo
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!descripcion
                  .TextMatrix(mj, 2) = NVL(mRec!Codigo, "")
               End With
         
               mRec.MoveNext
            Loop
         End If
         mRec.Close
      End If
      mRenglonMoAsign = 0
   End If
End Sub

Private Sub CommandPartes_Click(Index As Integer)
   Dim sListaPartesSeleccionados As String
   Dim mj As Integer
   Dim Origen As String
   Dim mTramo As String
   Dim mRamal As String
   Dim mNroComunicado As String
      
   sListaPartesSeleccionados = "-1"
   
   If Index = 0 Then
      If mRenglonPartes > 0 Then
         If Trim(FlexPartes.TextMatrix(mRenglonPartes, 1)) <> "" Then
            
            FlexPartAsignados.AddItem vbTab & FlexPartes.TextMatrix(mRenglonPartes, 1) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 2) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 3) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 4) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 5) & vbTab & FlexPartes.TextMatrix(mRenglonPartes, 6)
            'MsgBox FlexPartAsignados.Rows
         End If
         
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
   Else
      If FlexPartAsignados.Rows > 2 And mRenglonPartAsignados > 1 Then
         FlexPartAsignados.RemoveItem (mRenglonPartAsignados)
         
         If FlexPartAsignados.Rows > 2 Then
            For mj = 2 To FlexPartAsignados.Rows - 1
               sListaPartesSeleccionados = sListaPartesSeleccionados & "," & FlexPartAsignados.TextMatrix(mj, 1)
            Next
         End If
            
         mRenglonPartes = 0

         FlexPartes.Clear
         'Elimino los registros  de la grilla superior
         For mj = FlexPartes.Rows To 3 Step -1
            FlexPartes.RemoveItem mj
         Next
         
         With FlexPartes
            .TextMatrix(0, 1) = "Parte"
            .TextMatrix(0, 2) = "Fecha Solicitud"
            .TextMatrix(0, 3) = "Lugar"
            .TextMatrix(0, 4) = "Descripcion de la Solicitud"
            .TextMatrix(0, 5) = "Prioridad"
            
            .TextMatrix(0, 6) = "Sector Aire"
            
            .RowHeight(1) = 0
         End With
         
'         If Combo3.ListIndex = 0 Then
'            cargarGrillaConPartesOperaciones Trim(Left(Combo4.Text, 2)), sListaPartesSeleccionados
'         Else
'            cargarGrillaConPartesDeComunicado Trim(Combo3.Text), sListaPartesSeleccionados
'         End If
         
         
         
         If Combo3.ListIndex >= 0 Then
            
            Origen = Trim(Right(Combo3.Text, 3))
            Select Case Origen
               Case "OPE"
                  mTramo = Trim(Left(Combo4.Text, 2))
                  cargarGrillaConPartesOperaciones mTramo, sListaPartesSeleccionados
               Case "REL"
                  mRamal = Trim(Left(Combo4.Text, 50))
                  cargarGrillaConPartesDeRelevamientos mRamal, sListaPartesSeleccionados
               Case "COM"
                  mNroComunicado = Trim(Combo4.Text)
                  cargarGrillaConPartesDeComunicado mNroComunicado, sListaPartesSeleccionados
            End Select
         End If
         

      End If
      mRenglonPartAsignados = 0
   End If
End Sub

Private Sub cargarGrillaConPartesOperaciones(ByVal pTramo As String, ByVal plistaPartesSeleccionados As String)

   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer
                               
                                
'''   Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire,R.FechaIniAsist " & _
'''                                   " FROM Registros R " & _
'''                                       " Inner Join " & _
'''                                   " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                       " Left Join " & _
'''                                   " OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                       " Left Join " & _
'''                                   " COM_Comunicados_Det C ON C.Parte = R.Parte " & _
'''                                       " Left Join " & _
'''                                   " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                                " WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 AND OT.Finalizado = 'NO') OR (Cancelado = 0 AND Finalizado = 'NT')) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado = 'NO'))) " & _
'''                                " AND C.Parte IS NULL " & _
'''                                " AND CNL.Parte IS NULL " & _
'''                                " AND R.CodEdificio like '" & pTramo & "%' " & _
'''                                " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");")
                                
   Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire,R.FechaIniAsist " & _
                                   " FROM Registros R " & _
                                       " Inner Join " & _
                                   " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                       " Left Join " & _
                                   " OT_Partes OT ON OT.Parte = R.Parte " & _
                                       " Left Join " & _
                                   " COM_Comunicados_Det C ON C.Parte = R.Parte " & _
                                       " Left Join " & _
                                   " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                                " WHERE Estado NOT IN ('A', 'T') AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 AND OT.Finalizado = 'NO') OR (Cancelado = 0 AND Finalizado = 'NT')) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado = 'NO'))) " & _
                                " AND C.Parte IS NULL " & _
                                " AND CNL.Parte IS NULL " & _
                                " AND R.CodEdificio like '" & pTramo & "%' " & _
                                " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");")
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
               With FlexPartes
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!Parte
                  .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
                  .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
                  .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
                  .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
                  .TextMatrix(mj, 6) = NVL(mRec!FechaIniAsist, "")
                  .TextMatrix(mj, 6) = IIf(mRec!SectorAire = 1, "Si", "No")
               End With
               mRec.MoveNext
            Loop
         End If
         mRec.Close
End Sub

Private Sub cargarGrillaConPartesDeComunicado(ByVal pNroComunicado As String, ByVal plistaPartesSeleccionados As String)
   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer
                                
                                
'''   Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CD.NroComunicado,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
'''                                    " FROM Registros R " & _
'''                                    "     Inner Join " & _
'''                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                    "     Left Join " & _
'''                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " COM_Comunicados_Det CD ON CD.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " COM_Comunicados_H CH ON CD.NroComunicado = CH.NroComunicado " & _
'''                                        " Left Join " & _
'''                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                                    " WHERE Estado NOT IN ('A', 'T') " & _
'''                                    " AND CNL.Parte IS NULL " & _
'''                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
'''                                    " AND OT.Finalizado = 'NO') OR (Cancelado = 0 AND Finalizado = 'NT')) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado = 'NO'))) " & _
'''                                    " AND CH.NroComunicado = '" & pNroComunicado & "'" & _
'''                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");")
                                
                                
                                
   Set mRec = mObj.oEjecutarSelect("SELECT DISTINCT CD.NroComunicado,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
                                    " FROM Registros R " & _
                                    "     Inner Join " & _
                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                    "     Left Join " & _
                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " COM_Comunicados_Det CD ON CD.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " COM_Comunicados_H CH ON CD.NroComunicado = CH.NroComunicado " & _
                                        " Left Join " & _
                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                                    " WHERE Estado NOT IN ('A', 'T') " & _
                                    " AND CNL.Parte IS NULL " & _
                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
                                    " AND OT.Finalizado = 'NO') OR (Cancelado = 0 AND Finalizado = 'NT')) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado = 'NO'))) " & _
                                    " AND CH.NroComunicado = '" & pNroComunicado & "'" & _
                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");")
   If Not mRec.EOF Then
      mj = 1
      Do While Not mRec.EOF
         mj = mj + 1
         With FlexPartes
            .AddItem ""
            .TextMatrix(mj, 1) = mRec!Parte
            .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
            .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
            .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
            .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
            .TextMatrix(mj, 6) = NVL(mRec!FechaIniAsist, "")
            .TextMatrix(mj, 6) = IIf(mRec!SectorAire = 1, "Si", "No")
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub cargarGrillaConPartesDeRelevamientos(ByVal pDescRamal As String, ByVal plistaPartesSeleccionados As String)
   'IMPORTANTE: El parametro plistaPartesSeleccionados no puede venir vacio porque da error, en tal caso se lo fuerza con el parte = -1
   Dim mj As Integer
   Dim sSql As String
   
'''Backup sentencia igual a la siguiente
'''   sSql = "SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
'''                                    " FROM Registros R " & _
'''                                    "     Inner Join " & _
'''                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                    "     Left Join " & _
'''                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " REL_Relevamientos_Det RD ON RD.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
'''                                    "     Left Join " & _
'''                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                                    " WHERE Estado NOT IN ('A', 'T') " & _
'''                                    " AND CNL.Parte IS NULL " & _
'''                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
'''                                    " AND OT.Finalizado = 'NO') OR (Cancelado=0 AND Finalizado='NT' )) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado='NO'))) " & _
'''                                    " AND CodEdificio = '" & pDescRamal & "' " & _
'''                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ")"
'''   sSql = sSql & " UNION "
'''   sSql = sSql & " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.Descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
'''                                    " FROM Registros R " & _
'''                                    "     Inner Join " & _
'''                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
'''                                    "     Left Join " & _
'''                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " REL_Relevamientos_Det_Columnas RD ON RD.Parte = R.Parte " & _
'''                                    "     Inner Join " & _
'''                                    " REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
'''                                    "     Inner Join " & _
'''                                    " COM_Ramales CM ON CM.Codigo = RH.CodRamal " & _
'''                                    "     Left Join " & _
'''                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
'''                                    " WHERE Estado NOT IN ('A', 'T') " & _
'''                                    " AND CNL.Parte IS NULL " & _
'''                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
'''                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
'''                                    " AND OT.Finalizado = 'NO') OR (Cancelado=0 AND Finalizado='NT' )) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado='NO'))) " & _
'''                                    " AND CM.Descripcion ='" & pDescRamal & "'" & _
'''                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");"
   
   
   sSql = "SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
                                    " FROM Registros R " & _
                                    "     Inner Join " & _
                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                    "     Left Join " & _
                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " REL_Relevamientos_Det RD ON RD.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
                                    "     Left Join " & _
                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                                    " WHERE Estado NOT IN ('A', 'T') " & _
                                    " AND CNL.Parte IS NULL " & _
                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
                                    " AND OT.Finalizado = 'NO') OR (Cancelado=0 AND Finalizado='NT' )) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado='NO'))) " & _
                                    " AND CodEdificio = '" & pDescRamal & "' " & _
                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ")"
   sSql = sSql & " UNION "
   sSql = sSql & " SELECT DISTINCT RD.IdRele,R.Parte,FechaSolic,CodEdificio,R.Descripcion,Prioridad,R.SectorAire, FechaIniAsist " & _
                                    " FROM Registros R " & _
                                    "     Inner Join " & _
                                    " MailsxElectrico M ON R.SectorAire = M.SectorAire " & _
                                    "     Left Join " & _
                                    " OT_Partes OT ON OT.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " REL_Relevamientos_Det_Columnas RD ON RD.Parte = R.Parte " & _
                                    "     Inner Join " & _
                                    " REL_Relevamientos_H RH ON RD.IdRele = RH.IdRele " & _
                                    "     Inner Join " & _
                                    " COM_Ramales CM ON CM.Codigo = RH.CodRamal " & _
                                    "     Left Join " & _
                                    " Cancelaciones_Det CNL ON CNL.Parte = R.Parte " & _
                                    " WHERE Estado NOT IN ('A', 'T') " & _
                                    " AND CNL.Parte IS NULL " & _
                                    " AND Origen = 'O' AND M.codusuario = '" & Trim(Right(MDI.mUser, 20)) & "' " & _
                                    " AND (OT.Parte IS NULL OR (((OT.Cancelado = 1 " & _
                                    " AND OT.Finalizado = 'NO') OR (Cancelado=0 AND Finalizado='NT' )) AND NOT EXISTS(SELECT 1 FROM OT_Partes WHERE Parte = OT.Parte AND Cancelado = 0 AND Finalizado='NO'))) " & _
                                    " AND CM.Descripcion ='" & pDescRamal & "'" & _
                                    " AND R.Parte NOT IN (" & plistaPartesSeleccionados & ");"


   Set mRec = mObj.oEjecutarSelect(sSql)
   If Not mRec.EOF Then
      mj = 1
      Do While Not mRec.EOF
         mj = mj + 1
         With FlexPartes
            .AddItem ""
            .TextMatrix(mj, 1) = mRec!Parte
            .TextMatrix(mj, 2) = NVL(mRec!FechaSolic, "")
            .TextMatrix(mj, 3) = NVL(mRec!CodEdificio, "")
            .TextMatrix(mj, 4) = NVL(mRec!descripcion, "")
            .TextMatrix(mj, 5) = NVL(mRec!Prioridad, "")
            .TextMatrix(mj, 6) = NVL(mRec!FechaIniAsist, "")
            .TextMatrix(mj, 6) = IIf(mRec!SectorAire = 1, "Si", "No")
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub
















Private Sub CommandProd_Click(Index As Integer)

   Dim iStock As Double
   Dim mi As Integer
   Dim mRec1 As New ADODB.Recordset
   
   If Index = 0 Then
      If fValidaAsignaMateriales() Then
            FlexEgreso.AddItem vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 1) & vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 2) & vbTab & "" & vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 4) & vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 5) & vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 6) & vbTab & FlexProduct.TextMatrix(mRenglonProdDispo, 7)
            'FlexProduct.TextMatrix(mRenglonProdDispo, 3) = CDbl(Replace(Trim(FlexProduct.TextMatrix(mRenglonProdDispo, 3)), ".", ",")) - CDbl(Replace(Trim(Text2.Text), ".", ","))
            'Text2.Text = ""
            'Text2.SetFocus
      End If
   Else
'      For mI = 2 To FlexProduct.Rows - 1
'
'         If FlexProduct.TextMatrix(mI, 6) = FlexEgreso.TextMatrix(mRenglonEgreso, 6) And FlexProduct.TextMatrix(mI, 7) = FlexEgreso.TextMatrix(mRenglonEgreso, 7) Then
'            Set mRec1 = mObj.oEjecutarSelect("SELECT Fecha,CodProducto,CodUbicacion, IFNULL(Stock,0) As Stock " & _
'                                       " FROM Movimientos2 M " & _
'                                       " WHERE CodProducto  = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 6) & "' and CodUbicacion = '" & FlexEgreso.TextMatrix(mRenglonEgreso, 7) & "'" & _
'                                       " AND Fecha = (SELECT Max(Fecha) FROM Movimientos2 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion)")
'
'            If Not mRec1.EOF Then
'               iStock = mRec1!Stock
'            Else
'               iStock = 0
'            End If
'            mRec1.Close
'
'            FlexProduct.TextMatrix(mI, 3) = iStock
'
'            mI = 9999
'         End If
'      Next

      If FlexEgreso.Rows > 2 And mRenglonProdAsign > 1 Then
         '---------------CARGO VALOR TEXT EN GRILLA-------------------------
         Text2.Visible = False
         FlexEgreso.TextMatrix(filaAnt, columnAnt) = Text2.Text
         filaAnt = 0
         columnAnt = 0
         '---------------FIN: VALOR CARGO TEXT EN GRILLA--------------------
         FlexEgreso.RemoveItem (mRenglonProdAsign)
      End If

      mRenglonProdAsign = 0
   End If

End Sub

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

Private Sub CommandVeh_Click(Index As Integer)
   Dim sListaVehSeleccionados
   Dim mj As Integer
   sListaVehSeleccionados = ""
   
   If Index = 0 Then
      If mRenglonVehDispo > 0 Then
         If Trim(FlexVehDispo.TextMatrix(mRenglonVehDispo, 1)) <> "" Then
            
            FlexVehAsign.AddItem vbTab & FlexVehDispo.TextMatrix(mRenglonVehDispo, 1) & vbTab & FlexVehDispo.TextMatrix(mRenglonVehDispo, 2)

         End If
         
         If FlexVehDispo.Rows > 2 Then
            FlexVehDispo.RemoveItem mRenglonVehDispo
         
            mRenglonVehDispo = 0
         Else
            If Trim(FlexVehDispo.TextMatrix(mRenglonVehDispo, 1)) <> "" Then
               FlexVehDispo.TextMatrix(mRenglonVehDispo, 1) = ""
               FlexVehDispo.TextMatrix(mRenglonVehDispo, 2) = ""
         
               mRenglonVehDispo = 0
            End If
         End If
      End If
   Else
      If FlexVehAsign.Rows > 2 And mRenglonVehAsign > 1 Then
         
         FlexVehAsign.RemoveItem (mRenglonVehAsign)
         
         If FlexVehAsign.Rows > 2 Then
            For mj = 2 To FlexVehAsign.Rows - 1
               sListaVehSeleccionados = sListaVehSeleccionados & "'" & FlexVehAsign.TextMatrix(mj, 2) & "',"
            Next
            
            sListaVehSeleccionados = Left(sListaVehSeleccionados, Len(sListaVehSeleccionados) - 1)
         End If
            
         mRenglonVehDispo = 0

         FlexVehDispo.Clear
         'Elimino los registros  de la grilla superior
         For mj = FlexVehDispo.Rows To 3 Step -1
            FlexVehDispo.RemoveItem mj
         Next
         
         FlexVehDispo.Clear
         FlexVehDispo.Refresh
         
         With FlexVehDispo
            .TextMatrix(0, 1) = "Vehículo"
            .TextMatrix(0, 2) = "Codigo"
         
            .RowHeight(1) = 0
         End With
         
         If FlexVehAsign.Rows > 2 Then
            Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion " & _
                                   " From " & _
                                   "  Vehiculos V " & _
                                   " Left Join " & _
                                   "  Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where CodUbicacion Is Null " & _
                                   " AND V.Fecha_Baja IS NULL " & _
                                   " AND V.Codigo NOT IN (" & sListaVehSeleccionados & ");")
         Else
            Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion " & _
                                   " From " & _
                                   "  Vehiculos V " & _
                                   " Left Join " & _
                                   "  Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where CodUbicacion Is Null " & _
                                   " AND V.Fecha_Baja IS NULL; ")
         End If
         
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
      
               With FlexVehDispo
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!descripcion
                  .TextMatrix(mj, 2) = NVL(mRec!Codigo, "")
               End With
         
               mRec.MoveNext
            Loop
         End If
         mRec.Close
         
      End If
      mRenglonVehAsign = 0
   End If
End Sub

Private Sub CommandVehEsp_Click(Index As Integer)
   Dim sListaVehEspSeleccionados
   Dim mj As Integer
   sListaVehEspSeleccionados = ""
   
   If Index = 0 Then
      'Valido que solo se pueda seleccinar un vehiculo especial
      If FlexVehEspAsign.Rows < 3 Then
      
         If mRenglonVehEspDispo > 0 Then
            If Trim(FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 1)) <> "" Then
               
               FlexVehEspAsign.AddItem vbTab & FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 1) & vbTab & FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 2)
   
            End If
            
            If FlexVehEspDispo.Rows > 2 Then
               FlexVehEspDispo.RemoveItem mRenglonVehEspDispo
            
               mRenglonVehEspDispo = 0
            Else
               If Trim(FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 1)) <> "" Then
                  FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 1) = ""
                  FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 2) = ""
            
                  mRenglonVehEspDispo = 0
               End If
            End If
         End If
      Else
         MsgBox "Solo se puede seleccionar un vehículo especial por O.T.", vbExclamation
      End If
   Else
      
      If FlexVehEspAsign.Rows > 2 And mRenglonVehEspAsign > 1 Then
         
         FlexVehEspAsign.RemoveItem (mRenglonVehEspAsign)
         
         If FlexVehEspAsign.Rows > 2 Then
            For mj = 2 To FlexVehEspAsign.Rows - 1
               sListaVehEspSeleccionados = sListaVehEspSeleccionados & "'" & FlexVehEspAsign.TextMatrix(mj, 2) & "',"
            Next
            
            sListaVehEspSeleccionados = Left(sListaVehEspSeleccionados, Len(sListaVehEspSeleccionados) - 1)
        End If
            
         mRenglonVehEspDispo = 0

         FlexVehEspDispo.Clear
         'Elimino los registros  de la grilla superior
         For mj = FlexVehEspDispo.Rows To 3 Step -1
            FlexVehEspDispo.RemoveItem mj
         Next
         
         FlexVehEspDispo.Clear
         FlexVehEspDispo.Refresh
         
         With FlexVehEspDispo
            .TextMatrix(0, 1) = "Vehículo especial"
            .TextMatrix(0, 2) = "Codigo"
         
            .RowHeight(1) = 0
         End With
         

         If FlexVehEspAsign.Rows > 2 Then
         
            Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.Descripcion,  V.CodUbicacion " & _
                                            " From " & _
                                            " Vehiculos V " & _
                                            " Inner Join " & _
                                            " Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                            " Where V.Fecha_Baja IS NULL " & _
                                            " AND V.Codigo NOT IN (" & sListaVehEspSeleccionados & ");")
         Else
            Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.Descripcion,  V.CodUbicacion " & _
                                            " From " & _
                                            " Vehiculos V " & _
                                            " Inner Join " & _
                                            " Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                            " Where V.Fecha_Baja IS NULL;")
         End If
         
         If Not mRec.EOF Then
            mj = 1
            Do While Not mRec.EOF
               mj = mj + 1
      
               With FlexVehEspDispo
                  .AddItem ""
                  .TextMatrix(mj, 1) = mRec!descripcion
                  .TextMatrix(mj, 2) = NVL(mRec!Codigo, "")
               End With
         
               mRec.MoveNext
            Loop
         End If
         mRec.Close
         
      End If
      mRenglonVehEspAsign = 0
   End If
End Sub

Private Sub FlexEgreso_Click()
   Dim mi As Integer
   
   If FlexEgreso.MouseRow > 0 Then
   
      'En este caso 3 es la columna que seria editable
      If FlexEgreso.Col = 3 And FlexEgreso.Row <> 1 Then
         Text2.Text = FlexEgreso.Text
         Text2.Width = FlexEgreso.ColWidth(FlexEgreso.Col)
         Text2.Left = FlexEgreso.ColPos(FlexEgreso.Col) + FlexEgreso.Left + 30 'el valor treina termina de acomodar el textbox en la celda
         Text2.Top = FlexEgreso.Top + FlexEgreso.RowPos(FlexEgreso.Row)
         Text2.Visible = True
         Text2.SetFocus
      Else
         Text2.Visible = False
   
      End If
   
      filaAnt = FlexEgreso.Row
      columnAnt = FlexEgreso.Col
      
      If mRenglonProdAsign <> 0 Then
         If FlexEgreso.Rows > mRenglonProdAsign Then
            FlexEgreso.Row = mRenglonProdAsign
            For mi = 1 To FlexEgreso.Cols - 1
               FlexEgreso.Col = mi
               FlexEgreso.CellBackColor = vbWhite
            Next
         End If
      End If
      
      mRenglonProdAsign = FlexEgreso.MouseRow
      FlexEgreso.Row = mRenglonProdAsign
      For mi = 1 To FlexEgreso.Cols - 1
         FlexEgreso.Col = mi
         FlexEgreso.CellBackColor = &H80000003
      Next
                  
      
      If mRenglonProdAsign > 1 Then
          mCodProducto = FlexEgreso.TextMatrix(mRenglonProdAsign, 4)
      End If
   Else
      FlexEgreso.Row = mRenglonProdAsign
      For mi = 1 To FlexProduct.Cols - 1
         FlexEgreso.Col = mi
         FlexEgreso.CellBackColor = vbWhite
      Next
      mRenglonProdAsign = 0
   End If

End Sub

Private Sub FlexMoAsig_Click()
   Dim mi As Integer
   If FlexMoAsig.MouseRow > 0 Then
   
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
          mCodMO = FlexMoAsig.TextMatrix(mRenglonMoAsign, 2)
      End If
   Else
      FlexMoAsig.Row = mRenglonMoAsign
      For mi = 1 To FlexMoAsig.Cols - 1
         FlexMoAsig.Col = mi
         FlexMoAsig.CellBackColor = vbWhite
      Next
      mRenglonMoAsign = 0
   End If
End Sub

Private Sub FlexMoDispo_Click()
   Dim mi As Integer
   If FlexMoDispo.MouseRow > 0 Then
   
      If mRenglonMoDispo <> 0 Then
         FlexMoDispo.Row = mRenglonMoDispo
         For mi = 1 To FlexMoDispo.Cols - 1
            FlexMoDispo.Col = mi
            FlexMoDispo.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonMoDispo = FlexMoDispo.MouseRow
   
      FlexMoDispo.Row = mRenglonMoDispo
      For mi = 1 To FlexMoDispo.Cols - 1
         FlexMoDispo.Col = mi
         FlexMoDispo.CellBackColor = &H80000003
      Next
      
      If mRenglonMoDispo > 1 Then
          mCodMO = FlexMoDispo.TextMatrix(mRenglonMoDispo, 2)
      End If
   Else
      FlexMoDispo.Row = mRenglonMoDispo
      For mi = 1 To FlexMoDispo.Cols - 1
         FlexMoDispo.Col = mi
         FlexMoDispo.CellBackColor = vbWhite
      Next
      mRenglonMoDispo = 0
   End If
End Sub

Private Sub FlexPartAsignados_Click()
   Dim mi As Integer
   
   If FlexPartAsignados.MouseRow > 0 Then
   
      If mRenglonPartAsignados <> 0 Then
         FlexPartAsignados.Row = mRenglonPartAsignados
         For mi = 1 To FlexPartAsignados.Cols - 1
            FlexPartAsignados.Col = mi
            FlexPartAsignados.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonPartAsignados = FlexPartAsignados.MouseRow
   
      FlexPartAsignados.Row = mRenglonPartAsignados
      For mi = 1 To FlexPartAsignados.Cols - 1
         FlexPartAsignados.Col = mi
         FlexPartAsignados.CellBackColor = &H80000003
      Next
      
      If mRenglonPartAsignados > 1 Then
          mCodParte = FlexPartAsignados.TextMatrix(mRenglonPartAsignados, 1)
      End If
   Else
      FlexPartAsignados.Row = mRenglonPartAsignados
      For mi = 1 To FlexPartAsignados.Cols - 1
         FlexPartAsignados.Col = mi
         FlexPartAsignados.CellBackColor = vbWhite
      Next
      mRenglonPartAsignados = 0
   End If
End Sub

Private Sub FlexPartes_Click()
   Dim mi As Integer
   
   If FlexPartes.MouseRow > 0 Then
   
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
          mCodParte = FlexPartes.TextMatrix(mRenglonPartes, 1)
      End If
   Else
      FlexPartes.Row = mRenglonPartes
      For mi = 1 To FlexPartes.Cols - 1
         FlexPartes.Col = mi
         FlexPartes.CellBackColor = vbWhite
      Next
      mRenglonPartes = 0
   End If
End Sub

Private Sub FlexProduct_Click()
   Dim mi As Integer
   
   If FlexProduct.MouseRow > 0 Then
   
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
      For mi = 2 To FlexProduct.Cols - 1
         FlexProduct.Col = mi
         FlexProduct.CellBackColor = vbWhite
      Next
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
      For mi = 1 To FlexSubRubros.Cols - 1
         FlexSubRubros.Col = mi
         FlexSubRubros.CellBackColor = vbWhite
      Next
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
      For mi = 1 To FlexSubRubrosAsign.Cols - 1
         FlexSubRubrosAsign.Col = mi
         FlexSubRubrosAsign.CellBackColor = vbWhite
      Next
      mRenglonSubRubroAsign = 0
   End If
End Sub


Private Sub FlexVehAsign_Click()
   Dim mi As Integer
   
   If FlexVehAsign.MouseRow > 0 Then
   
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
          mCodVeh = FlexVehAsign.TextMatrix(mRenglonVehAsign, 2)
      End If
   Else
      FlexVehAsign.Row = mRenglonVehAsign
      For mi = 1 To FlexVehAsign.Cols - 1
         FlexVehAsign.Col = mi
         FlexVehAsign.CellBackColor = vbWhite
      Next
      mRenglonVehAsign = 0
   End If
End Sub

Private Sub FlexVehDispo_Click()
   Dim mi As Integer
   
   If FlexVehDispo.MouseRow > 0 Then
   
      If mRenglonVehDispo <> 0 Then
         FlexVehDispo.Row = mRenglonVehDispo
         For mi = 1 To FlexVehDispo.Cols - 1
            FlexVehDispo.Col = mi
            FlexVehDispo.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonVehDispo = FlexVehDispo.MouseRow
   
      FlexVehDispo.Row = mRenglonVehDispo
      For mi = 1 To FlexVehDispo.Cols - 1
         FlexVehDispo.Col = mi
         FlexVehDispo.CellBackColor = &H80000003
      Next
      
      If mRenglonVehDispo > 1 Then
          mCodVeh = FlexVehDispo.TextMatrix(mRenglonVehDispo, 2)
      End If
   Else
      FlexVehDispo.Row = mRenglonVehDispo
      For mi = 1 To FlexVehDispo.Cols - 1
         FlexVehDispo.Col = mi
         FlexVehDispo.CellBackColor = vbWhite
      Next
      mRenglonVehDispo = 0
   End If
End Sub

Private Sub FlexVehEspAsign_Click()
   Dim mi As Integer
   
   If FlexVehEspAsign.MouseRow > 0 Then
   
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
      For mi = 1 To FlexVehEspAsign.Cols - 1
         FlexVehEspAsign.Col = mi
         FlexVehEspAsign.CellBackColor = vbWhite
      Next
      mRenglonVehEspAsign = 0
   End If
End Sub

Private Sub FlexVehEspDispo_Click()
   Dim mi As Integer
   
   If FlexVehEspDispo.MouseRow > 0 Then
   
      If mRenglonVehEspDispo <> 0 Then
         FlexVehEspDispo.Row = mRenglonVehEspDispo
         For mi = 1 To FlexVehEspDispo.Cols - 1
            FlexVehEspDispo.Col = mi
            FlexVehEspDispo.CellBackColor = vbWhite
         Next
      End If
      
      mRenglonVehEspDispo = FlexVehEspDispo.MouseRow
   
      FlexVehEspDispo.Row = mRenglonVehEspDispo
      For mi = 1 To FlexVehEspDispo.Cols - 1
         FlexVehEspDispo.Col = mi
         FlexVehEspDispo.CellBackColor = &H80000003
      Next
      
      If mRenglonVehEspDispo > 1 Then
          mCodVehEsp = FlexVehEspDispo.TextMatrix(mRenglonVehEspDispo, 2)
      End If
   Else
      FlexVehEspDispo.Row = mRenglonVehEspDispo
      For mi = 1 To FlexVehEspDispo.Cols - 1
         FlexVehEspDispo.Col = mi
         FlexVehEspDispo.CellBackColor = vbWhite
      Next
      mRenglonVehEspDispo = 0
   End If
End Sub

Private Sub Form_Load()

   
   Me.Width = 17090
   Me.Height = 9920
   sAlinearForm Me
   
   Frame1(0).Visible = True
   Frame1(1).Visible = False
   Frame1(2).Visible = False
   Frame1(3).Visible = False
   Frame1(4).Visible = False
   
   'sLlenoCboComunicado
   sLlenoCboOrigen
   cboOrigenListIndex = -99
   cboDetalleListIndex = -99
 
   InicializoCboDetalle
   
   
   initPartes True
   initManoObra True
   initVehiculos True
   initVehiculosEspecial True
   initRubros_SubRubros True
   initMateriales
   
   
   
End Sub

Private Sub initPartes(pIniciaEncabezados As Boolean)
   Dim mi As Integer
   mRenglonPartes = 0
   mRenglonPartAsignados = 0

   If pIniciaEncabezados Then

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
         
         .ColAlignment(4) = flexAlignLeftCenter
         
         .RowHeight(1) = 0
      End With
      
      With FlexPartAsignados
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
         
         .ColAlignment(4) = flexAlignLeftCenter
         
         .RowHeight(1) = 0
      End With
   
   End If
   
 eliminoGrillaPartes
 eliminoGrillaPartesAsignados
 
  ' cargarGrillaConPartesOperaciones "-1"
   'cboOrigenListIndex = Combo3.ListIndex

End Sub

Private Sub initManoObra(pIniciaEncabezados As Boolean)
   Dim mi As Integer

   mRenglonMoDispo = 0
   mRenglonMoAsign = 0
   
   If pIniciaEncabezados Then
      With FlexMoDispo
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Técnico"
         .TextMatrix(0, 2) = "Codigo"
   
         .RowHeight(1) = 0
      End With
      With FlexMoAsig
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Técnico"
         .TextMatrix(0, 2) = "Codigo"
    
         .RowHeight(1) = 0
      End With
   End If
   
   
      
   'Elimino los registros de la grilla izquierda
   For mi = FlexMoDispo.Rows To 3 Step -1
      FlexMoDispo.RemoveItem mi
   Next
   'Elimino los registros de la grilla derecha
   For mi = FlexMoAsig.Rows To 3 Step -1
      FlexMoAsig.RemoveItem mi
   Next
   
   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM MO_Tecnicos Where Fecha_Baja IS NULL;")
                                
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
      
         With FlexMoDispo
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!descripcion
            .TextMatrix(mi, 2) = NVL(mRec!Codigo, "")
         End With
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub initVehiculos(pIniciaEncabezados As Boolean)
   Dim mi As Integer
   mRenglonVehDispo = 0
   mRenglonVehAsign = 0

   If pIniciaEncabezados Then
      With FlexVehDispo
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Vehículo"
         .TextMatrix(0, 2) = "Codigo"
   
         .RowHeight(1) = 0
      End With
      With FlexVehAsign
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Vehículo"
         .TextMatrix(0, 2) = "Codigo"
    
         .RowHeight(1) = 0
      End With
   End If
   
   'Elimino los registros de la grilla izquierda
   For mi = FlexVehDispo.Rows To 3 Step -1
      FlexVehDispo.RemoveItem mi
   Next
   'Elimino los registros de la grilla derecha
   For mi = FlexVehAsign.Rows To 3 Step -1
      FlexVehAsign.RemoveItem mi
   Next
   
   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion " & _
                                   " From " & _
                                   "  Vehiculos V " & _
                                   " Left Join " & _
                                   "  Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where CodUbicacion Is Null " & _
                                   " AND V.Fecha_Baja IS NULL; ")

   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
      
         With FlexVehDispo
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!descripcion
            .TextMatrix(mi, 2) = NVL(mRec!Codigo, "")
         End With
         
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub initVehiculosEspecial(pIniciaEncabezados As Boolean)
   Dim mi As Integer

   mRenglonVehEspDispo = 0
   mRenglonVehEspAsign = 0
   
   If pIniciaEncabezados Then
      With FlexVehEspDispo
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Vehículo especial"
         .TextMatrix(0, 2) = "Codigo"
   
         .RowHeight(1) = 0
      End With
      
      With FlexVehEspAsign
         .ColWidth(0) = 200
         .ColWidth(1) = 6200
         .ColWidth(2) = 0
         
         .TextMatrix(0, 1) = "Vehículo especial"
         .TextMatrix(0, 2) = "Codigo"
    
         .RowHeight(1) = 0
      End With
   End If
   
   'Elimino los registros de la grilla izquierda
   For mi = FlexVehEspDispo.Rows To 3 Step -1
      FlexVehEspDispo.RemoveItem mi
   Next
   'Elimino los registros de la grilla derecha
   For mi = FlexVehEspAsign.Rows To 3 Step -1
      FlexVehEspAsign.RemoveItem mi
   Next
   
   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.Descripcion,  V.CodUbicacion " & _
                                   " From " & _
                                   " Vehiculos V " & _
                                   " Inner Join " & _
                                   " Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
                                   " Where V.Fecha_Baja IS NULL;")
   
   If Not mRec.EOF Then
      mi = 1
      Do While Not mRec.EOF
         mi = mi + 1
         With FlexVehEspDispo
            .AddItem ""
            .TextMatrix(mi, 1) = mRec!descripcion
            .TextMatrix(mi, 2) = NVL(mRec!Codigo, "")
         End With
         mRec.MoveNext
      Loop
   End If
   mRec.Close
End Sub

Private Sub initRubros_SubRubros(pIniciaEncabezados As Boolean)
   Dim mi As Integer

   mRenglonSubRubroDispo = 0
   mRenglonSubRubroAsign = 0
   
   If pIniciaEncabezados Then
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
   End If

   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Rubros Where FechaBaja IS NULL;")
   Do While Not mRec.EOF
      'Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      Combo1.AddItem "" & mRec!descripcion & Space(50) & mRec!Codigo
      mRec.MoveNext
   Loop
   mRec.Close

   For mi = FlexSubRubros.Rows To 3 Step -1
      FlexSubRubros.RemoveItem mi
   Next
   For mi = FlexSubRubrosAsign.Rows To 3 Step -1
      FlexSubRubrosAsign.RemoveItem mi
   Next
End Sub

'Private Sub sLlenoCboComunicado()
'   Dim mRec1 As New ADODB.Recordset
'
'   Combo3.Clear
'   Set mRec1 = mObj.oEjecutarSelect("SELECT NroComunicado FROM MantElect.COM_Comunicados_H order by Fecha Desc; ")
'
'   Combo3.AddItem "NINGUN COMUNICADO"
'
'   If mObj.esSupervisorElectrico(Trim(Right(MDI.mUser, 20))) Then
'      Do While Not mRec1.EOF
'         Combo3.AddItem mRec1!NroComunicado
'         mRec1.MoveNext
'      Loop
'   End If
'   Combo3.ListIndex = 0
'
'   mRec1.Close
'   Set mRec1 = Nothing
'End Sub



Private Sub sLlenoCboOrigen()
   Combo3.Clear
   
   Combo3.AddItem "OPERACIONES" & Space(50) & "OPE"
   If mObj.esSupervisorElectrico(Trim(Right(MDI.mUser, 20))) Then
      Combo3.AddItem "RELEVAMIENTOS" & Space(50) & "REL"
      Combo3.AddItem "COMUNICADOS" & Space(50) & "COM"
   End If
   Combo3.ListIndex = -1
End Sub

Private Sub sLlenoCboDetalle()
   Dim mRec1 As New ADODB.Recordset
   Dim Origen As String
   
   Combo4.Enabled = True
   Combo4.Clear
   
   Origen = Trim(Right(Combo3.Text, 3))
   
   Select Case Origen
      Case "OPE"
         Set mRec1 = mObj.oEjecutarSelect("SELECT DISTINCT Tramo FROM MantElect.Edificios order by Tramo; ")
         Do While Not mRec1.EOF
            Combo4.AddItem mRec1!Tramo
            mRec1.MoveNext
         Loop
         mRec1.Close
      
      Case "REL"
         Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM COM_Ramales order by Descripcion; ")
         Do While Not mRec1.EOF
            Combo4.AddItem mRec1!descripcion & Space(50) & mRec1!Codigo
            mRec1.MoveNext
         Loop
         mRec1.Close
      Case "COM"
         Set mRec1 = mObj.oEjecutarSelect("SELECT NroComunicado FROM MantElect.COM_Comunicados_H order by Fecha Desc; ")
         Do While Not mRec1.EOF
            Combo4.AddItem mRec1!NroComunicado
            mRec1.MoveNext
         Loop
         mRec1.Close
   End Select
   
   Combo4.ListIndex = -1
End Sub


Private Sub InicializoCboDetalle()
   Combo4.Clear
   Combo4.Enabled = False
End Sub








Private Sub initMateriales()
    
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   With FlexProduct
      .ColWidth(0) = 200
      .ColWidth(1) = 10700
      .ColWidth(2) = 4500
      .ColWidth(3) = 1500
      .ColWidth(4) = 1900
      .ColWidth(5) = 1250
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      
      .TextMatrix(0, 1) = "Producto"
      .TextMatrix(0, 2) = "Ubicación"
      .TextMatrix(0, 3) = "Stock"
      .TextMatrix(0, 4) = "Unid.Medida"
      .TextMatrix(0, 5) = "Cód.Sap"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      
      .RowHeight(1) = 0
   End With

   With FlexEgreso
      .ColWidth(0) = 200
      .ColWidth(1) = 10700
      .ColWidth(2) = 4500
      .ColWidth(3) = 1500
      .ColWidth(4) = 1900
      .ColWidth(5) = 1250
      .ColWidth(6) = 0
      .ColWidth(7) = 0
      .ColWidth(8) = 0
      .ColWidth(9) = 0
      .ColWidth(10) = 0
   
      .TextMatrix(0, 1) = "Producto"
      .TextMatrix(0, 2) = "Ubicación"
      .TextMatrix(0, 3) = "Cantidad"
      .TextMatrix(0, 4) = "Unid.Medida"
      .TextMatrix(0, 5) = "Cód.Sap"
      .TextMatrix(0, 6) = "Cód. Producto"
      .TextMatrix(0, 7) = "Cód. Ubicacion"
      .TextMatrix(0, 8) = "CantidadOriginal"
      .TextMatrix(0, 9) = "StockActual"
      .TextMatrix(0, 10) = "YaEstaEnOT"

      .RowHeight(1) = 0
   End With
   
   
   'TODO: Debe traer las Ubicaciones correspondiente a vehiculos, no la ubicacion de Diego di pascual, en funcion de los permisos de acceso que tiene el usuario.
   'TODO: Cuando se pase a produccion recordar hardocdear esta quiere con los nuevos codigos, (acutalmente en desa: Bodega=0003,Ubicacion<>0016).
'        Set mRec = mObj.oEjecutarSelect(" SELECT Codigo,Descripcion,CodBodega " & _
'         " FROM Inventario.Ubicaciones  " & _
'         " WHERE CodBodega = '0003' " & _
'         " AND Codigo <> '0016' " & _
'         " AND Fecha_Baja IS NULL; ")
         
         
         Set mRec = mObj.oEjecutarSelect(" SELECT U.Codigo,U.Descripcion,U.CodBodega, V.CodUbicacion " & _
                                 " FROM " & _
                                 "   Inventario.Ubicaciones U " & _
                                 " INNER JOIN " & _
                                 "   Inventario.Usuario_AccesoBodega AB ON U.CodBodega = AB.CodBodega " & _
                                 " LEFT JOIN " & _
                                 "   MantElect.Vehiculos V ON V.CodUbicacion = U.Codigo " & _
                                 " Where V.CodUbicacion Is Not Null " & _
                                 " AND  AB.codusuario = '" & Trim(Right(MDI.mUser, 15)) & "' " & _
                                 " AND U.Fecha_Baja IS NULL; ")
                                 
         
         
         
   
   'Set mRec = mObj.oTabla("Bodegas", " where Fecha_Baja IS NULL and Codigo IN (SELECT CodBodega FROM Usuario_AccesoBodega WHERE CodUsuario = '" & Trim(Right(MDI.mUser, 15)) & "') order by Codigo")
      
   Do While Not mRec.EOF
      Combo2.AddItem "" & mRec!descripcion & Space(80) & mRec!Codigo & ""
      mRec.MoveNext
   Loop
   mRec.Close
   
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
   Dim mRet As Boolean
   Dim mMensajeError As String
   Dim mj As Integer
    
   mRet = True

   If mRenglonProdDispo = 0 Then
      mRet = False
      mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
   End If

   If mRet Then
      If mRenglonProdDispo <> 0 And FlexProduct.TextMatrix(mRenglonProdDispo, 1) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
      End If
   End If
   
   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
   If mRet Then
      For mj = 2 To FlexEgreso.Rows - 1
         If FlexEgreso.TextMatrix(mj, 6) = FlexProduct.TextMatrix(mRenglonProdDispo, 6) And FlexEgreso.TextMatrix(mj, 7) = FlexProduct.TextMatrix(mRenglonProdDispo, 7) Then
            mRet = False
            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
            mj = 999
         End If
      Next
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaAsignaMateriales = mRet
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
End Sub

Private Sub Text2_LostFocus()
   If FlexEgreso.Col <> columnAnt Or FlexEgreso.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         FlexEgreso.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub
