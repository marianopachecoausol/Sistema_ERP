VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MantElect07 
   Caption         =   "Nueva Orden de Trabajo"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   20340
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
      Height          =   495
      Index           =   1
      Left            =   11520
      TabIndex        =   27
      Top             =   12360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar O.T."
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   26
      Top             =   12360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1(4)"
      Height          =   20
      Index           =   4
      Left            =   22
      TabIndex        =   5
      Top             =   1200
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   20
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   21100
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
         Height          =   8505
         Left            =   70
         TabIndex        =   20
         Top             =   0
         Width           =   20895
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   278
            Left            =   18000
            TabIndex        =   34
            Top             =   6840
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
            Height          =   7065
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   20415
            _ExtentX        =   36010
            _ExtentY        =   12462
            _Version        =   327680
            Cols            =   8
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
      Height          =   20
      Index           =   2
      Left            =   22
      TabIndex        =   3
      Top             =   1200
      Width           =   21100
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   1
         Left            =   10795
         Picture         =   "MantElect07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5280
         Width           =   495
      End
      Begin VB.CommandButton CommandSubRubro 
         Height          =   495
         Index           =   0
         Left            =   9807
         Picture         =   "MantElect07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5280
         Width           =   495
      End
      Begin VB.Frame Frame7 
         Caption         =   "Rubros/Subrubros asignados"
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
         Height          =   3735
         Left            =   480
         TabIndex        =   12
         Top             =   6000
         Width           =   19895
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubrosAsign 
            Height          =   2655
            Left            =   1080
            TabIndex        =   13
            Top             =   720
            Width           =   18460
            _ExtentX        =   32570
            _ExtentY        =   4683
            _Version        =   327680
            Cols            =   5
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Rubros/Subrubros"
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
         Height          =   4815
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   19895
         Begin VB.ComboBox Combo1 
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
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   480
            Width           =   3255
         End
         Begin MSFlexGridLib.MSFlexGrid FlexSubRubros 
            Height          =   3135
            Left            =   1080
            TabIndex        =   11
            Top             =   1080
            Width           =   18465
            _ExtentX        =   32570
            _ExtentY        =   5530
            _Version        =   327680
            Cols            =   5
         End
         Begin VB.Label Label1 
            Caption         =   "Rubro:"
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
            Left            =   1080
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1(1)"
      Height          =   10020
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   21100
      Begin VB.Frame Frame9 
         Caption         =   "Vehículos Especiales asignados"
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
         Height          =   3015
         Left            =   7920
         TabIndex        =   18
         Top             =   240
         Width           =   7335
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   278
            Left            =   5160
            TabIndex        =   38
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
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
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3495
         Left            =   4440
         TabIndex        =   8
         Top             =   3480
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
            Size            =   12
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
            BackColor       =   &H80000013&
            Height          =   278
            Left            =   5160
            TabIndex        =   36
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
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
      Height          =   10020
      Index           =   0
      Left            =   22
      TabIndex        =   1
      Top             =   1200
      Width           =   21100
      Begin VB.Frame Frame12 
         Caption         =   "Partes"
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
         Height          =   7335
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   19935
         Begin MSFlexGridLib.MSFlexGrid FlexPartes 
            Height          =   6375
            Left            =   300
            TabIndex        =   25
            Top             =   600
            Width           =   19335
            _ExtentX        =   34105
            _ExtentY        =   11245
            _Version        =   327680
            Cols            =   7
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   21135
      _ExtentX        =   37280
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
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
            Caption         =   "OT - Asignación Materiales"
            Object.Tag             =   "SupervisorMantElectrico"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "OT - Abastecimiento"
            Object.Tag             =   "Bodeguero"
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
Attribute VB_Name = "MantElect07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim mObj As New clMantElect
'Dim mObjInven As New clInven
'Dim mRec As New ADODB.Recordset
'Dim mRenglonPartes As Integer
'Dim mRenglonVehAsign As Integer
'Dim mRenglonVehEspAsign As Integer
'Dim mRenglonMoAsign As Integer
'Dim mRenglonSubRubroDispo As Integer
'Dim mRenglonSubRubroAsign As Integer
'Dim mRenglonProdDispo As Integer
'
'Dim XLS As EXCEL.Application
'
'Dim filaAnt As Integer
'Dim columnAnt As Integer
'Dim filaAntVehAsign As Integer
'Dim columnAntVehAsign As Integer
'Dim filaAntVehAsignKmFinal As Integer
'Dim columnAntVehAsignKmFinal As Integer
'
'Dim filaAntVehEspAsignKmInicio As Integer
'Dim columnAntVehEspAsignKmInicio As Integer
'Dim filaAntVehEspAsignKmFinal As Integer
'Dim columnAntVehEspAsignKmFinal As Integer
'
'
''TODO: Ver si es necesario utilizar las siguientes variables:
'Dim mCodParte As Integer
'Dim mCodMO As String
'Dim mCodSubrubro As String
'Dim mCodVeh As String
'Dim mCodVehEsp As String
'Dim mCodProducto As String
'
'
'Private Sub Combo1_Click()
' Dim mI As Integer
' Dim sListaSubrubrosSeleccionados As String
'
'   'Elimino los registros  de la grilla superior
'   For mI = FlexSubRubros.Rows To 3 Step -1
'      FlexSubRubros.RemoveItem mI
'   Next
'
'   If FlexSubRubrosAsign.Rows > 2 Then
'      For mI = 2 To FlexSubRubrosAsign.Rows - 1
'         sListaSubrubrosSeleccionados = sListaSubrubrosSeleccionados & "'" & FlexSubRubrosAsign.TextMatrix(mI, 4) & "',"
'      Next
'      sListaSubrubrosSeleccionados = Left(sListaSubrubrosSeleccionados, Len(sListaSubrubrosSeleccionados) - 1)
'   End If
'
'   mRenglonSubRubroDispo = 0
'
'   If FlexSubRubrosAsign.Rows > 2 Then
'
'         Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
'         "  From " & _
'         " Rubros R " & _
'         "  Inner Join " & _
'         " SubRubros S ON S.CodRubro = R.Codigo " & _
'         " WHERE S.Codigo NOT IN (" & sListaSubrubrosSeleccionados & ")" & _
'         " AND R.Codigo ='" & Right(Combo1.Text, 8) & "';")
'   Else
'      Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
'         "  From " & _
'         " Rubros R " & _
'         "  Inner Join " & _
'         " SubRubros S ON S.CodRubro = R.Codigo" & _
'         " WHERE R.Codigo ='" & Right(Combo1.Text, 8) & "';")
'
'   End If
'
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexSubRubros
'            .AddItem ""
'            .TextMatrix(mI, 1) = mRec!RubroDesc
'            .TextMatrix(mI, 2) = mRec!SubRubroDesc
'            .TextMatrix(mI, 3) = mRec!CodRubro
'            .TextMatrix(mI, 4) = mRec!CodSubrubro
'         End With
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'End Sub
'
'Private Sub Combo3_Click()
'   Dim mI As Integer
'   Dim mIdOT As Integer
'   Dim mCodUbicacion As String
'
'   mIdOT = Left(Combo3.Text, 10)
'
'   '--PARTES-------------------------------------------------------------------------------------------------------------------
'   'Elimino los registros (de la consulta anterior) de la grilla superior
'
'   mRenglonPartes = 0
'
'   For mI = FlexPartes.Rows To 3 Step -1
'      FlexPartes.RemoveItem mI
'   Next
'
'   'TODO: VER SI EN ESTA QUERY NO TENGO QUE CONTEMPLAR DIFERENCIAR OT de aire y d electricos puro
'   Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,FechaSolic,CodEdificio,descripcion,Prioridad,R.SectorAire " & _
'                                       "FROM " & _
'                                          "OT_Partes OH " & _
'                                       "Inner Join " & _
'                                          "Registros R ON OH.Parte = R.Parte " & _
'                                        "where OH.IdOT = " & mIdOT & " " & _
'                                        "AND Cancelado = 0 " & _
'                                        "AND Finalizado = 'NO'; ")
'
'
'
'
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexPartes
'            .AddItem ""
'            .TextMatrix(mI, 0) = "X"
'            .TextMatrix(mI, 1) = mRec!Parte
'            .TextMatrix(mI, 2) = NVL(mRec!FechaSolic, "")
'            .TextMatrix(mI, 3) = NVL(mRec!CodEdificio, "")
'            .TextMatrix(mI, 4) = NVL(mRec!descripcion, "")
'            .TextMatrix(mI, 5) = NVL(mRec!Prioridad, "")
'
'            .TextMatrix(mI, 6) = IIf(mRec!SectorAire = 1, "Si", "No")
'
'         End With
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
''----------------------------------------------------------------------------------------------------------------------------------
'
'
'
'
''--VEHICULOS------------------------------------------------------------------------------------------------------------------------
'   'Elimino los registros (de la consulta anterior) de la grilla superior
'   For mI = FlexVehAsign.Rows To 3 Step -1
'      FlexVehAsign.RemoveItem mI
'   Next
'
'   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion " & _
'                                   " From " & _
'                                   "   Vehiculos V " & _
'                                   " Inner Join " & _
'                                   "   OT_Vehiculos OV ON V.Codigo = OV.CodVehiculo " & _
'                                   " Left Join " & _
'                                   "   Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
'                                   " Where OV.IdOT = " & mIdOT & _
'                                   " AND CodUbicacion Is Null; ")
'
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexVehAsign
'            .AddItem ""
'            .TextMatrix(mI, 0) = "X"
'            .TextMatrix(mI, 2) = mRec!descripcion
'            .TextMatrix(mI, 5) = NVL(mRec!Codigo, "")
'         End With
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
''--------------------------------------------------------------------------------------------------------------------------------------
'
'
'
''--VEHICULOS ESPECIAL (tambien completo grilla materiales)------------------------------------------------------------------------------------------------------------------
'   'Elimino los registros de la grilla
'   For mI = FlexVehEspAsign.Rows To 3 Step -1
'      FlexVehEspAsign.RemoveItem mI
'   Next
'
'   Set mRec = mObj.oEjecutarSelect(" SELECT V.Codigo, V.descripcion, V.CodUbicacion " & _
'                                   " From " & _
'                                   "   Vehiculos V " & _
'                                   " Inner Join " & _
'                                   "   OT_Vehiculos OV ON V.Codigo = OV.CodVehiculo " & _
'                                   " Left Join " & _
'                                   "   Inventario.Ubicaciones U ON V.CodUbicacion = U.Codigo " & _
'                                   " Where OV.IdOT = " & mIdOT & _
'                                   " AND CodUbicacion Is NOT Null; ")
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexVehEspAsign
'            .AddItem ""
'            .TextMatrix(mI, 1) = mRec!descripcion
'            .TextMatrix(mI, 4) = NVL(mRec!Codigo, "")
'         End With
'
'         mCodUbicacion = NVL(mRec!CodUbicacion, "")
'
'         Text1.Enabled = False
'         Text1 = mRec!descripcion & Space(100) & mRec!Codigo
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'
'
'
'   mRenglonProdDispo = 0
'   Text2.Text = ""
'   Text2.Visible = False
'   FlexProduct.ScrollBars = flexScrollBarVertical
'
'  'Elimino los registros  de la grilla
'  For mI = FlexProduct.Rows To 3 Step -1
'      FlexProduct.RemoveItem mI
'   Next
'
'
'  Set mRec = mObjInven.oEjecutarSelect(" SELECT IdMov, Fecha, CodTipoMovimiento, CodProducto, P.Descripcion AS Producto, CodUbicacion, " & _
'   " U.Descripcion AS Ubicacion, 0 AS Cantidad , Stock, Med.Descripcion AS UnidadMedida, P.CodigoSap, CodUsuario,   Observaciones " & _
'   " From " & _
'   "  Movimientos2 M " & _
'   " Inner Join " & _
'   "  Producto P ON M.CodProducto = P.Codigo " & _
'   " Inner Join " & _
'   "  Ubicaciones U ON  M.CodUbicacion = U.Codigo " & _
'   " Inner Join " & _
'   "  UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
'   " WHERE Fecha = (SELECT MAX(Fecha) " & _
'   "                From Movimientos2 " & _
'   "               WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
'   " AND U.Codigo ='" & mCodUbicacion & "';")
'
'
'   'Cargo la Grilla Superior con lo devuelto por el sp "getStockXBodegaConFiltroProducto"
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         FlexProduct.AddItem ""
'         FlexProduct.TextMatrix(mI, 1) = mRec!CodigoSap
'         FlexProduct.TextMatrix(mI, 2) = mRec!Producto
'         FlexProduct.TextMatrix(mI, 3) = mRec!Cantidad
'         FlexProduct.TextMatrix(mI, 4) = mRec!Stock
'         FlexProduct.TextMatrix(mI, 5) = mRec!UnidadMedida
'         FlexProduct.TextMatrix(mI, 6) = mRec!CodProducto
'         FlexProduct.TextMatrix(mI, 7) = mRec!CodUbicacion
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'
'
''--------------------------------------------------------------------------------------------------------------------------------------
'
'
'
''--TECNICOS----------------------------------------------------------------------------------------------------------------------------
'   'Elimino los registros (de la consulta anterior) de la grilla superior
'   For mI = FlexMoAsig.Rows To 3 Step -1
'      FlexMoAsig.RemoveItem mI
'   Next
'
'   Set mRec = mObj.oEjecutarSelect(" SELECT Codigo,Descripcion " & _
'                                   " From " & _
'                                   "   MO_Tecnicos M " & _
'                                   " Inner Join " & _
'                                   "  OT_MO_Tecnicos OM ON OM.CodMo_Tecnico = M.Codigo " & _
'                                   " WHERE OM.IdOT = " & mIdOT & ";")
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexMoAsig
'            .AddItem ""
'            .TextMatrix(mI, 0) = "X"
'            .TextMatrix(mI, 1) = mRec!descripcion
'            .TextMatrix(mI, 2) = NVL(mRec!Codigo, "")
'         End With
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'
'
'
''--SUBRUBROS--------------------------------------------------------------------------------------------------------------------------
'
'
'   Combo1.Clear
'   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Rubros Where FechaBaja IS NULL;")
'
'   Do While Not mRec.EOF
'      'Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
'      Combo1.AddItem "" & mRec!descripcion & Space(50) & mRec!Codigo
'      mRec.MoveNext
'   Loop
'   mRec.Close
'
'
'   'Elimino los registros de la grilla
'   For mI = FlexSubRubros.Rows To 3 Step -1
'      FlexSubRubros.RemoveItem mI
'   Next
'
'
'   'Elimino los registros de la grilla
'   For mI = FlexSubRubrosAsign.Rows To 3 Step -1
'      FlexSubRubrosAsign.RemoveItem mI
'   Next
'
'   Set mRec = mObj.oEjecutarSelect(" SELECT R.Codigo As CodRubro, R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc FROM " & _
'                                   " OT_Subrubros OS " & _
'                                   "   Inner Join " & _
'                                   " SubRubros S ON OS.CodSubrubro = S.Codigo " & _
'                                   "   Inner Join " & _
'                                   " Rubros R ON S.CodRubro = R.Codigo " & _
'                                   " WHERE OS.IdOT = " & mIdOT & ";")
'
'   If Not mRec.EOF Then
'      mI = 1
'      Do While Not mRec.EOF
'         mI = mI + 1
'
'         With FlexSubRubrosAsign
'            .AddItem ""
'            .TextMatrix(mI, 1) = mRec!RubroDesc
'            .TextMatrix(mI, 2) = mRec!SubRubroDesc
'            .TextMatrix(mI, 3) = mRec!CodRubro
'            .TextMatrix(mI, 4) = mRec!CodSubrubro
'
'         End With
'
'         mRec.MoveNext
'      Loop
'   End If
'   mRec.Close
'End Sub
'
'
'
'Private Sub Command2_Click(Index As Integer)
'   If Index = 0 Then
'      Dim OTgenerada As Integer
'      Dim vPartes_OT() As Double
'      Dim vVehiculos_OT() As String
'      Dim vVehiculosEsp_OT() As String
'      Dim vMO_Tecnicos_OT() As String
'      Dim vSubrubros_OT() As String
'      Dim mI As Integer
'      Dim fecOt As Date
'
'
'      fecOt = Now()
'
'      'TODO: Validar Fecha inicio fin, fechas validas, fechas mayores o iguales a fecha cracion ot
'      If fValidaOT Then
'
'         preparaArrayPartes vPartes_OT()
'         preparaArrayVehiculos vVehiculos_OT()
'         preparaArrayVehiculosEsp vVehiculosEsp_OT()
'         preparaArrayMO_Tecnicos vMO_Tecnicos_OT()
'         preparaArraySubrubros vSubrubros_OT()
'
'         OTgenerada = mObj.xinsOT("mpacheco", vPartes_OT(), vVehiculos_OT(), vVehiculosEsp_OT(), vMO_Tecnicos_OT(), vSubrubros_OT(), fecOt)
'         'TODO: REMPLAZAR POR LA SIGUIENTE LINEA COMENTADA
'         'OTgenerada = mObj.xinsOT(Trim(Right(MDI.mUser, 15)), vPartes_OT())
'
'
'         If OTgenerada <> 0 Then
'            MsgBox "Se ha generado la Orden de trabajo: " & OTgenerada
'            imprimirExcelOT OTgenerada, fecOt, Trim(Left(MDI.mUser, 40))
'         End If
'      End If
'   Else
'      'TODO: Ver el evento unload
'      Unload Me
'   End If
'End Sub
'
'Private Function fValidaOT() As Boolean
''
''
''   Dim mRet As Boolean
''   Dim mMensajeError As String
''   Dim mCodTipoVale As String
''   Dim mRec1 As New ADODB.Recordset
''
''   mRet = True
''
''   If FlexPartAsignados.Rows <= 2 Then
''      mRet = False
''
''      Frame1(0).Visible = True
''      Frame1(1).Visible = False
''      Frame1(2).Visible = False
''      Frame1(3).Visible = False
''      Frame1(4).Visible = False
''
''      mMensajeError = "Al menos se debe seleccionar un Parte"
''   End If
''
''   If mRet Then
''      If FlexMoAsig.Rows <= 2 Then
''         mRet = False
''
''         Frame1(0).Visible = False
''         Frame1(1).Visible = True
''         Frame1(2).Visible = False
''         Frame1(3).Visible = False
''         Frame1(4).Visible = False
''
''         mMensajeError = "Al menos se debe seleccionar un técnico"
''      End If
''   End If
''
''   If mRet Then
''      If FlexSubRubrosAsign.Rows <= 2 Then
''         mRet = False
''
''         Frame1(0).Visible = False
''         Frame1(1).Visible = False
''         Frame1(2).Visible = True
''         Frame1(3).Visible = False
''         Frame1(4).Visible = False
''
''         mMensajeError = "Al menos se debe seleccionar un Subrubro"
''      End If
''   End If
''
''   If Not mRet Then
''         MsgBox mMensajeError, vbCritical, "Atención"
''   End If
''
''   fValidaOT = mRet
'End Function
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
'
'
'
'
'
'
'
'Private Sub imprimirExcelOT(ByVal NroOT As Integer, FechaOT As Date, Supervisor As String)
'
'
'         sMsgEspere Me, "Procesando datos...", True
'         'mFechaEjec = Now()
'
'         Set XLS = CreateObject("Excel.Application")
'
'         sPlanilla1 NroOT, FechaOT, Supervisor
'
'
'         XLS.Worksheets(1).Select
'
'         sMsgEspere Me, "", False
'         XLS.Application.Visible = True
'
'End Sub
'
'Private Sub sPlanilla1(NroOT As Integer, FechaOT As Date, Supervisor As String)
''   mI = 10
'   sCabecera1 NroOT, FechaOT, Supervisor
'
''   Set mRec = mObj.oEjecutarSelect(" SELECT  CodProducto,CodigoSAP, P.Descripcion AS Producto, CodBodega, B.Descripcion AS Bodega, SUM(Stock) AS Stock, Med.Descripcion AS UnidadMedida " & _
''   "FROM  " & _
''   " Movimientos2 M " & _
''   "  INNER JOIN " & _
''   " Producto P ON M.CodProducto = P.Codigo " & _
''   "  INNER JOIN " & _
''   " Ubicaciones U ON  M.CodUbicacion = U.Codigo AND U.CodBodega = '" & Left(Combo1.Text, 4) & "' " & _
''   "  INNER JOIN " & _
''   " Bodegas B ON B.Codigo = U.CodBodega  " & _
''   "  INNER JOIN " & _
''   " UnidadMedida Med ON P.CodUnidadMedida = Med.Codigo " & _
''   " WHERE Fecha = (SELECT MAX(Fecha) " & _
''   "                 From Movimientos2 " & _
''   "                 WHERE CodProducto = M.CodProducto AND CodUbicacion = M.CodUbicacion) " & _
''   " GROUP BY   CodProducto, P.Descripcion,CodBodega, B.Descripcion,Med.Descripcion " & _
''   " ORDER BY   P.Descripcion ;")
''
''   Do While Not mRec.EOF
''      With XLS
''
''      .Cells(mI, 1).Formula = NVL(mRec!CodProducto, "")
''      .Cells(mI, 2).Formula = NVL(mRec!CodigoSap, "")
''      .Cells(mI, 3).Formula = NVL(mRec!Producto, "")
''      .Cells(mI, 4).Formula = NVL(mRec!CodBodega, "")
''      .Cells(mI, 5).Formula = NVL(mRec!Bodega, "")
''      .Cells(mI, 6).Formula = NVL(mRec!Stock, "")
''      .Cells(mI, 7).Formula = NVL(mRec!UnidadMedida, "")
''
''      End With
''      mRec.MoveNext
''      mI = mI + 1
''   Loop
''   mRec.Close
'End Sub
'
'Private Sub sCabecera1(NroOT As Integer, FechaOT As Date, Supervisor As String)
'   Dim mI As Integer
'   Dim primerColumna As Boolean
'
'   mI = 10
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
'         .Range("B" & mI & ":H" & mI).Select
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
'         .Range("E" & mI & ":E" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         If primerColumna Then
'            .Cells(mI, 2).Formula = NVL(mRec!descripcion, "")
'            primerColumna = False
'         Else
'            .Cells(mI, 5).Formula = NVL(mRec!descripcion, "")
'            primerColumna = True
'            mI = mI + 1
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
'      mI = mI + 2
'
'      .Range("B" & mI & ":H" & (mI + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mI & ":H" & (mI + 1)).Select
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
'      .Range("E" & (mI + 1) & ":H" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("G" & (mI + 1) & ":G" & (mI + 1)).Select
'       With .Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'       End With
'
'      .Cells(mI, 4).Formula = "VEHICULOS QUE INTERVIENEN"
'      mI = mI + 1
'      .Cells(mI, 2).Formula = "Vehículo"
'      .Cells(mI, 5).Formula = "Km Inicial"
'      .Cells(mI, 7).Formula = "Km Final"
'
''-----------------------------------------------------------------------------------------------------
'
'
''---------------------------------DETALLE VEHICULOS------------------------------------------------
'      mI = mI + 1
'      Set mRec = mObj.oEjecutarSelect("SELECT Codigo,Descripcion FROM " & _
'                                          "OT_Vehiculos O " & _
'                                              "Inner Join " & _
'                                          "Vehiculos V ON O.CodVehiculo = Codigo " & _
'                                      "WHERE IdOT = '" & NroOT & "'; ")
'
'      Do While Not mRec.EOF
'         .Range("B" & mI & ":H" & mI).Select
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
'         .Range("E" & mI & ":E" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("G" & mI & ":G" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mI, 2).Formula = NVL(mRec!descripcion, "")
'         End With
'         mRec.MoveNext
'         mI = mI + 1
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
'
'
'
''---------------------------------ENCABEZADO TAREAS--------------------------------------------------
'      mI = mI + 2
'
'      .Cells(mI, 5).Formula = "TAREAS"
'      .Cells(mI + 1, 2).Formula = "Parte"
'      .Cells(mI + 1, 3).Formula = "Lugar"
'      .Cells(mI + 1, 4).Formula = "Descripcion"
'      .Cells(mI + 1, 9).Formula = "¿Finalizado?"
'
'      .Range("B" & mI & ":I" & (mI + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mI & ":I" & (mI + 1)).Select
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
'      .Range("C" & (mI + 1) & ":C" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("D" & (mI + 1) & ":D" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("I" & (mI + 1) & ":I" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
''---------------------------------DETALLE TAREAS------------------------------------------------------
'      mI = mI + 2
'      Set mRec = mObj.oEjecutarSelect("SELECT R.Parte,R.CodEdificio, R.Descripcion FROM " & _
'                                          "OT_Partes O " & _
'                                              "Inner Join " & _
'                                          "Registros R ON O.Parte = R.Parte " & _
'                                          "WHERE IDOT = '" & NroOT & "' " & _
'                                          "ORDER BY R.parte; ")
'
'      Do While Not mRec.EOF
'         .Range("B" & mI & ":I" & mI).Select
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
'         .Range("C" & mI & ":C" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("D" & mI & ":D" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("I" & mI & ":I" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mI, 2).Formula = NVL(mRec!Parte, "")
'            .Cells(mI, 3).Formula = NVL(mRec!CodEdificio, "")
'            .Cells(mI, 4).Formula = NVL(mRec!descripcion, "")
'         End With
'         mRec.MoveNext
'         mI = mI + 1
'      Loop
'      mRec.Close
''-----------------------------------------------------------------------------------------------------
''---------------------------------ENCABEZADO SUBRUBROS------------------------------------------------
'      mI = mI + 2
'
'      .Cells(mI, 5).Formula = "FALLAS"
'      .Cells(mI + 1, 2).Formula = "Subrubro"
'      .Cells(mI + 1, 6).Formula = "Subrubro"
'
'      .Range("B" & mI & ":I" & (mI + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mI & ":I" & (mI + 1)).Select
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
'      .Range("E" & mI + 1 & ":E" & (mI + 1)).Select
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
'      .Range("I" & mI + 1 & ":I" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
''---------------------------------DETALLE SUBRUBROS-----------------------------------------------
'      mI = mI + 2
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
'            .Range("B" & mI & ":I" & mI).Select
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
'            .Range("E" & mI & ":E" & mI).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Range("F" & mI & ":F" & mI).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlMedium
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Range("I" & mI & ":I" & mI).Select
'            With .Selection.Borders(xlEdgeLeft)
'              .LineStyle = xlContinuous
'              .Weight = xlThin
'              .ColorIndex = xlAutomatic
'            End With
'
'            .Cells(mI, 2).Formula = NVL(mRec!descripcion, "")
'            primerColumna = False
'         Else
'            .Cells(mI, 6).Formula = NVL(mRec!descripcion, "")
'            primerColumna = True
'            mI = mI + 1
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
'      mI = mI + 2
'      .Cells(mI, 4).Formula = "                 MATERIALES"
'      .Cells(mI + 1, 2).Formula = "Cód.Sap"
'      .Cells(mI + 1, 3).Formula = "Descripción"
'      .Cells(mI + 1, 6).Formula = "Consumido"
'      .Cells(mI + 1, 7).Formula = "Unid. Media"
'
'      .Range("B" & mI & ":H" & (mI + 1)).Select
'      .Selection.Interior.ColorIndex = 15
'
'      .Range("B" & mI & ":H" & (mI + 1)).Select
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
'      .Range("C" & (mI + 1) & ":C" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("F" & (mI + 1) & ":F" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("G" & (mI + 1) & ":G" & (mI + 1)).Select
'      With .Selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
''-----------------------------------------------------------------------------------------------------
'
'
''---------------------------------DETALLE Materiales--------------------------------------------------
'      mI = mI + 2
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
'                                          "AND U.Codigo = '0006'" & _
'                                          "and OV.IDOT = '" & NroOT & "' and stock > 0; ")
'
'      Do While Not mRec.EOF
'         .Range("B" & mI & ":H" & mI).Select
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
'         .Range("C" & mI & ":C" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("F" & mI & ":F" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         .Range("G" & mI & ":G" & mI).Select
'         With .Selection.Borders(xlEdgeLeft)
'           .LineStyle = xlContinuous
'           .Weight = xlThin
'           .ColorIndex = xlAutomatic
'         End With
'
'         With XLS
'            .Cells(mI, 2).Formula = NVL(mRec!CodigoSap, "")
'            .Cells(mI, 3).Formula = NVL(mRec!descripcion, "")
'            .Cells(mI, 7).Formula = NVL(mRec!UnidadMedidad, "")
'         End With
'         mRec.MoveNext
'         mI = mI + 1
'      Loop
'      mRec.Close
'
''-----------------------------------------------------------------------------------------------------
'
'
''----------------------------------------------OBSERVACIONES------------------------------------------
'      mI = mI + 2
'      .Cells(mI, 2).Formula = "OBSERVACIONES"
'      mI = mI + 1
'      .Range("B" & mI & ":I" & (mI + 4)).Select
'      .Selection.RowHeight = 16.5
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
'      mI = mI + 8
'      .Cells(mI, 3).Formula = "              SUPERVISOR"
'      .Cells(mI, 6).Formula = "     ENCARGADO BODEGA"
'
'      .Range("C" & mI & ":C" & mI).Select
'      With .Selection.Borders(xlTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'      End With
'
'      .Range("F" & mI & ":G" & mI).Select
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
'
'Private Sub preparaArrayPartes(ByRef pvPartes_OT() As Double)
'
'   Dim mJ As Integer
'   Dim cantPartes As Integer
'
'   cantPartes = FlexPartAsignados.Rows - 2
'   If cantPartes > 0 Then
'      ReDim pvPartes_OT(0 To cantPartes - 1) As Double
'
'      For mJ = 2 To FlexPartAsignados.Rows - 1
'         pvPartes_OT(mJ - 2) = FlexPartAsignados.TextMatrix(mJ, 1)
'      Next
'   Else
'      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
'      ReDim pvPartes_OT(0)
'      pvPartes_OT(0) = 0
'   End If
'End Sub
'
'Private Sub preparaArrayVehiculos(ByRef pvVehiculos_OT() As String)
'   Dim mJ As Integer
'   Dim cantVehiculos As Integer
'
'   cantVehiculos = FlexVehAsign.Rows - 2
'   If cantVehiculos > 0 Then
'      ReDim pvVehiculos_OT(0 To cantVehiculos - 1) As String
'
'      For mJ = 2 To FlexVehAsign.Rows - 1
'         pvVehiculos_OT(mJ - 2) = FlexVehAsign.TextMatrix(mJ, 2)
'      Next
'   Else
'      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
'      ReDim pvVehiculos_OT(0)
'      pvVehiculos_OT(0) = "00"
'   End If
'End Sub
'
'Private Sub preparaArrayVehiculosEsp(ByRef pvVehiculosEsp_OT() As String)
'   Dim mJ As Integer
'   Dim cantVehiculosEsp As Integer
'
'   cantVehiculosEsp = FlexVehEspAsign.Rows - 2
'   If cantVehiculosEsp > 0 Then
'      ReDim pvVehiculosEsp_OT(0 To cantVehiculosEsp - 1) As String
'
'      For mJ = 2 To FlexVehEspAsign.Rows - 1
'         pvVehiculosEsp_OT(mJ - 2) = FlexVehEspAsign.TextMatrix(mJ, 2)
'      Next
'   Else
'      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
'      ReDim pvVehiculosEsp_OT(0)
'      pvVehiculosEsp_OT(0) = "00"
'   End If
'End Sub
'
'Private Sub preparaArrayMO_Tecnicos(ByRef pvMO_Tecnicos_OT() As String)
'   Dim mJ As Integer
'   Dim cantMO_Tecnicos As Integer
'
'   cantMO_Tecnicos = FlexMoAsig.Rows - 2
'   If cantMO_Tecnicos > 0 Then
'      ReDim pvMO_Tecnicos_OT(0 To cantMO_Tecnicos - 1) As String
'
'      For mJ = 2 To FlexMoAsig.Rows - 1
'         pvMO_Tecnicos_OT(mJ - 2) = FlexMoAsig.TextMatrix(mJ, 2)
'      Next
'   Else
'      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
'      ReDim pvMO_Tecnicos_OT(0)
'      pvMO_Tecnicos_OT(0) = "00"
'   End If
'End Sub
'
'Private Sub preparaArraySubrubros(ByRef pvSubrubros_OT() As String)
'   Dim mJ As Integer
'   Dim cantSubrubros As Integer
'
'   cantSubrubros = FlexSubRubrosAsign.Rows - 2
'   If cantSubrubros > 0 Then
'      ReDim pvSubrubros_OT(0 To cantSubrubros - 1) As String
'
'      For mJ = 2 To FlexSubRubrosAsign.Rows - 1
'         pvSubrubros_OT(mJ - 2) = FlexSubRubrosAsign.TextMatrix(mJ, 4)
'      Next
'   Else
'      'Esta linea sirve como flag para avisar en el procedimiento xinsOT que no tiene registros y por ende no me de error de intervalos al usar Ubound y Lubound
'      ReDim pvSubrubros_OT(0)
'      pvSubrubros_OT(0) = "000000"
'   End If
'End Sub
'
'Private Sub CommandProd_Click(Index As Integer)
'
'End Sub
'
'Private Sub CommandSubRubro_Click(Index As Integer)
'   Dim sListaSubrubrosSeleccionados
'   Dim mJ As Integer
'   sListaSubrubrosSeleccionados = ""
'
'   If Index = 0 Then
'      If mRenglonSubRubroDispo > 0 Then
'         If Trim(FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1)) <> "" Then
'            FlexSubRubrosAsign.AddItem vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 2) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 3) & vbTab & FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 4)
'         End If
'
'         If FlexSubRubros.Rows > 2 Then
'            FlexSubRubros.RemoveItem mRenglonSubRubroDispo
'
'            mRenglonSubRubroDispo = 0
'         Else
'            If Trim(FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1)) <> "" Then
'               FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 1) = ""
'               FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 2) = ""
'
'               mRenglonSubRubroDispo = 0
'            End If
'         End If
'      End If
'   Else
'      If FlexSubRubrosAsign.Rows > 2 And mRenglonSubRubroAsign > 1 Then
'
'         FlexSubRubrosAsign.RemoveItem (mRenglonSubRubroAsign)
'
'         If FlexSubRubrosAsign.Rows > 2 Then
'            For mJ = 2 To FlexSubRubrosAsign.Rows - 1
'               sListaSubrubrosSeleccionados = sListaSubrubrosSeleccionados & "'" & FlexSubRubrosAsign.TextMatrix(mJ, 4) & "',"
'            Next
'            sListaSubrubrosSeleccionados = Left(sListaSubrubrosSeleccionados, Len(sListaSubrubrosSeleccionados) - 1)
'        End If
'
'         mRenglonSubRubroDispo = 0
'
'         FlexSubRubros.Clear
'         'Elimino los registros  de la grilla superior
'         For mJ = FlexSubRubros.Rows To 3 Step -1
'            FlexSubRubros.RemoveItem mJ
'         Next
'
'         With FlexSubRubros
'            .TextMatrix(0, 1) = "Técnico"
'            .TextMatrix(0, 2) = "Codigo"
'
'            .RowHeight(1) = 0
'         End With
'
'         If FlexSubRubrosAsign.Rows > 2 Then
'            Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
'               "  From " & _
'               " Rubros R " & _
'               "  Inner Join " & _
'               " SubRubros S ON S.CodRubro = R.Codigo " & _
'               " WHERE S.Codigo NOT IN (" & sListaSubrubrosSeleccionados & ")" & _
'               " AND R.Codigo ='" & Right(Combo1.Text, 8) & "';")
'         Else
'            Set mRec = mObj.oEjecutarSelect("SELECT R.Codigo As CodRubro,  R.Descripcion AS RubroDesc,  S.Codigo As CodSubrubro,  S.Descripcion AS SubRubroDesc " & _
'               "  From " & _
'               " Rubros R " & _
'               "  Inner Join " & _
'               " SubRubros S ON S.CodRubro = R.Codigo" & _
'               " WHERE R.Codigo ='" & Right(Combo1.Text, 8) & "';")
'         End If
'
'         If Not mRec.EOF Then
'            mJ = 1
'            Do While Not mRec.EOF
'               mJ = mJ + 1
'
'               With FlexSubRubros
'                  .AddItem ""
'                  .TextMatrix(mJ, 1) = mRec!RubroDesc
'                  .TextMatrix(mJ, 2) = mRec!SubRubroDesc
'                  .TextMatrix(mJ, 3) = mRec!CodRubro
'                  .TextMatrix(mJ, 4) = mRec!CodSubrubro
'               End With
'
'               mRec.MoveNext
'            Loop
'         End If
'         mRec.Close
'
'      End If
'      mRenglonSubRubroAsign = 0
'   End If
'End Sub
'
'
'
'Private Sub FlexMoAsig_Click()
'   Dim mI As Integer
'   Dim resultado As String
'
'   If FlexMoAsig.MouseCol = 0 And FlexMoAsig.MouseRow > 0 Then
'      If mRenglonMoAsign <> 0 Then
'         FlexMoAsig.Row = mRenglonMoAsign
'         For mI = 1 To FlexMoAsig.Cols - 1
'            FlexMoAsig.Col = mI
'            FlexMoAsig.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonMoAsign = FlexMoAsig.MouseRow
'
'      FlexMoAsig.Row = mRenglonMoAsign
'      For mI = 1 To FlexMoAsig.Cols - 1
'         FlexMoAsig.Col = mI
'         FlexMoAsig.CellBackColor = &H80000003
'      Next
'
'      If mRenglonMoAsign > 1 Then
'         If FlexMoAsig.Rows > 3 Then
'            mCodMO = FlexMoAsig.TextMatrix(mRenglonMoAsign, 1)
'            resultado = MsgBox(" ¿ Desea quitar al Técnico  " & mCodMO & " de esta Orden de Trabajo ?", vbOKCancel, "Quitar Técnico de OT")
'
'            If resultado = vbOK Then
'
'               If FlexMoAsig.Rows > 2 Then
'                  FlexMoAsig.RemoveItem mRenglonMoAsign
'
'                  mRenglonMoAsign = 0
'               Else
'                  If Trim(FlexMoAsig.TextMatrix(mRenglonMoAsign, 1)) <> "" Then
'                     FlexMoAsig.TextMatrix(mRenglonMoAsign, 1) = ""
'                     FlexMoAsig.TextMatrix(mRenglonMoAsign, 2) = ""
'
'                     mRenglonMoAsign = 0
'                  End If
'               End If
'
'            End If
'         Else
'            MsgBox "No es posible quitar todos los Técnicos de una OT"
'         End If
'      End If
'   Else
'      FlexMoAsig.Row = mRenglonMoAsign
'      If FlexMoAsig.Row > 0 Then
'         For mI = 1 To FlexMoAsig.Cols - 1
'            FlexMoAsig.Col = mI
'            FlexMoAsig.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonMoAsign = 0
'   End If
'End Sub
'
'
'Private Sub FlexPartAsignados_Click()
'
'End Sub
'
'Private Sub FlexMoDispo_Click()
'
'End Sub
'
'Private Sub FlexPartes_Click()
'   Dim mI As Integer
'   Dim resultado As String
'
'   If FlexPartes.MouseCol = 0 And FlexPartes.MouseRow > 0 Then
'      If mRenglonPartes <> 0 Then
'         FlexPartes.Row = mRenglonPartes
'         For mI = 1 To FlexPartes.Cols - 1
'            FlexPartes.Col = mI
'            FlexPartes.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonPartes = FlexPartes.MouseRow
'
'      FlexPartes.Row = mRenglonPartes
'      For mI = 1 To FlexPartes.Cols - 1
'         FlexPartes.Col = mI
'         FlexPartes.CellBackColor = &H80000003
'      Next
'
'      If mRenglonPartes > 1 Then
'         If FlexPartes.Rows > 3 Then
'            mCodParte = FlexPartes.TextMatrix(mRenglonPartes, 1)
'            resultado = MsgBox(" ¿ Desea eliminar el Parte número " & mCodParte & " de esta Orden de Trabajo ?", vbOKCancel, "Eliminar Parte de OT")
'
'            If resultado = vbOK Then
'
'               If FlexPartes.Rows > 2 Then
'                  FlexPartes.RemoveItem mRenglonPartes
'
'                  mRenglonPartes = 0
'               Else
'                  If Trim(FlexPartes.TextMatrix(mRenglonPartes, 1)) <> "" Then
'                     FlexPartes.TextMatrix(mRenglonPartes, 1) = ""
'                     FlexPartes.TextMatrix(mRenglonPartes, 2) = ""
'
'                     mRenglonPartes = 0
'                  End If
'               End If
'
'            End If
'         Else
'            MsgBox "No es posible eliminar todos los partes de una OT"
'         End If
'       End If
'   Else
'      FlexPartes.Row = mRenglonPartes
'      If FlexPartes.Row <> 0 Then
'         For mI = 1 To FlexPartes.Cols - 1
'            FlexPartes.Col = mI
'            FlexPartes.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonPartes = 0
'   End If
'End Sub
'
'Private Sub FlexProduct_Click()
'   Dim mI As Integer
'
'   If FlexProduct.MouseRow > 0 Then
'
'
'
'      'En este caso 3 es la columna que seria editable
'      If FlexProduct.Col = 3 And FlexProduct.Row <> 1 Then
'         Text2.Text = FlexProduct.Text
'         Text2.Width = FlexProduct.ColWidth(FlexProduct.Col)
'         Text2.Left = FlexProduct.ColPos(FlexProduct.Col) + FlexProduct.Left + 30 'el valor treina termina de acomodar el textbox en la celda
'         Text2.Top = FlexProduct.Top + FlexProduct.RowPos(FlexProduct.Row)
'         Text2.Visible = True
'         Text2.SetFocus
'         FlexProduct.ScrollBars = flexScrollBarNone
'      Else
'         Text2.Visible = False
'         FlexProduct.ScrollBars = flexScrollBarVertical
'      End If
'
'      filaAnt = FlexProduct.Row
'      columnAnt = FlexProduct.Col
'
'
'      If mRenglonProdDispo <> 0 Then
'         FlexProduct.Row = mRenglonProdDispo
'         For mI = 1 To FlexProduct.Cols - 1
'            FlexProduct.Col = mI
'            FlexProduct.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonProdDispo = FlexProduct.MouseRow
'
'      FlexProduct.Row = mRenglonProdDispo
'      For mI = 1 To FlexProduct.Cols - 1
'         FlexProduct.Col = mI
'         FlexProduct.CellBackColor = &H80000003
'      Next
'
'      If mRenglonProdDispo > 1 Then
'          mCodProducto = FlexProduct.TextMatrix(mRenglonProdDispo, 4)
'      End If
'   Else
'      FlexProduct.Row = mRenglonProdDispo
'      If FlexProduct.Row > 0 Then
'         For mI = 1 To FlexProduct.Cols - 1
'            FlexProduct.Col = mI
'            FlexProduct.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonProdDispo = 0
'   End If
'
'End Sub
'
'Private Sub FlexSubRubros_Click()
'   Dim mI As Integer
'
'   If FlexSubRubros.MouseRow > 0 Then
'
'      If mRenglonSubRubroDispo <> 0 Then
'         FlexSubRubros.Row = mRenglonSubRubroDispo
'         For mI = 1 To FlexSubRubros.Cols - 1
'            FlexSubRubros.Col = mI
'            FlexSubRubros.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonSubRubroDispo = FlexSubRubros.MouseRow
'
'      FlexSubRubros.Row = mRenglonSubRubroDispo
'      For mI = 1 To FlexSubRubros.Cols - 1
'         FlexSubRubros.Col = mI
'         FlexSubRubros.CellBackColor = &H80000003
'      Next
'
'      If mRenglonSubRubroDispo > 1 Then
'          mCodSubrubro = FlexSubRubros.TextMatrix(mRenglonSubRubroDispo, 4)
'      End If
'   Else
'      FlexSubRubros.Row = mRenglonSubRubroDispo
'      If FlexSubRubros.Row > 0 Then
'         For mI = 1 To FlexSubRubros.Cols - 1
'            FlexSubRubros.Col = mI
'            FlexSubRubros.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonSubRubroDispo = 0
'   End If
'End Sub
'
'Private Sub FlexSubRubrosAsign_Click()
'   Dim mI As Integer
'   If FlexSubRubrosAsign.MouseRow > 0 Then
'
'      If mRenglonSubRubroAsign <> 0 Then
'         FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
'         For mI = 1 To FlexSubRubrosAsign.Cols - 1
'            FlexSubRubrosAsign.Col = mI
'            FlexSubRubrosAsign.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonSubRubroAsign = FlexSubRubrosAsign.MouseRow
'
'      FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
'      For mI = 1 To FlexSubRubrosAsign.Cols - 1
'         FlexSubRubrosAsign.Col = mI
'         FlexSubRubrosAsign.CellBackColor = &H80000003
'      Next
'
'      If mRenglonSubRubroAsign > 1 Then
'          mCodSubrubro = FlexSubRubrosAsign.TextMatrix(mRenglonSubRubroAsign, 4)
'      End If
'   Else
'      FlexSubRubrosAsign.Row = mRenglonSubRubroAsign
'      If FlexSubRubrosAsign.Row > 0 Then
'         For mI = 1 To FlexSubRubrosAsign.Cols - 1
'            FlexSubRubrosAsign.Col = mI
'            FlexSubRubrosAsign.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonSubRubroAsign = 0
'   End If
'End Sub
'
'Private Sub FlexVehAsign_Click()
'   Dim mI As Integer
'   Dim mColVehAsign As Integer
'   Dim resultado As String
'
'   mColVehAsign = FlexVehAsign.Col
'
'   If FlexVehAsign.MouseRow > 0 Then
'       'En este caso 3 es la columna que seria editable
'      If FlexVehAsign.Col = 3 And FlexVehAsign.Row <> 1 Then
'         Text4.Text = FlexVehAsign.Text
'         Text4.Width = FlexVehAsign.ColWidth(FlexVehAsign.Col)
'         Text4.Left = FlexVehAsign.ColPos(FlexVehAsign.Col) + FlexVehAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
'         Text4.Top = FlexVehAsign.Top + FlexVehAsign.RowPos(FlexVehAsign.Row)
'         Text4.Visible = True
'         Text4.SetFocus
'         FlexVehAsign.ScrollBars = flexScrollBarNone
'      Else
'         Text4.Visible = False
'         If FlexVehAsign.Col <> 4 Then
'            FlexVehAsign.ScrollBars = flexScrollBarVertical
'         End If
'      End If
'      filaAntVehAsign = FlexVehAsign.Row
'      columnAntVehAsign = FlexVehAsign.Col
'
'       'En este caso 4 es la columna que seria editable
'      If FlexVehAsign.Col = 4 And FlexVehAsign.Row <> 1 Then
'         Text5.Text = FlexVehAsign.Text
'         Text5.Width = FlexVehAsign.ColWidth(FlexVehAsign.Col)
'         Text5.Left = FlexVehAsign.ColPos(FlexVehAsign.Col) + FlexVehAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
'         Text5.Top = FlexVehAsign.Top + FlexVehAsign.RowPos(FlexVehAsign.Row)
'         Text5.Visible = True
'         Text5.SetFocus
'         FlexVehAsign.ScrollBars = flexScrollBarNone
'      Else
'         Text5.Visible = False
'         If FlexVehAsign.Col <> 3 Then
'            FlexVehAsign.ScrollBars = flexScrollBarVertical
'         End If
'      End If
'      filaAntVehAsignKmFinal = FlexVehAsign.Row
'      columnAntVehAsignKmFinal = FlexVehAsign.Col
'
'      If mRenglonVehAsign <> 0 Then
'         FlexVehAsign.Row = mRenglonVehAsign
'         For mI = 1 To FlexVehAsign.Cols - 1
'            FlexVehAsign.Col = mI
'            FlexVehAsign.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonVehAsign = FlexVehAsign.MouseRow
'
'      FlexVehAsign.Row = mRenglonVehAsign
'      For mI = 1 To FlexVehAsign.Cols - 1
'         FlexVehAsign.Col = mI
'         FlexVehAsign.CellBackColor = &H80000003
'      Next
'
'      If mRenglonVehAsign > 1 Then
'         If mColVehAsign = 1 Then
'           mCodVeh = FlexVehAsign.TextMatrix(mRenglonVehAsign, 2)
'           resultado = MsgBox(" ¿ Desea eliminar el Vehiculo  " & mCodVeh & " de esta Orden de Trabajo ?", vbOKCancel, "Eliminar Vehículo de OT")
'           If resultado = vbOK Then
'              If FlexVehAsign.Rows > 2 Then
'                 FlexVehAsign.RemoveItem mRenglonVehAsign
'                 mRenglonVehAsign = 0
'              Else
'                 If Trim(FlexVehAsign.TextMatrix(mRenglonVehAsign, 1)) <> "" Then
'                    FlexVehAsign.TextMatrix(mRenglonVehAsign, 1) = ""
'                    FlexVehAsign.TextMatrix(mRenglonVehAsign, 2) = ""
'                    mRenglonVehAsign = 0
'                 End If
'              End If
'           End If
'         End If
'      End If
'   Else
'      FlexVehAsign.Row = mRenglonVehAsign
'      If FlexVehAsign.Row > 0 Then
'         For mI = 1 To FlexVehAsign.Cols - 1
'            FlexVehAsign.Col = mI
'            FlexVehAsign.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonVehAsign = 0
'   End If
'End Sub
'
'Private Sub Text4_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 46 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   End If
'
'   If KeyAscii = 13 Then
'      FlexVehAsign.TextMatrix(filaAntVehAsign, columnAntVehAsign) = Text4.Text
'      Text4.Visible = False
'      FlexVehAsign.ScrollBars = flexScrollBarVertical
'   End If
'End Sub
'
'Private Sub Text4_LostFocus()
'   If FlexVehAsign.Col <> columnAntVehAsign Or FlexVehAsign.Row <> filaAntVehAsign Then
'      'En este caso 3 es la columna que seria editable
'      If columnAntVehAsign = 3 Then
'         FlexVehAsign.TextMatrix(filaAntVehAsign, columnAntVehAsign) = Text4.Text
'      End If
'   End If
'End Sub
'
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 46 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   End If
'   If KeyAscii = 13 Then
'      FlexVehAsign.TextMatrix(filaAntVehAsignKmFinal, columnAntVehAsignKmFinal) = Text5.Text
'      Text5.Visible = False
'      FlexVehAsign.ScrollBars = flexScrollBarVertical
'   End If
'End Sub
'
'Private Sub Text5_LostFocus()
'   If FlexVehAsign.Col <> columnAntVehAsignKmFinal Or FlexVehAsign.Row <> filaAntVehAsignKmFinal Then
'      'En este caso 4 es la columna que seria editable
'      If columnAntVehAsignKmFinal = 4 Then
'         FlexVehAsign.TextMatrix(filaAntVehAsignKmFinal, columnAntVehAsignKmFinal) = Text5.Text
'      End If
'   End If
'End Sub
'
'
'
'
'
'
'
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 46 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   End If
'   If KeyAscii = 13 Then
'      FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmInicio, columnAntVehEspAsignKmInicio) = Text6.Text
'      Text6.Visible = False
'      FlexVehEspAsign.ScrollBars = flexScrollBarVertical
'   End If
'End Sub
'
'Private Sub Text6_LostFocus()
'   If FlexVehEspAsign.Col <> columnAntVehEspAsignKmInicio Or FlexVehEspAsign.Row <> filaAntVehEspAsignKmInicio Then
'      'En este caso 2 es la columna que seria editable
'      If columnAntVehEspAsignKmInicio = 2 Then
'         FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmInicio, columnAntVehEspAsignKmInicio) = Text6.Text
'      End If
'   End If
'End Sub
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
'Private Sub Text7_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 46 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   End If
'   If KeyAscii = 13 Then
'      FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmFinal, columnAntVehEspAsignKmFinal) = Text7.Text
'      Text7.Visible = False
'      FlexVehEspAsign.ScrollBars = flexScrollBarVertical
'   End If
'End Sub
'
'Private Sub Text7_LostFocus()
'   If FlexVehEspAsign.Col <> columnAntVehEspAsignKmFinal Or FlexVehEspAsign.Row <> filaAntVehEspAsignKmFinal Then
'      'En este caso 3 es la columna que seria editable
'      If columnAntVehEspAsignKmFinal = 3 Then
'         FlexVehEspAsign.TextMatrix(filaAntVehEspAsignKmFinal, columnAntVehEspAsignKmFinal) = Text7.Text
'      End If
'   End If
'End Sub
'
'Private Sub FlexVehEspAsign_Click()
'   Dim mI As Integer
'
'   If FlexVehEspAsign.MouseRow > 0 Then
'      'En este caso 2 es la columna que seria editable
'      If FlexVehEspAsign.Col = 2 And FlexVehEspAsign.Row <> 1 Then
'         Text6.Text = FlexVehEspAsign.Text
'         Text6.Width = FlexVehEspAsign.ColWidth(FlexVehEspAsign.Col)
'         Text6.Left = FlexVehEspAsign.ColPos(FlexVehEspAsign.Col) + FlexVehEspAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
'         Text6.Top = FlexVehEspAsign.Top + FlexVehEspAsign.RowPos(FlexVehEspAsign.Row)
'         Text6.Visible = True
'         Text6.SetFocus
'         FlexVehEspAsign.ScrollBars = flexScrollBarNone
'      Else
'         Text6.Visible = False
'         If FlexVehEspAsign.Col <> 3 Then
'            FlexVehEspAsign.ScrollBars = flexScrollBarVertical
'         End If
'      End If
'      filaAntVehEspAsignKmInicio = FlexVehEspAsign.Row
'      columnAntVehEspAsignKmInicio = FlexVehEspAsign.Col
'
'       'En este caso 3 es la columna que seria editable
'      If FlexVehEspAsign.Col = 3 And FlexVehEspAsign.Row <> 1 Then
'         Text7.Text = FlexVehEspAsign.Text
'         Text7.Width = FlexVehEspAsign.ColWidth(FlexVehEspAsign.Col)
'         Text7.Left = FlexVehEspAsign.ColPos(FlexVehEspAsign.Col) + FlexVehEspAsign.Left + 30 'el valor treina termina de acomodar el textbox en la celda
'         Text7.Top = FlexVehEspAsign.Top + FlexVehEspAsign.RowPos(FlexVehEspAsign.Row)
'         Text7.Visible = True
'         Text7.SetFocus
'         FlexVehEspAsign.ScrollBars = flexScrollBarNone
'      Else
'         Text7.Visible = False
'         If FlexVehEspAsign.Col <> 2 Then
'            FlexVehEspAsign.ScrollBars = flexScrollBarVertical
'         End If
'      End If
'      filaAntVehEspAsignKmFinal = FlexVehEspAsign.Row
'      columnAntVehEspAsignKmFinal = FlexVehEspAsign.Col
'
'      If mRenglonVehEspAsign <> 0 Then
'         FlexVehEspAsign.Row = mRenglonVehEspAsign
'         For mI = 1 To FlexVehEspAsign.Cols - 1
'            FlexVehEspAsign.Col = mI
'            FlexVehEspAsign.CellBackColor = vbWhite
'         Next
'      End If
'
'      mRenglonVehEspAsign = FlexVehEspAsign.MouseRow
'
'      FlexVehEspAsign.Row = mRenglonVehEspAsign
'      For mI = 1 To FlexVehEspAsign.Cols - 1
'         FlexVehEspAsign.Col = mI
'         FlexVehEspAsign.CellBackColor = &H80000003
'      Next
'
'      If mRenglonVehEspAsign > 1 Then
'          mCodVeh = FlexVehEspAsign.TextMatrix(mRenglonVehEspAsign, 2)
'      End If
'   Else
'      FlexVehEspAsign.Row = mRenglonVehEspAsign
'      If FlexVehEspAsign.Row > 0 Then
'         For mI = 1 To FlexVehEspAsign.Cols - 1
'            FlexVehEspAsign.Col = mI
'            FlexVehEspAsign.CellBackColor = vbWhite
'         Next
'      End If
'      mRenglonVehEspAsign = 0
'   End If
'End Sub
'
'Private Sub Form_Load()
'
'   Set mRec = mObj.oEjecutarSelect("SELECT CONVERT( CONCAT(LPAD(IdOT,10,'0'),' - ',Date_Format(Fecha,'%d/%m/%Y')),char(23)) AS OT_Fecha, O.Fecha " & _
'                           " FROM MantElect.OT_H O " & _
'                           " ORDER BY IdOT DESC; ")
'
'   Do While Not mRec.EOF
'      Combo3.AddItem mRec!OT_Fecha
'      mRec.MoveNext
'   Loop
'   mRec.Close
'
'
'   Me.Width = 21270
'   Me.Height = 13950
'   sAlinearForm Me
'
'   Frame1(0).Visible = True
'   Frame1(1).Visible = False
'   Frame1(2).Visible = False
'   Frame1(3).Visible = False
'   Frame1(4).Visible = False
'
'   initPartes
'   initManoObra
'   initVehiculos
'   initVehiculosEspecial
'   initRubros_SubRubros
'   initMateriales
'
'End Sub
'
'Private Sub initPartes()
'   mRenglonPartes = 0
'
'   With FlexPartes
'      .ColWidth(0) = 200
'      .ColWidth(1) = 500
'      .ColWidth(2) = 2000
'      .ColWidth(3) = 3000
'      .ColWidth(4) = 11900
'      .ColWidth(5) = 750
'
'      .ColWidth(6) = 0
'
'
'      .TextMatrix(0, 1) = "Parte"
'      .TextMatrix(0, 2) = "Fecha Solicitud"
'      .TextMatrix(0, 3) = "Lugar"
'      .TextMatrix(0, 4) = "Descripcion de la Solicitud"
'      .TextMatrix(0, 5) = "Prioridad"
'
'      .TextMatrix(0, 6) = "Sector Aire"
'
'      .RowHeight(1) = 0
'   End With
'End Sub
'
'Private Sub initManoObra()
'   Dim mI As Integer
'
'   mRenglonMoAsign = 0
'   With FlexMoAsig
'      .ColWidth(0) = 200
'      .ColWidth(1) = 5000
'      .ColWidth(2) = 0
'
'      .TextMatrix(0, 1) = "Técnico"
'      .TextMatrix(0, 2) = "Codigo"
'
'      .RowHeight(1) = 0
'   End With
'End Sub
'
'Private Sub initVehiculos()
'   mRenglonVehAsign = 0
'
'   filaAntVehAsign = 0
'   columnAntVehAsign = 0
'   Text4.Visible = False
'
'   filaAntVehAsignKmFinal = 0
'   columnAntVehAsignKmFinal = 0
'   Text5.Visible = False
'
'   With FlexVehAsign
'      .ColWidth(0) = 200
'      .ColWidth(1) = 0
'      .ColWidth(2) = 3000
'      .ColWidth(3) = 1500
'      .ColWidth(4) = 1500
'      .ColWidth(5) = 0
'
'      .TextMatrix(0, 2) = "Vehículo"
'      .TextMatrix(0, 3) = "Km. Inicial"
'      .TextMatrix(0, 4) = "Km Final"
'      .TextMatrix(0, 5) = "Codigo"
'
'      .RowHeight(1) = 0
'   End With
'
'End Sub
'
'Private Sub initVehiculosEspecial()
'   mRenglonVehEspAsign = 0
'
'   filaAntVehEspAsignKmInicio = 0
'   columnAntVehEspAsignKmInicio = 0
'   Text6.Visible = False
'
'   filaAntVehEspAsignKmFinal = 0
'   columnAntVehEspAsignKmFinal = 0
'   Text7.Visible = False
'
'   With FlexVehEspAsign
'      .ColWidth(0) = 200
'      .ColWidth(1) = 3000
'      .ColWidth(2) = 1500
'      .ColWidth(3) = 1500
'      .ColWidth(4) = 0
'
'
'      .TextMatrix(0, 1) = "Vehículo especial"
'      .TextMatrix(0, 2) = "Km. Inicial"
'      .TextMatrix(0, 3) = "Km Final"
'      .TextMatrix(0, 4) = "Codigo"
'
'      .RowHeight(1) = 0
'   End With
'End Sub
'
'Private Sub initRubros_SubRubros()
'   Dim mI As Integer
'
'
'
'     For mI = FlexSubRubros.Rows To 3 Step -1
'            FlexSubRubros.RemoveItem mI
'         Next
'
'   Set mRec = mObj.oEjecutarSelect("SELECT Codigo, Descripcion FROM Rubros Where FechaBaja IS NULL;")
'
'
'   Do While Not mRec.EOF
'      'Combo1.AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
'      Combo1.AddItem "" & mRec!descripcion & Space(50) & mRec!Codigo
'      mRec.MoveNext
'   Loop
'   mRec.Close
'
'
'   mRenglonSubRubroDispo = 0
'
'   With FlexSubRubros
'      .ColWidth(0) = 200
'      .ColWidth(1) = 6000
'      .ColWidth(2) = 10750
'      .ColWidth(3) = 0
'      .ColWidth(4) = 0
'
'      .TextMatrix(0, 1) = "Rubro"
'      .TextMatrix(0, 2) = "SubRubro"
'      .TextMatrix(0, 3) = "CodRubro"
'      .TextMatrix(0, 4) = "CodSubRubro"
'
'      .RowHeight(1) = 0
'   End With
'
'   With FlexSubRubrosAsign
'      .ColWidth(0) = 200
'      .ColWidth(1) = 6000
'      .ColWidth(2) = 10750
'      .ColWidth(3) = 0
'      .ColWidth(4) = 0
'
'      .TextMatrix(0, 1) = "Rubro"
'      .TextMatrix(0, 2) = "SubRubro"
'      .TextMatrix(0, 3) = "CodRubro"
'      .TextMatrix(0, 4) = "CodSubRubro"
'
'      .RowHeight(1) = 0
'   End With
'
'
'End Sub
'
'
'
'Private Sub initMateriales()
'   filaAnt = 0
'   columnAnt = 0
'   Text2.Visible = False
'
'   With FlexProduct
'      .ColWidth(0) = 200
'      .ColWidth(1) = 1250
'      .ColWidth(2) = 10700
'      .ColWidth(3) = 1500
'      .ColWidth(4) = 1500
'      .ColWidth(5) = 1900
'      .ColWidth(6) = 0
'      .ColWidth(7) = 0
'
'      .TextMatrix(0, 1) = "Cód.Sap"
'      .TextMatrix(0, 2) = "Producto"
'      .TextMatrix(0, 3) = "Cantidad"
'      .TextMatrix(0, 4) = "Stock"
'      .TextMatrix(0, 5) = "Unid.Medida"
'      .TextMatrix(0, 6) = "Cód. Producto"
'      .TextMatrix(0, 7) = "Cód. Ubicacion"
'
'      .RowHeight(1) = 0
'   End With
'End Sub
'
'
'
'Private Sub TabStrip1_Click()
'   Dim i As Integer
'   Dim j As Integer
'
'    i = TabStrip1.SelectedItem.Index
'
'
'   For j = 1 To TabStrip1.Tabs.Count
'      If j = i Then
'         Frame1(j - 1).Visible = True
'
'      Else
'         Frame1(j - 1).Visible = False
'      End If
'   Next
'End Sub
'
'
'Private Function fValidaAsignaMateriales() As Boolean
''   Dim mRet As Boolean
''   Dim mMensajeError As String
''   Dim mJ As Integer
''
''   mRet = True
''
''   If mRenglonProdDispo = 0 Then
''      mRet = False
''      mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
''   End If
''
''   If mRet Then
''      If mRenglonProdDispo <> 0 And FlexProduct.TextMatrix(mRenglonProdDispo, 1) = "" Then
''         mRet = False
''         mMensajeError = "Debe seleccionar un Producto de la Grilla superior"
''      End If
''   End If
''
''   'Valido si el Producto y la Ubicacion seleccionados coinciden con un Egreso ya registrado
''   If mRet Then
''      For mJ = 2 To FlexEgreso.Rows - 1
''         If FlexEgreso.TextMatrix(mJ, 6) = FlexProduct.TextMatrix(mRenglonProdDispo, 6) And FlexEgreso.TextMatrix(mJ, 7) = FlexProduct.TextMatrix(mRenglonProdDispo, 7) Then
''            mRet = False
''            mMensajeError = "El Producto y la Ubicación elegidos ya han sido seleccionados"
''            mJ = 999
''         End If
''      Next
''   End If
''
''   If Not mRet Then
''         MsgBox mMensajeError, vbCritical, "Atención"
''   End If
''   fValidaAsignaMateriales = mRet
'End Function
'
'Private Sub Text2_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 46 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   End If
'
'   If KeyAscii = 13 Then
'      FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
'      Text2.Visible = False
'      FlexProduct.ScrollBars = flexScrollBarVertical
'   End If
'End Sub
'
'Private Sub Text2_LostFocus()
'   If FlexProduct.Col <> columnAnt Or FlexProduct.Row <> filaAnt Then
'      'En este caso 3 es la columna que seria editable
'      If columnAnt = 3 Then
'         FlexProduct.TextMatrix(filaAnt, columnAnt) = Text2.Text
'      End If
'   End If
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'Set mObj = Nothing
'Set mRec = Nothing
'ShowMenu 47, True, False
'End Sub
'
'Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
'   KeyAscii = fDateKeyPress(Text3(Index), KeyAscii)
'End Sub
