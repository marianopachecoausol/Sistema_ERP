VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RAcc5_frm 
   Caption         =   "Form8"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   9075
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   600
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   5160
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   8655
         Begin VB.CommandButton Command1 
            Caption         =   "&Agregar"
            Height          =   375
            Index           =   4
            Left            =   2880
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   5
            Left            =   3840
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   6
            Left            =   4800
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Buscar"
            Height          =   375
            Index           =   7
            Left            =   5760
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   375
            Index           =   8
            Left            =   6720
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Volver"
            Height          =   375
            Index           =   9
            Left            =   7680
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   3
            Left            =   1920
            Picture         =   "Form5.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   2
            Left            =   1320
            Picture         =   "Form5.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   1
            Left            =   720
            Picture         =   "Form5.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "Form5.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5160
         TabIndex        =   5
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         TabIndex        =   4
         Top             =   1320
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Vehículo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   3
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
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
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "RAcc5_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObj As New clRAcc
Dim mObjAcc As New clAccess
Dim mRec As ADODB.Recordset
Dim mRec1 As ADODB.Recordset
Dim ManText As String
Dim manVehi As String
Public mFlagRAccd As Boolean
Dim mI As Integer

Private Sub Form_Load()
   Me.Height = 4000
   Me.Width = 9200
   Me.Caption = "Administración de la Tabla Marcas"
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   Label1.AutoSize = True
   Label1.Caption = "Tabla de Marcas"
   Label1.Left = (Me.Width - Label1.Width) / 2
   For mI = 0 To Text1.UBound
      Text1(mI).Text = ""
   Next
   Combo1.Enabled = False
   Text1(0).Enabled = False
   Text1(1).Enabled = False
   Set mRec = mObj.oTabla("TipoVehiculo", "")
   Do While Not mRec.EOF
      Combo1.AddItem "" & mRec!CodTipoVehic & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   Set mRec = mObj.oTablaDina("Marca", "order by codtipovehic, codmarca")
   If Not mRec.EOF Then
      sCargarText mRec!CodTipoVehic, mRec!CodMarca
   Else
      For mI = 5 To Command1.UBound - 1
         Command1(mI).Enabled = False
      Next
   End If
'   If RAcc2_frm.mAgregar Then
'      RAcc2_frm.Enabled = False
'   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec1 = Nothing
   Set mRec = Nothing
   If mFlagRAccd Then
      RAcc1beta.Enabled = True
   Else
      ShowMenu 2, True, False
   End If
'   If RAcc2_frm.mAgregar Then
'      RAcc2_frm.mAgregar = False
'      RAcc2_frm.Enabled = True
'   End If
End Sub

Private Sub Combo1_Click()
Dim mRec1 As New ADODB.Recordset

   If Command1(4).Caption = "&Grabar" Then
      Set mRec1 = mObj.oTablaDina("Marca", "where codtipovehic='" & Trim(Left(Combo1.Text, 2)) & "' order by codmarca")
      If Not mRec1.EOF Then
         mRec1.MoveLast
         Text1(0).Text = Format((Val(mRec1!CodMarca) + 1), "00")
         Text1(1).Enabled = True
         Text1(1).Text = ""
      Else
         Text1(0).Text = "01"
         Text1(1).Enabled = True
         Text1(1).Text = ""
         Text1(1).SetFocus
      End If
      mRec1.Close
   End If
   Set mRec1 = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim mAuxi
 Dim mSalir As Boolean
 Dim mBorrar As Boolean
 Dim mCodigo As String
 Dim CodMarca As Integer
   Select Case Index
      Case 0
         If Not mRec.EOF Then
            mRec.MoveFirst
            sCargarText mRec!CodTipoVehic, mRec!CodMarca
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 1
         If Not mRec.EOF Then
            mRec.MovePrevious
            If mRec.BOF Then
               MsgBox "No hay Registros Anteriores!!!", vbExclamation, sMessage
               mRec.MoveFirst
            Else
               sCargarText mRec!CodTipoVehic, mRec!CodMarca
            End If
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
   
      Case 2
         If Not mRec.EOF Then
            mRec.MoveNext
            If mRec.EOF Then
               MsgBox "No hay Registros Posteriores!!!", vbExclamation, sMessage
               mRec.MoveLast
            Else
               sCargarText mRec!CodTipoVehic, mRec!CodMarca
            End If
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 3
         If Not mRec.EOF Then
            mRec.MoveLast
            sCargarText mRec!CodTipoVehic, mRec!CodMarca
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
   
      Case 4
         If Command1(Index).Caption = "&Agregar" Then
            ManText = Text1(0).Text
            manVehi = Left(Combo1.Text, 2)
            For mI = 0 To Command1.Count - 2
               If mI <> Index Then
                  Command1(mI).Enabled = False
               End If
            Next
            Text1(1).Text = ""
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Set mRec1 = mObj.oTablaDina("Marca", "WHERE CodTipoVehic = '" & Left(Combo1.Text, 2) & "' ORDER BY CodMarca")
            If Not mRec1.EOF Then
               mRec1.MoveLast
               CodMarca = mRec1!CodMarca
               Text1(0).Text = Format((CodMarca + 1), "00")
               Text1(1).Enabled = True
               Text1(1).Text = ""
            Else
               Text1(0).Text = "01"
               Text1(1).Enabled = True
               Text1(1).Text = ""
               Text1(1).SetFocus
            End If
            mRec1.Close
            Combo1.Enabled = True
         Else
            If Combo1.ListIndex <> -1 And Text1(1).Text <> "" Then
               Set mRec1 = mObj.oTabla("Marca", "WHERE CodTipoVehic = '" & Left(Combo1.Text, 2) & "' AND  CodMarca = '" & Text1(0).Text & "'")
               If mRec1.EOF Then
                  mObj.xInsMarcas Left(Combo1.Text, 2), Text1(0).Text, Text1(1).Text
                  mRec.Requery
                  sCargarText Left(Combo1.Text, 2), Text1(0).Text
                  For mI = 0 To Command1.Count - 2
                     Command1(mI).Enabled = True
                  Next
                  Command1(Index).Caption = "&Agregar"
                  Command1(9).Caption = "&Volver"
                  For mI = 0 To Text1.UBound
                     Text1(mI).Enabled = False
                  Next
                  Combo1.Enabled = False
               Else
                  MsgBox "Código Existente!!!", vbExclamation, sMessage
               End If
               mRec1.Close
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, sMessage
            End If
         End If
   
      Case 5
         If Command1(Index).Caption = "&Modificar" Then
            manVehi = Left(Combo1.Text, 2)
            ManText = Text1(0).Text
            For mI = 0 To Command1.Count - 1
               Command1(mI).Enabled = False
            Next
            Text1(1).Enabled = True
            Text1(1).SetFocus
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Command1(Index).Enabled = True
            Command1(9).Enabled = True
         Else
            If Text1(1).Text <> "" Then
               mObj.xUpdMarcas Left(Combo1.Text, 2), Text1(0).Text, Trim(Text1(1).Text)
               mRec.Requery
               sCargarText Left(Combo1.Text, 2), Text1(0).Text
               For mI = 0 To Command1.Count - 1
                  Command1(mI).Enabled = True
               Next
               Command1(Index).Caption = "&Modificar"
               Command1(9).Caption = "&Volver"
               Text1(1).Enabled = False
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, "Atención"
            End If
         End If
      
      Case 6
         Set mRec1 = mObj.oTabla("VehiculosInvolucr", "WHERE CodTipoVehic = '" & Left(Combo1.Text, 2) & "' AND CodMarca = '" & Text1(0).Text & "'")
         If Not mRec1.EOF Then
            If MsgBox("¿Está seguro de Eliminar este Registro?", vbYesNo, "Atención") = vbYes Then
               mObj.xDelTabla "Marca", "WHERE CodTipoVehic = '" & Left(Combo1.Text, 2) & "' AND CodMarca = '" & Text1(0).Text & "'"
               mRec.Requery
               If Not mRec.EOF Then
                  sCargarText mRec!CodTipoVehic, mRec!CodMarca
               Else
                  Form_Load
               End If
            End If
         Else
            MsgBox "NO puede Eliminar este Registro!!! " & vbCrLf & "Existe En Vehículos Involucrados!!!", vbCritical, sMessage
         End If
         mRec1.Close
      
      Case 7
         If Command1(Index).Caption = "&Buscar" Then
            For mI = 0 To Command1.Count - 2
               Command1(mI).Enabled = False
            Next
            Text1(0).Enabled = True
            Command1(Index).Caption = "C&onfirmar"
            Command1(9).Caption = "&Cancelar"
            Command1(Index).Enabled = True
            Combo1.Enabled = True
            ManText = Text1(0).Text
            manVehi = Left(Combo1.Text, 2)
         Else
            mRec.MoveFirst
            mSalir = False
            Do While Not mRec.EOF And Not mSalir
               If mRec!CodTipoVehic = Left(Combo1.Text, 2) And mRec!CodMarca = Text1(0).Text Then
                  mSalir = True
               Else
                  mRec.MoveNext
               End If
            Loop
            If Not mSalir Then
               MsgBox "Registro Inexistente", vbExclamation, sMessage
               mRec.Requery
               mRec.MoveFirst
               Do While Not mRec.EOF And Not mSalir
                  If mRec!CodTipoVehic = manVehi And mRec!CodMarca = ManText Then
                     mSalir = True
                  Else
                     mRec.MoveNext
                  End If
               Loop
               sCargarText mRec!CodTipoVehic, mRec!CodMarca

            Else
               sCargarText Left(Combo1.Text, 2), Text1(0).Text
            End If
            For mI = 0 To Command1.Count - 1
               Command1(mI).Enabled = True
            Next
            Command1(Index).Caption = "&Buscar"
            Command1(9).Caption = "&Volver"
            Text1(0).Enabled = False
            Combo1.Enabled = False
         End If
         
      Case 8
         mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
         mData.Execute ("CREATE TABLE Auxi (Codigo1 TEXT, Codigo2 TEXT, Descripcion TEXT)")
         If MsgBox("¿Imprimir Solo El Tipo de Vehículo Seleccionado?", vbYesNo, sMessage) = vbYes Then
            Set mRec1 = mObj.oTabla("Marca", "WHERE CodTipoVehic = '" & Left(Combo1.Text, 2) & "' ORDER BY Descripcion ")
            CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: Marcas- Vehículo: " & Trim(Mid(Combo1.Text, 4, 20)) & "'"
         Else
            Set mRec1 = mObj.oTabla("Marca", "")
            CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: Marcas'"
         End If
         Do While Not mRec1.EOF
            mData.Execute ("INSERT INTO Auxi (Codigo1,Codigo2,Descripcion) VALUES ('" & mRec1!CodTipoVehic & "','" & mRec1!CodMarca & "','" & mRec1!descripcion & "')")
            mRec1.MoveNext
         Loop
         mRec1.Close
         CrystalReport1.ReportFileName = App.Path & "\RegAccidentes\rep05.rpt"
         CrystalReport1.WindowTitle = "Tabla Marcas"
         CrystalReport1.DataFiles(0) = App.Path & "\RegAccidentes\FichaAccid.mdb"
         CrystalReport1.Action = 1
         CrystalReport1.Formulas(0) = ""
         
      Case 9
         If Command1(Index).Caption = "&Cancelar" Then
            Command1(Index).Caption = "&Volver"
            Command1(4).Caption = "&Agregar"
            Command1(5).Caption = "&Modificar"
            Command1(7).Caption = "&Buscar"
            For mI = 0 To Command1.Count - 1
               Command1(mI).Enabled = True
            Next
            For mI = 0 To Text1.Count - 1
               Text1(mI).Enabled = False
            Next
            Combo1.Enabled = False
            If Not mRec.EOF Then
               sCargarText Trim(manVehi), Trim(ManText)
            Else
               Form_Load
            End If
         Else
            Unload Me
         End If
   End Select
End Sub

Private Sub sCargarText(ByVal pCodVehic As String, ByVal pParam As String)
Dim mSalir As Boolean
Dim mI As Integer

mSalir = False
mRec.MoveFirst
Do While Not mRec.EOF And Not mSalir
   If pCodVehic <> mRec.Fields(0) Or pParam <> mRec.Fields(1) Then
      mRec.MoveNext
   Else
      mSalir = True
   End If
Loop
If mSalir Then
   For mI = 0 To Combo1.ListCount - 1
      Combo1.ListIndex = mI
      If Left(Combo1.Text, 2) = pCodVehic Then
         mI = 100
      End If
   Next
   Text1(0).Text = NVL(mRec.Fields(1), "")
   Text1(1).Text = NVL(mRec.Fields(2), "")
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 0 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   Else
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   End If
End Sub
