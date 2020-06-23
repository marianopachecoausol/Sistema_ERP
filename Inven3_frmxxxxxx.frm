VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Inven3_frmold 
   Caption         =   "Form8"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   14070
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13895
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   4160
         MaxLength       =   5
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2200
         Width           =   1000
      End
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
         Index           =   0
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2200
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2660
         MaxLength       =   90
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1380
         Width           =   8500
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1380
         Width           =   1000
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   13655
         Begin VB.CommandButton Command1 
            Caption         =   "&Agregar"
            Height          =   375
            Index           =   4
            Left            =   5380
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Modificarr"
            Height          =   375
            Index           =   5
            Left            =   6340
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   6
            Left            =   7300
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Buscar"
            Height          =   375
            Index           =   7
            Left            =   8260
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   375
            Index           =   8
            Left            =   9220
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Volver"
            Height          =   375
            Index           =   9
            Left            =   10180
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   3
            Left            =   1920
            Picture         =   "Inven3_frm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   2
            Left            =   1320
            Picture         =   "Inven3_frm.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   1
            Left            =   720
            Picture         =   "Inven3_frm.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "Inven3_frm.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código Sap"
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
         Left            =   4155
         TabIndex        =   20
         Top             =   1845
         Width           =   1080
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
         Left            =   2660
         TabIndex        =   5
         Top             =   1020
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
         Left            =   600
         TabIndex        =   4
         Top             =   1020
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Medida"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1845
         Width           =   1200
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
Attribute VB_Name = "Inven3_frmold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mData As Database
Dim mObj As New clInven
Dim mObjAcc As New clAccess
Dim mRec As ADODB.Recordset
Dim mRec1 As ADODB.Recordset
Dim ManText As String
Dim manVehi As String
Dim CodUMedidaText As String
Dim Reporte As String

Public mFlagRAccd As Boolean
Dim mi As Integer

Private Sub Form_Load()
   Me.Height = 4200
   Me.Width = 14190
   Me.Caption = "Productos"
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\Inventario\Inventario.mdb")
   Label1.AutoSize = True
   Label1.Caption = "Tabla de Productos"
   Label1.Left = (Me.Width - Label1.Width) / 2
   For mi = 0 To Text1.UBound
      Text1(mi).Text = ""
   Next
   Combo1(0).Enabled = False

   Text1(0).Enabled = False
   Text1(1).Enabled = False
   Text1(2).Enabled = False
   
   Set mRec = mObj.oTabla("UnidadMedida", "")
   Do While Not mRec.EOF
      Combo1(0).AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   
   Set mRec = mObj.oTablaDina("Producto", " where Fecha_Baja IS NULL order by Codigo")
   If Not mRec.EOF Then
      sCargarText mRec!CodUnidadMedida, mRec!Codigo
   Else
      For mi = 5 To Command1.UBound - 1
         Command1(mi).Enabled = False
      Next
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec1 = Nothing
   Set mRec = Nothing
      ShowMenu 12, True, False
End Sub

Private Sub Combo1_Click(Index As Integer)
'Dim mRec1 As New ADODB.Recordset

'   If Command1(4).Caption = "&Grabar" Then
'      Set mRec1 = mObj.oTablaDina("Marca", "where codtipovehic='" & Trim(Left(Combo1(0).Text, 2)) & "' order by codmarca")
'      If Not mRec1.EOF Then
'         mRec1.MoveLast
'         Text1(0).Text = Format((Val(mRec1!CodMarca) + 1), "00")
'         Text1(1).Enabled = True
'         Text1(1).Text = ""
'      Else
'         Text1(0).Text = "01"
'         Text1(1).Enabled = True
'         Text1(1).Text = ""
'         Text1(1).SetFocus
'      End If
'      mRec1.Close
'   End If
'   Set mRec1 = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
 Dim mAuxi
 Dim mSalir As Boolean
 Dim mBorrar As Boolean
 Dim mCodigo As String
 Dim CodMarca As Integer
 Dim CodProducto As Integer
   Select Case Index
      Case 0
         If Not mRec.EOF Then
            mRec.MoveFirst
            sCargarText mRec!CodUnidadMedida, mRec!Codigo
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
               sCargarText mRec!CodUnidadMedida, mRec!Codigo
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
               sCargarText mRec!CodUnidadMedida, mRec!Codigo
            End If
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 3
         If Not mRec.EOF Then
            mRec.MoveLast
            sCargarText mRec!CodUnidadMedida, mRec!Codigo
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
   
      Case 4 'Agregar
         If Command1(Index).Caption = "&Agregar" Then
            ManText = Text1(0).Text
            CodUMedidaText = Left(Combo1(0).Text, 4)
            For mi = 0 To Command1.Count - 2
               If mi <> Index Then
                  Command1(mi).Enabled = False
               End If
            Next
            Text1(1).Text = ""
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Set mRec1 = mObj.oTablaDina("Producto", "ORDER BY Codigo")
            If Not mRec1.EOF Then
               mRec1.MoveLast
               CodProducto = mRec1!Codigo
               Text1(0).Text = Format((CodProducto + 1), "000000")
               Text1(1).Enabled = True
               Text1(1).Text = ""
               Text1(1).SetFocus
            Else
               Text1(0).Text = "000001"
               Text1(1).Enabled = True
               Text1(1).Text = ""
               Text1(1).SetFocus
            End If
            
            Text1(2).Enabled = True
            Text1(2).Text = ""
            
            mRec1.Close
            Combo1(0).Enabled = True
            Combo1(0).ListIndex = 0
         Else
            If Combo1(0).ListIndex <> -1 And Text1(1).Text <> "" Then
                  Set mRec1 = mObj.oTabla("Producto", "WHERE Codigo = '" & Text1(0).Text & "'")
                  If mRec1.EOF Then
                     mObj.xInsProducto Text1(0).Text, Text1(1).Text, Text1(2).Text, Left(Combo1(0).Text, 4)
                     mRec.Requery
                     sCargarText Left(Combo1(0).Text, 4), Text1(0).Text
                     For mi = 0 To Command1.Count - 2
                        Command1(mi).Enabled = True
                     Next
                     Command1(Index).Caption = "&Agregar"
                     Command1(9).Caption = "&Volver"
                     For mi = 0 To Text1.UBound
                        Text1(mi).Enabled = False
                     Next
                     Combo1(0).Enabled = False
                  Else
                     MsgBox "Código Existente!!!", vbExclamation, sMessage
                  End If
                  mRec1.Close
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, sMessage
            End If
         End If
   
      Case 5 'Modificar
         If Command1(Index).Caption = "&Modificar" Then
            ManText = Text1(0).Text
            CodUMedidaText = Left(Combo1(0).Text, 4)
            
            For mi = 0 To Command1.Count - 1
               Command1(mi).Enabled = False
            Next

            Text1(1).Enabled = True
            Text1(1).SetFocus
            
            Text1(2).Enabled = True
            Combo1(0).Enabled = True
            
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Command1(Index).Enabled = True
            Command1(9).Enabled = True
         Else
            If Text1(1).Text <> "" Then
               If Combo1(0).ListIndex <> -1 And Text1(1).Text <> "" Then
                  mObj.xUpdProducto Text1(0).Text, Text1(1).Text, Text1(2).Text, Left(Combo1(0).Text, 4)
                  mRec.Requery
                  
                  sCargarText Left(Combo1(0).Text, 4), Text1(0).Text
                  For mi = 0 To Command1.Count - 1
                     Command1(mi).Enabled = True
                  Next
                  Command1(Index).Caption = "&Modificar"
                  Command1(9).Caption = "&Volver"
                  
                  Text1(1).Enabled = False
                  Text1(2).Enabled = False
                
                  Combo1(0).Enabled = False
               Else
                  MsgBox "El 'Stock Mínimo' debe tener valor numérico !!!", vbExclamation, sMessage
               End If
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, "Atención"
            End If
         End If
      
      Case 6 'Eliminar
         'Set mRec1 = mObj.oTabla("VehiculosInvolucr", "WHERE CodTipoVehic = '" & Left(Combo1(0).Text, 2) & "' AND CodMarca = '" & Text1(0).Text & "'")
         'If Not mRec1.EOF Then
            If MsgBox("¿Está seguro de dar de baja el producto?", vbYesNo, "Atención") = vbYes Then
'               mObj.xDelTabla "Marca", "WHERE CodTipoVehic = '" & Left(Combo1(0).Text, 2) & "' AND CodMarca = '" & Text1(0).Text & "'"
               mObj.xUpdBaJaProducto Text1(0).Text
               mRec.Requery
               If Not mRec.EOF Then
                  'sCargarText mRec!CodTipoVehic, mRec!CodMarca
                  sCargarText mRec!CodUnidadMedida, mRec!Codigo
               Else
                  Form_Load
               End If
            End If
         'Else
          '  MsgBox "NO puede Eliminar este Registro!!! " & vbCrLf & "Existe En Vehículos Involucrados!!!", vbCritical, sMessage
         'End If
         'mRec1.Close
      
      Case 7 'Buscar
      If Command1(Index).Caption = "&Buscar" Then
         For mi = 0 To Command1.Count - 2
            Command1(mi).Enabled = False
         Next
         Text1(0).Enabled = True
         Command1(Index).Caption = "C&onfirmar"
         Command1(9).Caption = "&Cancelar"
         Command1(Index).Enabled = True
         ManText = Text1(0).Text
         CodUMedidaText = Left(Combo1(0).Text, 2)
         Text1(0).SetFocus
      Else
         mRec.MoveFirst
         mSalir = False
         Do While Not mRec.EOF And Not mSalir
            If mRec!Codigo = Text1(0).Text Then
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
               If mRec!Codigo = ManText Then
                  mSalir = True
               Else
                  mRec.MoveNext
               End If
            Loop
            sCargarText CodUMedidaText, ManText
         Else
            sCargarText mRec!CodUnidadMedida, mRec!Codigo
         End If
         For mi = 0 To Command1.Count - 1
            Command1(mi).Enabled = True
         Next
         Command1(Index).Caption = "&Buscar"
         Command1(9).Caption = "&Volver"
         Text1(0).Enabled = False
      End If

      Case 8
      
      mObjAcc.mBorrarAuxi "\Inventario\Inventario", "Reportes"
      Set mRec1 = mObj.oProductos()
         mData.Execute ("CREATE TABLE Reportes (Codigo TEXT,Descripcion TEXT,Stock TEXT,Stock_Min TEXT, UnidadMedida TEXT, Sector TEXT)")
         Do While Not mRec1.EOF
            mData.Execute "insert into Reportes (Codigo,Descripcion,Stock, Stock_Min, UnidadMedida, Sector) values ('" & mRec1!Codigo & "','" & mRec1!descripcion & "','" & mRec1!Stock & "','" & mRec1!Stock_Min & "','" & mRec1!UnidadMedida & "','" & mRec1!Sector & "')"
            mRec1.MoveNext
         Loop
      mRec1.Close
      Set mAuxi = mData.OpenRecordset("SELECT * FROM Reportes")
      mAuxi.Close
      Reporte = "Rep02.rpt"
      CrystalReport1.ReportFileName = "" & App.Path & "\Inventario\" & Reporte
      CrystalReport1.WindowTitle = "Productos "
      CrystalReport1.DataFiles(0) = App.Path & "\Inventario\Inventario.mdb"
      CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: Productos'"
      CrystalReport1.Action = 1

      Case 9
         If Command1(Index).Caption = "&Cancelar" Then
            Command1(Index).Caption = "&Volver"
            Command1(4).Caption = "&Agregar"
            Command1(5).Caption = "&Modificar"
            Command1(7).Caption = "&Buscar"
            For mi = 0 To Command1.Count - 1
               Command1(mi).Enabled = True
            Next
            For mi = 0 To Text1.Count - 1
               Text1(mi).Enabled = False
            Next
            Combo1(0).Enabled = False

            If Not mRec.EOF Then
                sCargarText CodUMedidaText, ManText
            Else
               Form_Load
            End If
         Else
            Unload Me
         End If
   End Select
End Sub

Private Sub sCargarText(ByVal pCodUMed As String, ByVal pParam As String)
Dim mSalir As Boolean
Dim mi As Integer

mSalir = False
mRec.MoveFirst
Do While Not mRec.EOF And Not mSalir
   If pParam <> mRec.Fields(0) Then
      mRec.MoveNext
   Else
      mSalir = True
   End If
Loop

If mSalir Then
   For mi = 0 To Combo1(0).ListCount - 1
      Combo1(0).ListIndex = mi
      If Left(Combo1(0).Text, 4) = pCodUMed Then
         mi = 10000
      End If
   Next
   
   Text1(0).Text = NVL(mRec.Fields(0), "")
   Text1(1).Text = NVL(mRec.Fields(1), "")
   Text1(2).Text = NVL(mRec.Fields(2), "")
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'   If Index = 0 Then
'      KeyAscii = fNumeroKeyPress(KeyAscii)
'   Else
'      KeyAscii = fAlfaNumKeyPress(KeyAscii)
'   End If

   If (Index = 2 Or Index = 3) Then
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
   Else
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   End If
End Sub


