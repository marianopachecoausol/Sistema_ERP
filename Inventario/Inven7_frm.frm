VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Inven7_frm 
   Caption         =   "Form8"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14070
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3720
   ScaleWidth      =   14070
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13895
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
         Left            =   7060
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1380
         Width           =   3000
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2060
         MaxLength       =   90
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   1380
         Width           =   4250
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1380
         Width           =   795
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
            Caption         =   "&Modificar"
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
            Picture         =   "Inven7_frm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   2
            Left            =   1320
            Picture         =   "Inven7_frm.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   1
            Left            =   720
            Picture         =   "Inven7_frm.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "Inven7_frm.frx":091E
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
         Left            =   2060
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
         Caption         =   "Almacen"
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
         Left            =   7065
         TabIndex        =   3
         Top             =   1020
         Width           =   795
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
Attribute VB_Name = "Inven7_frm"
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
Dim CodAlmacenText As String
Dim Reporte As String

Public mFlagRAccd As Boolean
Dim mi As Integer

Private Sub Form_Load()
   Me.Height = 4000
   Me.Width = 14190
   Me.Caption = "Bodegas"
   sAlinearForm Me
   Set mData = OpenDatabase(App.Path & "\Inventario\Inventario.mdb")
   Label1.AutoSize = True
   Label1.Caption = "Tabla de Bodegas"
   Label1.Left = (Me.Width - Label1.Width) / 2
   For mi = 0 To Text1.UBound
      Text1(mi).Text = ""
   Next
   Combo1(0).Enabled = False
   Text1(0).Enabled = False
   Text1(1).Enabled = False
   Set mRec = mObj.oTabla("Almacenes", " where Fecha_Baja IS NULL order by Codigo")
   Do While Not mRec.EOF
      Combo1(0).AddItem "" & mRec!Codigo & " " & mRec!descripcion & ""
      mRec.MoveNext
   Loop
   mRec.Close
   
   Set mRec = mObj.oTablaDina("Bodegas", " where Fecha_Baja IS NULL order by Codigo")
   If Not mRec.EOF Then
      
      sCargarText mRec!CodAlmacen, mRec!Codigo
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

Private Sub Command1_Click(Index As Integer)
 Dim mAuxi
 Dim mSalir As Boolean
 Dim mBorrar As Boolean
 Dim mCodigo As String
 Dim CodMarca As Integer
 Dim CodBodega As Integer
   Select Case Index
      Case 0
         If Not mRec.EOF Then
            mRec.MoveFirst
            sCargarText mRec!CodAlmacen, mRec!Codigo
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
               sCargarText mRec!CodAlmacen, mRec!Codigo
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
               sCargarText mRec!CodAlmacen, mRec!Codigo
            End If
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 3
         If Not mRec.EOF Then
            mRec.MoveLast
            sCargarText mRec!CodAlmacen, mRec!Codigo
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
   
      Case 4 'Agregar
         If Command1(Index).Caption = "&Agregar" Then
            ManText = Text1(0).Text
            CodAlmacenText = Left(Combo1(0).Text, 4)
            
            For mi = 0 To Command1.Count - 2
               If mi <> Index Then
                  Command1(mi).Enabled = False
               End If
            Next
            Text1(1).Text = ""
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Set mRec1 = mObj.oTablaDina("Bodegas", "ORDER BY Codigo")
            If Not mRec1.EOF Then
               mRec1.MoveLast
               CodBodega = mRec1!Codigo 'mp2020
               Text1(0).Text = Format((CodBodega + 1), "0000")
               Text1(1).Enabled = True
               Text1(1).Text = ""
               Text1(1).SetFocus
            Else
               Text1(0).Text = "0001"
               Text1(1).Enabled = True
               Text1(1).Text = ""
               Text1(1).SetFocus
               
            End If
            
            mRec1.Close
            
            Combo1(0).Enabled = True
            Combo1(0).ListIndex = 0
         Else 'mp2020
            If Combo1(0).ListIndex <> -1 And Text1(1).Text <> "" Then
               
                  Set mRec1 = mObj.oTabla("Bodegas", "WHERE Codigo = '" & Text1(0).Text & "'")
               
                  If mRec1.EOF Then
                     'mp2020
                     mObj.xInsBodega Text1(0).Text, Text1(1).Text, Left(Combo1(0).Text, 4)
                     
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
            CodAlmacenText = Left(Combo1(0).Text, 4)
            
            For mi = 0 To Command1.Count - 1
               Command1(mi).Enabled = False
            Next

            Text1(1).Enabled = True
            Text1(1).SetFocus
            
            Combo1(0).Enabled = True
           
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            Command1(Index).Enabled = True
            Command1(9).Enabled = True
         Else
            If Text1(1).Text <> "" Then
                  'mp2020
                  mObj.xUpdBodega Text1(0).Text, Text1(1).Text, Left(Combo1(0).Text, 4)
                  
                  mRec.Requery
                  sCargarText Left(Combo1(0).Text, 4), Text1(0).Text
                  For mi = 0 To Command1.Count - 1
                     Command1(mi).Enabled = True
                  Next
                  Command1(Index).Caption = "&Modificar"
                  Command1(9).Caption = "&Volver"
                  
                  Text1(1).Enabled = False
                  
                  Combo1(0).Enabled = False
                  
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, "Atención"
            End If
         End If
      
      Case 6 'Eliminar
         Set mRec1 = mObj.oTabla("Ubicaciones", "WHERE CodBodega = '" & Text1(0).Text & "' AND Fecha_Baja IS NULL")
         If mRec1.EOF Then
            If MsgBox("¿Está seguro de dar de baja la Bodega?", vbYesNo, "Atención") = vbYes Then
               'mp2020
               mObj.xUpdBaJaBodega Text1(0).Text
               mRec.Requery
               If Not mRec.EOF Then
                  sCargarText mRec!CodAlmacen, mRec!Codigo
               Else
                  Form_Load
               End If
            End If
         Else
            MsgBox "NO puede eliminar esta Bodega!!! " & vbCrLf & "La Bodega aún tiene Ubicaciones asociadas!!!", vbCritical, sMessage
         End If
         mRec1.Close
      
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
         'mp2020
         CodAlmacenText = Left(Combo1(0).Text, 4)
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
            sCargarText CodAlmacenText, ManText
         Else
            sCargarText mRec!CodAlmacen, mRec!Codigo
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
      Set mRec1 = mObj.oBodegas()
         mData.Execute ("CREATE TABLE Reportes (Codigo TEXT,Descripcion TEXT, Almacen TEXT, Fecha_Baja TEXT)")
         Do While Not mRec1.EOF
            mData.Execute "insert into Reportes (Codigo,Descripcion, Almacen, Fecha_Baja) values ('" & mRec1!Codigo & "','" & mRec1!descripcion & "','" & mRec1!Almacen & "','" & mRec1!Fecha_Baja & "' )"
            mRec1.MoveNext
         Loop
      mRec1.Close
      Set mAuxi = mData.OpenRecordset("SELECT * FROM Reportes")
      mAuxi.Close
      Reporte = "Rep02.rpt"
      CrystalReport1.ReportFileName = "" & App.Path & "\Inventario\" & Reporte
      CrystalReport1.WindowTitle = "Bodegas "
      CrystalReport1.DataFiles(0) = App.Path & "\Inventario\Inventario.mdb"
      CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: Bodegas'"
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
               sCargarText CodAlmacenText, ManText
            Else
               Form_Load
            End If
         Else
            Unload Me
         End If
   End Select
End Sub

Private Sub sCargarText(ByVal pCodAlmacen As String, ByVal pParam As String)
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
      If Left(Combo1(0).Text, 4) = pCodAlmacen Then
         mi = 10000
      End If
   Next
   
   Text1(0).Text = NVL(mRec.Fields(0), "")
   Text1(1).Text = NVL(mRec.Fields(1), "")
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If (Index = 2 Or Index = 3) Then
      If KeyAscii <> 46 Then
         KeyAscii = fNumeroKeyPress(KeyAscii)
      End If
   Else
      KeyAscii = fAlfaNumKeyPress(KeyAscii)
   End If
End Sub


