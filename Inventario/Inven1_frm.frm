VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form Inven1_frm 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   360
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   8655
         Begin VB.CommandButton Command1 
            Caption         =   "&VolvER"
            Height          =   375
            Index           =   9
            Left            =   7680
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   375
            Index           =   8
            Left            =   6720
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Buscar"
            Height          =   375
            Index           =   7
            Left            =   5760
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   6
            Left            =   4800
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   5
            Left            =   3840
            TabIndex        =   12
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Agregar"
            Height          =   375
            Index           =   4
            Left            =   2880
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   3
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ir a Final"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   2
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Posterior"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   1
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Anterior"
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ir a Principio"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   3480
         TabIndex        =   4
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label2 
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
         Left            =   1440
         TabIndex        =   2
         Top             =   1680
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Height          =   615
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
   End
End
Attribute VB_Name = "Inven1_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mData As Database
Dim mObj As New clInven
Dim mObjAcc As New clAccess
Dim mRec As New ADODB.Recordset
Dim mI As Integer
Dim ManText As String
Dim Codigo As String
Dim CodName As String
Dim Descrip As String
Dim Reporte As String
Dim mTable1 As String
Public mTabla As String
Public mFlagRAccd As Boolean

Private Sub Form_Load()
   Me.Height = 4005
   Me.Width = 9200
  ' Set mData = OpenDatabase(App.Path & "\RegAccidentes\FichaAccid.mdb")
   Label1.AutoSize = True
   sAlinearForm Me
   For mI = 0 To Text1.UBound
      Text1(mI).Text = ""
   Next
   Text1(0).Enabled = False
   Text1(1).Enabled = False
   Set mRec = mObj.oTablaDina(mTabla, "")
   If Not mRec.EOF Then
      Select Case mTabla
         Case "CiaSeguros"
             Codigo = mRec!CodCiaSeguro
             mTable1 = "VehiculosInvolucr"
'             If RAcc2_frm.mAgregar Then
'                RAcc2_frm.Enabled = False
'             End If
         Case "LugarTrasl"
             Codigo = mRec!codlugartrasl
             mTable1 = "VictimasInvolucr"
         Case "Patrullero"
             Codigo = mRec!CodPatrullero
             mTable1 = "Ficha"
         Case "TipoVehiculo"
             Codigo = mRec!CodTipoVehic
             mTable1 = "VehiculosInvolucr"
      End Select
      CodName = mRec.Fields(0).Name
      Descrip = mRec.Fields(1).Name
      Reporte = App.Path & "\RegAccidentes\" & "Rep01.rpt"
      sCargarText (Codigo)
   Else
      For mI = 5 To Command1.UBound - 1
         Command1(mI).Enabled = False
      Next
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   If RAcc2_frm.mAgregar Then
'      RAcc2_frm.mAgregar = False
'      RAcc2_frm.Enabled = True
'   End If
   If mFlagRAccd Then
      RAcc1beta.Enabled = True
   Else
      ShowMenu 2, True, False
   End If
   mRec.Close
   mData.Close
   Set mData = Nothing
   Set mObj = Nothing
   Set mObjAcc = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mRec1 As New ADODB.Recordset
Dim mSalir As Boolean
Dim mBorrar As Boolean
Dim mCodigo As Integer
   
   Select Case Index
      Case 0
         If Not mRec.EOF Then
            mRec.MoveFirst
            mCodigo = mRec.Fields(0)
            sCargarText (Codigo)
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
               mCodigo = mRec.Fields(0)
               sCargarText (Format(mCodigo, "00"))
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
               mCodigo = mRec.Fields(0)
               sCargarText (Format(mCodigo, "00"))
            End If
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 3
         If Not mRec.EOF Then
            mRec.MoveLast
            mCodigo = mRec.Fields(0)
            sCargarText mCodigo
         Else
            MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
         End If
      
      Case 4
         If Command1(Index).Caption = "&Agregar" Then
            ManText = Text1(0).Text
            For mI = 0 To Command1.Count - 2
               If mI <> Index Then
                  Command1(mI).Enabled = False
               End If
            Next
            Text1(1).Enabled = True
            Text1(1).Text = ""
            Text1(1).SetFocus
            Command1(Index).Caption = "&Grabar"
            Command1(9).Caption = "&Cancelar"
            mRec.MoveLast
            mCodigo = mRec.Fields(0)
            Text1(0) = mCodigo + 1
         Else
            If Text1(0).Text <> "" And Text1(1).Text <> "" Then
               Set mRec1 = mObj.oTabla(mTabla, " where " & CodName & " = '" & Trim(Text1(0).Text) & "'")
               If mRec1.EOF Then
                  mObj.xInsTabla mTabla, Trim(Text1(0).Text), Trim(Text1(1).Text)
                  mRec.Requery
                  sCargarText (Text1(0).Text)
                  For mI = 0 To Command1.Count - 2
                     Command1(mI).Enabled = True
                  Next
                  Command1(Index).Caption = "&Agregar"
                  Command1(9).Caption = "&Volver"
                  For mI = 0 To Text1.UBound
                     Text1(mI).Enabled = False
                  Next
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
               'mObj.xUpdTabla mTabla, CodName & " = '" & Trim(Text1(0).Text) & "'", Descrip & " = '" & Trim(Text1(1).Text) & "'"
               mRec.Requery
               sCargarText (Text1(0).Text)
               For mI = 0 To Command1.Count - 1
                 Command1(mI).Enabled = True
               Next
               Command1(Index).Caption = "&Modificar"
               Command1(9).Caption = "&Volver"
               Text1(1).Enabled = False
            Else
               MsgBox "Faltan Ingresar Datos!!!", vbExclamation, sMessage
            End If
         End If
         
      Case 6
         Set mRec1 = mObj.oTabla(mTable1, "where " & CodName & " = '" & Trim(Text1(0).Text) & "'")
         If mRec1.EOF Then
            If MsgBox("¿Está seguro de Eliminar este Registro?", vbYesNo, sMessage) = vbYes Then
               'mObj.xDelTabla mTabla, CodName & " = '" & Trim(Text1(0).Text) & "'"
               mRec.Requery
               If Not mRec.EOF Then
                  Codigo = mRec.Fields(0)
                  sCargarText (Codigo)
               Else
                  Form_Load
               End If
            End If
         Else
            MsgBox "NO puede Eliminar este Registro!!! " & vbCrLf & "Existe en " & mTable1 & "!!!", vbCritical, sMessage
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
            ManText = Text1(0).Text
            Text1(0).SetFocus
         Else
            mRec.MoveFirst
            mSalir = False
            Do While Not mRec.EOF And Not mSalir
               Codigo = mRec.Fields(0)
               If Codigo = Text1(0).Text Then
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
                  Codigo = mRec.Fields(0)
                  If Codigo = ManText Then
                     mSalir = True
                  Else
                     mRec.MoveNext
                  End If
               Loop
               sCargarText (ManText)
            Else
               sCargarText (Text1(0).Text)
            End If
            For mI = 0 To Command1.Count - 1
               Command1(mI).Enabled = True
            Next
            Command1(Index).Caption = "&Buscar"
            Command1(9).Caption = "&Volver"
            Text1(0).Enabled = False
         End If
         
      Case 8
         sMsgEspere Me, "Generando reporte...", True
         mObjAcc.mBorrarAuxi "\RegAccidentes\FichaAccid", "Auxi"
         Set mRec1 = mObj.oTabla(mTabla, "")
         mData.Execute ("CREATE TABLE Auxi (Codigo TEXT, Descripcion TEXT)")
         Do While Not mRec1.EOF
            mData.Execute ("INSERT INTO Auxi (Codigo,Descripcion) VALUES ('" & mRec1.Fields(0) & "','" & mRec1.Fields(1) & "')")
            mRec1.MoveNext
         Loop
         mRec1.Close
         CrystalReport1.ReportFileName = "" & Reporte & ""
         CrystalReport1.WindowTitle = "Tabla " & mTabla
         CrystalReport1.DataFiles(0) = App.Path & "\RegAccidentes\" & "FichaAccid.mdb"
         CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: " & mTabla & "'"
         CrystalReport1.Action = 1
         CrystalReport1.Formulas(0) = ""
         sMsgEspere Me, "", False
         
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
            If Not mRec.EOF Then
               sCargarText (ManText)
            Else
               Form_Load
            End If
         Else
            Unload Me
         End If
   End Select
   Set mRec1 = Nothing
End Sub

Private Sub sCargarText(ByVal pParam As String)
   Text1(0).Text = pParam
   Text1(1).Text = mObj.sTablaDescr(mTabla, CodName & "='" & pParam & "'", 1)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
End Sub
