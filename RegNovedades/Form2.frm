VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form RNov2_frm 
   Caption         =   "Módulo ABM de Tablas."
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   2  'Dot
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9300
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   480
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6120
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   14
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   13
         Top             =   1920
         Width           =   735
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   8895
         Begin VB.CommandButton Command1 
            Caption         =   "&Volver"
            Height          =   375
            Index           =   9
            Left            =   7920
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Imprimir"
            Height          =   375
            Index           =   8
            Left            =   6960
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Buscar"
            Height          =   375
            Index           =   7
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Eliminar"
            Height          =   375
            Index           =   6
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   5
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Agregar"
            Height          =   375
            Index           =   4
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   3
            Left            =   2040
            Picture         =   "Form2.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   2
            Left            =   1440
            Picture         =   "Form2.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   1
            Left            =   840
            Picture         =   "Form2.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Height          =   495
            Index           =   0
            Left            =   240
            Picture         =   "Form2.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   6120
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Height          =   300
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   7305
      End
   End
End
Attribute VB_Name = "RNov2_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clRNov
Dim mObjAcc As New clAccess
Dim mData As Database
Dim mRec As New ADODB.Recordset
Dim mCodigo As String
Dim ManText As String
Dim Reporte As String
Public mTabla As String
Public mFromAccid As Boolean

Private Sub Form_Load()
sInitForm
End Sub

Private Sub Form_Unload(Cancel As Integer)
mData.Close
Set mData = Nothing
Set mObj = Nothing
Set mRec = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim mRec1 As New ADODB.Recordset
Dim mI As Integer
Dim mSalir As Boolean
Dim mAuxi
Dim Flag As Boolean
Select Case Index
   Case 0
      If Not mRec.EOF Then
         mRec.MoveFirst
         sCargarText (mRec!Codigo)
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
            sCargarText (mRec!Codigo)
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
            sCargarText (mRec!Codigo)
         End If
      Else
         MsgBox "No hay Registros en la Tabla!!!", vbExclamation, sMessage
      End If
   Case 3
      If Not mRec.EOF Then
         mRec.MoveLast
         sCargarText (mRec!Codigo)
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
         Text1(0).Enabled = True
         Text1(0).Text = ""
         Text1(1).Enabled = True
         Text1(1).Text = ""
         Text1(0).SetFocus
         If Combo1.Visible Then
            Combo1.Enabled = True
            Combo1.ListIndex = -1
         End If
         Command1(Index).Caption = "&Grabar"
         Command1(9).Caption = "&Cancelar"
         mRec.MoveLast
      Else
         Flag = False 'ADD Validación de Texto Vacios
         If Text1(0).Text <> "" And Text1(1).Text <> "" Then
            Flag = True
            If Combo1.Visible Then
               If Combo1.Text = "" Then Flag = False
            End If
         End If
         If Flag Then
            Set mRec1 = mObj.oTabla(mTabla, "where Codigo = '" & Text1(0).Text & "'")
            If mRec1.EOF Then
               If Combo1.Visible Then
                  mObj.xInsTabla mTabla, "(Codigo,Descripcion,CodTipOtro)", "('" & Text1(0).Text & "','" & Text1(1).Text & "','" & Left(Combo1.Text, 3) & "')"
               Else
                  mObj.xInsTabla mTabla, "", "('" & Text1(0).Text & "','" & Text1(1).Text & "','0000-00-00 00:00:00')"
               End If
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
               If Combo1.Visible Then
                  Combo1.Enabled = False
               End If
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
            If mTabla = "patrulleros" Then
               mObj.xUpTablaDescr mTabla, "nombre='" & Trim(Text1(1).Text) & "'", "codigo='" & Text1(0).Text & "'"
            Else
               mObj.xUpTablaDescr mTabla, "descripcion='" & Trim(Text1(1).Text) & "'", "codigo='" & Text1(0).Text & "'"
            End If
            mRec.Requery
            sCargarText (Text1(0).Text)
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
      If MsgBox("¿Está seguro de Eliminar este Registro?", vbYesNo, "Atención") = vbYes Then
         mObj.xUpTablaDescr mTabla, "Fecha_Baja = current_timestamp", " Codigo = '" & Text1(0).Text & "'"
         mRec.Requery
         If Not mRec.EOF Then
            sCargarText (mRec!Codigo)
         Else
            Form_Load
         End If
      End If
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
      mObjAcc.mBorrarAuxi "\RegNovedades\RegNovPlus", "Reportes"
      Set mRec = mObj.oTabla(mTabla, "where fecha_baja is null order by codigo asc")
      If mTabla = "otros" Then
         mData.Execute ("CREATE TABLE Reportes (Codigo TEXT,Descripcion TEXT,CodTipoOtro TEXT)")
         Do While Not mRec.EOF
            mData.Execute "insert into Reportes (codigo,descripcion,CodTipoOtro) values ('" & mRec.Fields(0) & "','" & mRec.Fields(1) & "','" & mRec.Fields(2) & "'"
            mRec.MoveNext
         Loop
      Else
         mData.Execute ("CREATE TABLE Reportes (Codigo TEXT,Descripcion TEXT)")
         Do While Not mRec.EOF
            mData.Execute "insert into Reportes (codigo,descripcion) values ('" & mRec.Fields(0) & "','" & mRec.Fields(1) & "')"
            mRec.MoveNext
         Loop
      End If
      mRec.Close
      Set mAuxi = mData.OpenRecordset("SELECT * FROM Reportes")
      mAuxi.Close
      CrystalReport1.ReportFileName = "" & App.Path & "\RegNovedades\" & Reporte & ""
      CrystalReport1.WindowTitle = "Tabla " & mTabla
      CrystalReport1.DataFiles(0) = App.Path & "\RegNovedades\RegNovPlus.mdb"
      CrystalReport1.Formulas(0) = "Listado = 'Listado de la Tabla: " & mTabla & "'"
      CrystalReport1.Action = 1
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
         If Combo1.Visible Then
            Combo1.Enabled = False
         End If
         If Not mRec.EOF Then
            sCargarText (ManText)
         Else
            Form_Load
         End If
      Else
         If mFromAccid Then
            RAcc1beta.Enabled = True
         Else
            ShowMenu 1, True, False
         End If
         Unload Me
      End If
End Select
End Sub

Private Sub sCargarText(ByVal pCodigo As String)
Dim mSalir As Boolean
Dim mI As Integer
mSalir = False
mRec.MoveFirst
Do While Not mRec.EOF And Not mSalir
   If pCodigo <> mRec.Fields(0) Then
      mRec.MoveNext
   Else
      mSalir = True
   End If
Loop
If mSalir Then
   Text1(0).Text = NVL(mRec.Fields(0), "")
   Text1(1).Text = NVL(mRec.Fields(1), "")
   If mTabla = "otros" Then
      Combo1.ListIndex = -1
      For mI = 0 To Combo1.ListCount - 1
         If Left(Combo1.List(mI), 3) = mRec.Fields(2) Then
            Combo1.ListIndex = mI
         End If
      Next
   End If
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = fAlfaNumSimbKeyPress(KeyAscii)
KeyAscii = fUcaseKeyPress(KeyAscii)
End Sub

Private Sub sInitForm()
Dim mRecB As New ADODB.Recordset
Dim mI As Integer
Set mData = OpenDatabase(App.Path & "\RegNovedades\RegNovPlus.mdb")
RNov2_frm.Height = 4695
RNov2_frm.Width = 9420
sAlinearForm Me
For mI = 0 To Text1.UBound
   Text1(mI).Text = ""
Next
Set mRec = mObj.oTablaDina(mTabla, "where fecha_baja is null")
If Not mRec.EOF Then
   Reporte = "Rep01.rpt"
   If mTabla = "otros" Then
      Set mRecB = mObj.oTablaNull("tipootros")
      Do While Not mRecB.EOF
         Combo1.AddItem mRecB!Codigo & "-" & mRecB!descripcion
         mRecB.MoveNext
      Loop
      mRecB.Close
      Combo1.Visible = True
      Text1(0).Left = 1320
      Label2.Left = 1320
      Text1(1).Left = 2400
      Label3.Left = 2400
      Label4.Visible = True
      Reporte = "Rep02.rpt"
   End If
   sCargarText (mRec!Codigo)
Else
   For mI = 5 To Command1.UBound - 1
      Command1(mI).Enabled = False
   Next
End If
Set mRecB = Nothing
End Sub
