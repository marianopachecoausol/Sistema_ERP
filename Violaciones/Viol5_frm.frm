VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Viol5_frm 
   BackColor       =   &H00CECECE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Pasadas por Patente"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5970
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlex1 
      Height          =   3015
      Left            =   45
      TabIndex        =   5
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   327680
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   4
      Left            =   4560
      TabIndex        =   4
      Top             =   555
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Cant. Pasadas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   555
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00CECECE&
      Caption         =   "Patente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CECECE&
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4050
   End
End
Attribute VB_Name = "Viol5_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim mObj As New clViolaciones
Dim mObjPea As New clPeaje
Dim mRec As New ADODB.Recordset
Public mpFechaD As String
Public mpFechaH As String
Public mpPatente As String
Dim mEstaciones(18) As String
Dim mI As Integer

Private Sub Form_Load()
Dim mCnt As Integer
Me.Height = 4590
Me.Width = 6090
sAlinearForm Me
sMsgEspere Me, "Buscando datos...", True
With MSFlex1
   .ColWidth(0) = 500
   .ColWidth(1) = 1000
   .ColWidth(2) = 700
   .ColWidth(3) = 2500
   .ColWidth(4) = 800
   .TextMatrix(0, 0) = "N°"
   .TextMatrix(0, 1) = "Fecha"
   .TextMatrix(0, 2) = "Hora"
   .TextMatrix(0, 3) = "Estación"
   .TextMatrix(0, 4) = "Vía"
   .Font = "Arial"
   For mI = 0 To 4
      .CellFontBold = True
   Next
End With
Label1(0).Caption = Label1(0).Caption & " del " & mpFechaD & " al " & mpFechaH
Label1(2).Caption = mpPatente
Set mRec = mObjPea.oEstaciones("")
Do While Not mRec.EOF
   mEstaciones(mRec!CODIGO_ESTACION) = Trim(mRec!Descripcion_Estacion)
   mRec.MoveNext
Loop
mRec.Close
mCnt = 0
Set mRec = mObj.oViolFechasPatente(mpFechaD, mpFechaH, mpPatente, Viol4_frm.mTipo)
Do While Not mRec.EOF
   mCnt = mCnt + 1
   MSFlex1.AddItem " " & vbTab & " " & mRec!Fecha & vbTab & " " & mRec!Hora & vbTab & " " & mRec!Estacion & "-" & mEstaciones(Int(Val(mRec!Estacion))) & vbTab & " " & mRec!Via
   mRec.MoveNext
Loop
mRec.Close
sSetFlex2Colors Viol5_frm.MSFlex1, &HFFFFFF, &HE0E0E0
If MSFlex1.Rows > 2 Then
   MSFlex1.RemoveItem 1
End If
sSetFlexNroFila Viol5_frm.MSFlex1, 0
Label1(4).Caption = mCnt
sMsgEspere Me, "", False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mObj = Nothing
Set mObjPea = Nothing
Set mRec = Nothing
Viol4_frm.Enabled = True
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
