VERSION 5.00
Begin VB.Form RAcc13 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F7F7F7&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1290
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2775
      MaxLength       =   5
      TabIndex        =   1
      Top             =   750
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Borrar Fichas de Accidentes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   225
      TabIndex        =   4
      Top             =   150
      Width           =   3990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Ficha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1650
      TabIndex        =   0
      Top             =   825
      Width           =   1005
   End
End
Attribute VB_Name = "RAcc13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
   sAlinearForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ShowMenu 2, True, False
End Sub

Private Sub Command1_Click(Index As Integer)
   If Index = 0 Then
      If Trim(Text1.Text) <> "" Then
         sRollBack Text1.Text
      Else
         MsgBox "Ingresar un número de ficha", vbInformation, sMessage
      End If
   Else
      Unload Me
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = fNumeroKeyPress(KeyAscii)
End Sub

Private Sub sRollBack(ByVal pNroOrden As String)
Dim mObj As New clRAcc
Dim mFecha As String
Dim mIP As String

   If mObj.bExistFicha(pNroOrden) = True Then
      mFecha = mObj.sTablaDescr("Ficha", "nroorden='" & pNroOrden & "'", 4)
      mFecha = mFecha & " " & mObj.sTablaDescr("Ficha", "nroorden='" & pNroOrden & "'", 5) & ":00"
      If MsgBox("Seguro de Borrar la ficha " & Text1.Text & "?", vbYesNo, sMessage) = vbYes Then
         mObj.xDeleteTable "Ficha", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "fichadescr", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "interterceros", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "intergco", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "VehiculosInvolucr", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "VictimasInvolucr", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "daniosgco", " nroorden='" & pNroOrden & "' "
         mObj.xDeleteTable "fichaobs", " nroorden='" & pNroOrden & "' "
         'insertar el borrado de ficha, fecha de borrado y codusuario, y fecha de ficha
         mObj.xInLogDel pNroOrden, mFecha, Trim(Right(MDI.mUser, 20)), MDI.mPCname
      End If
   Else
      MsgBox "No existe el número de ficha", vbExclamation, sMessage
   End If
End Sub
