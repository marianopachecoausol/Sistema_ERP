VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form MantElect16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Relevamiento"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16965
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   16965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Relevamiento"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   16680
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1800
         MaxLength       =   90
         TabIndex        =   9
         Top             =   960
         Width           =   8535
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Prioridad:"
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
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Ramal:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Problema:"
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
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CommandButton CommandProd 
      Height          =   375
      Left            =   8377
      Picture         =   "MantElect16.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1775
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Confirmar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   3
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Frame Frame10 
      Caption         =   "Partes generados"
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
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   16680
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   278
         Left            =   13680
         TabIndex        =   1
         Top             =   7555
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid FlexPartes 
         Height          =   5640
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   16440
         _ExtentX        =   28998
         _ExtentY        =   9948
         _Version        =   327680
         Cols            =   3
      End
   End
End
Attribute VB_Name = "MantElect16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mObj As New clMantElect
Dim mRec As New ADODB.Recordset

Dim filaAnt As Integer
Dim columnAnt As Integer

Dim mvMat_CodProd_Orig() As String
Dim mvMat_CodUbic_Orig() As String
Dim mvMat_Cantidad_Orig() As Double
Dim mvMat_CantidadBD_Orig() As Double


Dim cboRamalListIndex As Integer

Private Sub Combo1_Click()
   Dim mi As Integer
   If cboRamalListIndex <> Combo1.ListIndex Then
      
      If cboRamalListIndex <> -1 And FlexPartes.Rows > 2 Then
         If MsgBox("Si selecciona otro Ramal se perderán los datos cargados hasta el momento. ¿ Desea continuar ? ", vbYesNo, "Cambio de Ramal") = vbYes Then
            limpiarFormularioParcial
            'Elimino los registros grilla inferior
            For mi = FlexPartes.Rows To 3 Step -1
               FlexPartes.RemoveItem mi
            Next
         Else
            Combo1.ListIndex = cboRamalListIndex
         End If
         cboRamalListIndex = Combo1.ListIndex
      
      
      Else
         limpiarFormularioParcial
         cboRamalListIndex = Combo1.ListIndex
      End If
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   
   Dim vParte_DescProblema() As String
   Dim vParte_Prioridad() As String
   Dim vParte_NroParte() As Integer
   Dim mCodRamal As String
   
   If Index = 0 Then
      If fValidaRelevamiento() Then

         preparaArrayPartes vParte_NroParte(), vParte_DescProblema(), vParte_Prioridad()

         mCodRamal = Right(Combo1.Text, 1)

         mObj.xGenerarRelevamiento mCodRamal, vParte_NroParte(), vParte_DescProblema(), vParte_Prioridad(), Trim(Right(MDI.mUser, 15))

         limpiarFormularioTotal


      End If
   Else
      Unload Me
   End If
End Sub

Private Sub preparaArrayPartes(ByRef pvParte_NroParte() As Integer, ByRef pvParte_DescProblema() As String, ByRef pvParte_Prioridad() As String)
   Dim mj As Integer
   Dim cantPartes As Integer

   cantPartes = FlexPartes.Rows - 2
   If cantPartes > 0 Then
      
      ReDim pvParte_DescProblema(0 To cantPartes - 1) As String
      ReDim pvParte_Prioridad(0 To cantPartes - 1) As String
      ReDim pvParte_NroParte(0 To cantPartes - 1) As Integer
      
      
      For mj = 2 To FlexPartes.Rows - 1
        pvParte_DescProblema(mj - 2) = FlexPartes.TextMatrix(mj, 1)
        pvParte_Prioridad(mj - 2) = FlexPartes.TextMatrix(mj, 2)
        pvParte_NroParte(mj - 2) = 0
      Next
   End If
End Sub

Private Function fValidaRelevamiento() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String

   mRet = True
   
   If Trim(Combo1.Text) = "" Then
      mRet = False
      mMensajeError = "Debe seleccionar un Ramal"
   End If

   If mRet Then
      If FlexPartes.Rows <= 2 Then
         mRet = False
         mMensajeError = "Al menos se debe crear un Parte"
      End If
   End If
   
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If

   fValidaRelevamiento = mRet
End Function


Private Sub CommandProd_Click()
   If fValidaParte Then
      generarUnParte
      limpiarFormularioParcial
   End If
End Sub
Private Function fValidaParte() As Boolean
   Dim mRet As Boolean
   Dim mMensajeError As String
    
   mRet = True
      
   If Trim(Combo1.Text) = "" Then
      mRet = False
      mMensajeError = "Debe seleccionar un Ramal"
   End If
      
      
   If mRet Then
         If Trim(Text1.Text) = "" Then
            mRet = False
            mMensajeError = "Debe completar el campo 'Problema'."
         End If
   End If
     
   If mRet Then
      If Trim(Combo2.Text) = "" Then
         mRet = False
         mMensajeError = "Debe seleccionar la Prioridad."
      End If
   End If
      
   If Not mRet Then
         MsgBox mMensajeError, vbCritical, "Atención"
   End If
   fValidaParte = mRet
End Function

Private Sub generarUnParte()
   
   Dim mDescProblema As String
   mDescProblema = Trim(Text1.Text)
   
   FlexPartes.AddItem "X" & vbTab & mDescProblema & vbTab & Trim(Combo2.Text)
 
End Sub

Private Sub limpiarFormularioParcial()
   Text1.Text = ""
End Sub
Private Sub limpiarFormularioTotal()
   Dim mi As Integer
   Text1.Text = ""
   sLlenoRamal
   sLlenoPrioridad
   limpiarFormularioParcial
   'Elimino los registros grilla inferior
   For mi = FlexPartes.Rows To 3 Step -1
      FlexPartes.RemoveItem mi
   Next
End Sub

Private Sub FlexPartes_Click()
   If FlexPartes.MouseCol = 0 And FlexPartes.MouseRow > 0 Then
      If FlexPartes.Rows > 2 Then
         FlexPartes.RemoveItem FlexPartes.MouseRow
         FlexPartes.Row = 0
      End If
   End If
End Sub


Private Sub Form_Load()
   Me.Width = 17085
   Me.Height = 9920
   sAlinearForm Me


   sLlenoRamal
   cboRamalListIndex = Combo1.ListIndex

   sLlenoPrioridad
   initPartes
End Sub

Private Sub sLlenoRamal()
   Dim mRec1 As New ADODB.Recordset
   
   Combo1.Clear
   Set mRec1 = mObj.oEjecutarSelect("SELECT Codigo, Descripcion From COM_Ramales ORDER BY Descripcion; ")

   Do While Not mRec1.EOF
      Combo1.AddItem mRec1!descripcion & Space(100) & mRec1!Codigo
      mRec1.MoveNext
   Loop
   
   mRec1.Close
   Set mRec1 = Nothing
End Sub

Private Sub sLlenoPrioridad()
   Combo2.Clear
   Combo2.AddItem "Alta"
   Combo2.AddItem "Media"
   Combo2.AddItem "Baja"
   Combo2.ListIndex = 1
End Sub

Private Sub initPartes()
   filaAnt = 0
   columnAnt = 0
   Text2.Visible = False
   
   With FlexPartes
      .ColWidth(0) = 200
      .ColWidth(1) = 14190 '9470
      .ColWidth(2) = 1650

      .TextMatrix(0, 1) = "Problema"
      .TextMatrix(0, 2) = "Prioridad"
      
      .ColAlignment(1) = flexAlignLeftCenter

      .RowHeight(1) = 0
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set mObj = Nothing
   Set mRec = Nothing
   ShowMenu 47, True, False
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 46 Then
      KeyAscii = fNumeroKeyPress(KeyAscii)
   End If
   
   If KeyAscii = 13 Then
      FlexPartes.TextMatrix(filaAnt, columnAnt) = Text2.Text
      Text2.Visible = False
      FlexPartes.ScrollBars = flexScrollBarVertical
   End If
End Sub

Private Sub Text2_LostFocus()
   If FlexPartes.Col <> columnAnt Or FlexPartes.Row <> filaAnt Then
      'En este caso 3 es la columna que seria editable
      If columnAnt = 3 Then
         FlexPartes.TextMatrix(filaAnt, columnAnt) = Text2.Text
      End If
   End If
End Sub
